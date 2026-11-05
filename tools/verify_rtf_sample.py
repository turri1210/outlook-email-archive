"""Offline MSG evidence checks; no Outlook, HTML rendering or network access.

Reads only item-NNN pairs corresponding to report.Items with a DestinationId.
Outputs sizes, hashes and known RTF marker counts, never message content.
The result does not enable source deletion or change the application's policy.
"""
from __future__ import annotations

import argparse
import hashlib
import importlib.metadata
import json
import os
from pathlib import Path
import re
import struct
import sys

MAX_BYTES = 32 * 1024 * 1024
RTF_STREAM = "__substg1.0_10090102"
HTML_STREAM = "__substg1.0_10130102"


class EvidenceError(Exception):
    """Contains a fixed diagnostic code, never message-derived content."""


def check(condition: bool, code: str) -> None:
    if not condition:
        raise EvidenceError(code)


def validate_tokens(payload: bytes, declared_size: int) -> None:
    """Bound expansion BEFORE calling the library, which ignores raw_size.

    LZFu has eight tokens per control byte, low bit first. References use
    a 12-bit dictionary offset and a 4-bit length plus two. A reference to
    the current write offset is the end marker (MS-OXRTFCP).
    """
    position, produced, write_offset = 0, 0, 207
    initialized = [True] * 207 + [False] * (4096 - 207)
    while position < len(payload):
        control = payload[position]
        position += 1
        for bit in range(8):
            if control & (1 << bit):
                check(position + 2 <= len(payload), "truncated-reference")
                token = int.from_bytes(payload[position:position + 2], "big")
                position += 2
                offset, length = token >> 4, (token & 15) + 2
                if offset == write_offset:
                    check(produced == declared_size, "end-marker-size-mismatch")
                    return  # Remaining bytes are padding, covered by the CRC.
                check(produced + length <= declared_size, "expansion-exceeds-declared-size")
                for step in range(length):
                    check(initialized[(offset + step) % 4096], "uninitialized-dictionary-reference")
                    initialized[write_offset] = True
                    write_offset = (write_offset + 1) % 4096
                produced += length
            else:
                check(position < len(payload), "truncated-literal")
                check(produced < declared_size, "expansion-exceeds-declared-size")
                position += 1
                produced += 1
                initialized[write_offset] = True
                write_offset = (write_offset + 1) % 4096
    raise EvidenceError("missing-end-marker")


def inspect_msg(path: Path) -> tuple[bytes, bytes, dict]:
    import olefile
    from compressed_rtf import decompress
    from compressed_rtf.crc32 import crc32

    with olefile.OleFileIO(str(path), write_mode=False) as msg:
        check(msg.exists(RTF_STREAM), "rtf-stream-missing")
        check(16 <= msg.get_size(RTF_STREAM) <= MAX_BYTES, "compressed-stream-size-limit")
        with msg.openstream(RTF_STREAM) as stream:
            compressed = stream.read()
        html_size = msg.get_size(HTML_STREAM) if msg.exists(HTML_STREAM) else None

    check(len(compressed) >= 16, "truncated-header")
    comp_size, raw_size, magic, expected_crc = struct.unpack("<II4sI", compressed[:16])
    check(comp_size >= 12 and comp_size + 4 == len(compressed), "compressed-size-mismatch")
    check(raw_size <= MAX_BYTES, "decompressed-size-limit")
    payload = compressed[16:]
    if magic == b"LZFu":
        check(crc32(payload) == expected_crc, "crc-mismatch")
        validate_tokens(payload, raw_size)
    elif magic == b"MELA":
        check(expected_crc == 0, "uncompressed-crc-not-zero")
        check(len(payload) == raw_size, "uncompressed-size-mismatch")
    else:
        raise EvidenceError("unsupported-compression-magic")
    raw = decompress(compressed)
    check(len(raw) == raw_size, "decompressed-size-mismatch")
    check(raw.startswith(b"{\\rtf"), "rtf-header-missing")
    next_group = raw.find(b"{", 1)
    header = raw[:next_group] if next_group >= 0 else raw
    from_html = len(re.findall(rb"(?<!\\)\\fromhtml1\b", raw))
    html_tags = len(re.findall(rb"(?<!\\)\\htmltag\d*\b", raw))
    result = {
        "compressed_bytes": len(compressed),
        "compressed_sha256": hashlib.sha256(compressed).hexdigest(),
        "declared_uncompressed_bytes": raw_size,
        "actual_uncompressed_bytes": len(raw),
        "decompressed_sha256": hashlib.sha256(raw).hexdigest(),
        "magic": magic.decode("ascii"),
        "crc_valid": True,
        "fromhtml1_count": from_html,
        "fromhtml1_in_header": bool(re.search(rb"(?<!\\)\\fromhtml1\b", header)),
        "fromtext_count": len(re.findall(rb"(?<!\\)\\fromtext\b", raw)),
        "htmltag_count": html_tags,
        "htmlrtf_count": len(re.findall(rb"(?<!\\)\\htmlrtf\d*\b", raw)),
        "separate_html_stream_bytes": html_size,
    }
    return compressed, raw, result


def analyze(run: Path) -> dict:
    report_path = run / "report.json"
    check(report_path.stat().st_size <= MAX_BYTES, "report-size-limit")
    report = json.loads(report_path.read_text(encoding="utf-8-sig"))
    items = report.get("Items")
    check(isinstance(items, list), "report-items-invalid")
    check(0 < len(items) <= 10000, "report-item-count-invalid")
    rows, skipped = [], []
    for index, item in enumerate(items, 1):
        check(isinstance(item, dict), "report-item-invalid")
        destination_id = item.get("DestinationId")
        if not isinstance(destination_id, str) or not destination_id.strip():
            skipped.append(index)
            continue
        row = {"index": index, "status": "review-required"}
        try:
            directory = run / f"item-{index:03d}"
            lc, lr, ls = inspect_msg(directory / "message.msg")
            rc, rr, rs = inspect_msg(directory / "destination-check.msg")
            encapsulated = all(s["fromhtml1_in_header"] and s["htmltag_count"] > 0 for s in (ls, rs))
            same_compressed, same_raw = lc == rc, lr == rr
            row.update({
                "left": ls, "right": rs,
                "compressed_rtf_equal": same_compressed,
                "decompressed_rtf_equal": same_raw,
                "html_encapsulation_markers_both": encapsulated,
                "status": "equal-encapsulated-rtf" if same_raw and encapsulated else "review-required",
            })
        except Exception as error:
            row["error_code"] = str(error) if isinstance(error, EvidenceError) else type(error).__name__
        rows.append(row)
    all_equal = bool(rows) and all(row["status"] == "equal-encapsulated-rtf" for row in rows)
    return {
        "schema_version": 1,
        "offline": True,
        "libraries": {name: importlib.metadata.version(name) for name in ("olefile", "compressed-rtf")},
        "limit_bytes_per_compressed_or_decompressed_stream": MAX_BYTES,
        "report_items": len(items),
        "recorded_destination_items": len(rows),
        "skipped_indices_without_destination_id": skipped,
        "eligible_items_valid_and_equal": all_equal,
        "all_report_items_verified": all_equal and not skipped,
        "items": rows,
        "limits": "Verifica pares MSG históricos e representação RTF encapsulando HTML; não verifica servidor, renderização, recursos remotos, anexos ou propriedades MAPI completas. Não habilita exclusão, repetição de transferência nem aprovação automática.",
    }


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--run-dir", required=True, type=Path)
    parser.add_argument("--output", required=True, type=Path)
    args = parser.parse_args()
    try:
        check(not args.output.exists(), "output-already-exists")
        result = analyze(args.run_dir.resolve(strict=True))
        # Exclusive creation protects earlier evidence, including a racing writer.
        with args.output.open("x", encoding="utf-8") as output:
            json.dump(result, output, ensure_ascii=False, indent=2)
            output.flush()
            os.fsync(output.fileno())
        print(json.dumps({"recorded_destination_items": result["recorded_destination_items"], "all_report_items_verified": result["all_report_items_verified"]}))
        return 0 if result["all_report_items_verified"] else 3
    except Exception as error:
        code = str(error) if isinstance(error, EvidenceError) else type(error).__name__
        print("Falha na verificação offline: " + code, file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
