"""Persistent, offline MSG body fingerprint worker for the archive engine.

Input/output is one JSON object per line. Never renders content, opens Outlook,
executes message scripts or downloads resources. A parse failure is unsupported,
not permission to ignore the body.
"""
from __future__ import annotations
import hashlib
import importlib.metadata
import json
from pathlib import Path
import sys

VERSION = "relevant-message-v1"
EXPECTED = {"RTFDE": "0.1.2.2", "olefile": "0.47", "compressed-rtf": "1.0.7", "html5lib": "1.1"}
sys.path.insert(0, str(Path(__file__).resolve().parent))
if len(sys.argv) != 2:
    raise SystemExit(2)
sys.path.insert(0, str(Path(sys.argv[1]).resolve(strict=True)))
for package, expected in EXPECTED.items():
    if importlib.metadata.version(package) != expected:
        raise RuntimeError("dependency-version-mismatch:" + package)

from compare_relevant_message import html_profile, read_html, digest, self_test
import olefile

BODY_STREAMS = ("1000001F", "1000001E", "10130102", "1013001F", "10090102")
MAX_BODY = 32 * 1024 * 1024


def profile(path: Path, report_equivalence=False, repair_equivalence=False) -> dict[str, str]:
    if not path.is_absolute() or path.suffix.lower() != ".msg" or not path.is_file():
        raise ValueError("msg-path-invalid")
    if path.stat().st_size > 512 * 1024 * 1024:
        raise ValueError("msg-size-limit")
    # The exact fallback is useful for genuine RTF/plain-text messages. It does
    # not approximate arbitrary RTF as text or approve a failed HTML extraction.
    exact = {}
    plain_streams = []
    unicode_stream = None
    ansi_stream = None
    with olefile.OleFileIO(str(path), write_mode=False) as message:
        if repair_equivalence:
            from signed_mime_profile import signed_mime_profile
            signed = signed_mime_profile(message)
            if signed is not None:return signed
        for tag in BODY_STREAMS:
            name = "__substg1.0_" + tag
            if not message.exists(name):
                exact[tag] = "missing"
                continue
            if message.get_size(name) > MAX_BODY:
                raise ValueError("body-size-limit")
            data = message.openstream(name).read()
            exact[tag] = hashlib.sha256(data).hexdigest()
            if tag in ("1000001F", "1000001E"): plain_streams.append(data)
            if tag == "1000001F": unicode_stream = data
            if tag == "1000001E": ansi_stream = data
        has_html = any(exact[tag] != "missing" for tag in ("10130102", "1013001F"))
        has_rtf = exact["10090102"] != "missing"
        if not has_html and not has_rtf:
            if exact["1000001F"] == "missing" and exact["1000001E"] == "missing":
                raise ValueError("recognized-body-missing")
            unicode_empty = unicode_stream is None or not unicode_stream.decode("utf-16-le", errors="strict").strip("\r\n")
            ansi_empty = ansi_stream is None or not ansi_stream.strip(b"\r\n")
            if unicode_empty and ansi_empty:
                return {"ProfileVersion": VERSION, "BodyMode": "empty-body-v2", **{"Body-" + tag: hashlib.sha256(("validated-empty-body-v2:" + tag).encode()).hexdigest() for tag in BODY_STREAMS}}
            if repair_equivalence and unicode_stream is not None and ansi_stream is None:
                from rtf_representation import canonical_fromtext_unicode
                exact["1000001F"] = hashlib.sha256(b"repair-plain-unicode-v1:" + canonical_fromtext_unicode(unicode_stream)).hexdigest()
                return {"ProfileVersion": VERSION, "BodyMode": "repair-plain-unicode-v1", **{"Body-" + tag: value for tag, value in exact.items()}}
            return {"ProfileVersion": VERSION, "BodyMode": "native-exact-v1", **{"Body-" + tag: value for tag, value in exact.items()}}
    try:
        html, _ = read_html(path, restore_rtf_literal_emails=report_equivalence)
    except Exception:
        # Direct HTML must be parsed successfully. Encapsulated HTML with an
        # invalid CRC/decompression is also never silently reduced to plaintext.
        if has_html:
            raise
        from verify_rtf_sample import inspect_msg
        _, raw, validation = inspect_msg(path)
        if validation.get("fromhtml1_in_header"):
            raise
        if repair_equivalence:
            from rtf_representation import validated_repair_text_representation
            normalized = validated_repair_text_representation(raw, unicode_stream, ansi_stream)
            if normalized is not None:
                # Do not normalize this RTF hash: it preserves the exact source
                # text when Outlook regenerates its Unicode presentation.
                exact["1000001F"] = hashlib.sha256(b"repair-text-representation-v1:" + normalized).hexdigest()
                return {"ProfileVersion": VERSION, "BodyMode": "repair-text-representation-v1", **{"Body-" + tag: value for tag, value in exact.items()}}
        from rtf_representation import validated_fromtext_rtf, validated_empty_fromtext_rtf, empty_fromtext_rtf
        if validated_empty_fromtext_rtf(raw, unicode_stream, ansi_stream):
            return {"ProfileVersion": VERSION, "BodyMode": "empty-body-v2", **{"Body-" + tag: hashlib.sha256(("validated-empty-body-v2:" + tag).encode()).hexdigest() for tag in BODY_STREAMS}}
        if validated_fromtext_rtf(raw, unicode_stream):
            from rtf_representation import canonical_fromtext_unicode
            normalized = canonical_fromtext_unicode(unicode_stream)
            exact["1000001F"] = hashlib.sha256(b"validated-unicode-text-v3:" + normalized).hexdigest()
            exact["10090102"] = hashlib.sha256(b"validated-text-derived-rtf-v3:" + normalized).hexdigest()
            return {"ProfileVersion": VERSION, "BodyMode": "text-derived-rtf-v3", **{"Body-" + tag: value for tag, value in exact.items()}}
        if empty_fromtext_rtf(raw, plain_streams):
            exact["10090102"] = hashlib.sha256(b"validated-empty-fromtext-rtf-v1").hexdigest()
            return {"ProfileVersion": VERSION, "BodyMode": "empty-text-rtf-v1", **{"Body-" + tag: value for tag, value in exact.items()}}
        return {"ProfileVersion": VERSION, "BodyMode": "native-exact-v1", **{"Body-" + tag: value for tag, value in exact.items()}}
    parsed = html_profile(html, normalize_layout_nbsp=report_equivalence, qualified_css=repair_equivalence)
    if parsed["Unsupported"]:
        # Special markup/CSS is retained in full, never discarded or executed.
        # Bytes and strings use distinct encodings: cross-representation matches
        # remain blocked unless the full representation is exactly identical.
        payload = b"bytes:" + html if isinstance(html, bytes) else b"unicode:" + html.encode("utf-8")
        return {"ProfileVersion": VERSION, "BodyMode": "html-exact-v1", "Body-HTML": hashlib.sha256(payload).hexdigest()}
    components = ("Text", "Links", "Images", "Resources", "Tables", "Lists", "Preformatted", "ActivePayloads")
    mode = "html-controls-v1" if parsed["Controls"] is not None else "html-relevant-v1"
    if mode == "html-controls-v1": components += ("Controls",)
    if parsed["InertCSS"] is not None:
        mode = "html-static-v1"
        if "Controls" not in components: components += ("Controls",)
        components += ("InertCSS",)
    return {"ProfileVersion": VERSION, "BodyMode": mode, **{"Body-" + key: digest(parsed[key]) for key in components}}


for line in sys.stdin:
    try:
        if len(line) > 1024 * 1024:
            raise ValueError("request-size-limit")
        request = json.loads(line)
        if request.get("Operation") == "ping":
            response = {"Supported": True, "Version": VERSION}
        elif request.get("Operation") == "self-test":
            from rtf_representation import unicode_fallback_self_test, fromtext_self_test, empty_fromtext_self_test, canonical_unicode_self_test
            response = {"Supported": True, "Version": VERSION, "Tests": self_test()["Passed"] + unicode_fallback_self_test() + fromtext_self_test() + empty_fromtext_self_test() + canonical_unicode_self_test()}
        elif request.get("Operation") == "profile":
            response = {"Supported": True, "Version": VERSION, "Fields": profile(Path(request["Path"]), request.get("ReportEquivalence") is True, request.get("RepairEquivalence") is True)}
        elif request.get("Operation") == "recover-encapsulated-html":
            html, evidence = read_html(Path(request["Path"]), restore_rtf_literal_emails=True)
            if not isinstance(html,str) or evidence.get("Representation") != "rtfde-recovered-html":
                raise ValueError("encapsulated-html-required")
            response = {"Supported": True, "Version": VERSION, "Html": html}
        else:
            raise ValueError("unknown-operation")
    except Exception as error:
        # Message-derived data and exception strings must never reach logs.
        response = {"Supported": False, "Version": VERSION, "ErrorType": type(error).__name__}
    print(json.dumps(response, ensure_ascii=True, separators=(",", ":")), flush=True)
