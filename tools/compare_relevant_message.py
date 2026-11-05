"""Offline diagnostic of relevant information in existing MSG evidence.

Never renders HTML, follows links, opens Outlook or changes removal policy.
RTFDE de-encapsulates original HTML; it does not approximate arbitrary RTF.
Use --deps-dir for the isolated pinned dependencies, never a global install.
"""
from __future__ import annotations
import argparse
import hashlib
import importlib.metadata
import json
import os
from pathlib import Path
import re
import sys

MAX_BYTES = 32 * 1024 * 1024
ASCII_SPACE = re.compile(r"[ \t\r\n\f]+")
BLOCKS = {"p", "div", "li", "tr", "td", "th", "table", "blockquote", "h1", "h2", "h3", "h4", "h5", "h6", "ul", "ol", "pre", "hr", "section", "article"}
IGNORED_TEXT = {"style", "script", "head", "title", "meta"}
RESOURCE_ATTRS = {"href", "src", "srcset", "background", "poster", "data", "action", "formaction", "longdesc", "cite"}
UNSUPPORTED = {"iframe", "object", "embed", "form", "textarea", "select", "video", "audio", "canvas", "svg", "math"}

def norm(text: str) -> str:
    return ASCII_SPACE.sub(" ", text.replace("\r\n", "\n").replace("\r", "\n")).strip(" \t\r\n\f")

def digest(value) -> str:
    return hashlib.sha256(json.dumps(value, ensure_ascii=False, sort_keys=True, separators=(",", ":")).encode("utf-8")).hexdigest()

def visible_text(node, control_tokens=False) -> str:
    if not isinstance(node.tag, str) or node.tag in IGNORED_TEXT:
        return ""
    if node.tag == "input" and control_tokens:
        return "\ufff0" + digest(["input", sorted(node.attrib.items())]) + "\ufff1"
    def project(text):
        text = text or ""
        return text.replace("\ufff0", "\ufff0\ufff0").replace("\ufff1", "\ufff1\ufff1") if control_tokens else text
    pieces = [project(node.text)]
    for child in node:
        if child.tag == "br": pieces.append("\n")
        elif child.tag in BLOCKS: pieces.extend(["\n", visible_text(child, control_tokens), "\n"])
        else: pieces.append(visible_text(child, control_tokens))
        pieces.append(project(child.tail))
    return "".join(pieces)

def html_profile(html: bytes | str, *, normalize_layout_nbsp=False, qualified_css=False) -> dict:
    def normalize(text):
        if qualified_css:
            import unicodedata
            text=unicodedata.normalize('NFC',text)
        return norm(text.replace("\u00a0", " ") if normalize_layout_nbsp else text)
    import html5lib
    dom = html5lib.parse(html, treebuilder="etree", namespaceHTMLElements=False)
    body = dom.find("body")
    if body is None: raise ValueError("html-body-unavailable")
    normalized_markers = 0
    # Word's de-encapsulation can retain a numeric fallback which duplicates
    # the HTML ordered-list counter. Remove only the exact standalone N.<TAB>
    # text in a Word LI whose real OL counter is independently retained below.
    # Ordinary numeric text, wrong counters and other prefixes remain compared.
    for ordered in body.iter("ol"):
        number = int(ordered.get("start", "1"))
        step = -1 if "reversed" in ordered.attrib else 1
        for li in (child for child in ordered if child.tag == "li"):
            number = int(li.get("value", str(number)))
            if (ordered.get("type", "1") == "1" and "MsoListParagraph" in li.get("class", "").split()
                and "mso-list:" in li.get("style", "") and li.text == f"{number}.\t"):
                li.text = ""
                normalized_markers += 1
            number += step
    if qualified_css:
        from rtf_representation import remove_proven_word_markers
        normalized_markers += remove_proven_word_markers(body)
    text = normalize(visible_text(body))
    links, images, resources, tables, lists, preformatted, unsupported, active_payloads = [], [], [], [], [], [], [], []
    for element in body.iter():
        if not isinstance(element.tag, str): continue
        tag, attrs = element.tag, element.attrib
        if tag in UNSUPPORTED: unsupported.append("dynamic-or-special-element:" + tag)
        if "href" in attrs:
            links.append((tag, attrs["href"], normalize(visible_text(element)), attrs.get("title"), attrs.get("target")))
        if tag == "img" or tag.endswith(":imagedata"):
            images.append((tag, attrs.get("src"), attrs.get("srcset"), attrs.get("alt"), attrs.get("title")))
        semantic_attrs = [(key, value) for key, value in attrs.items() if key in RESOURCE_ATTRS or key.endswith("}href")]
        if semantic_attrs: resources.append((tag, sorted(semantic_attrs)))
        if tag == "table":
            rows = []
            for row in element.iter("tr"):
                cells = [(cell.tag, cell.get("colspan", "1"), cell.get("rowspan", "1"), cell.get("headers"), cell.get("scope"), normalize(visible_text(cell))) for cell in row if cell.tag in {"td", "th"}]
                rows.append(cells)
            tables.append((element.get("summary"), rows))
        if tag in {"ol", "ul"}:
            lists.append((tag, attrs.get("type", "1" if tag == "ol" else "disc"), attrs.get("start", "1"), "reversed" in attrs,
                          [(li.get("value"), normalize(visible_text(li))) for li in element if li.tag == "li"]))
        if tag == "pre": preformatted.append(visible_text(element).replace("\r\n", "\n").replace("\r", "\n"))
    # Textually generated CSS and active/dynamic documents are outside this
    # profile. Never approve their information by merely ignoring the markup.
    for element in dom.iter():
        if not isinstance(element.tag, str): continue
        if element.tag == "script": active_payloads.append((sorted(element.attrib.items()), (element.text or "").replace("\r\n", "\n").replace("\r", "\n")))
        css = element.text or "" if element.tag == "style" else element.get("style", "")
        if re.search(r"(?:^|[;{])\s*content\s*:", css, re.I): unsupported.append("css-generated-content")
    inert_css = None
    if "css-generated-content" in unsupported:
        from static_css_profile import inert_generated_css
        inert_css = inert_generated_css(dom)
        if inert_css is None and qualified_css:
            from static_css_profile import qualified_generated_css
            inert_css = qualified_generated_css(dom, lambda e: normalize(visible_text(e)))
        if inert_css is not None: unsupported = [reason for reason in unsupported if reason != "css-generated-content"]
    inputs = [sorted(element.attrib.items()) for element in body.iter("input")]
    controls = {"Inputs": inputs, "AnchoredText": normalize(visible_text(body, True)),
                "Labels": [(sorted(element.attrib.items()), normalize(visible_text(element, True))) for element in body.iter("label")]} if inputs else None
    return {"Text": text, "Links": links, "Images": images, "Resources": resources, "Tables": tables, "Lists": lists,
            "Preformatted": preformatted, "ActivePayloads": active_payloads, "Controls": controls, "InertCSS": inert_css, "Unsupported": sorted(set(unsupported)), "NormalizedWordListMarkers": normalized_markers}

def read_html(path: Path, *, restore_rtf_literal_emails=False) -> tuple[bytes | str, dict]:
    import olefile
    from rtf_representation import html_deencapsulator
    from RTFDE.deencapsulate import DeEncapsulator
    make_reader = html_deencapsulator if restore_rtf_literal_emails else DeEncapsulator
    from verify_rtf_sample import inspect_msg
    with olefile.OleFileIO(str(path), write_mode=False) as message:
        for stream, encoding in [("__substg1.0_10130102", None), ("__substg1.0_1013001F", "utf-16-le")]:
            if message.exists(stream):
                if message.get_size(stream) > MAX_BYTES: raise ValueError("html-size-limit")
                data = message.openstream(stream).read()
                value = data.decode(encoding, errors="strict") if encoding else data
                return value, {"Representation": "direct-html", "Bytes": len(data), "Sha256": hashlib.sha256(data).hexdigest()}
    _, raw, validated = inspect_msg(path)  # CRC + bounded LZFu validation before expansion.
    from rtf_representation import explicit_unicode_space_fallbacks
    raw = explicit_unicode_space_fallbacks(raw)
    encapsulated = make_reader(raw)
    encapsulated.deencapsulate()
    if encapsulated.content_type != "html": raise ValueError("arbitrary-rtf-not-converted")
    # RTFDE decodes the RTF codepage and emits UTF-8. Passing those bytes to an
    # HTML legacy charset sniffer would corrupt accented characters a second time.
    html = encapsulated.html.decode("utf-8", errors="strict")
    from rtf_representation import strip_rtf_list_fallbacks, only_redundant_list_markers
    clean, markers = strip_rtf_list_fallbacks(raw)
    if markers:
        recovered = make_reader(clean); recovered.deencapsulate()
        candidate = recovered.html.decode("utf-8", errors="strict")
        if only_redundant_list_markers(html, candidate, markers): html = candidate
    if len(html) > MAX_BYTES: raise ValueError("recovered-html-size-limit")
    return html, {"Representation": "rtfde-recovered-html", "Bytes": len(encapsulated.html),
                  "Sha256": hashlib.sha256(encapsulated.html).hexdigest(), "RtfValidation": validated}

def compare_profiles(left: dict, right: dict) -> dict:
    names = ["Text", "Links", "Images", "Resources", "Tables", "Lists", "Preformatted", "ActivePayloads", "Controls", "InertCSS"]
    components = {name: {"Equal": left[name] == right[name], "LeftSha256": digest(left[name]), "RightSha256": digest(right[name])} for name in names}
    supported = not left["Unsupported"] and not right["Unsupported"]
    return {"ComparedRelevantHtmlInformationEqual": supported and all(value["Equal"] for value in components.values()),
            "Supported": supported, "Components": components, "LeftUnsupported": left["Unsupported"], "RightUnsupported": right["Unsupported"],
            "LeftNormalizedWordListMarkers": left["NormalizedWordListMarkers"], "RightNormalizedWordListMarkers": right["NormalizedWordListMarkers"],
            "LeftTextCharacters": len(left["Text"]), "RightTextCharacters": len(right["Text"]),
            "LeftCounts": {name: len(left[name] or []) for name in names if name != "Text"}, "RightCounts": {name: len(right[name] or []) for name in names if name != "Text"}}

def metadata(directory: Path) -> dict:
    fields = json.loads((directory / "field-hashes.json").read_text(encoding="utf-8-sig"))
    return {key: value for key, value in fields.items() if not key.startswith("NativeBody-")}

def analyze(run: Path) -> dict:
    sources = []
    for index in range(5):
        directory = run / f"s{index:02}"
        html, representation = read_html(directory / "message.msg")
        sources.append((metadata(directory), html_profile(html), representation))
    rows = []
    for target_index in range(5):
        directory = run / f"n{target_index:02}"
        target_metadata = metadata(directory)
        matches = [index for index, source in enumerate(sources) if source[0] == target_metadata]
        html, representation = read_html(directory / "native-copy.msg")
        profile = html_profile(html)
        comparisons = []
        for index in matches:
            comparisons.append({"SourceIndex": index, "SourceRepresentation": sources[index][2], **compare_profiles(sources[index][1], profile)})
        rows.append({"TargetIndex": target_index, "TargetRepresentation": representation, "MetadataAndAttachmentCandidates": matches, "Comparisons": comparisons})
    # A perfect one-to-one match preserves duplicate multiplicity. Never collapse
    # equal content to a single message or accept only a per-target 'some match'.
    edges = [[candidate["SourceIndex"] for candidate in row["Comparisons"] if candidate["ComparedRelevantHtmlInformationEqual"]] for row in rows]
    def match(position, used):
        if position == len(edges): return []
        for source in edges[position]:
            if source not in used:
                rest = match(position + 1, used | {source})
                if rest is not None: return [source] + rest
        return None
    assignment = match(0, set())
    return {"SchemaVersion": 1, "Offline": True, "DiagnosticOnly": True, "ProductionPolicyChanged": False, "RemovalAllowed": False,
            "ComparedRelevantInformationEqualForFive": assignment is not None, "OneToOneSourceAssignment": assignment,
            "DuplicateMultiplicityPreserved": assignment is not None, "Items": rows,
            "Libraries": {name: importlib.metadata.version(name) for name in ("RTFDE", "olefile", "compressed-rtf", "html5lib")},
            "Profile": "ASCII HTML whitespace normalization; independently retained ordered-list counters allow exact Word N.<TAB> fallback removal; text, hyperlink labels/targets, image resources/alt text, resource attributes, table cells/spans, lists and preformatted text compared. Existing metadata and attachment hashes must also match.",
            "Limits": "Does not certify rendered appearance/CSS layout, remote-resource contents, all MAPI properties, server completeness or removal safety. Arbitrary RTF, special bodies and generated CSS content are unsupported. Script code/attributes are compared statically without execution. Extended SMTP recipients and UTC dates require the separate metadata evidence. No links/images are fetched."}

def self_test() -> dict:
    from rtf_representation import self_test as representation_tests
    representation_count = representation_tests()
    from rtf_representation import html_text_destination_self_test
    representation_count += html_text_destination_self_test()
    tests = []
    def check(name, a, b, expected, ndr=False):
        actual = compare_profiles(html_profile(a, normalize_layout_nbsp=ndr), html_profile(b, normalize_layout_nbsp=ndr))["ComparedRelevantHtmlInformationEqual"]
        if actual != expected: raise AssertionError(name)
        tests.append(name)
    check("format-and-normal-html-whitespace", "<p>Olá <b>mundo</b>!</p>", "<div>Olá <span>mundo</span>!</div>", True)
    check("ordinary-mail-nbsp-profile-unchanged", "<p>A&nbsp; B</p>", "<p>A B</p>", False)
    check("nbsp-layout-equivalence", "<p>A&nbsp; B</p>", "<p>A B</p>", True, ndr=True)
    check("nbsp-word-gap-preserved", "<p>A&nbsp;B</p>", "<p>AB</p>", False, ndr=True)
    check("nbsp-preformatted-preserved", "<pre>A&nbsp;B</pre>", "<pre>A B</pre>", False, ndr=True)
    check("nbsp-attribute-preserved", "<a href='a&#160;b'>A</a>", "<a href='a b'>A</a>", False, ndr=True)
    check("changed-amount", "<p>Total: 100</p>", "<p>Total: 900</p>", False)
    check("accent-preserved", "<p>ação</p>", "<p>acao</p>", False)
    check("changed-link-target", "<a href='https://example.invalid/a'>A</a>", "<a href='https://example.invalid/b'>A</a>", False)
    check("changed-link-label", "<a href='https://example.invalid/a'>A</a>", "<a href='https://example.invalid/a'>B</a>", False)
    check("inline-image-content-id", "<img src='cid:a' alt='foto'>", "<img src='cid:b' alt='foto'>", False)
    check("lost-inline-image", "<img src='cid:a' alt='foto'>", "", False)
    check("changed-image-alt", "<img src='cid:a' alt='A'>", "<img src='cid:a' alt='B'>", False)
    check("table-cell-order", "<table><tr><td>A</td><td>B</td></tr></table>", "<table><tr><td>B</td><td>A</td></tr></table>", False)
    check("table-cell-span", "<table><tr><td colspan='2'>A</td></tr></table>", "<table><tr><td>A</td></tr></table>", False)
    check("list-start-value", "<ol start='2'><li>A</li></ol>", "<ol start='3'><li>A</li></ol>", False)
    check("word-duplicate-counter", "<ol start='2'><li class='MsoListParagraph' style='mso-list:l1'>A</li></ol>", "<ol start='2'><li class='MsoListParagraph' style='mso-list:l1'>2.\tA</li></ol>", False)
    # Exact standalone fallback is a separate text node before the content span.
    check("word-standalone-counter", "<ol start='2'><li class='MsoListParagraph' style='mso-list:l1'><span>A</span></li></ol>", "<ol start='2'><li class='MsoListParagraph' style='mso-list:l1'>2.\t<span>A</span></li></ol>", True)
    check("wrong-word-counter", "<ol start='2'><li class='MsoListParagraph' style='mso-list:l1'><span>A</span></li></ol>", "<ol start='2'><li class='MsoListParagraph' style='mso-list:l1'>3.\t<span>A</span></li></ol>", False)
    check("ordinary-numeric-text", "<ol start='2'><li>2.\t<span>A</span></li></ol>", "<ol start='2'><li><span>A</span></li></ol>", False)
    check("preformatted-spacing", "<pre>A  B</pre>", "<pre>A B</pre>", False)
    check("script-preserved-without-executing", "<script>x()</script><p>A</p>", "<script>x()</script><p>A</p>", True)
    check("script-change-detected-without-executing", "<script>x()</script><p>A</p>", "<script>y()</script><p>A</p>", False)
    check("generated-css-unsupported", "<style>p:after{content:'B'}</style><p>A</p>", "<style>p:after{content:'B'}</style><p>A</p>", False)
    check("input-format-only", "<p>A <input type='text' value='Ana'> B</p>", "<div>A <span><input value='Ana' type='text'></span> B</div>", True)
    check("input-value-changed", "<input value='Ana'>", "<input value='Bia'>", False)
    check("input-date-changed", "<input type='date' value='2024-12-31'>", "<input type='date' value='2025-01-01'>", False)
    check("input-checked-changed", "<input type='checkbox' checked>", "<input type='checkbox'>", False)
    check("input-removed", "<p>A<input value='X'>B</p>", "<p>AB</p>", False)
    check("input-moved", "<p>A<input value='X'>B</p>", "<p>AB<input value='X'></p>", False)
    check("input-label-binding-changed", "<label for='a'>Nome</label><input id='a'>", "<label for='b'>Nome</label><input id='a'>", False)
    check("input-type-changed", "<input type='text' value='A'>", "<input type='hidden' value='A'>", False)
    check("input-handler-changed", "<input onclick='a()'>", "<input onclick='b()'>", False)
    check("input-order-changed", "<input name='a'><input name='b'>", "<input name='b'><input name='a'>", False)
    check("form-still-unsupported", "<form><input value='A'></form>", "<form><input value='A'></form>", False)
    token = "\ufff0" + digest(["input", [("value", "X")]]) + "\ufff1"
    check("input-literal-marker-collision", f"A<input value='X'>{token}B", f"A{token}<input value='X'>B", False)
    inert = "<style>@media (max-width:600px){.unused{content:'X';}}</style><p>A</p>"
    check("unused-generated-css", inert, inert, True)
    check("unused-css-format-only", inert, inert.replace('<p>A</p>', '<div>A</div>'), True)
    check("unused-css-value-changed", inert, inert.replace("content:'X'", "content:'Y'"), False)
    check("unused-css-class-introduced", inert, inert.replace('<p>', '<p class="unused">'), False)
    active = inert.replace('<p>', '<p class="unused">')
    check("active-generated-css-blocked", active, active, False)
    upper_class = active.replace('class="unused"', 'class="UNUSED"')
    check("generated-case-insensitive-class-blocked", upper_class, upper_class, False)
    for selector in ('#unused', 'p .unused', '.unused:before', '.a,.unused'):
        value = inert.replace('.unused{', selector+'{')
        check("generated-selector-blocked-"+selector, value, value, False)
    extra = inert.replace("content:'X';", "content:'X';color:red;")
    check("generated-extra-declarations-blocked", extra, extra, False)
    dynamic = inert.replace("content:'X'", "content:attr(data-info)")
    check("generated-dynamic-value-blocked", dynamic, dynamic, False)
    scripted = inert + "<script>x()</script>"
    check("generated-script-context-blocked", scripted, scripted, False)
    inline = "<p style=\"content:'X'\">A</p>"
    check("generated-inline-blocked", inline, inline, False)
    escaped = inert + r"<style>.active{co\6e tent:'Y';}</style>"
    check("generated-escaped-identifier-blocked", escaped, escaped, False)
    return {"Passed": len(tests) + representation_count, "Tests": tests, "Offline": True, "ProductionPolicyChanged": False}

def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--deps-dir", type=Path)
    parser.add_argument("--run-dir", type=Path)
    parser.add_argument("--output", type=Path)
    parser.add_argument("--self-test", action="store_true")
    args = parser.parse_args()
    if args.deps_dir: sys.path.insert(0, str(args.deps_dir.resolve(strict=True)))
    if args.self_test:
        print(json.dumps(self_test(), ensure_ascii=False))
        return 0
    if not args.run_dir or not args.output: parser.error("--run-dir e --output são obrigatórios")
    if args.output.exists(): raise FileExistsError("output-already-exists")
    report = analyze(args.run_dir.resolve(strict=True))
    with args.output.open("x", encoding="utf-8") as output:
        json.dump(report, output, ensure_ascii=False, indent=2)
        output.flush(); os.fsync(output.fileno())
    print(json.dumps({"ComparedRelevantInformationEqualForFive": report["ComparedRelevantInformationEqualForFive"], "Assignment": report["OneToOneSourceAssignment"], "DiagnosticOnly": True}))
    return 0 if report["ComparedRelevantInformationEqualForFive"] else 3

if __name__ == "__main__":
    try: raise SystemExit(main())
    except Exception as error:
        print(json.dumps({"ErrorType": type(error).__name__, "ProductionPolicyChanged": False}), file=sys.stderr)
        raise SystemExit(1)
