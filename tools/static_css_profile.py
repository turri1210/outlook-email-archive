"""Conservative proof for unused, literal generated-content CSS rules.

No CSS evaluation. Every generated-content declaration must be a complete
simple .class rule containing only a literal content string, and that class
must be absent from the document. All stylesheet bytes remain fingerprinted.
"""
import re

CONTENT = re.compile(r'(?:^|[;{])\s*content\s*:', re.I)
RULE = re.compile(r'''(?P<selector>\.[a-zA-Z_][a-zA-Z0-9_-]*)\s*\{\s*(?P<property>content)\s*:\s*(?:'(?:[^'\\]|\\.)*'|"(?:[^"\\]|\\.)*")\s*;?\s*\}''', re.I)

def masked_css(css):
    value = list(css); i = 0
    while i < len(css):
        if css.startswith('/*', i):
            end = css.find('*/', i+2)
            if end < 0: return None
            value[i:end+2] = ' ' * (end+2-i); i = end+2
        elif css[i] in "\"'":
            start, quote = i, css[i]; i += 1
            while i < len(css):
                if css[i] == '\\': i += 2; continue
                if css[i] == quote: break
                if css[i] in '\r\n': return None
                i += 1
            if i >= len(css): return None
            i += 1; value[start:i] = ' ' * (i-start)
        else: i += 1
    result = ''.join(value)
    # Escaped identifiers require a full CSS parser; never guess their meaning.
    return None if '\\' in result else result

def inert_generated_css(dom):
    elements = [e for e in dom.iter() if isinstance(e.tag, str)]
    if any(e.tag == 'script' or any(k.lower().startswith('on') for k in e.attrib) for e in elements): return None
    classes = {value.casefold() for e in elements for value in e.get('class', '').split()}
    styles, inactive = [], []
    for element in elements:
        inline = element.get('style', '')
        mask = masked_css(inline)
        if mask is None or CONTENT.search(mask): return None
        if element.tag != 'style': continue
        css = element.text or ''; mask = masked_css(css)
        if mask is None: return None
        declarations = {match.start() + match.group().lower().index('content') for match in CONTENT.finditer(mask)}
        proven = set()
        for match in RULE.finditer(css):
            prefix = mask[:match.start()].rstrip()
            if prefix and prefix[-1] not in '{};': continue
            # A match inside a comment/string is not a real CSS rule.
            if mask[match.start():match.start()+1] != '.': continue
            selector = match['selector'][1:]
            if selector.casefold() in classes: return None
            proven.add(match.start('property')); inactive.append(selector)
        if declarations != proven: return None
        styles.append((sorted(element.attrib.items()), css.replace('\r\n', '\n').replace('\r', '\n')))
    return {'Stylesheets': styles, 'InactiveClasses': inactive}

def qualified_generated_css(dom, text_of):
    """Finite static CSS selectors with exact retained styles and bindings.

    No script/event context, inline content, counters, URLs or dynamic values.
    All generated declarations must belong to one of the selectors below.
    """
    elements = [e for e in dom.iter() if isinstance(e.tag, str)]
    # Reviewed survey customization consists solely of an if(false) block.
    # Pin its complete source, rather than interpreting arbitrary JavaScript.
    import hashlib
    for element in elements:
        if element.tag == 'script':
            code = (element.text or '').replace('\r\n','\n').replace('\r','\n').strip()
            if element.attrib != {'type':'text/javascript'} or hashlib.sha256(code.encode('utf-8')).hexdigest() != 'fd9b2167848fa488d39dff0b017fbf94c06443e0ecf81bf9579d2328965485df': return None
        for name, value in element.attrib.items():
            if name.lower().startswith('on') and not (element.tag == 'img' and name == 'onerror' and value == "this.style.display='none'"): return None
    for e in elements:
        inline = masked_css(e.get('style', ''))
        if inline is None or CONTENT.search(inline): return None
    anchors = [e for e in elements if e.tag == 'a']
    anchor_index = {id(e): i for i, e in enumerate(anchors)}
    blocks = re.compile(r'([^{}]+)\{([^{}]*)\}')
    content_value = re.compile(r'''\s*(?:"(?:\\.|[^"\\])*"|'(?:\\.|[^'\\])*'|none|normal)\s*(?:!\s*important\s*)?''', re.I | re.S)
    class_selector = re.compile(r'''span\[class=(?:"([A-Za-z0-9_-]+)"|'([A-Za-z0-9_-]+)'|([A-Za-z0-9_-]+))\]\s+a::?(before|after)''')
    def bindings(selector):
        match = class_selector.fullmatch(selector)
        if match:
            cls = next(v for v in match.groups()[:3] if v is not None)
            found = {id(a): a for e in elements if e.tag == 'span' and e.get('class') == cls for a in e.iter('a')}
            return [('a', anchor_index[id(a)], sorted(a.attrib.items()), text_of(a)) for a in anchors if id(a) in found]
        match = re.fullmatch(r'(blockquote|q)(?:\s+(footer|small|\.small))?::?(before|after)', selector)
        if not match: return None
        tag, descendant, _ = match.groups()
        found = []
        for e in elements:
            if e.tag != tag: continue
            if descendant:
                for child in e.iter():
                    if child is e: continue
                    if (descendant == '.small' and 'small' in child.get('class', '').split()) or child.tag == descendant: found.append(child)
            else: found.append(e)
        result, used = [], set()
        for e in found:
            if id(e) in used: continue
            used.add(id(e)); result.append((e.tag, sorted(e.attrib.items()), text_of(e), [anchor_index[id(a)] for a in e.iter('a')]))
        return result
    styles, proof = [], []
    for style in (e for e in elements if e.tag == 'style'):
        css = style.text or ''; mask = masked_css(css)
        if mask is None: return None
        positions = {m.start() + m.group().lower().index('content') for m in CONTENT.finditer(mask)}
        if not positions: continue
        covered = set()
        for rule in blocks.finditer(mask):
            start, end = rule.span(2)
            declarations = [p for p in positions if start <= p < end]
            if not declarations: continue
            selector_text = css[rule.start(1):rule.end(1)].strip()
            selectors = [re.sub(r'\s+', ' ', value.strip()) for value in selector_text.split(',')]
            mapped = [(selector, bindings(selector)) for selector in selectors]
            if any(value is None for _, value in mapped): return None
            for pos in declarations:
                colon = mask.find(':', pos, end); finish = mask.find(';', colon + 1, end)
                if colon < 0: return None
                if finish < 0: finish = end
                if not content_value.fullmatch(css[colon+1:finish]): return None
                covered.add(pos)
            proof.extend(mapped)
        if covered != positions: return None
        styles.append((sorted(style.attrib.items()), css.replace('\r\n','\n').replace('\r','\n')))
    if not styles: return None
    return {'Version': 'qualified-generated-css-v1', 'Stylesheets': styles, 'Bindings': proof}
