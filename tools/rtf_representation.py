"""Bounded representation checks; never erase literal HTML list text."""
import re

def remove_proven_word_markers(body):
    """Remove only standalone Word fallbacks with an independently kept counter."""
    def marker(number,kind):
        if not 1 <= number <= 3999:return None
        if kind=='1':return str(number)+'.'
        if kind in ('a','A'):
            value=''
            while number:
                number,rem=divmod(number-1,26);value=chr(97+rem)+value
            return (value.upper() if kind=='A' else value)+'.'
        if kind in ('i','I'):
            value=''
            for n,s in [(1000,'M'),(900,'CM'),(500,'D'),(400,'CD'),(100,'C'),(90,'XC'),(50,'L'),(40,'XL'),(10,'X'),(9,'IX'),(5,'V'),(4,'IV'),(1,'I')]:
                count,number=divmod(number,n);value+=s*count
            return (value.lower() if kind=='i' else value)+'.'
        return None
    changed=0
    for parent in body.iter():
        if parent.tag not in ('ol','ul'):continue
        number=int(parent.get('start','1'));step=-1 if 'reversed' in parent.attrib else 1
        for li in parent:
            if li.tag!='li':continue
            number=int(li.get('value',str(number)))
            expected=marker(number,parent.get('type','1')) if parent.tag=='ol' else ('*' if parent.get('type','disc') in ('disc','circle','square') else None)
            if expected and 'MsoListParagraph' in li.get('class','').split() and 'mso-list:' in li.get('style','') and li.text in (expected+'\t',expected+'\t\n',expected+'\t\r\n'):
                li.text='';changed+=1
            number+=step
    return changed

def explicit_unicode_space_fallbacks(raw):
    """Spell a genuine uc1 space fallback as hex before RTFDE parses it.

    The delimiter after a numeric control word is not its fallback character.
    RTFDE can consume later visible characters when that fallback is a space.
    Hex spelling preserves the same fallback byte and Unicode value. Lexical
    scope, escaped backslashes and binary data must remain untouched.
    """
    stack, spans, i = [1], [], 0
    while i < len(raw):
        if raw[i] == 123:
            stack.append(stack[-1]); i += 1; continue
        if raw[i] == 125:
            if len(stack) == 1: raise ValueError('rtf-group-underflow')
            stack.pop(); i += 1; continue
        if raw[i] != 92:
            i += 1; continue
        start = i; i += 1
        if i == len(raw): raise ValueError('rtf-control-truncated')
        if raw[i] in b'\\{}': i += 1; continue
        if raw[i] == 39:
            if not re.fullmatch(rb'[0-9a-fA-F]{2}', raw[i+1:i+3]): raise ValueError('rtf-hex-invalid')
            i += 3; continue
        match = re.match(rb'([a-zA-Z]+)(-?\d+)?( ?)', raw[i:])
        if not match: i += 1; continue
        i += match.end()
        word, number, delimiter = match.groups()
        if word == b'uc':
            value = int(number or b'-1')
            if value < 0: raise ValueError('rtf-uc-invalid')
            stack[-1] = value
        elif word == b'bin':
            size = int(number or b'-1')
            if size < 0 or i + size > len(raw): raise ValueError('rtf-binary-size')
            i += size
        elif word == b'u' and number is not None and stack[-1] == 1 and delimiter == b' ' and raw[i:i+1] == b' ':
            spans.append((start, i+1, b'\\u' + number + b"\\'20"))
            i += 1
    if len(stack) != 1: raise ValueError('rtf-group-unclosed')
    for start, end, replacement in reversed(spans): raw = raw[:start] + replacement + raw[end:]
    return raw

def unicode_fallback_self_test():
    cases = [
        (br'{\rtf1\uc1\u8194  (Para)}', br"{\rtf1\uc1\u8194\'20(Para)}"),
        (br'{\rtf1\\u8194  literal}', br'{\rtf1\\u8194  literal}'),
        (br'{\rtf1\uc0\u8194  keep}', br'{\rtf1\uc0\u8194  keep}'),
        (br'{\rtf1\uc0{\uc1\u8194  A}\u8194  B}', br"{\rtf1\uc0{\uc1\u8194\'20A}\u8194  B}"),
        (br'{\rtf1\u8194? text}', br'{\rtf1\u8194? text}'),
    ]
    payload = br'\u8194  {}'
    binary = b'{\\rtf1\\bin' + str(len(payload)).encode() + b' ' + payload + b'}'
    cases.append((binary, binary))
    for original, expected in cases:
        if explicit_unicode_space_fallbacks(original) != expected: raise AssertionError('unicode-fallback-boundary')
    for invalid in (b'{', b'}', b'{\\bin50 x}'):
        try: explicit_unicode_space_fallbacks(invalid)
        except ValueError: pass
        else: raise AssertionError('invalid-rtf-accepted')
    return len(cases) + 3

MARKER = re.compile(rb"\{\\pntext ([0-9]+\.|\*)\\tab\}")

def strip_rtf_list_fallbacks(raw):
    """Remove only complete, genuine RTF pntext groups, not escaped text/bin."""
    stack, spans, markers = [], [], []
    i = 0
    while i < len(raw):
        c = raw[i]
        if c == 123:
            stack.append(i); i += 1
        elif c == 125:
            if not stack: raise ValueError('rtf-group-underflow')
            start = stack.pop(); matched = MARKER.fullmatch(raw[start:i+1])
            if matched:
                spans.append((start, i+1)); markers.append(matched[1].decode('ascii'))
            i += 1
        elif c == 92:
            i += 1
            if i == len(raw): raise ValueError('rtf-control-truncated')
            if raw[i] in b"\\{}": i += 1; continue
            if raw[i] == 39: i += 3; continue
            matched = re.match(rb'([a-zA-Z]+)(-?\d+)? ?', raw[i:])
            if not matched: i += 1; continue
            i += matched.end()
            if matched[1] == b'bin':
                size = int(matched[2] or b'-1')
                if size < 0 or i + size > len(raw): raise ValueError('rtf-binary-size')
                i += size
        else: i += 1
    if stack: raise ValueError('rtf-group-unclosed')
    value = raw
    for start, end in reversed(spans): value = value[:start] + value[end:]
    return value, markers

def only_redundant_list_markers(before, after, markers):
    import html5lib
    left = html5lib.parse(before, treebuilder='etree', namespaceHTMLElements=False)
    right = html5lib.parse(after, treebuilder='etree', namespaceHTMLElements=False)
    allowed = {}
    for parent in right.iter():
        if parent.tag not in ('ol','ul'): continue
        counter = int(parent.get('start','1'))
        step = -1 if 'reversed' in parent.attrib else 1
        for li in parent:
            if li.tag != 'li': continue
            counter = int(li.get('value',str(counter)))
            if parent.tag == 'ul' and parent.get('type','disc') in ('disc', 'circle', 'square'): allowed[id(li)] = '*'
            if parent.tag == 'ol' and parent.get('type','1') == '1': allowed[id(li)] = str(counter)+'.'
            counter += step
    changed = []
    def equal(a,b):
        if a.tag != b.tag or a.attrib != b.attrib or a.tail != b.tail or len(a) != len(b): return False
        if a.text != b.text:
            marker = allowed.get(id(b))
            if marker is None or a.text != marker+'\t'+(b.text or ''): return False
            changed.append(marker)
        return all(equal(x,y) for x,y in zip(a,b))
    return equal(left,right) and changed == markers

EMPTY_TEXT_CONTROLS = {'rtf','ansi','ansicpg','fromtext','deff','deftab','f','fbidis','fcharset','fonttbl','colortbl',
    'red','green','blue','fswiss','fmodern','fnil','fs','uc','pard','plain','viewkind','viewscale','generator',
    'formatConverter','bkmkstart','bkmkend'}

def empty_fromtext_rtf(raw, plain_streams):
    """Only a validated text-derived EMPTY body, never arbitrary rich text."""
    if len(raw) > 16384 or not plain_streams or any(value != b'' for value in plain_streams): return False
    words = set(m[1].decode('ascii') for m in re.finditer(rb'\\([a-zA-Z]+)',raw))
    if not words <= EMPTY_TEXT_CONTROLS or 'fromtext' not in words: return False
    from RTFDE.deencapsulate import DeEncapsulator
    parsed = DeEncapsulator(raw); parsed.deencapsulate()
    return parsed.content_type == 'text' and parsed.text == b''

def validated_fromtext_rtf(raw, unicode_stream):
    """Only text-derived RTF whose decoded body matches its Unicode stream.

    Font/color tables and paragraph serialization can differ; rich objects,
    fields, pictures, hidden text and arbitrary formatting remain unsupported.
    The Unicode stream's exact bytes remain separately compared by the caller.
    """
    if unicode_stream is None or len(raw) > 32 * 1024 * 1024: return False
    words = {m[1].decode('ascii') for m in re.finditer(rb'\\([a-zA-Z]+)', raw)}
    if not words <= EMPTY_TEXT_CONTROLS | {'line', 'par', 'tab'} or 'fromtext' not in words: return False
    from RTFDE.deencapsulate import DeEncapsulator
    try:
        parsed = DeEncapsulator(explicit_unicode_space_fallbacks(raw)); parsed.deencapsulate()
        if parsed.content_type != 'text': return False
        text = unicode_stream.decode('utf-16-le', errors='strict').rstrip('\0')
        decoded = parsed.text.decode('utf-8', errors='strict')
    except Exception:
        return False
    normalize = lambda value: value.replace('\r\n', '\n').replace('\r', '\n').rstrip('\n')
    return normalize(decoded) == normalize(text)

def validated_empty_fromtext_rtf(raw, unicode_stream, ansi_stream):
    words = {m[1].decode('ascii') for m in re.finditer(rb'\\([a-zA-Z]+)', raw)}
    if len(raw) > 16384 or not words <= EMPTY_TEXT_CONTROLS | {'line', 'par', 'tab'} or 'fromtext' not in words: return False
    try:
        if unicode_stream is not None and unicode_stream.decode('utf-16-le', errors='strict').strip('\r\n'): return False
        if ansi_stream is not None and ansi_stream.strip(b'\r\n'): return False
        from RTFDE.deencapsulate import DeEncapsulator
        parsed = DeEncapsulator(explicit_unicode_space_fallbacks(raw)); parsed.deencapsulate()
        return parsed.content_type == 'text' and not parsed.text.decode('utf-8', errors='strict').strip('\r\n')
    except Exception: return False

def canonical_fromtext_unicode(data):
    text = data.decode('utf-16-le', errors='strict').replace('\r\n', '\n').replace('\r', '\n')
    text = text.rstrip('\n')  # Structural final paragraphs only; keep spaces/tabs and all interior lines.
    return text.encode('utf-8')

def canonical_unicode_self_test():
    C = lambda text: canonical_fromtext_unicode(text.encode('utf-16-le'))
    assert C('A\r\nB') == C('A\nB') == C('A\rB')
    assert C('\r\n') == C('\r\n\r\n') == b''
    assert C('A\r\n') == b'A'
    assert C('A\r\n') == C('A')
    assert C('A\n\nB') != C('A\nB')
    assert C('A  B') != C('A B')
    assert C('\t') != b''
    assert C('\0') != b''
    assert C('ação') != C('acao')
    assert C('A') != C('B')
    try: canonical_fromtext_unicode(b'\x01')
    except UnicodeDecodeError: pass
    else: raise AssertionError('invalid-unicode-accepted')
    return 11

def empty_fromtext_self_test():
    prefix = br'{\rtf1\ansi\ansicpg1252\fromtext\deff0{\fonttbl{\f0\fnil Arial;}}\pard\plain\f0 '
    assert validated_empty_fromtext_rtf(prefix+br'\par}', None, None)
    assert validated_empty_fromtext_rtf(prefix+br'\par}', '\r\n'.encode('utf-16-le'), None)
    assert validated_empty_fromtext_rtf(prefix+br'\par\par}', '\r\n\r\n'.encode('utf-16-le'), b'\r\n')
    assert not validated_empty_fromtext_rtf(prefix+b'A}', None, None)
    assert not validated_empty_fromtext_rtf(prefix+b'}', 'A'.encode('utf-16-le'), None)
    assert not validated_empty_fromtext_rtf(prefix+b'}', None, b'A')
    assert not validated_empty_fromtext_rtf(prefix+b'}', ' '.encode('utf-16-le'), None)
    assert not validated_empty_fromtext_rtf(prefix+b'}', '\t'.encode('utf-16-le'), None)
    assert not validated_empty_fromtext_rtf(prefix+br'\pict ABC}', None, None)
    assert not validated_empty_fromtext_rtf(br'{\rtf1\fromhtml1}', None, None)
    return 10

def fromtext_self_test():
    text = 'A\r\nB'.encode('utf-16-le')
    prefix = br'{\rtf1\ansi\ansicpg1252\fromtext\deff0{\fonttbl{\f0\fnil Arial;}}\pard\plain\f0 '
    assert validated_fromtext_rtf(prefix + br'A\line B}', text)
    assert validated_fromtext_rtf(prefix + br'A\par B\par}', text)
    assert not validated_fromtext_rtf(prefix + br'A\line C}', text)
    for control in (br'\pict', br'\object', br'\field', br'\v', br'\deleted', br'\b', br'\htmltag'):
        assert not validated_fromtext_rtf(prefix+control+br' A}', b'A\0')
    assert not validated_fromtext_rtf(prefix + b'A}', None)
    assert not validated_fromtext_rtf(br'{\rtf1\fromhtml1 A}', b'A\0')
    return 12

def self_test():
    a='<ul><li><b>A</b></li></ul>'; b='<ul><li>*\t<b>A</b></li></ul>'
    assert only_redundant_list_markers(b,a,['*'])
    assert not only_redundant_list_markers(b,a,['1.'])
    assert not only_redundant_list_markers('<p>*\tA</p>','<p>A</p>',['*'])
    assert not only_redundant_list_markers('<ol><li>2.\tA</li></ol>','<ol><li>A</li></ol>',['2.'])
    assert not only_redundant_list_markers('<ul><li>*\tA</li></ul>','<ul><li>B</li></ul>',['*'])
    assert not only_redundant_list_markers('<ul><li>*\t<a href="x">A</a></li></ul>','<ul><li><a href="y">A</a></li></ul>',['*'])
    for bullet in ('disc', 'circle', 'square'):
        before = f'<ul type="{bullet}"><li>*\tA</li></ul>'
        after = f'<ul type="{bullet}"><li>A</li></ul>'
        assert only_redundant_list_markers(before, after, ['*'])
        assert not only_redundant_list_markers(before, after.replace('A</li>', 'B</li>'), ['*'])
        assert not only_redundant_list_markers(before, after.replace(f'type="{bullet}"', 'type="custom"'), ['*'])
    assert not only_redundant_list_markers('<ul type="custom"><li>*\tA</li></ul>', '<ul type="custom"><li>A</li></ul>', ['*'])
    mixed_before = '<ul><li>*\tA</li></ul><ul type="circle"><li>*\tB</li></ul>'
    mixed_after = '<ul><li>A</li></ul><ul type="circle"><li>B</li></ul>'
    assert only_redundant_list_markers(mixed_before, mixed_after, ['*', '*'])
    assert not only_redundant_list_markers(mixed_before, mixed_after.replace('B</li>', 'C</li>'), ['*', '*'])
    clean, markers = strip_rtf_list_fallbacks(br'{\rtf1 {\pntext *\tab} A}')
    assert clean == br'{\rtf1  A}' and markers == ['*']
    value = br'{\rtf1 \bin16 {\pntext *\tab}xx}'
    assert strip_rtf_list_fallbacks(value)[1] == []
    assert not empty_fromtext_rtf(br'{\rtf1\fromtext {\pict ABC}}',[b''])
    assert not empty_fromtext_rtf(br'{\rtf1\fromtext\field link}',[b''])
    assert not empty_fromtext_rtf(br'{\rtf1\fromtext}',[b'content'])
    return 23

def html_deencapsulator(raw):
    """Restore literal angle-bracket email addresses proven by RTF text tokens.

    Markup destinations, attribute values, comments and script/style raw text
    are untouched. RTFDE otherwise exposes these literal addresses as HTML tags.
    """
    from RTFDE.deencapsulate import DeEncapsulator
    from lark import Tree, Token
    from collections import Counter
    from html.parser import HTMLParser
    from html import escape
    email = re.compile(rb"<[a-zA-Z0-9.!#$%&'*+/=?^_`{|}~-]+@[a-zA-Z0-9-]+(?:\.[a-zA-Z0-9-]+)+>")
    class Reader(DeEncapsulator):
        def strip_htmlrtf_tokens(self):
            tree = super().strip_htmlrtf_tokens()
            self.literal_emails = Counter()
            if self.get_content_type() != 'html': return tree
            from copy import deepcopy
            from RTFDE.transformers import StripControlWords, StripUnusedSpecialCharacters, StripNonVisibleRTFGroups
            visible = StripNonVisibleRTFGroups().transform(StripUnusedSpecialCharacters().transform(StripControlWords().transform(deepcopy(tree))))
            pending, visited, parts = [(visible, 0)], 0, []
            while pending:
                node, depth = pending.pop(); visited += 1
                if visited > 2000000 or depth > 256: raise ValueError('rtf-html-token-limit')
                if isinstance(node, Tree):
                    if str(node.data) in {'htmltag_group', 'mhtmltag_group'}:
                        parts.append(b'\0'); continue
                    pending.extend((child, depth + 1) for child in reversed(node.children))
                elif isinstance(node, Token):
                    if node.type in {'STRING', 'SPACE_SAVE'}:
                        if not isinstance(node.value, bytes): raise ValueError('rtf-decoded-text-type')
                        parts.append(node.value)
                    elif node.type == 'CONTROLWORD' and node.value in {br'\par', br'\tab', br'\line'}:
                        parts.append(b'\0')
            self.literal_emails.update(value.decode('ascii') for value in email.findall(b''.join(parts)))
            return tree
        def deencapsulate(self):
            super().deencapsulate()
            if self.content_type != 'html' or not self.literal_emails: return
            html = self.html.decode('utf-8', errors='strict')
            lines = [0]
            for index, char in enumerate(html):
                if char == '\n': lines.append(index + 1)
            allowed, spans = self.literal_emails, []
            class Locator(HTMLParser):
                def handle_starttag(self, tag, attrs):
                    text = self.get_starttag_text()
                    if not attrs and text in allowed:
                        line, column = self.getpos(); spans.append((lines[line-1] + column, text))
            locator = Locator(convert_charrefs=False); locator.feed(html); locator.close()
            counts = Counter(text for _, text in spans)
            if any(counts[text] > allowed[text] for text in counts): raise ValueError('ambiguous-literal-email-tag')
            for offset, text in reversed(spans): html = html[:offset] + escape(text) + html[offset+len(text):]
            self.html = html.encode('utf-8')
    return Reader(raw)


def html_text_destination_self_test():
    from compare_relevant_message import html_profile, compare_profiles
    template = br'{\rtf1\ansi\ansicpg1252\fromhtml1{\fonttbl{\f0\fnil Arial;}}\f0{\*\htmltag64 <html><body><p>}PAYLOAD{\*\htmltag64 </p></body></html>}}'
    cases = [(b'Nome <email@example.invalid> & texto', '<p>Nome &lt;email@example.invalid&gt; &amp; texto</p>'),
             (br'Nome \'3cemail@example.invalid\'3e', '<p>Nome &lt;email@example.invalid&gt;</p>'),
             (br'Nome \uc1\u60?email@example.invalid\u62?', '<p>Nome &lt;email@example.invalid&gt;</p>')]
    for payload, expected in cases:
        reader = html_deencapsulator(template.replace(b'PAYLOAD', payload)); reader.deencapsulate()
        actual = html_profile(reader.html.decode('utf-8'))
        if not compare_profiles(html_profile(expected), actual)['ComparedRelevantHtmlInformationEqual']: raise AssertionError('literal-rtf-html-text-lost')
    return len(cases)

def validated_repair_text_representation(raw, unicode_stream, ansi_stream=None):
    """Allow Unicode presentation differences only with a retained exact RTF.

    The caller MUST keep the compressed RTF hash exact in its profile. Thus
    nonbreaking characters remain recoverable even when Outlook regenerates
    the alternate Unicode representation with ordinary spaces.
    """
    if unicode_stream is None or ansi_stream is not None or len(raw) > 32 * 1024 * 1024: return None
    words = {m[1].decode('ascii') for m in re.finditer(rb'\\([a-zA-Z]+)', raw)}
    if 'fromtext' not in words or not words <= EMPTY_TEXT_CONTROLS | {'line', 'par', 'tab', 'u'}: return None
    from RTFDE.deencapsulate import DeEncapsulator
    def normalize(text):
        if text.endswith('\0'): text = text[:-1]
        return text.replace('\r\n','\n').replace('\r','\n').replace('\u00a0',' ').rstrip('\n')
    try:
        reader = DeEncapsulator(explicit_unicode_space_fallbacks(raw)); reader.deencapsulate()
        if reader.content_type != 'text': return None
        decoded = normalize(reader.text.decode('utf-8', errors='strict'))
        if decoded != normalize(unicode_stream.decode('utf-16-le', errors='strict')): return None
        return decoded.encode('utf-16-le')
    except Exception: return None
