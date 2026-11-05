import pathlib,sys
sys.path[:0]=[str(pathlib.Path(__file__).resolve().parent),str(pathlib.Path(sys.argv[1]).resolve())]
from compare_relevant_message import html_profile,compare_profiles
from rtf_representation import validated_repair_text_representation

def compare(a,b):return compare_profiles(html_profile(a,qualified_css=True),html_profile(b,qualified_css=True))['ComparedRelevantHtmlInformationEqual']
def check(name,actual,expected):
 if actual!=expected:raise AssertionError(name)
style='<style>blockquote footer:before,blockquote small:before,blockquote .small:before{content:"x"}q:before,q:after{content:"";content:none}</style>'
check('inert-qualified-selector',compare(style+'<p>A</p>',style+'<div>A</div>'),True)
check('changed-generated-css',compare(style+'A',style.replace('"x"','"y"')+'A'),False)
check('activated-quote',compare(style+'A',style+'<q>A</q>'),False)
css='<style>span[class=show1] a:after{content:" - satisfeito"}</style>'
a='<span class="show1"><a href="https://example.invalid">Voto</a></span><span><a href="https://example.invalid">Voto</a></span>'
b='<span><a href="https://example.invalid">Voto</a></span><span class="show1"><a href="https://example.invalid">Voto</a></span>'
check('same-generated-label',compare(css+a,css+a),True)
check('generated-label-position',compare(css+a,css+b),False)
check('changed-generated-link',compare(css+a,css+a.replace('example.invalid','other.invalid')),False)
check('dynamic-generated-value',compare(css.replace('" - satisfeito"','attr(title)')+a,css.replace('" - satisfeito"','attr(title)')+a),False)
check('unknown-selector',compare(css.replace('span[class=show1] a','#unknown a')+a,css.replace('span[class=show1] a','#unknown a')+a),False)
check('foreign-executable-script',compare(style+'<script>run()</script>A',style+'<script>run()</script>A'),False)
check('legacy-css-unchanged',html_profile(style+'A')['Unsupported']!=[],True)
check('canonical-unicode-composition',compare('<p>du\u0301vidas</p>','<p>dúvidas</p>'),True)
check('corrupted-combining-accent-blocked',compare('<p>du\u0301vidas</p>','<p>duìvidas</p>'),False)
check('different-accent-blocked',compare('<p>avó</p>','<p>avô</p>'),False)
for kind,label in [('a','a.'),('A','A.'),('i','i.'),('I','I.'),('1','1.')]:
 plain=f'<ol type="{kind}"><li class="MsoListParagraph" style="mso-list:l1">Text</li></ol>'
 fallback=plain.replace('>Text</li>','>'+label+'\t\nText</li>')
 # The real fallback is a standalone LI text node before its content span.
 plain=plain.replace('>Text</li>','><span>Text</span></li>')
 fallback=fallback.replace('Text</li>','<span>Text</span></li>')
 check('qualified-counter-'+kind,compare(plain,fallback),True)
 check('wrong-counter-'+kind,compare(plain,fallback.replace(label+'\t','wrong.\t')),False)
 check('literal-label-'+kind,compare(plain,fallback.replace('MsoListParagraph','ordinary')),False)
raw=br'{\rtf1\ansi\ansicpg1252\fromtext{\fonttbl{\f0\fnil Arial;}}\f0\uc1 A\u160?B\par}'
check('unicode-variant-qualified',validated_repair_text_representation(raw,'A B\r\n'.encode('utf-16-le')) is not None,True)
check('unicode-original-qualified',validated_repair_text_representation(raw,'A\u00a0B\r\n\0'.encode('utf-16-le')) is not None,True)
check('unicode-text-change-blocked',validated_repair_text_representation(raw,'A C\r\n'.encode('utf-16-le')) is None,True)
check('interior-null-preserved',validated_repair_text_representation(raw,'A\0 B\r\n'.encode('utf-16-le')) is None,True)
check('rich-rtf-blocked',validated_repair_text_representation(raw.replace(b'A\\u160',b'\\field A\\u160'),'A B'.encode('utf-16-le')) is None,True)
print('Reparo: cenários de CSS, posição dos rótulos, contadores, scripts e Unicode aprovados.')

from signed_mime_profile import signed_mime_profile
from io import BytesIO
class SignedFixture:
 def __init__(self,payload):
  base='__attach_version1.0_#00000000/'
  self.data={'__substg1.0_001A001F':'IPM.Note.SMIME.MultipartSigned'.encode('utf-16-le'),base+'__substg1.0_3707001F':'smime.p7m'.encode('utf-16-le'),base+'__substg1.0_370E001F':'multipart/signed'.encode('utf-16-le'),base+'__substg1.0_37010102':payload}
 def exists(self,path):return path in self.data
 def openstream(self,path):return BytesIO(self.data[path])
 def get_size(self,path):return len(self.data[path])
 def listdir(self,**kwargs):return [['__attach_version1.0_#00000000']]
mime=b'MIME-Version: 1.0\r\nContent-Type: multipart/signed; boundary="boundary"; protocol="application/pkcs7-signature"\r\n\r\n--boundary\r\nContent-Type: text/plain\r\n\r\nOriginal body\r\n--boundary\r\nContent-Type: application/pkcs7-signature\r\n\r\nSignature bytes\r\n--boundary--\r\n'
first=SignedFixture(mime);second=SignedFixture(mime)
second.data['__substg1.0_1000001F']='Derived rendering'.encode('utf-16-le')
check('signed-render-cache-independent',signed_mime_profile(first)==signed_mime_profile(second),True)
check('signed-body-change-blocked',signed_mime_profile(first)==signed_mime_profile(SignedFixture(mime.replace(b'Original body',b'Changed body'))),False)
check('signed-signature-change-blocked',signed_mime_profile(first)==signed_mime_profile(SignedFixture(mime.replace(b'Signature bytes',b'Changed signature'))),False)
try:
 signed_mime_profile(SignedFixture(mime.replace(b'application/pkcs7-signature',b'application/unknown')))
 raise AssertionError('invalid-signed-protocol')
except ValueError:pass
print('S/MIME: conteúdo e assinatura preservados byte a byte; alterações e protocolo desconhecido bloqueados.')
