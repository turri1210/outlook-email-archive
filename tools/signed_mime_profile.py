"""Preserve the complete native multipart/signed payload, without trusting a signer."""
import hashlib
from email.parser import BytesParser
from email import policy

def signed_mime_profile(message):
    tag='__substg1.0_001A001F'
    if not message.exists(tag) or message.openstream(tag).read().decode('utf-16-le',errors='strict').rstrip('\0')!='IPM.Note.SMIME.MultipartSigned':return None
    payloads=[]
    for entry in message.listdir(streams=False,storages=True):
        if len(entry)!=1 or not entry[0].startswith('__attach_version1.0_#'):continue
        base=entry[0]
        required={name:base+'/__substg1.0_'+name for name in ['3707001F','370E001F','37010102']}
        if not all(message.exists(path) for path in required.values()):continue
        filename=message.openstream(required['3707001F']).read().decode('utf-16-le',errors='strict').rstrip('\0')
        mime=message.openstream(required['370E001F']).read().decode('utf-16-le',errors='strict').rstrip('\0')
        if filename!='smime.p7m' or mime!='multipart/signed':continue
        if message.get_size(required['37010102'])>32*1024*1024:raise ValueError('signed-mime-size-limit')
        payload=message.openstream(required['37010102']).read();parsed=BytesParser(policy=policy.default).parsebytes(payload)
        protocol=parsed.get_param('protocol')
        if parsed.get_content_type()!='multipart/signed' or protocol not in ['application/pkcs7-signature','application/x-pkcs7-signature'] or not parsed.is_multipart() or parsed.defects:raise ValueError('signed-mime-envelope-invalid')
        parts=list(parsed.iter_parts())
        if len(parts)!=2 or parts[1].get_content_type() not in ['application/pkcs7-signature','application/x-pkcs7-signature'] or any(part.defects for part in parts):raise ValueError('signed-mime-parts-invalid')
        payloads.append(hashlib.sha256(payload).hexdigest())
    if len(payloads)!=1:raise ValueError('signed-mime-payload-not-unique')
    return {'ProfileVersion':'relevant-message-v1','BodyMode':'signed-mime-exact-v1','Body-SignedMime':payloads[0]}
