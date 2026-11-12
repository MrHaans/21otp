#!/usr/bin/env python3
"""21OTP Backend — IMAP XOAUTH2 for Render"""
import json, os, re, email, imaplib, base64, socket, urllib.request, urllib.parse, urllib.error, asyncio
from email.header import decode_header
from aiohttp import web

socket.setdefaulttimeout(120)

HTML = os.path.join(os.path.dirname(os.path.abspath(__file__)), '21otp_v2.html')
PORT = int(os.environ.get('PORT', 8900))
TOKEN_URL = "https://login.microsoftonline.com/common/oauth2/v2.0/token"
DEFAULT_CLIENT = "9e5f94bc-e8a4-4e73-b8be-63364c29d753"
SCOPE = "https://outlook.office.com/IMAP.AccessAsUser.All offline_access"
routes = web.RouteTableDef()

def find_otp(text):
    ps=[r'\b(\d{4,8})\b',r'\b([A-Z0-9]{6,8})\b',r'code[:\s]*(\d{4,8})',r'otp[:\s]*(\d{4,8})',r'verification[:\s]*(\d{4,8})',r'(\d{3}[- ]?\d{3}[- ]?\d{3})']
    fs=set()
    for p in ps:
        for m in re.finditer(p,text,re.IGNORECASE):
            c=(m.group(1)or m.group(0)).replace('-','').replace(' ','')
            if 4<=len(c)<=8:fs.add(c)
    return list(fs)

def decode_mime(s):
    if not s: return ''
    parts=decode_header(s)
    return ''.join(p[0].decode(p[1]or'utf-8',errors='replace') if isinstance(p[0],bytes) else str(p[0]) for p in parts)

def get_token(rt, cid=None):
    cid=cid or DEFAULT_CLIENT
    d=urllib.parse.urlencode({'client_id':cid,'refresh_token':rt,'grant_type':'refresh_token','scope':SCOPE}).encode()
    r=urllib.request.Request(TOKEN_URL,data=d,headers={'Content-Type':'application/x-www-form-urlencoded'})
    t=json.loads(urllib.request.urlopen(r,timeout=30).read())
    return t.get('access_token'),t.get('refresh_token')

def fetch_imap(email_addr, access_token, limit=30):
    auth_str = f"user={email_addr}\x01auth=Bearer {access_token}\x01\x01"
    auth_b64 = base64.b64encode(auth_str.encode()).decode()
    conn = imaplib.IMAP4_SSL('outlook.office365.com', 993)
    conn.authenticate('XOAUTH2', lambda x: auth_b64)
    conn.select('INBOX')
    _, msgs = conn.search(None, 'ALL')
    msg_ids = msgs[0].split()[-limit:]
    emails = []
    for mid in reversed(msg_ids):
        _, data = conn.fetch(mid, '(RFC822)')
        raw = data[0][1]
        msg = email.message_from_bytes(raw)
        subject = decode_mime(msg.get('Subject',''))
        from_addr = decode_mime(msg.get('From',''))
        date = msg.get('Date','')
        body = ''
        if msg.is_multipart():
            for part in msg.walk():
                if part.get_content_type() == 'text/plain':
                    try: body = part.get_payload(decode=True).decode('utf-8',errors='replace')
                    except: body = str(part.get_payload())
                    break
        else:
            try: body = msg.get_payload(decode=True).decode('utf-8',errors='replace')
            except: body = str(msg.get_payload())
        emails.append({'id':mid.decode(),'from':from_addr,'subject':subject,'preview':body[:300],'date':date,'otps':find_otp(body)})
    conn.logout()
    return emails

@routes.get('/api/fetch')
async def api_fetch(request):
    em=request.query.get('email',''); tk=request.query.get('token',''); ci=request.query.get('client',DEFAULT_CLIENT); lm=int(request.query.get('limit','30'))
    if not em or not tk: return web.json_response({'error':'Missing params'},status=400)
    try:
        at,nrt = await asyncio.to_thread(get_token, tk, ci)
        emails = await asyncio.to_thread(fetch_imap, em, at, lm)
        resp = {'emails':emails,'total':len(emails)}
        if nrt and nrt!=tk: resp['newToken']=nrt
        return web.json_response(resp)
    except Exception as e: return web.json_response({'error':str(e)},status=500)

@routes.get('/api/health')
async def health(r): return web.json_response({'ok':1})

@routes.get('/')
async def index(r):
    if os.path.exists(HTML): return web.FileResponse(HTML)
    return web.Response(text='21OTP — missing HTML',status=500)

app=web.Application(); app.add_routes(routes)

if __name__=='__main__':
    print(f'21OTP → http://0.0.0.0:{PORT}')
    web.run_app(app,host='0.0.0.0',port=PORT)