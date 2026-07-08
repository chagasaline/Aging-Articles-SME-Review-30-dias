import os
import io
import uuid
import logging
from flask import Flask, request, jsonify, render_template, session, redirect, url_for
from dotenv import load_dotenv
import openpyxl
import msal
import requests

load_dotenv()

app = Flask(__name__)
app.secret_key = os.urandom(24)
logging.basicConfig(level=logging.INFO)
log = logging.getLogger("kb-mailer-oauth")

# ---------------------------------------------------------------------------
# Configurações do .env
# ---------------------------------------------------------------------------
CLIENT_ID = os.getenv("CLIENT_ID")
CLIENT_SECRET = os.getenv("CLIENT_SECRET")
TENANT_ID = os.getenv("TENANT_ID", "common")
REDIRECT_URI = os.getenv("REDIRECT_URI", "http://localhost:5000/callback")
MAIL_FROM = os.getenv("MAIL_FROM", "")
MAIL_FROM_NAME = os.getenv("MAIL_FROM_NAME", "IT Knowledge Management")
CC_LIST = os.getenv("CC_LIST", "")
SUBJECT = os.getenv("SUBJECT", "KB Mgmt. - Aging Articles (SME Review more than 30 days)")
ALLOW_SEND_AS_SHARED = os.getenv("ALLOW_SEND_AS_SHARED", "true").lower() == "true"

FONT_CSS = "font-family:Calibri, 'Times New Roman';font-size: 15px;"
BATCHES = {}

# ---------------------------------------------------------------------------
# Autenticação Microsoft
# ---------------------------------------------------------------------------
AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
SCOPES = ["Mail.Send"]

def get_msal_app():
    return msal.ConfidentialClientApplication(
        CLIENT_ID, authority=AUTHORITY, client_credential=CLIENT_SECRET
    )

def get_token_from_cache():
    return session.get("token")

def build_auth_url():
    return get_msal_app().get_authorization_request_url(SCOPES, redirect_uri=REDIRECT_URI)

@app.route("/login")
def login():
    return redirect(build_auth_url())

@app.route("/callback")
def callback():
    code = request.args.get("code")
    if not code:
        return "Erro: código não recebido", 400
    result = get_msal_app().acquire_token_by_authorization_code(
        code, scopes=SCOPES, redirect_uri=REDIRECT_URI
    )
    if "access_token" not in result:
        return f"Erro ao obter token: {result.get('error_description')}", 400
    session["token"] = result
    return redirect(url_for("index"))

@app.route("/logout")
def logout():
    session.clear()
    return redirect(url_for("index"))

# ---------------------------------------------------------------------------
# Funções da planilha (mesmas de antes)
# ---------------------------------------------------------------------------
def parse_workbook(file_stream):
    wb = openpyxl.load_workbook(file_stream, data_only=True)
    ws = wb.worksheets[0]
    groups = {}
    order = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        doc_id, doc_type, title, owner_name, owner_email = (row + (None,) * 5)[:5]
        if not owner_email:
            continue
        email = str(owner_email).strip().lower()
        owner_name = str(owner_name).strip() if owner_name else ""
        if email not in groups:
            groups[email] = {"owner_name": owner_name, "articles": []}
            order.append(email)
        groups[email]["articles"].append({
            "doc_id": str(doc_id).strip() if doc_id else "",
            "type": str(doc_type).strip() if doc_type else "",
            "title": str(title).strip() if title else "",
            "owner_name": owner_name,
        })
    return [{"email": e, **groups[e]} for e in order]

def build_html_table(articles):
    rows = "".join(
        f"<tr><td style='background-color:#50ba33;color:#fff;'>{a['doc_id']}</td>"
        f"<td style='background-color:#f8fbff;color:#000;'>{a['type']}</td>"
        f"<td style='background-color:#f8fbff;color:#000;'>{a['title']}</td>"
        f"<td style='background-color:#f8fbff;color:#000;'>{a['owner_name']}</td></tr>"
        for a in articles
    )
    return f"""<table border="1" cellspacing="0" cellpadding="4" style="{FONT_CSS}">
        <tr style='background-color:#007e7a;color:#fff;'>
          <th>DocId</th><th>KnowledgeSourceName</th><th>ArticleTitle</th><th>ArticleOwnerName</th>
        </tr>
        {rows}
        </table>"""

def build_text_table(articles):
    lines = ["DocId | Type | Article Title | Owner"]
    lines += [f"{a['doc_id']} | {a['type']} | {a['title']} | {a['owner_name']}" for a in articles]
    return "\n".join(lines)

def build_message(owner_name, articles):
    pt = (f"Olá {owner_name} existe pelo menos um artigo de conhecimento sob sua revisão no Vale Support Center há mais de 30 dias.\n"
          f"Você tem alguma dúvida do que fazer? Por favor entre em contato: IT.Knowledge.management@vale.com.")
    en = (f"Hi {owner_name} there is a least one knowledge article under your review in Vale Support Center for more than 30 days.\n"
          f"Do you have questions? Please, contact: IT.Knowledge.management@vale.com.")
    text_body = f"{pt}\n\nPORTUGUESE VERSION\n\n{build_text_table(articles)}\n\n{en}\n\nENGLISH VERSION"
    html_body = f"<div style='{FONT_CSS}'>{pt}<br><br><b>PORTUGUESE VERSION</b><br><br>{build_html_table(articles)}<br><br>{en}<br><br><b>ENGLISH VERSION</b></div>"
    return text_body, html_body

# ---------------------------------------------------------------------------
# Envio via Microsoft Graph
# ---------------------------------------------------------------------------
def send_via_graph(to_addr, cc_addrs, subject, html_body, token, send_as_shared=False):
    headers = {"Authorization": f"Bearer {token['access_token']}", "Content-Type": "application/json"}
    msg = {
        "message": {
            "subject": subject,
            "body": {"contentType": "HTML", "content": html_body},
            "toRecipients": [{"emailAddress": {"address": to_addr}}],
            "ccRecipients": [{"emailAddress": {"address": addr}} for addr in cc_addrs if addr],
        },
        "saveToSentItems": "true"
    }
    endpoint = f"https://graph.microsoft.com/v1.0/users/{MAIL_FROM}/sendMail" if (send_as_shared and MAIL_FROM) else "https://graph.microsoft.com/v1.0/me/sendMail"
    resp = requests.post(endpoint, headers=headers, json=msg)
    if resp.status_code == 202:
        return True, "Enviado com sucesso via Graph API"
    else:
        return False, f"Erro: {resp.status_code} - {resp.text}"

# ---------------------------------------------------------------------------
# Rotas principais
# ---------------------------------------------------------------------------
@app.route("/")
def index():
    token = get_token_from_cache()
    user_info = token.get("id_token_claims") if token else None
    return render_template("index.html",
        user_info=user_info,
        mail_from=MAIL_FROM,
        cc_list=CC_LIST,
        subject=SUBJECT,
        allow_send_as=ALLOW_SEND_AS_SHARED,
    )

@app.route("/upload", methods=["POST"])
def upload():
    file = request.files.get("file")
    if not file:
        return jsonify({"error": "Nenhum arquivo enviado"}), 400
    try:
        groups = parse_workbook(io.BytesIO(file.read()))
    except Exception as e:
        return jsonify({"error": f"Erro ao ler a planilha: {e}"}), 400
    if not groups:
        return jsonify({"error": "Nenhuma linha com e-mail do dono foi encontrada"}), 400
    batch_id = str(uuid.uuid4())
    BATCHES[batch_id] = groups
    preview = [
        {
            "email": g["email"],
            "owner_name": g["owner_name"],
            "article_count": len(g["articles"]),
            "articles": [a["title"] for a in g["articles"]],
        }
        for g in groups
    ]
    return jsonify({"batch_id": batch_id, "groups": preview})

@app.route("/send", methods=["POST"])
def send():
    data = request.get_json(force=True)
    batch_id = data.get("batch_id")
    selected_emails = set(data.get("emails", []))
    dry_run = bool(data.get("dry_run", False))
    test_email = (data.get("test_email") or "").strip()
    send_as_shared = bool(data.get("send_as_shared", False)) and ALLOW_SEND_AS_SHARED

    token = get_token_from_cache()
    if not token:
        return jsonify({"error": "Usuário não autenticado. Faça login novamente."}), 401

    groups = BATCHES.get(batch_id)
    if not groups:
        return jsonify({"error": "Lote não encontrado. Faça upload novamente."}), 400

    cc_addrs = [c.strip() for c in CC_LIST.split(";") if c.strip()]
    results = []
    for g in groups:
        if selected_emails and g["email"] not in selected_emails:
            continue
        text_body, html_body = build_message(g["owner_name"], g["articles"])
        to_addr = test_email if test_email else g["email"]

        if dry_run:
            results.append({
                "email": g["email"],
                "sent_to": to_addr,
                "owner_name": g["owner_name"],
                "success": True,
                "info": "dry-run (nenhum e-mail enviado)"
            })
        else:
            ok, info = send_via_graph(to_addr, cc_addrs, SUBJECT, html_body, token, send_as_shared)
            results.append({
                "email": g["email"],
                "sent_to": to_addr,
                "owner_name": g["owner_name"],
                "success": ok,
                "info": info
            })
    return jsonify({"results": results})

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=True)