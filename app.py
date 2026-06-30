from flask import Flask, jsonify, request, redirect
from flask_cors import CORS
import requests
import os

app = Flask(__name__)
CORS(app)

CLIENT_ID = os.getenv('AZURE_CLIENT_ID')
CLIENT_SECRET = os.getenv('AZURE_CLIENT_SECRET')
TENANT_ID = os.getenv('AZURE_TENANT_ID')
REDIRECT_URI = os.getenv('REDIRECT_URI')
CLAUDE_API_KEY = os.getenv('CLAUDE_API_KEY')
SUPABASE_URL = os.getenv('SUPABASE_URL')
SUPABASE_KEY = os.getenv('SUPABASE_KEY')
CRON_SECRET = os.getenv('CRON_SECRET', 'betty2026')
FRONTEND_URL = 'https://superlative-macaron-311744.netlify.app'

BETTY_PROMPT = "Tu es BETTY, assistante comptable francaise. Tu aides avec les emails, factures et paiements. Sois concise, professionnelle et chaleureuse. Reponds toujours en francais."


def supabase_headers():
    return {
        'apikey': SUPABASE_KEY,
        'Authorization': 'Bearer ' + SUPABASE_KEY,
        'Content-Type': 'application/json'
    }


def save_refresh_token(refresh_token):
    requests.delete(
        SUPABASE_URL + '/rest/v1/outlook_tokens?user_id=eq.betty_principal',
        headers=supabase_headers()
    )
    requests.post(
        SUPABASE_URL + '/rest/v1/outlook_tokens',
        headers=supabase_headers(),
        json={'user_id': 'betty_principal', 'refresh_token': refresh_token}
    )


def get_refresh_token():
    res = requests.get(
        SUPABASE_URL + '/rest/v1/outlook_tokens?user_id=eq.betty_principal&order=updated_at.desc&limit=1',
        headers=supabase_headers()
    )
    data = res.json()
    if isinstance(data, list) and len(data) > 0:
        return data[0]['refresh_token']
    return None


def save_resumen(resumen, count):
    requests.post(
        SUPABASE_URL + '/rest/v1/resumenes',
        headers=supabase_headers(),
        json={'user_id': 'betty_principal', 'resumen': resumen, 'email_count': count}
    )


def get_access_token_from_refresh():
    refresh_token = get_refresh_token()
    if not refresh_token:
        return None

    token_url = 'https://login.microsoftonline.com/common/oauth2/v2.0/token'
    data = {
        'client_id': CLIENT_ID,
        'client_secret': CLIENT_SECRET,
        'refresh_token': refresh_token,
        'grant_type': 'refresh_token',
        'scope': 'https://graph.microsoft.com/Mail.Read https://graph.microsoft.com/Mail.Send https://graph.microsoft.com/User.Read offline_access'
    }
    response = requests.post(token_url, data=data)
    tokens = response.json()

    if 'refresh_token' in tokens:
        save_refresh_token(tokens['refresh_token'])

    return tokens.get('access_token')


def read_emails(access_token):
    headers = {'Authorization': 'Bearer ' + access_token}
    url = 'https://graph.microsoft.com/v1.0/me/mailFolders/inbox/messages?$top=20&$select=subject,from,receivedDateTime,bodyPreview'
    response = requests.get(url, headers=headers)
    if response.status_code != 200:
        return None
    return response.json().get('value', [])


def build_email_context(emails):
    lines = []
    for i, e in enumerate(emails):
        frm = 'Inconnu'
        if e.get('from') and e['from'].get('emailAddress'):
            ea = e['from']['emailAddress']
            frm = ea.get('name') or ea.get('address') or 'Inconnu'
        subject = e.get('subject') or '(sans objet)'
        preview = (e.get('bodyPreview') or '')[:150]
        lines.append(str(i + 1) + '. De: ' + frm + ' | Objet: ' + subject + ' | Apercu: ' + preview)
    return '\n'.join(lines)


def ask_claude(message, email_context=''):
    full_message = message
    if email_context:
        full_message = 'Voici mes emails recents:\n' + email_context + '\n\nQuestion: ' + message

    headers = {
        'Content-Type': 'application/json',
        'x-api-key': CLAUDE_API_KEY,
        'anthropic-version': '2023-06-01'
    }
    payload = {
        'model': 'claude-haiku-4-5-20251001',
        'max_tokens': 600,
        'system': BETTY_PROMPT,
        'messages': [{'role': 'user', 'content': full_message}]
    }
    response = requests.post('https://api.anthropic.com/v1/messages', headers=headers, json=payload)
    result = response.json()
    if 'content' in result:
        return result['content'][0]['text']
    return None


@app.route('/api/outlook/callback', methods=['GET'])
def outlook_callback():
    code = request.args.get('code')
    if not code:
        return 'Error: no code received', 400

    token_url = 'https://login.microsoftonline.com/common/oauth2/v2.0/token'
    data = {
        'client_id': CLIENT_ID,
        'client_secret': CLIENT_SECRET,
        'code': code,
        'redirect_uri': REDIRECT_URI,
        'grant_type': 'authorization_code',
        'scope': 'https://graph.microsoft.com/Mail.Read https://graph.microsoft.com/Mail.Send https://graph.microsoft.com/User.Read offline_access'
    }

    response = requests.post(token_url, data=data)
    tokens = response.json()

    if 'access_token' not in tokens:
        print('TOKEN ERROR:', tokens)
        return jsonify(tokens), 400

    if 'refresh_token' in tokens:
        save_refresh_token(tokens['refresh_token'])
        print('REFRESH TOKEN GUARDADO')

    access_token = tokens['access_token']
    return redirect(FRONTEND_URL + '/?token=' + access_token)


@app.route('/api/outlook/emails', methods=['POST'])
def get_emails():
    token = request.json.get('access_token')
    if not token:
        return jsonify({'error': {'message': 'No token provided'}}), 400

    headers = {'Authorization': 'Bearer ' + token}
    url = 'https://graph.microsoft.com/v1.0/me/mailFolders/inbox/messages?$top=20&$select=subject,from,receivedDateTime,bodyPreview'

    try:
        response = requests.get(url, headers=headers)
        if response.status_code != 200:
            print('GRAPH ERROR:', response.status_code)
            print('GRAPH BODY:', response.text)
            return jsonify({'error': {'message': 'Erreur Graph', 'status': response.status_code, 'detail': response.text}}), 200
        return jsonify(response.json())
    except Exception as e:
        print('EMAILS EXCEPTION:', str(e))
        return jsonify({'error': {'message': str(e)}}), 200


@app.route('/api/chat', methods=['POST'])
def chat():
    data = request.json
    user_message = data.get('message', '')
    email_context = data.get('email_context', '')
    reply = ask_claude(user_message, email_context)
    if reply:
        return jsonify({'reply': reply})
    return jsonify({'reply': 'Erreur IA'}), 400


@app.route('/api/resumen/actuel', methods=['GET'])
def resumen_actuel():
    try:
        res = requests.get(
            SUPABASE_URL + '/rest/v1/resumenes?user_id=eq.betty_principal&order=created_at.desc&limit=1',
            headers=supabase_headers()
        )
        data = res.json()
        if isinstance(data, list) and len(data) > 0:
            return jsonify({'resumen': data[0]['resumen'], 'count': data[0].get('email_count', 0), 'date': data[0].get('created_at', '')})
        return jsonify({'resumen': None})
    except Exception as e:
        return jsonify({'error': str(e)}), 200


@app.route('/api/resumen/generer', methods=['POST'])
def resumen_generer():
    secret = request.args.get('secret', '')
    if secret != CRON_SECRET:
        return jsonify({'error': 'Non autorise'}), 401

    access_token = get_access_token_from_refresh()
    if not access_token:
        return jsonify({'error': 'Pas de connexion Outlook'}), 200

    emails = read_emails(access_token)
    if emails is None:
        return jsonify({'error': 'Impossible de lire les emails'}), 200

    email_context = build_email_context(emails)
    prompt = 'Resume mes emails en UNE seule phrase tres courte (max 15 mots), style notification. Exemple: "3 factures en attente, 2 demandes de paiement, 5 publicites". Juste les categories importantes.'
    resumen = ask_claude(prompt, email_context)
    if not resumen:
        resumen = str(len(emails)) + ' emails dans la boite'

    resumen = resumen.replace('\n', ' ').strip()
    save_resumen(resumen, len(emails))
    return jsonify({'resumen': resumen, 'count': len(emails)})


@app.route('/', methods=['GET'])
def health():
    return jsonify({'status': 'ok'})


if __name__ == '__main__':
    app.run(host='0.0.0.0', port=int(os.getenv('PORT', 5000)))
