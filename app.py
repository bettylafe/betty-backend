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
FRONTEND_URL = 'https://superlative-macaron-311744.netlify.app'

BETTY_PROMPT = "Tu es BETTY, assistante comptable francaise. Tu aides avec les emails, factures et paiements. Sois concise, professionnelle et chaleureuse. Reponds toujours en francais."


@app.route('/api/outlook/callback', methods=['GET'])
def outlook_callback():
    code = request.args.get('code')
    if not code:
        return 'Error: no code received', 400

    token_url = 'https://login.microsoftonline.com/' + TENANT_ID + '/oauth2/v2.0/token'
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

    access_token = tokens['access_token']
    print('TOKEN OBTENU - scope:', tokens.get('scope', 'AUCUN'))
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
        # Si Microsoft devuelve error (token expirado, etc), lo pasamos como JSON
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

    full_message = user_message
    if email_context:
        full_message = 'Voici mes emails recents:\n' + email_context + '\n\nQuestion: ' + user_message

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
        return jsonify({'reply': result['content'][0]['text']})
    else:
        print('CLAUDE ERROR:', result)
        return jsonify({'reply': 'Erreur IA', 'debug': result}), 400


@app.route('/', methods=['GET'])
def health():
    return jsonify({'status': 'ok'})


if __name__ == '__main__':
    app.run(host='0.0.0.0', port=int(os.getenv('PORT', 5000)))


