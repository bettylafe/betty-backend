from flask import Flask, jsonify, request, redirect
import requests
import os

app = Flask(__name__)

CLIENT_ID = os.getenv('AZURE_CLIENT_ID')
CLIENT_SECRET = os.getenv('AZURE_CLIENT_SECRET')
TENANT_ID = os.getenv('AZURE_TENANT_ID')
REDIRECT_URI = os.getenv('REDIRECT_URI')
FRONTEND_URL = 'https://superlative-macaron-311744.netlify.app'

@app.route('/api/outlook/callback', methods=['GET'])
def outlook_callback():
    code = request.args.get('code')
    if not code:
        return 'Error: no code received', 400

    token_url = f'https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token'
    data = {
        'client_id': CLIENT_ID,
        'client_secret': CLIENT_SECRET,
        'code': code,
        'redirect_uri': REDIRECT_URI,
        'grant_type': 'authorization_code',
        'scope': 'Mail.Read Mail.Send offline_access'
    }

    response = requests.post(token_url, data=data)
    tokens = response.json()
    access_token = tokens.get('access_token', '')

    # Redirige al frontend con el token
    return redirect(f'{FRONTEND_URL}/?token={access_token}')

@app.route('/api/outlook/emails', methods=['POST'])
def get_emails():
    token = request.json.get('access_token')
    headers = {'Authorization': f'Bearer {token}'}
    response = requests.get(
        'https://graph.microsoft.com/v1.0/me/mailFolders/inbox/messages?$top=20',
        headers=headers
    )
    return jsonify(response.json())

@app.route('/', methods=['GET'])
def health():
    return jsonify({'status': 'ok'})

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=int(os.getenv('PORT', 5000)))
