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

BETTY_PROMPT = "Tu es BETTY, assistante comptable française. Tu aides avec les emails, factures et paiements. Sois concise, professionnelle et chaleureuse. Réponds toujours en français."

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
        'scope': 'https://graph.microsoft.com/Mail.Read https://graph.microsoft.com/Mail.Send offline_access'
    }

    response = requests.post(token_url, data=data)
    tokens = response.json()

    if 'access_token' not in tokens:
        print('TOKEN ERROR:', tokens)
        return jsonify(tokens), 400

    access_token = tokens['access_token']
    return redirect(f'{FRONTEND_URL}/?token={access_token}')

@app.route('/api/outlook/emails', methods=['POST'])
def get_emails():
    token = request.json.get('access_token')
    headers = {'Authorization': f'Bearer {token}'}
    response = requests.get(
        'https://graph.microsoft.com/v1.0/me/mailFolders/inbox/messages?$top=20&$select=subject,from,receivedDateTime,bodyPreview',
        headers=headers
    )
    return jsonify(response.json())

@app.route('/api/chat', methods=['POST'])
def chat():
    data = request.json
    user_message = data.get('message', '')
    email_context = data.get('email_context', '')

    full_message = user_message
    if email_context:
        full_message = f"Voici mes emails récents:\n{email_context}\n\nQuestion: {user_message}"

    response =
