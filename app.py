from flask import Flask, jsonify, request
import requests
import os

app = Flask(__name__)

CLIENT_ID = os.getenv('AZURE_CLIENT_ID')
CLIENT_SECRET = os.getenv('AZURE_CLIENT_SECRET')
TENANT_ID = os.getenv('AZURE_TENANT_ID')

@app.route('/api/outlook/auth', methods=['POST'])
def get_outlook_token():
    code = request.json.get('code')
    token_url = f'https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token'
    
    data = {
        'client_id': CLIENT_ID,
        'client_secret': CLIENT_SECRET,
        'code': code,
        'redirect_uri': os.getenv('REDIRECT_URI'),
        'grant_type': 'authorization_code',
        'scope': 'Mail.Read Mail.Send offline_access'
    }
    
    response = requests.post(token_url, data=data)
    return jsonify(response.json())

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
