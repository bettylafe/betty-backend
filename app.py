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
    url = 'https://graph.microsoft.com/v1.0/me/mailFolders/inbox/messages?$top=20&$select=id,subject,from,receivedDateTime,bodyPreview'

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


@app.route('/api/outlook/email-complet', methods=['POST'])
def email_complet():
    token = request.json.get('access_token')
    email_id = request.json.get('email_id')
    if not token or not email_id:
        return jsonify({'error': 'Token ou id manquant'}), 400
    headers = {'Authorization': 'Bearer ' + token}
    url = 'https://graph.microsoft.com/v1.0/me/messages/' + email_id + '?$select=subject,from,toRecipients,body,receivedDateTime,hasAttachments'
    try:
        response = requests.get(url, headers=headers)
        if response.status_code != 200:
            return jsonify({'error': 'Erreur Graph', 'detail': response.text}), 200
        result = response.json()
        # Si l'email a des pieces jointes, recuperer leurs metadonnees (nom, type, taille) - gratuit, pas de contenu
        result['attachments_liste'] = []
        if result.get('hasAttachments'):
            att_url = 'https://graph.microsoft.com/v1.0/me/messages/' + email_id + '/attachments?$select=name,contentType,size'
            att_res = requests.get(att_url, headers=headers)
            if att_res.status_code == 200:
                att_data = att_res.json()
                for a in att_data.get('value', []):
                    result['attachments_liste'].append({
                        'name': a.get('name', 'fichier'),
                        'contentType': a.get('contentType', ''),
                        'size': a.get('size', 0)
                    })
        return jsonify(result)
    except Exception as e:
        return jsonify({'error': str(e)}), 200


@app.route('/api/redaction', methods=['POST'])
def redaction():
    data = request.json
    email_original = data.get('email_original', '')
    expediteur = data.get('expediteur', '')
    objet = data.get('objet', '')
    instructions = data.get('instructions', '')

    prompt = ("Tu dois rediger une reponse professionnelle a cet email recu.\n\n"
              "Expediteur: " + expediteur + "\n"
              "Objet: " + objet + "\n"
              "Contenu de l'email:\n" + email_original + "\n\n")
    if instructions:
        prompt += "Instructions specifiques pour la reponse: " + instructions + "\n\n"
    prompt += ("Redige UNIQUEMENT le corps de la reponse, en francais, ton professionnel et chaleureux. "
               "Pas d'objet, pas de signature automatique generique. Commence directement par la salutation. "
               "Sois concis et clair.")

    reply = ask_claude(prompt, '')
    if reply:
        return jsonify({'redaction': reply.strip()})
    return jsonify({'error': 'Erreur IA'}), 400


@app.route('/api/outlook/envoyer', methods=['POST'])
def envoyer_email():
    data = request.json
    token = data.get('access_token')
    destinataire = data.get('destinataire')
    objet = data.get('objet', '')
    corps = data.get('corps', '')

    if not token or not destinataire or not corps:
        return jsonify({'error': 'Donnees manquantes (token, destinataire ou corps)'}), 400

    # Nettoyage du corps: enlever espaces en trop et lignes vides multiples
    lignes = [ligne.rstrip() for ligne in corps.split('\n')]
    corps_propre = '\n'.join(lignes).strip()
    # Reduire 3+ sauts de ligne consecutifs a 2 maximum
    while '\n\n\n' in corps_propre:
        corps_propre = corps_propre.replace('\n\n\n', '\n\n')
    corps = corps_propre

    headers = {
        'Authorization': 'Bearer ' + token,
        'Content-Type': 'application/json'
    }
    message = {
        'message': {
            'subject': objet,
            'body': {'contentType': 'Text', 'content': corps},
            'toRecipients': [{'emailAddress': {'address': destinataire}}]
        },
        'saveToSentItems': True
    }
    try:
        response = requests.post(
            'https://graph.microsoft.com/v1.0/me/sendMail',
            headers=headers,
            json=message
        )
        print('SENDMAIL status:', response.status_code, 'body:', response.text[:300])
        if response.status_code == 202:
            return jsonify({'success': True, 'message': 'Email envoye a ' + destinataire})
        return jsonify({'success': False, 'error': 'Erreur Graph', 'status': response.status_code, 'detail': response.text}), 200
    except Exception as e:
        print('SENDMAIL EXCEPTION:', str(e))
        return jsonify({'success': False, 'error': str(e)}), 200


@app.route('/', methods=['GET'])
def health():
    return jsonify({'status': 'ok'})


if __name__ == '__main__':
    app.run(host='0.0.0.0', port=int(os.getenv('PORT', 5000)))
