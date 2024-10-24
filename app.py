from flask import Flask, render_template, request, redirect, url_for, session
import sqlite3
import requests
import os
import random
import string
from requests_oauthlib import OAuth2Session
from dotenv import load_dotenv

app = Flask(__name__)
app.secret_key = 'your_secret_key'  # Replace with a secure key
load_dotenv()

# Configuration
DATABASE = 'pedidos_vendas.db'
AUTHORIZATION_BASE_URL = 'https://jeanrabelo.github.io/J-lia/criptografia/index_animado_crip_1_automatico.html'
TOKEN_URL = 'https://developer.bling.com.br/api/bling/oauth/token'
CLIENT_ID = os.getenv('CLIENT_ID')
CLIENT_SECRET = os.getenv('CLIENT_SECRET')
REDIRECT_URI = 'http://localhost:5000/callback'  # Update this with your actual redirect URI

# Initialize the database
def init_db():
    conn = sqlite3.connect(DATABASE)
    cursor = conn.cursor()
    cursor.execute('''CREATE TABLE IF NOT EXISTS pedidos_vendas
                      (id INTEGER PRIMARY KEY AUTOINCREMENT, numero INTEGER, nome TEXT, total REAL, data TEXT, dataPrevista TEXT)''')
    conn.commit()
    conn.close()

# Generate a random state string
def generate_state(length=16):
    return ''.join(random.choices(string.ascii_uppercase + string.digits, k=length))

@app.route('/')
def index():
    conn = sqlite3.connect(DATABASE)
    cursor = conn.cursor()
    cursor.execute('SELECT * FROM pedidos_vendas')
    orders = cursor.fetchall()
    conn.close()
    return render_template('index.html', orders=orders)

@app.route('/fetch', methods=['POST'])
def fetch():
    # Step 1: Redirect user to Bling's authorization page to get the authorization code
    state = generate_state()
    bling = OAuth2Session(CLIENT_ID, redirect_uri=REDIRECT_URI, state=state)
    authorization_url, state = bling.authorization_url(AUTHORIZATION_BASE_URL)
    print(f'autorization_url:\n{authorization_url}\nstate:\n{state}')

    # Save the state in the session to validate later
    session['oauth_state'] = state

    return redirect(authorization_url)

@app.route('/callback', methods=['GET'])
def callback():
    # Step 2: Exchange the authorization code for an access token
    bling = OAuth2Session(CLIENT_ID, redirect_uri=REDIRECT_URI, state=session['oauth_state'])
    token = bling.fetch_token(TOKEN_URL, client_secret=CLIENT_SECRET, authorization_response=request.url)
    print(f'token:\n{token}')

    # Save the token in the session for later use
    session['oauth_token'] = token

    # Fetch sales orders after successful authentication
    fetch_sales_orders(token['access_token'])
    
    return redirect(url_for('index'))

def fetch_sales_orders(token):
    # Step 3: Fetch sales orders from Bling API and store in SQLite
    orders_url = 'https://developer.bling.com.br/api/bling/pedidos/vendas'
    params = {
        'pagina': 1,
        'limite': 100,
        'dataInicial': '2024-05-01',
        'dataFinal': '2025-01-15'
    }
    headers = {
        'accept': 'application/json',
        'authorization': f'Bearer {token}'
    }
    response = requests.get(orders_url, headers=headers, params=params)
    orders_data = response.json()['data']

    conn = sqlite3.connect(DATABASE)
    cursor = conn.cursor()
    for order in orders_data:
        cursor.execute('INSERT INTO pedidos_vendas (numero, nome, total, data, dataPrevista) VALUES (?, ?, ?, ?, ?)',
                       (order['numero'], order['contato']['nome'], order['total'], order['data'], order['dataPrevista']))
    conn.commit()
    conn.close()

if __name__ == '__main__':
    init_db()
    app.run(debug=True)

