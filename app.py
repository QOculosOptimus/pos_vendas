from flask import Flask, render_template, request, redirect, url_for
import sqlite3
import requests
import datetime

app = Flask(__name__)

# SQLite database setup
DATABASE = 'pedidos_vendas.db'

def init_db():
    conn = sqlite3.connect(DATABASE)
    cursor = conn.cursor()
    cursor.execute('''CREATE TABLE IF NOT EXISTS pedidos_vendas
                      (id INTEGER PRIMARY KEY AUTOINCREMENT, numero INTEGER, nome TEXT, total REAL, data TEXT, dataPrevista TEXT)''')
    conn.commit()
    conn.close()

# Get access token (refresh if needed)
def get_access_token():
    token_url = 'https://developer.bling.com.br/api/bling/oauth/token'
    payload = {
        'grant_type': 'authorization_code',
        'code': 'your_code',  # Replace with your code
        'client_id': 'your_client_id',  # Replace with your client_id
        'client_secret': 'your_client_secret',  # Replace with your client_secret
        'redirect_uri': 'https://developer.bling.com.br/oauth/redirect'
    }
    headers = {
        'accept': 'application/json',
        'content-type': 'application/x-www-form-urlencoded'
    }
    response = requests.post(token_url, data=payload, headers=headers)
    token_data = response.json()
    return token_data['access_token']

# Fetch sales orders from Bling API and store in SQLite
def fetch_sales_orders(token):
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
    token = get_access_token()
    fetch_sales_orders(token)
    return redirect(url_for('index'))

if __name__ == '__main__':
    init_db()
    app.run(debug=True)

