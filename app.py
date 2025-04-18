from flask import Flask, request
from utils import gerar_planilha_personalizada, enviar_email_com_anexo

app = Flask(__name__)

@app.route('/webhook', methods=['POST'])
def webhook():
    data = request.json
    nome = data['purchase']['buyer']['name']
    cpf = data['purchase']['buyer']['document'].replace('.', '').replace('-', '')
    email = data['purchase']['buyer']['email']

    planilha = gerar_planilha_personalizada(nome, cpf)
    enviar_email_com_anexo(nome, email, planilha)
    return "Planilha enviada com sucesso!", 200

if __name__ == '__main__':
    app.run(debug=True)