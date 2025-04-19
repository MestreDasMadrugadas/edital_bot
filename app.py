import os
import pickle
from flask import Flask, request, jsonify
from openpyxl import Workbook, load_workbook
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload

# Definindo a configuração básica do Flask
app = Flask(__name__)

# Caminho para o arquivo modelo
ARQUIVO_EXCEL = "PMCE_Edital_Verticalizado_MODELO.xlsx"

# Definindo escopos do Google Drive
SCOPES = ['https://www.googleapis.com/auth/drive.file']

# Função de autenticação e criação do serviço do Google Drive
def autenticar_google_drive():
    creds = None
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)

        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)

    service = build('drive', 'v3', credentials=creds)
    return service

# Função para fazer o upload do arquivo no Google Drive
def upload_no_drive(caminho_arquivo, nome_arquivo):
    service = autenticar_google_drive()

    media = MediaFileUpload(caminho_arquivo, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    file_metadata = {'name': nome_arquivo, 'mimeType': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'}
    arquivo = service.files().create(body=file_metadata, media_body=media, fields='id').execute()

    file_id = arquivo.get('id')
    link_arquivo = f'https://drive.google.com/file/d/{file_id}/view?usp=sharing'

    return link_arquivo

# Função para salvar os dados na planilha
def salvar_dados_em_excel(nome, email, produto, status, cpf):
    nome_arquivo = f"resultado-{nome.replace(' ', '_')}-{cpf}.xlsx"
    wb = Workbook()
    ws = wb.active
    ws.title = "Compras"
    ws.append(["Nome", "Email", "Produto", "Status", "CPF"])
    ws.append([nome, email, produto, status, cpf])

    caminho_arquivo = os.path.join("temp", nome_arquivo)
    
    if not os.path.exists('temp'):
        os.makedirs('temp')

    wb.save(caminho_arquivo)

    # Fazendo upload para o Google Drive
    link = upload_no_drive(caminho_arquivo, nome_arquivo)
    
    # Apaga o arquivo temporário após o upload
    os.remove(caminho_arquivo)

    return link

# Endpoint principal
@app.route("/", methods=["GET"])
def home():
    return "✅ API do Edital Bot está no ar!"

# Endpoint webhook para receber dados da Kiwify
@app.route("/webhook", methods=["POST"])
def webhook():
    try:
        data = request.get_json()
        print("📦 Webhook recebido:", data)

        nome = data["purchase"]["buyer"]["name"]
        email = data["purchase"]["buyer"]["email"]
        cpf = data["purchase"]["buyer"].get("document", "CPF não informado")
        cpf = cpf.replace(".", "").replace("-", "")
        produto = data["purchase"]["product"]["name"]
        status = data["purchase"]["status"]

        # Salva os dados no Excel e retorna o link do arquivo no Google Drive
        link_arquivo = salvar_dados_em_excel(nome, email, produto, status, cpf)

        return jsonify({
            "status": "ok",
            "mensagem": "Dados recebidos e salvos com sucesso!",
            "link": link_arquivo
        })

    except Exception as e:
        print("❌ Erro ao processar webhook:", str(e))
        return jsonify({"status": "error", "message": str(e)}), 400

# Endpoint para disponibilizar o arquivo na área de membros
@app.route("/download", methods=["GET"])
def download_arquivo():
    cpf = request.args.get('cpf')
    
    if not cpf:
        return "CPF não informado", 400

    # A função que chama o link de download vai ser a mesma de salvar_dados_em_excel
    # Aqui apenas passamos os dados necessários para gerar o arquivo
    nome = "Nome do comprador"  # Você deve buscar esses dados em algum lugar, possivelmente na base de dados
    email = "email@comprador.com"
    produto = "Produto da compra"
    status = "Aprovado"

    link_arquivo = salvar_dados_em_excel(nome, email, produto, status, cpf)

    return jsonify({
        "status": "ok",
        "mensagem": "Arquivo gerado com sucesso!",
        "link": link_arquivo
    })

# Rodando a aplicação
if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
