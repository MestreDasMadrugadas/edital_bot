from flask import Flask, request, jsonify
from openpyxl import Workbook, load_workbook
import os

app = Flask(__name__)

ARQUIVO_EXCEL = "PMCE_Edital_Verticalizado_MODELO.xlsx"

def salvar_dados_em_excel(nome, email, produto, status, cpf):
    # Se o arquivo não existir, cria com cabeçalhos
    if not os.path.exists(ARQUIVO_EXCEL):
        wb = Workbook()
        ws = wb.active
        ws.title = "Compras"
        ws.append(["Nome", "Email", "Produto", "Status", "CPF"])
    else:
        wb = load_workbook(ARQUIVO_EXCEL)
        ws = wb.active

    # Adiciona nova linha com os dados recebidos
    ws.append([nome, email, produto, status, cpf])
    wb.save(ARQUIVO_EXCEL)
    print(f"✅ Dados salvos na planilha: {ARQUIVO_EXCEL}")

@app.route("/", methods=["GET"])
def home():
    return "✅ API do Edital Bot está no ar!"

@app.route("/webhook", methods=["POST"])
def webhook():
    try:
        data = request.get_json()
        print("📦 Webhook recebido:", data)

        # Acessando dados da compra (estrutura Kiwify)
        nome = data["purchase"]["buyer"]["name"]
        email = data["purchase"]["buyer"]["email"]
        cpf = data["purchase"]["buyer"].get("document", "CPF não informado")
        cpf = cpf.replace(".", "").replace("-", "")
        produto = data["purchase"]["product"]["name"]
        status = data["purchase"]["status"]

        # Salva na planilha
        salvar_dados_em_excel(nome, email, produto, status, cpf)

        return jsonify({
            "status": "ok",
            "mensagem": "Dados recebidos e salvos com sucesso!",
            "nome": nome,
            "email": email,
            "produto": produto,
            "status": status,
            "cpf": cpf
        })

    except Exception as e:
        print("❌ Erro ao processar webhook:", str(e))
        return jsonify({"status": "error", "message": str(e)}), 400

if __name__ == "__main__":
    # Isso permite rodar tanto localmente quanto na Render
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)


