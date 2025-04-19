from flask import Flask, request, jsonify
from openpyxl import load_workbook
import os
import re

app = Flask(__name__)

ARQUIVO_MODELO = "PMCE_Edital_Verticalizado_MODELO.xlsx"
PASTA_RESULTADOS = "resultados"

# Garante que a pasta de saída exista
if not os.path.exists(PASTA_RESULTADOS):
    os.makedirs(PASTA_RESULTADOS)

def substituir_placeholders(nome, cpf):
    try:
        # Carrega o modelo
        wb = load_workbook(ARQUIVO_MODELO)
        ws = wb.active

        # Substitui {{nome}}, {{cpf}} no conteúdo das células
        for row in ws.iter_rows():
            for cell in row:
                if cell.value and isinstance(cell.value, str):
                    cell.value = cell.value.replace("{{nome}}", nome).replace("{{cpf}}", cpf)

        # Cria um nome de arquivo limpo com base no nome do comprador
        nome_arquivo = re.sub(r"[^\w\s-]", "", nome).strip().replace(" ", "_")
        caminho_arquivo = os.path.join(PASTA_RESULTADOS, f"resultado_{nome_arquivo}.xlsx")

        # Salva como novo arquivo
        wb.save(caminho_arquivo)
        print(f"✅ Arquivo personalizado salvo em: {caminho_arquivo}")

        return caminho_arquivo

    except Exception as e:
        print("❌ Erro ao gerar arquivo personalizado:", e)
        return None

@app.route("/", methods=["GET"])
def home():
    return "✅ API do Edital Bot está no ar!"

@app.route("/webhook", methods=["POST"])
def webhook():
    try:
        data = request.get_json()
        print("📦 Webhook recebido:", data)

        nome = data["purchase"]["buyer"]["name"]
        email = data["purchase"]["buyer"]["email"]
        cpf = data["purchase"]["buyer"].get("document", "000.000.000-00").replace(".", "").replace("-", "")
        produto = data["purchase"]["product"]["name"]
        status = data["purchase"]["status"]

        # Substitui os dados no Excel e gera um arquivo novo
        caminho_arquivo = substituir_placeholders(nome, cpf)

        return jsonify({
            "status": "ok",
            "mensagem": "Arquivo gerado com sucesso!",
            "arquivo_salvo_em": caminho_arquivo
        })

    except Exception as e:
        print("❌ Erro ao processar webhook:", str(e))
        return jsonify({"status": "error", "message": str(e)}), 400

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)



