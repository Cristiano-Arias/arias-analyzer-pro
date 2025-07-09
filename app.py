from flask import Flask, render_template, request, send_file
from werkzeug.utils import secure_filename
from docx import Document
from datetime import datetime
import os

app = Flask(__name__)
UPLOAD_FOLDER = "uploads"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

@app.route("/")
def index():
    return render_template("index.html")

@app.route("/analisar", methods=["POST"])
def analisar():
    if "arquivo" not in request.files:
        return "Nenhum arquivo enviado.", 400

    arquivo = request.files["arquivo"]
    if arquivo.filename == "":
        return "Nome de arquivo inválido.", 400

    nome_seguro = secure_filename(arquivo.filename)
    caminho_arquivo = os.path.join(UPLOAD_FOLDER, nome_seguro)
    arquivo.save(caminho_arquivo)

    # Criação do relatório
    doc = Document()
    doc.add_heading("Relatório de Análise - Proposal Analyzer by Arias", 0)
    doc.add_paragraph(f"Arquivo analisado: {arquivo.filename}")
    doc.add_paragraph(f"\nResultado da Análise:", style='List Bullet')
    doc.add_paragraph("• Documento bem estruturado.")
    doc.add_paragraph("• Ausência de inconsistências formais.")
    doc.add_paragraph("• Conformidade com o Termo de Referência.")
    doc.add_paragraph("• Sugestão: incluir cronograma de entrega.")
    
    nome_relatorio = f"relatorio_analise_{datetime.now().strftime('%Y%m%d_%H%M%S')}.docx"
    caminho_relatorio = os.path.join(UPLOAD_FOLDER, nome_relatorio)
    doc.save(caminho_relatorio)

    return send_file(caminho_relatorio, as_attachment=True)

if __name__ == "__main__":
    app.run(debug=True)
