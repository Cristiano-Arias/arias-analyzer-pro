from flask import Flask, render_template_string, request, send_file
from flask_cors import CORS
import os
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT

app = Flask(__name__)
CORS(app)

# Diretório para uploads
UPLOAD_FOLDER = 'uploads'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

# HTML mínimo
HTML = """
<!DOCTYPE html>
<html>
<head>
    <title>Arias Analyzer Pro</title>
    <style>
        body {
            font-family: Arial, sans-serif;
            padding: 2rem;
            background: #f4f4f4;
        }
        h1 {
            color: #333;
        }
        form {
            background: white;
            padding: 2rem;
            border-radius: 8px;
        }
        input[type=file], input[type=submit] {
            display: block;
            margin: 1rem 0;
        }
    </style>
</head>
<body>
    <h1>Proposal Analyzer by Arias</h1>
    <form method="post" enctype="multipart/form-data">
        <label for="file">Selecione um arquivo para análise:</label>
        <input type="file" name="file" required>
        <input type="submit" value="Analisar">
    </form>
</body>
</html>
"""

# Função para gerar relatório Word
def gerar_relatorio_word(nome_arquivo_analisado, resultado_analise, nome_arquivo_saida):
    document = Document()

    # Título
    titulo = document.add_heading('Relatório de Análise - Proposal Analyzer by Arias', level=1)
    titulo.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER

    # Arquivo analisado
    par_arquivo = document.add_paragraph()
    par_arquivo.add_run('Arquivo analisado: ').bold = True
    par_arquivo.add_run(nome_arquivo_analisado)

    document.add_paragraph('')

    # Resultado
    subtitulo = document.add_paragraph()
    subtitulo.add_run('Resultado da Análise:').bold = True

    for item in resultado_analise.split('\n'):
        document.add_paragraph(item, style='Normal')

    document.save(nome_arquivo_saida)

# Página principal
@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        uploaded_file = request.files['file']
        if uploaded_file.filename != '':
            path = os.path.join(UPLOAD_FOLDER, uploaded_file.filename)
            uploaded_file.save(path)

            # Simulação de resultado da análise
            resultado_analise = (
                "• Documento bem estruturado.\n"
                "• Ausência de inconsistências formais.\n"
                "• Conformidade com o Termo de Referência.\n"
                "• Sugestão: incluir cronograma de entrega."
            )

            output_path = os.path.join(UPLOAD_FOLDER, 'relatorio_analise.docx')
            gerar_relatorio_word(uploaded_file.filename, resultado_analise, output_path)

            return send_file(output_path, as_attachment=True)

    return render_template_string(HTML)

if __name__ == '__main__':
    app.run(debug=True)
