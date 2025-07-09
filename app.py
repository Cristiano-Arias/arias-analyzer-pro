from flask import Flask, request, render_template, send_file
from werkzeug.utils import secure_filename
import os
from docx import Document

app = Flask(__name__)
UPLOAD_FOLDER = 'uploads'
REPORT_PATH = 'report.docx'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/analyze', methods=['POST'])
def analyze():
    uploaded_file = request.files['file']
    filename = secure_filename(uploaded_file.filename)
    filepath = os.path.join(UPLOAD_FOLDER, filename)
    uploaded_file.save(filepath)

    # Análise simulada (por enquanto)
    analysis_result = f"Análise simulada para o arquivo: {filename}"

    # Criar relatório Word
    doc = Document()
    doc.add_heading('Relatório de Análise - Proposal Analyzer by Arias', 0)
    doc.add_paragraph(f'Arquivo analisado: {filename}')
    doc.add_paragraph('Resultado da análise:')
    doc.add_paragraph(analysis_result)
    doc.save(REPORT_PATH)

    return render_template('result.html')

@app.route('/download')
def download_report():
    return send_file(REPORT_PATH, as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True)
Atualiza app.py com upload e geração de relatório Word
