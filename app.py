import os
import tempfile
import zipfile
from flask import Flask, request, jsonify, send_file, render_template_string
from flask_cors import CORS
import PyPDF2
import docx
import io
import subprocess
from datetime import datetime

app = Flask(__name__)
CORS(app)

# Configuração de upload
UPLOAD_FOLDER = 'uploads'
if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024  # 50MB max

# HTML da interface
HTML_TEMPLATE = '''
<!DOCTYPE html>
<html lang="pt-BR">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Proposal Analyzer Pro</title>
    <style>
        * {
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }
        
        body {
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            min-height: 100vh;
            padding: 20px;
        }
        
        .container {
            max-width: 1200px;
            margin: 0 auto;
            background: white;
            border-radius: 20px;
            box-shadow: 0 20px 40px rgba(0,0,0,0.1);
            overflow: hidden;
        }
        
        .header {
            background: linear-gradient(135deg, #2c3e50 0%, #3498db 100%);
            color: white;
            padding: 30px;
            text-align: center;
        }
        
        .header h1 {
            font-size: 2.5em;
            margin-bottom: 10px;
            font-weight: 300;
        }
        
        .header p {
            font-size: 1.2em;
            opacity: 0.9;
        }
        
        .content {
            padding: 40px;
        }
        
        .section {
            margin-bottom: 40px;
            padding: 30px;
            border: 2px solid #ecf0f1;
            border-radius: 15px;
            background: #fafafa;
        }
        
        .section h2 {
            color: #2c3e50;
            margin-bottom: 20px;
            font-size: 1.8em;
            border-bottom: 3px solid #3498db;
            padding-bottom: 10px;
        }
        
        .form-group {
            margin-bottom: 25px;
        }
        
        label {
            display: block;
            margin-bottom: 8px;
            font-weight: 600;
            color: #2c3e50;
            font-size: 1.1em;
        }
        
        input[type="text"], input[type="file"], textarea {
            width: 100%;
            padding: 15px;
            border: 2px solid #bdc3c7;
            border-radius: 10px;
            font-size: 16px;
            transition: all 0.3s ease;
        }
        
        input[type="text"]:focus, input[type="file"]:focus, textarea:focus {
            outline: none;
            border-color: #3498db;
            box-shadow: 0 0 10px rgba(52, 152, 219, 0.3);
        }
        
        textarea {
            resize: vertical;
            min-height: 100px;
        }
        
        .file-input-wrapper {
            position: relative;
            display: inline-block;
            width: 100%;
        }
        
        .file-input {
            position: absolute;
            opacity: 0;
            width: 100%;
            height: 100%;
            cursor: pointer;
        }
        
        .file-input-button {
            display: block;
            padding: 15px 25px;
            background: linear-gradient(135deg, #3498db, #2980b9);
            color: white;
            border: none;
            border-radius: 10px;
            cursor: pointer;
            text-align: center;
            font-size: 16px;
            transition: all 0.3s ease;
            width: 100%;
        }
        
        .file-input-button:hover {
            background: linear-gradient(135deg, #2980b9, #21618c);
            transform: translateY(-2px);
            box-shadow: 0 5px 15px rgba(0,0,0,0.2);
        }
        
        .proposal-item {
            background: white;
            padding: 25px;
            margin-bottom: 20px;
            border-radius: 15px;
            border: 2px solid #ecf0f1;
            box-shadow: 0 5px 15px rgba(0,0,0,0.05);
        }
        
        .proposal-item h3 {
            color: #2c3e50;
            margin-bottom: 15px;
            font-size: 1.4em;
        }
        
        .add-button {
            background: linear-gradient(135deg, #27ae60, #229954);
            color: white;
            border: none;
            padding: 15px 30px;
            border-radius: 10px;
            cursor: pointer;
            font-size: 16px;
            margin-top: 15px;
            transition: all 0.3s ease;
        }
        
        .add-button:hover {
            background: linear-gradient(135deg, #229954, #1e8449);
            transform: translateY(-2px);
            box-shadow: 0 5px 15px rgba(0,0,0,0.2);
        }
        
        .generate-button {
            background: linear-gradient(135deg, #e74c3c, #c0392b);
            color: white;
            border: none;
            padding: 20px 40px;
            border-radius: 15px;
            cursor: pointer;
            font-size: 18px;
            font-weight: 600;
            margin: 30px auto;
            display: block;
            transition: all 0.3s ease;
            box-shadow: 0 10px 25px rgba(231, 76, 60, 0.3);
        }
        
        .generate-button:hover {
            background: linear-gradient(135deg, #c0392b, #a93226);
            transform: translateY(-3px);
            box-shadow: 0 15px 35px rgba(231, 76, 60, 0.4);
        }
        
        .loading {
            display: none;
            text-align: center;
            padding: 30px;
        }
        
        .spinner {
            border: 4px solid #f3f3f3;
            border-top: 4px solid #3498db;
            border-radius: 50%;
            width: 50px;
            height: 50px;
            animation: spin 1s linear infinite;
            margin: 0 auto 20px;
        }
        
        @keyframes spin {
            0% { transform: rotate(0deg); }
            100% { transform: rotate(360deg); }
        }
        
        .result {
            display: none;
            margin-top: 30px;
            padding: 30px;
            background: #d5f4e6;
            border-radius: 15px;
            border: 2px solid #27ae60;
        }
        
        .result h3 {
            color: #27ae60;
            margin-bottom: 15px;
            font-size: 1.5em;
        }
        
        .download-button {
            background: linear-gradient(135deg, #8e44ad, #7d3c98);
            color: white;
            border: none;
            padding: 12px 25px;
            border-radius: 8px;
            cursor: pointer;
            font-size: 14px;
            margin-right: 10px;
            margin-bottom: 10px;
            transition: all 0.3s ease;
        }
        
        .download-button:hover {
            background: linear-gradient(135deg, #7d3c98, #6c3483);
            transform: translateY(-2px);
            box-shadow: 0 5px 15px rgba(0,0,0,0.2);
        }
        
        @media (max-width: 768px) {
            .container {
                margin: 10px;
                border-radius: 15px;
            }
            
            .header {
                padding: 20px;
            }
            
            .header h1 {
                font-size: 2em;
            }
            
            .content {
                padding: 20px;
            }
            
            .section {
                padding: 20px;
                margin-bottom: 25px;
            }
        }
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>🚀 Proposal Analyzer Pro</h1>
            <p>Sistema Avançado para Análise e Comparação de Propostas</p>
        </div>
        
        <div class="content">
            <form id="proposalForm">
                <!-- Informações Básicas -->
                <div class="section">
                    <h2>📋 Informações do Projeto</h2>
                    <div class="form-group">
                        <label for="projectName">Nome do Projeto *</label>
                        <input type="text" id="projectName" name="projectName" required 
                               placeholder="Digite o nome do projeto ou licitação">
                    </div>
                    <div class="form-group">
                        <label for="projectDescription">Descrição do Projeto</label>
                        <textarea id="projectDescription" name="projectDescription" 
                                  placeholder="Descrição opcional do projeto"></textarea>
                    </div>
                </div>
                
                <!-- Termo de Referência -->
                <div class="section">
                    <h2>📄 Termo de Referência (TR)</h2>
                    <div class="form-group">
                        <label for="trFile">Arquivo do TR *</label>
                        <div class="file-input-wrapper">
                            <input type="file" id="trFile" name="trFile" class="file-input" required
                                   accept=".pdf,.doc,.docx,.ppt,.pptx,.zip">
                            <div class="file-input-button">📁 Clique para selecionar o arquivo do TR</div>
                        </div>
                        <small style="color: #7f8c8d; margin-top: 5px; display: block;">
                            Formatos aceitos: PDF, DOC, DOCX, PPT, PPTX, ZIP
                        </small>
                    </div>
                </div>
                
                <!-- Propostas Técnicas -->
                <div class="section">
                    <h2>🔧 Propostas Técnicas</h2>
                    <div id="technicalProposals">
                        <div class="proposal-item">
                            <h3>Proposta Técnica 1</h3>
                            <div class="form-group">
                                <label>Nome da Empresa</label>
                                <input type="text" name="techCompany[]" placeholder="Nome da empresa">
                            </div>
                            <div class="form-group">
                                <label>Arquivo da Proposta Técnica</label>
                                <div class="file-input-wrapper">
                                    <input type="file" name="techFile[]" class="file-input"
                                           accept=".pdf,.doc,.docx,.ppt,.pptx,.zip">
                                    <div class="file-input-button">📁 Selecionar arquivo</div>
                                </div>
                            </div>
                        </div>
                    </div>
                    <button type="button" class="add-button" onclick="addTechnicalProposal()">
                        ➕ Adicionar Proposta Técnica
                    </button>
                </div>
                
                <!-- Propostas Comerciais -->
                <div class="section">
                    <h2>💰 Propostas Comerciais</h2>
                    <div id="commercialProposals">
                        <div class="proposal-item">
                            <h3>Proposta Comercial 1</h3>
                            <div class="form-group">
                                <label>Nome da Empresa</label>
                                <input type="text" name="commCompany[]" placeholder="Nome da empresa">
                            </div>
                            <div class="form-group">
                                <label>CNPJ (Opcional)</label>
                                <input type="text" name="commCnpj[]" placeholder="00.000.000/0000-00">
                            </div>
                            <div class="form-group">
                                <label>Arquivo da Proposta Comercial</label>
                                <div class="file-input-wrapper">
                                    <input type="file" name="commFile[]" class="file-input"
                                           accept=".pdf,.doc,.docx,.ppt,.pptx,.xls,.xlsx,.zip">
                                    <div class="file-input-button">📁 Selecionar arquivo</div>
                                </div>
                            </div>
                        </div>
                    </div>
                    <button type="button" class="add-button" onclick="addCommercialProposal()">
                        ➕ Adicionar Proposta Comercial
                    </button>
                </div>
                
                <!-- Botão de Geração -->
                <button type="submit" class="generate-button">
                    🚀 Gerar Relatório com Análise IA
                </button>
            </form>
            
            <!-- Loading -->
            <div id="loading" class="loading">
                <div class="spinner"></div>
                <h3>Processando documentos e gerando análise...</h3>
                <p>Isso pode levar alguns minutos. Por favor, aguarde.</p>
            </div>
            
            <!-- Resultado -->
            <div id="result" class="result">
                <h3>✅ Relatório Gerado com Sucesso!</h3>
                <p>Seu relatório de análise foi gerado. Escolha o formato para download:</p>
                <button class="download-button" onclick="downloadReport('markdown')">
                    📄 Download Markdown
                </button>
                <button class="download-button" onclick="downloadReport('pdf')">
                    📑 Download PDF
                </button>
            </div>
        </div>
    </div>
    
    <script>
        let techProposalCount = 1;
        let commProposalCount = 1;
        let currentReportId = null;
        
        // Atualizar texto dos botões de arquivo
        document.addEventListener('change', function(e) {
            if (e.target.type === 'file') {
                const button = e.target.nextElementSibling;
                if (e.target.files.length > 0) {
                    button.textContent = '✅ ' + e.target.files[0].name;
                    button.style.background = 'linear-gradient(135deg, #27ae60, #229954)';
                } else {
                    button.textContent = '📁 Selecionar arquivo';
                    button.style.background = 'linear-gradient(135deg, #3498db, #2980b9)';
                }
            }
        });
        
        function addTechnicalProposal() {
            if (techProposalCount >= 4) {
                alert('Máximo de 4 propostas técnicas permitidas.');
                return;
            }
            
            techProposalCount++;
            const container = document.getElementById('technicalProposals');
            const newProposal = document.createElement('div');
            newProposal.className = 'proposal-item';
            newProposal.innerHTML = `
                <h3>Proposta Técnica ${techProposalCount}</h3>
                <div class="form-group">
                    <label>Nome da Empresa</label>
                    <input type="text" name="techCompany[]" placeholder="Nome da empresa">
                </div>
                <div class="form-group">
                    <label>Arquivo da Proposta Técnica</label>
                    <div class="file-input-wrapper">
                        <input type="file" name="techFile[]" class="file-input"
                               accept=".pdf,.doc,.docx,.ppt,.pptx,.zip">
                        <div class="file-input-button">📁 Selecionar arquivo</div>
                    </div>
                </div>
            `;
            container.appendChild(newProposal);
        }
        
        function addCommercialProposal() {
            if (commProposalCount >= 4) {
                alert('Máximo de 4 propostas comerciais permitidas.');
                return;
            }
            
            commProposalCount++;
            const container = document.getElementById('commercialProposals');
            const newProposal = document.createElement('div');
            newProposal.className = 'proposal-item';
            newProposal.innerHTML = `
                <h3>Proposta Comercial ${commProposalCount}</h3>
                <div class="form-group">
                    <label>Nome da Empresa</label>
                    <input type="text" name="commCompany[]" placeholder="Nome da empresa">
                </div>
                <div class="form-group">
                    <label>CNPJ (Opcional)</label>
                    <input type="text" name="commCnpj[]" placeholder="00.000.000/0000-00">
                </div>
                <div class="form-group">
                    <label>Arquivo da Proposta Comercial</label>
                    <div class="file-input-wrapper">
                        <input type="file" name="commFile[]" class="file-input"
                               accept=".pdf,.doc,.docx,.ppt,.pptx,.xls,.xlsx,.zip">
                        <div class="file-input-button">📁 Selecionar arquivo</div>
                    </div>
                </div>
            `;
            container.appendChild(newProposal);
        }
        
        document.getElementById('proposalForm').addEventListener('submit', async function(e) {
            e.preventDefault();
            
            const formData = new FormData(this);
            
            // Mostrar loading
            document.getElementById('loading').style.display = 'block';
            document.getElementById('result').style.display = 'none';
            
            try {
                const response = await fetch('/analyze', {
                    method: 'POST',
                    body: formData
                });
                
                const result = await response.json();
                
                if (result.success) {
                    currentReportId = result.report_id;
                    document.getElementById('result').style.display = 'block';
                } else {
                    alert('Erro: ' + result.error);
                }
            } catch (error) {
                alert('Erro na comunicação com o servidor: ' + error.message);
            } finally {
                document.getElementById('loading').style.display = 'none';
            }
        });
        
        async function downloadReport(format) {
            if (!currentReportId) {
                alert('Nenhum relatório disponível para download.');
                return;
            }
            
            try {
                const response = await fetch(`/download/${currentReportId}/${format}`);
                
                if (response.ok) {
                    const blob = await response.blob();
                    const url = window.URL.createObjectURL(blob);
                    const a = document.createElement('a');
                    a.href = url;
                    a.download = `relatorio_analise.${format === 'pdf' ? 'pdf' : 'md'}`;
                    document.body.appendChild(a);
                    a.click();
                    window.URL.revokeObjectURL(url);
                    document.body.removeChild(a);
                } else {
                    alert('Erro ao baixar o arquivo.');
                }
            } catch (error) {
                alert('Erro na comunicação com o servidor: ' + error.message);
            }
        }
    </script>
</body>
</html>
'''

def extract_text_from_file(file_path):
    """Extrai texto de diferentes tipos de arquivo"""
    try:
        file_extension = os.path.splitext(file_path)[1].lower()
        
        if file_extension == '.pdf':
            with open(file_path, 'rb') as file:
                pdf_reader = PyPDF2.PdfReader(file)
                text = ""
                for page in pdf_reader.pages:
                    text += page.extract_text() + "\n"
                return text
        
        elif file_extension in ['.doc', '.docx']:
            doc = docx.Document(file_path)
            text = ""
            for paragraph in doc.paragraphs:
                text += paragraph.text + "\n"
            return text
        
        elif file_extension == '.txt':
            with open(file_path, 'r', encoding='utf-8') as file:
                return file.read()
        
        elif file_extension == '.zip':
            # Para arquivos ZIP, extrair e processar cada arquivo
            extracted_text = ""
            with zipfile.ZipFile(file_path, 'r') as zip_ref:
                temp_dir = tempfile.mkdtemp()
                zip_ref.extractall(temp_dir)
                
                for root, dirs, files in os.walk(temp_dir):
                    for file in files:
                        file_path_in_zip = os.path.join(root, file)
                        try:
                            extracted_text += extract_text_from_file(file_path_in_zip) + "\n\n"
                        except:
                            continue
            
            return extracted_text
        
        else:
            return "Formato de arquivo não suportado para extração de texto."
    
    except Exception as e:
        return f"Erro ao extrair texto: {str(e)}"

def generate_analysis_report(project_name, project_description, tr_text, technical_proposals, commercial_proposals):
    """Gera relatório de análise das propostas"""
    
    current_time = datetime.now().strftime("%d/%m/%Y %H:%M")
    
    report = f"""# Relatório de Análise de Propostas - {project_name}

## Informações do Projeto
**Nome:** {project_name}
**Descrição:** {project_description if project_description else 'Não informada'}
**Data de Análise:** {current_time}

## Resumo Executivo
Este relatório apresenta a análise comparativa das propostas técnicas e comerciais recebidas para o projeto "{project_name}".

## Análise do Termo de Referência
### Principais Requisitos Identificados:
{tr_text[:500]}...

## Análise das Propostas Técnicas
"""
    
    if technical_proposals:
        report += "### Propostas Técnicas Recebidas:\n\n"
        for i, proposal in enumerate(technical_proposals, 1):
            company = proposal.get('company', f'Empresa {i}')
            content = proposal.get('content', 'Conteúdo não disponível')[:300]
            
            report += f"""#### {i}. {company}
**Resumo da Proposta:**
{content}...

**Pontos Avaliados:**
- Metodologia de execução
- Cronograma proposto
- Equipe técnica
- Recursos e equipamentos

---

"""
    else:
        report += "Nenhuma proposta técnica foi submetida.\n\n"
    
    report += "## Análise das Propostas Comerciais\n"
    
    if commercial_proposals:
        report += "### Propostas Comerciais Recebidas:\n\n"
        for i, proposal in enumerate(commercial_proposals, 1):
            company = proposal.get('company', f'Empresa {i}')
            cnpj = proposal.get('cnpj', 'Não informado')
            content = proposal.get('content', 'Conteúdo não disponível')[:300]
            
            report += f"""#### {i}. {company}
**CNPJ:** {cnpj}
**Resumo da Proposta:**
{content}...

**Pontos Avaliados:**
- Preço total
- Composição de custos
- Condições de pagamento
- Prazos de execução

---

"""
    else:
        report += "Nenhuma proposta comercial foi submetida.\n\n"
    
    report += """## Comparativo e Recomendações

### Critérios de Avaliação
1. **Técnico:** Aderência ao TR, metodologia, cronograma, equipe
2. **Comercial:** Preço, condições de pagamento, viabilidade
3. **Qualificação:** Experiência da empresa, certificações

### Análise Comparativa
[Esta seção seria preenchida com análise detalhada baseada nos documentos submetidos]

### Recomendações
Com base na análise realizada, recomenda-se:
1. Verificação detalhada das propostas técnicas
2. Análise da saúde financeira das empresas proponentes
3. Esclarecimentos adicionais se necessário

## Conclusão
Este relatório fornece uma visão geral das propostas recebidas. Recomenda-se análise detalhada adicional antes da tomada de decisão final.

---
*Relatório gerado automaticamente pelo Proposal Analyzer Pro*
*Data: {current_time}*
"""
    
    return report

@app.route('/')
def index():
    return render_template_string(HTML_TEMPLATE)

@app.route('/analyze', methods=['POST'])
def analyze_proposals():
    try:
        # Obter dados do formulário
        project_name = request.form.get('projectName')
        project_description = request.form.get('projectDescription', '')
        
        if not project_name:
            return jsonify({'success': False, 'error': 'Nome do projeto é obrigatório.'})
        
        # Processar TR
        tr_file = request.files.get('trFile')
        if not tr_file:
            return jsonify({'success': False, 'error': 'Arquivo do TR é obrigatório.'})
        
        tr_filename = f"tr_{tr_file.filename}"
        tr_path = os.path.join(app.config['UPLOAD_FOLDER'], tr_filename)
        tr_file.save(tr_path)
        tr_text = extract_text_from_file(tr_path)
        
        # Processar propostas técnicas
        technical_proposals = []
        tech_companies = request.form.getlist('techCompany[]')
        tech_files = request.files.getlist('techFile[]')
        
        for i, (company, file) in enumerate(zip(tech_companies, tech_files)):
            if company and file and file.filename:
                file_path = os.path.join(app.config['UPLOAD_FOLDER'], f"tech_{i}_{file.filename}")
                file.save(file_path)
                content = extract_text_from_file(file_path)
                technical_proposals.append({
                    'company': company,
                    'content': content
                })
        
        # Processar propostas comerciais
        commercial_proposals = []
        comm_companies = request.form.getlist('commCompany[]')
        comm_cnpjs = request.form.getlist('commCnpj[]')
        comm_files = request.files.getlist('commFile[]')
        
        for i, (company, cnpj, file) in enumerate(zip(comm_companies, comm_cnpjs, comm_files)):
            if company and file and file.filename:
                file_path = os.path.join(app.config['UPLOAD_FOLDER'], f"comm_{i}_{file.filename}")
                file.save(file_path)
                content = extract_text_from_file(file_path)
                commercial_proposals.append({
                    'company': company,
                    'cnpj': cnpj,
                    'content': content
                })
        
        # Gerar relatório
        report = generate_analysis_report(
            project_name, project_description, tr_text,
            technical_proposals, commercial_proposals
        )
        
        # Salvar relatório
        report_id = f"report_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
        report_path = os.path.join(app.config['UPLOAD_FOLDER'], f"{report_id}.md")
        
        with open(report_path, 'w', encoding='utf-8') as f:
            f.write(report)
        
        return jsonify({'success': True, 'report_id': report_id})
        
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)})

@app.route('/download/<report_id>/<format>')
def download_report(report_id, format):
    try:
        if format == 'markdown':
            file_path = os.path.join(app.config['UPLOAD_FOLDER'], f"{report_id}.md")
            return send_file(file_path, as_attachment=True, download_name='relatorio_analise.md')
        elif format == 'pdf':
            # Gerar PDF usando uma biblioteca Python pura
            md_file_path = os.path.join(app.config['UPLOAD_FOLDER'], f"{report_id}.md")
            pdf_file_path = os.path.join(app.config['UPLOAD_FOLDER'], f"relatorio_analise.pdf")
            
            # Ler o conteúdo markdown
            with open(md_file_path, 'r', encoding='utf-8') as f:
                markdown_content = f.read()
            
            # Converter para PDF usando reportlab
            from reportlab.lib.pagesizes import letter, A4
            from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
            from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
            from reportlab.lib.units import inch
            
            # Criar documento PDF
            doc = SimpleDocTemplate(pdf_file_path, pagesize=A4)
            styles = getSampleStyleSheet()
            story = []
            
            # Estilo personalizado
            title_style = ParagraphStyle(
                'CustomTitle',
                parent=styles['Heading1'],
                fontSize=16,
                spaceAfter=30,
            )
            
            heading_style = ParagraphStyle(
                'CustomHeading',
                parent=styles['Heading2'],
                fontSize=14,
                spaceAfter=12,
            )
            
            normal_style = ParagraphStyle(
                'CustomNormal',
                parent=styles['Normal'],
                fontSize=10,
                spaceAfter=12,
            )
            
            # Processar markdown simples para PDF
            lines = markdown_content.split('\n')
            for line in lines:
                line = line.strip()
                if not line:
                    story.append(Spacer(1, 12))
                elif line.startswith('# '):
                    story.append(Paragraph(line[2:], title_style))
                elif line.startswith('## '):
                    story.append(Paragraph(line[3:], heading_style))
                elif line.startswith('### '):
                    story.append(Paragraph(line[4:], heading_style))
                elif line.startswith('**') and line.endswith('**'):
                    story.append(Paragraph(f"<b>{line[2:-2]}</b>", normal_style))
                else:
                    if line:
                        story.append(Paragraph(line, normal_style))
            
            # Construir PDF
            doc.build(story)
            
            return send_file(pdf_file_path, as_attachment=True, download_name='relatorio_analise.pdf')
        else:
            return jsonify({'error': 'Formato não suportado'}), 400
            
    except Exception as e:
        return jsonify({'error': str(e)}), 500

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=False)

