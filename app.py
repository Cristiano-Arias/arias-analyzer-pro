import os
import tempfile
import zipfile
from flask import Flask, request, jsonify, send_file, render_template_string
from flask_cors import CORS
import PyPDF2
import docx
import pandas as pd
import openpyxl
import io
import re
from datetime import datetime
from collections import defaultdict
import gc

app = Flask(__name__)
CORS(app)

# Configuração de upload
UPLOAD_FOLDER = 'uploads'
if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024  # 50MB

# HTML da interface simplificada (sem TR)
HTML_TEMPLATE = '''
<!DOCTYPE html>
<html lang="pt-BR">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Proposal Analyzer Pro - Comparação de Propostas</title>
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
        
        .info-box {
            background: #e8f4fd;
            border: 2px solid #3498db;
            border-radius: 10px;
            padding: 20px;
            margin-bottom: 30px;
        }
        
        .info-box h3 {
            color: #2c3e50;
            margin-bottom: 10px;
        }
        
        .info-box p {
            color: #34495e;
            line-height: 1.6;
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
            <p>Análise Comparativa Inteligente de Propostas Técnicas e Comerciais</p>
        </div>
        
        <div class="content">
            <div class="info-box">
                <h3>📊 Como Funciona</h3>
                <p>Este sistema compara automaticamente suas propostas técnicas e comerciais, extraindo e analisando metodologias, cronogramas, equipes, recursos, preços, composição de custos e muito mais. Não é necessário Termo de Referência - a análise é feita comparando as propostas entre si.</p>
            </div>
            
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
                
                <!-- Propostas Técnicas -->
                <div class="section">
                    <h2>🔧 Propostas Técnicas</h2>
                    <p style="margin-bottom: 20px; color: #7f8c8d;">
                        Adicione as propostas técnicas para análise comparativa de metodologias, cronogramas, equipes, recursos e muito mais.
                    </p>
                    <div id="technicalProposals">
                        <div class="proposal-item">
                            <h3>Proposta Técnica 1</h3>
                            <div class="form-group">
                                <label>Nome da Empresa *</label>
                                <input type="text" name="techCompany[]" placeholder="Nome da empresa" required>
                            </div>
                            <div class="form-group">
                                <label>Arquivo da Proposta Técnica *</label>
                                <div class="file-input-wrapper">
                                    <input type="file" name="techFile[]" class="file-input" required
                                           accept=".pdf,.doc,.docx,.ppt,.pptx,.zip">
                                    <div class="file-input-button">📁 Selecionar arquivo</div>
                                </div>
                            </div>
                        </div>
                        <div class="proposal-item">
                            <h3>Proposta Técnica 2</h3>
                            <div class="form-group">
                                <label>Nome da Empresa *</label>
                                <input type="text" name="techCompany[]" placeholder="Nome da empresa" required>
                            </div>
                            <div class="form-group">
                                <label>Arquivo da Proposta Técnica *</label>
                                <div class="file-input-wrapper">
                                    <input type="file" name="techFile[]" class="file-input" required
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
                    <p style="margin-bottom: 20px; color: #7f8c8d;">
                        Adicione as propostas comerciais para análise de preços, composição de custos, BDI, condições de pagamento e garantias.
                    </p>
                    <div id="commercialProposals">
                        <div class="proposal-item">
                            <h3>Proposta Comercial 1</h3>
                            <div class="form-group">
                                <label>Nome da Empresa *</label>
                                <input type="text" name="commCompany[]" placeholder="Nome da empresa" required>
                            </div>
                            <div class="form-group">
                                <label>CNPJ (Opcional)</label>
                                <input type="text" name="commCnpj[]" placeholder="00.000.000/0000-00">
                            </div>
                            <div class="form-group">
                                <label>Arquivo da Proposta Comercial *</label>
                                <div class="file-input-wrapper">
                                    <input type="file" name="commFile[]" class="file-input" required
                                           accept=".pdf,.doc,.docx,.ppt,.pptx,.xls,.xlsx,.zip">
                                    <div class="file-input-button">📁 Selecionar arquivo</div>
                                </div>
                            </div>
                        </div>
                        <div class="proposal-item">
                            <h3>Proposta Comercial 2</h3>
                            <div class="form-group">
                                <label>Nome da Empresa *</label>
                                <input type="text" name="commCompany[]" placeholder="Nome da empresa" required>
                            </div>
                            <div class="form-group">
                                <label>CNPJ (Opcional)</label>
                                <input type="text" name="commCnpj[]" placeholder="00.000.000/0000-00">
                            </div>
                            <div class="form-group">
                                <label>Arquivo da Proposta Comercial *</label>
                                <div class="file-input-wrapper">
                                    <input type="file" name="commFile[]" class="file-input" required
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
                    🚀 Gerar Análise Comparativa com IA
                </button>
            </form>
            
            <!-- Loading -->
            <div id="loading" class="loading">
                <div class="spinner"></div>
                <h3>Processando propostas e gerando análise comparativa...</h3>
                <p>Analisando metodologias, cronogramas, equipes, recursos, preços e muito mais. Aguarde alguns minutos.</p>
            </div>
            
            <!-- Resultado -->
            <div id="result" class="result">
                <h3>✅ Análise Comparativa Gerada com Sucesso!</h3>
                <p>Seu relatório de análise comparativa foi gerado. Escolha o formato para download:</p>
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
        let techProposalCount = 2;
        let commProposalCount = 2;
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
            if (techProposalCount >= 6) {
                alert('Máximo de 6 propostas técnicas permitidas.');
                return;
            }
            
            techProposalCount++;
            const container = document.getElementById('technicalProposals');
            const newProposal = document.createElement('div');
            newProposal.className = 'proposal-item';
            newProposal.innerHTML = `
                <h3>Proposta Técnica ${techProposalCount}</h3>
                <div class="form-group">
                    <label>Nome da Empresa *</label>
                    <input type="text" name="techCompany[]" placeholder="Nome da empresa" required>
                </div>
                <div class="form-group">
                    <label>Arquivo da Proposta Técnica *</label>
                    <div class="file-input-wrapper">
                        <input type="file" name="techFile[]" class="file-input" required
                               accept=".pdf,.doc,.docx,.ppt,.pptx,.zip">
                        <div class="file-input-button">📁 Selecionar arquivo</div>
                    </div>
                </div>
            `;
            container.appendChild(newProposal);
        }
        
        function addCommercialProposal() {
            if (commProposalCount >= 6) {
                alert('Máximo de 6 propostas comerciais permitidas.');
                return;
            }
            
            commProposalCount++;
            const container = document.getElementById('commercialProposals');
            const newProposal = document.createElement('div');
            newProposal.className = 'proposal-item';
            newProposal.innerHTML = `
                <h3>Proposta Comercial ${commProposalCount}</h3>
                <div class="form-group">
                    <label>Nome da Empresa *</label>
                    <input type="text" name="commCompany[]" placeholder="Nome da empresa" required>
                </div>
                <div class="form-group">
                    <label>CNPJ (Opcional)</label>
                    <input type="text" name="commCnpj[]" placeholder="00.000.000/0000-00">
                </div>
                <div class="form-group">
                    <label>Arquivo da Proposta Comercial *</label>
                    <div class="file-input-wrapper">
                        <input type="file" name="commFile[]" class="file-input" required
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
            
            // Validar se há pelo menos 2 propostas técnicas e 2 comerciais
            const techCompanies = formData.getAll('techCompany[]').filter(c => c.trim());
            const commCompanies = formData.getAll('commCompany[]').filter(c => c.trim());
            
            if (techCompanies.length < 2) {
                alert('É necessário pelo menos 2 propostas técnicas para comparação.');
                return;
            }
            
            if (commCompanies.length < 2) {
                alert('É necessário pelo menos 2 propostas comerciais para comparação.');
                return;
            }
            
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
                    a.download = `analise_comparativa.${format === 'pdf' ? 'pdf' : 'md'}`;
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
    """Extrai texto de arquivos de forma otimizada"""
    try:
        file_extension = os.path.splitext(file_path)[1].lower()
        
        if file_extension == '.pdf':
            return extract_pdf_text(file_path)
        elif file_extension in ['.doc', '.docx']:
            return extract_docx_text(file_path)
        elif file_extension in ['.xls', '.xlsx']:
            return extract_excel_text(file_path)
        elif file_extension == '.txt':
            return extract_txt_text(file_path)
        elif file_extension == '.zip':
            return extract_zip_text(file_path)
        else:
            return "Formato não suportado"
    
    except Exception as e:
        return f"Erro na extração: {str(e)}"

def extract_pdf_text(file_path):
    """Extração de texto de PDF"""
    text = ""
    try:
        with open(file_path, 'rb') as file:
            pdf_reader = PyPDF2.PdfReader(file)
            for page in pdf_reader.pages:
                text += page.extract_text() + "\n"
        return text
    except Exception as e:
        return f"Erro PDF: {str(e)}"

def extract_docx_text(file_path):
    """Extração de texto de DOCX"""
    try:
        doc = docx.Document(file_path)
        text = ""
        for paragraph in doc.paragraphs:
            text += paragraph.text + "\n"
        return text
    except Exception as e:
        return f"Erro DOCX: {str(e)}"

def extract_txt_text(file_path):
    """Extração de texto de TXT"""
    try:
        with open(file_path, 'r', encoding='utf-8') as file:
            return file.read()
    except Exception as e:
        return f"Erro TXT: {str(e)}"

def extract_excel_text(file_path):
    """Extração de dados de Excel"""
    try:
        wb = openpyxl.load_workbook(file_path, read_only=True)
        extracted_data = []
        
        for sheet_name in wb.sheetnames:
            try:
                df = pd.read_excel(file_path, sheet_name=sheet_name)
                sheet_text = f"\n=== ABA: {sheet_name} ===\n"
                
                for index, row in df.iterrows():
                    row_text = " | ".join([str(cell) if pd.notna(cell) else "" for cell in row])
                    if row_text.strip() and not row_text.replace(" | ", "").strip() == "":
                        sheet_text += f"Linha {index + 1}: {row_text}\n"
                
                extracted_data.append(sheet_text)
                
            except Exception as e:
                extracted_data.append(f"Erro na aba {sheet_name}: {str(e)}")
        
        wb.close()
        
        combined_text = f"ARQUIVO EXCEL: {os.path.basename(file_path)}\n"
        combined_text += "\n".join(extracted_data)
        
        return combined_text
        
    except Exception as e:
        return f"Erro Excel: {str(e)}"

def extract_zip_text(file_path):
    """Extração de texto de ZIP"""
    try:
        extracted_text = ""
        with zipfile.ZipFile(file_path, 'r') as zip_ref:
            temp_dir = tempfile.mkdtemp()
            
            relevant_extensions = ['.pdf', '.docx', '.doc', '.xlsx', '.xls', '.txt']
            
            for file_info in zip_ref.filelist:
                file_ext = os.path.splitext(file_info.filename)[1].lower()
                if file_ext in relevant_extensions:
                    try:
                        zip_ref.extract(file_info, temp_dir)
                        file_path_in_zip = os.path.join(temp_dir, file_info.filename)
                        
                        file_text = extract_text_from_file(file_path_in_zip)
                        extracted_text += f"\n=== ARQUIVO: {file_info.filename} ===\n"
                        extracted_text += file_text + "\n\n"
                        
                        os.remove(file_path_in_zip)
                        
                    except:
                        continue
            
            import shutil
            shutil.rmtree(temp_dir)
        
        return extracted_text
        
    except Exception as e:
        return f"Erro ZIP: {str(e)}"

def analyze_technical_proposal_detailed(proposal_text, company_name):
    """Análise técnica detalhada e aprofundada"""
    analysis = {
        'empresa': company_name,
        'cnpj': '',
        'metodologia_execucao': {
            'descricao': '',
            'fases': [],
            'ferramentas': [],
            'abordagem': '',
            'inovacoes': []
        },
        'histograma_mao_obra': {
            'presente': False,
            'detalhes': [],
            'qualificacoes': [],
            'quantidade_total': ''
        },
        'histograma_equipamentos': {
            'presente': False,
            'equipamentos': [],
            'tecnologias': [],
            'quantidade_total': ''
        },
        'lista_materiais': {
            'presente': False,
            'materiais': [],
            'quantidades': [],
            'especificacoes': []
        },
        'obrigacoes': {
            'principais': [],
            'responsabilidades': [],
            'compromissos': []
        },
        'canteiro': {
            'informacoes': [],
            'logistica': '',
            'infraestrutura': '',
            'organizacao': ''
        },
        'exclusoes': {
            'itens_excluidos': [],
            'limitacoes': [],
            'nao_inclusos': []
        },
        'prazo_cronograma': {
            'prazo_total': '',
            'marcos_principais': [],
            'fases_cronograma': [],
            'viabilidade': ''
        },
        'equipes_recursos': {
            'estrutura_equipe': [],
            'coordenador': '',
            'especialistas': [],
            'recursos_humanos': [],
            'alocacao': []
        },
        'pontos_fortes': [],
        'pontos_fracos': [],
        'score_geral': 0
    }
    
    # Extrair CNPJ
    cnpj_patterns = [
        r'(\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2})',
        r'CNPJ[:\s]*(\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2})',
        r'(\d{2}\d{3}\d{3}\d{4}\d{2})'
    ]
    
    for pattern in cnpj_patterns:
        cnpj_match = re.search(pattern, proposal_text)
        if cnpj_match:
            analysis['cnpj'] = cnpj_match.group(1)
            break
    
    # Dividir texto em seções para análise
    lines = proposal_text.split('\n')
    
    # Análise de Metodologia de Execução
    metodologia_keywords = ['metodologia', 'método', 'execução', 'abordagem', 'estratégia', 'procedimento']
    metodologia_sections = []
    
    for i, line in enumerate(lines):
        if any(keyword in line.lower() for keyword in metodologia_keywords) and len(line.strip()) > 30:
            # Capturar contexto (linha atual + próximas 3)
            context = []
            for j in range(i, min(i+4, len(lines))):
                if lines[j].strip():
                    context.append(lines[j].strip())
            metodologia_sections.extend(context)
    
    if metodologia_sections:
        analysis['metodologia_execucao']['descricao'] = ' '.join(metodologia_sections[:3])
        
        # Extrair fases
        fase_patterns = [r'fase\s+(\d+)', r'etapa\s+(\d+)', r'(\d+)ª?\s*fase', r'(\d+)ª?\s*etapa']
        for section in metodologia_sections:
            for pattern in fase_patterns:
                matches = re.findall(pattern, section.lower())
                for match in matches:
                    if f"Fase {match}" not in analysis['metodologia_execucao']['fases']:
                        analysis['metodologia_execucao']['fases'].append(f"Fase {match}")
        
        # Extrair ferramentas
        ferramenta_keywords = ['ferramenta', 'software', 'equipamento', 'tecnologia', 'sistema']
        for section in metodologia_sections:
            for keyword in ferramenta_keywords:
                if keyword in section.lower():
                    analysis['metodologia_execucao']['ferramentas'].append(section[:100])
                    break
        
        analysis['pontos_fortes'].append('Metodologia de execução apresentada')
    else:
        analysis['pontos_fracos'].append('Metodologia de execução não detalhada')
    
    # Análise de Histograma de Mão de Obra
    mao_obra_keywords = ['mão de obra', 'mao de obra', 'pessoal', 'funcionários', 'trabalhadores', 'equipe', 'histograma']
    mao_obra_sections = []
    
    for line in lines:
        if any(keyword in line.lower() for keyword in mao_obra_keywords) and len(line.strip()) > 20:
            mao_obra_sections.append(line.strip())
    
    if mao_obra_sections:
        analysis['histograma_mao_obra']['presente'] = True
        analysis['histograma_mao_obra']['detalhes'] = mao_obra_sections[:5]
        
        # Extrair qualificações
        qualif_keywords = ['engenheiro', 'técnico', 'especialista', 'coordenador', 'supervisor', 'operador']
        for section in mao_obra_sections:
            for keyword in qualif_keywords:
                if keyword in section.lower():
                    analysis['histograma_mao_obra']['qualificacoes'].append(section[:80])
        
        # Extrair quantidades
        qty_matches = re.findall(r'(\d+)\s*(?:pessoas|funcionários|trabalhadores)', ' '.join(mao_obra_sections).lower())
        if qty_matches:
            analysis['histograma_mao_obra']['quantidade_total'] = f"{sum(int(q) for q in qty_matches)} pessoas"
        
        analysis['pontos_fortes'].append('Histograma de mão de obra presente')
    else:
        analysis['pontos_fracos'].append('Histograma de mão de obra não apresentado')
    
    # Análise de Histograma de Equipamentos
    equip_keywords = ['equipamento', 'máquina', 'veículo', 'ferramenta', 'instrumento', 'aparelho']
    equip_sections = []
    
    for line in lines:
        if any(keyword in line.lower() for keyword in equip_keywords) and len(line.strip()) > 20:
            equip_sections.append(line.strip())
    
    if equip_sections:
        analysis['histograma_equipamentos']['presente'] = True
        analysis['histograma_equipamentos']['equipamentos'] = equip_sections[:8]
        
        # Extrair tecnologias
        tech_keywords = ['gps', 'laser', 'digital', 'automatizado', 'computadorizado', 'eletrônico']
        for section in equip_sections:
            for keyword in tech_keywords:
                if keyword in section.lower():
                    analysis['histograma_equipamentos']['tecnologias'].append(section[:80])
        
        analysis['pontos_fortes'].append('Histograma de equipamentos presente')
    else:
        analysis['pontos_fracos'].append('Histograma de equipamentos não apresentado')
    
    # Análise de Lista de Materiais
    material_keywords = ['material', 'insumo', 'produto', 'componente', 'item', 'especificação']
    material_sections = []
    
    for line in lines:
        if any(keyword in line.lower() for keyword in material_keywords) and len(line.strip()) > 25:
            material_sections.append(line.strip())
    
    if material_sections:
        analysis['lista_materiais']['presente'] = True
        analysis['lista_materiais']['materiais'] = material_sections[:10]
        
        # Extrair quantidades
        qty_patterns = [r'(\d+(?:[.,]\d+)?)\s*(?:m³|m²|m|kg|ton|unid|pç)', r'(\d+(?:[.,]\d+)?)\s*(?:metros|quilos|toneladas|unidades)']
        for section in material_sections:
            for pattern in qty_patterns:
                matches = re.findall(pattern, section.lower())
                analysis['lista_materiais']['quantidades'].extend(matches[:3])
        
        # Extrair especificações
        spec_keywords = ['especificação', 'norma', 'padrão', 'qualidade', 'tipo', 'modelo']
        for section in material_sections:
            if any(keyword in section.lower() for keyword in spec_keywords):
                analysis['lista_materiais']['especificacoes'].append(section[:100])
        
        analysis['pontos_fortes'].append('Lista de materiais presente')
    else:
        analysis['pontos_fracos'].append('Lista de materiais não apresentada')
    
    # Análise de Obrigações
    obrig_keywords = ['obrigação', 'responsabilidade', 'compromisso', 'dever', 'incumbência', 'atribuição']
    obrig_sections = []
    
    for line in lines:
        if any(keyword in line.lower() for keyword in obrig_keywords) and len(line.strip()) > 30:
            obrig_sections.append(line.strip())
    
    if obrig_sections:
        analysis['obrigacoes']['principais'] = obrig_sections[:5]
        
        # Categorizar responsabilidades
        resp_keywords = ['responsável', 'encarregado', 'incumbido']
        for section in obrig_sections:
            if any(keyword in section.lower() for keyword in resp_keywords):
                analysis['obrigacoes']['responsabilidades'].append(section[:120])
        
        analysis['pontos_fortes'].append('Obrigações claramente definidas')
    else:
        analysis['pontos_fracos'].append('Obrigações não especificadas')
    
    # Análise de Canteiro
    canteiro_keywords = ['canteiro', 'obra', 'instalação', 'infraestrutura', 'logística', 'organização']
    canteiro_sections = []
    
    for line in lines:
        if any(keyword in line.lower() for keyword in canteiro_keywords) and len(line.strip()) > 25:
            canteiro_sections.append(line.strip())
    
    if canteiro_sections:
        analysis['canteiro']['informacoes'] = canteiro_sections[:6]
        
        # Extrair informações específicas
        if any('logística' in s.lower() for s in canteiro_sections):
            analysis['canteiro']['logistica'] = 'Informações de logística apresentadas'
        
        if any('infraestrutura' in s.lower() for s in canteiro_sections):
            analysis['canteiro']['infraestrutura'] = 'Infraestrutura do canteiro detalhada'
        
        analysis['pontos_fortes'].append('Informações sobre canteiro apresentadas')
    else:
        analysis['pontos_fracos'].append('Informações sobre canteiro não apresentadas')
    
    # Análise de Exclusões
    exclusao_keywords = ['exclusão', 'excluído', 'não incluso', 'não incluído', 'fora do escopo', 'limitação']
    exclusao_sections = []
    
    for line in lines:
        if any(keyword in line.lower() for keyword in exclusao_keywords) and len(line.strip()) > 20:
            exclusao_sections.append(line.strip())
    
    if exclusao_sections:
        analysis['exclusoes']['itens_excluidos'] = exclusao_sections[:5]
        
        # Categorizar limitações
        limit_keywords = ['limitação', 'restrição', 'condição']
        for section in exclusao_sections:
            if any(keyword in section.lower() for keyword in limit_keywords):
                analysis['exclusoes']['limitacoes'].append(section[:100])
        
        analysis['pontos_fortes'].append('Exclusões claramente especificadas')
    else:
        analysis['pontos_fracos'].append('Exclusões não especificadas')
    
    # Análise de Prazo e Cronograma
    prazo_keywords = ['prazo', 'cronograma', 'tempo', 'duração', 'período', 'dias', 'meses']
    prazo_sections = []
    
    for line in lines:
        if any(keyword in line.lower() for keyword in prazo_keywords) and len(line.strip()) > 20:
            prazo_sections.append(line.strip())
    
    if prazo_sections:
        # Extrair prazo total
        prazo_patterns = [r'(\d+)\s*dias?', r'(\d+)\s*meses?', r'(\d+)\s*semanas?']
        for section in prazo_sections:
            for pattern in prazo_patterns:
                matches = re.findall(pattern, section.lower())
                if matches:
                    analysis['prazo_cronograma']['prazo_total'] = f"{matches[0]} {pattern.split('s')[0].split('\\')[1]}"
                    break
        
        # Extrair marcos
        marco_keywords = ['marco', 'milestone', 'entrega', 'conclusão']
        for section in prazo_sections:
            if any(keyword in section.lower() for keyword in marco_keywords):
                analysis['prazo_cronograma']['marcos_principais'].append(section[:100])
        
        analysis['prazo_cronograma']['fases_cronograma'] = prazo_sections[:4]
        analysis['prazo_cronograma']['viabilidade'] = 'Cronograma apresentado'
        analysis['pontos_fortes'].append('Prazo e cronograma definidos')
    else:
        analysis['prazo_cronograma']['viabilidade'] = 'Cronograma não apresentado'
        analysis['pontos_fracos'].append('Prazo e cronograma não definidos')
    
    # Análise de Equipes e Recursos
    equipe_keywords = ['equipe', 'time', 'grupo', 'coordenador', 'gerente', 'responsável técnico']
    equipe_sections = []
    
    for line in lines:
        if any(keyword in line.lower() for keyword in equipe_keywords) and len(line.strip()) > 25:
            equipe_sections.append(line.strip())
    
    if equipe_sections:
        analysis['equipes_recursos']['estrutura_equipe'] = equipe_sections[:6]
        
        # Extrair coordenador
        coord_patterns = [r'coordenador[:\s]*([^,\n]+)', r'gerente[:\s]*([^,\n]+)', r'responsável[:\s]*([^,\n]+)']
        for section in equipe_sections:
            for pattern in coord_patterns:
                match = re.search(pattern, section.lower())
                if match:
                    analysis['equipes_recursos']['coordenador'] = match.group(1)[:50]
                    break
        
        # Extrair especialistas
        espec_keywords = ['especialista', 'engenheiro', 'técnico', 'consultor']
        for section in equipe_sections:
            for keyword in espec_keywords:
                if keyword in section.lower():
                    analysis['equipes_recursos']['especialistas'].append(section[:80])
        
        analysis['pontos_fortes'].append('Estrutura de equipe apresentada')
    else:
        analysis['pontos_fracos'].append('Estrutura de equipe não apresentada')
    
    # Calcular score geral
    criterios_atendidos = 0
    total_criterios = 9
    
    if analysis['metodologia_execucao']['descricao']:
        criterios_atendidos += 1
    if analysis['histograma_mao_obra']['presente']:
        criterios_atendidos += 1
    if analysis['histograma_equipamentos']['presente']:
        criterios_atendidos += 1
    if analysis['lista_materiais']['presente']:
        criterios_atendidos += 1
    if analysis['obrigacoes']['principais']:
        criterios_atendidos += 1
    if analysis['canteiro']['informacoes']:
        criterios_atendidos += 1
    if analysis['exclusoes']['itens_excluidos']:
        criterios_atendidos += 1
    if analysis['prazo_cronograma']['prazo_total']:
        criterios_atendidos += 1
    if analysis['equipes_recursos']['estrutura_equipe']:
        criterios_atendidos += 1
    
    analysis['score_geral'] = round((criterios_atendidos / total_criterios) * 100, 1)
    
    return analysis

def analyze_commercial_proposal_detailed(proposal_text, company_name, cnpj):
    """Análise comercial detalhada e aprofundada"""
    analysis = {
        'empresa': company_name,
        'cnpj': cnpj,
        'tabela_servicos': {
            'presente': False,
            'itens': [],
            'estrutura': '',
            'detalhamento': []
        },
        'composicao_custo': {
            'mao_obra': {'valor': '', 'percentual': '', 'detalhes': []},
            'materiais': {'valor': '', 'percentual': '', 'detalhes': []},
            'equipamentos': {'valor': '', 'percentual': '', 'detalhes': []},
            'bdi': {'valor': '', 'percentual': '', 'detalhes': []}
        },
        'preco_total': '',
        'condicoes_pagamento': {
            'forma': '',
            'parcelas': [],
            'prazos': [],
            'adiantamento': ''
        },
        'garantia': {
            'prazo': '',
            'cobertura': '',
            'termos': []
        },
        'treinamento': {
            'oferecido': False,
            'detalhes': [],
            'duracao': '',
            'local': ''
        },
        'seguros': {
            'tipos': [],
            'coberturas': [],
            'valores': []
        },
        'outras_informacoes': {
            'validade_proposta': '',
            'observacoes': [],
            'condicoes_especiais': []
        }
    }
    
    # Se é Excel, processar diferente
    if "ARQUIVO EXCEL:" in proposal_text:
        return analyze_excel_commercial_detailed(proposal_text, company_name, cnpj)
    
    # Extrair CNPJ se não fornecido
    if not analysis['cnpj']:
        cnpj_patterns = [
            r'(\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2})',
            r'CNPJ[:\s]*(\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2})'
        ]
        
        for pattern in cnpj_patterns:
            cnpj_match = re.search(pattern, proposal_text)
            if cnpj_match:
                analysis['cnpj'] = cnpj_match.group(1)
                break
    
    lines = proposal_text.split('\n')
    
    # Análise de Tabela de Serviços
    servico_keywords = ['serviço', 'item', 'atividade', 'tarefa', 'trabalho']
    servico_sections = []
    
    for line in lines:
        if any(keyword in line.lower() for keyword in servico_keywords) and len(line.strip()) > 20:
            servico_sections.append(line.strip())
    
    if servico_sections:
        analysis['tabela_servicos']['presente'] = True
        analysis['tabela_servicos']['itens'] = servico_sections[:10]
        analysis['tabela_servicos']['estrutura'] = 'Tabela de serviços identificada'
        
        # Extrair detalhamento
        for section in servico_sections:
            if any(char.isdigit() for char in section) and ('R$' in section or 'valor' in section.lower()):
                analysis['tabela_servicos']['detalhamento'].append(section[:150])
    
    # Análise de Preços
    price_patterns = [
        r'R\$\s*[\d.,]+',
        r'valor\s*total[:\s]*R\$\s*[\d.,]+',
        r'preço[:\s]*R\$\s*[\d.,]+',
        r'[\d.,]+\s*reais'
    ]
    
    prices_found = []
    for pattern in price_patterns:
        matches = re.findall(pattern, proposal_text, re.IGNORECASE)
        prices_found.extend(matches)
    
    if prices_found:
        # Converter para comparação e pegar o maior
        prices_with_values = []
        for price in prices_found:
            clean_price = re.sub(r'[^\d,.]', '', price)
            try:
                if ',' in clean_price and '.' in clean_price:
                    clean_price = clean_price.replace('.', '').replace(',', '.')
                elif ',' in clean_price:
                    clean_price = clean_price.replace(',', '.')
                float_value = float(clean_price)
                if float_value > 1000:  # Filtrar valores muito pequenos
                    prices_with_values.append((price, float_value))
            except:
                continue
        
        if prices_with_values:
            analysis['preco_total'] = max(prices_with_values, key=lambda x: x[1])[0]
    
    # Análise de Composição de Custo
    # Mão de Obra
    mao_obra_patterns = [
        r'mão\s*de\s*obra[:\s]*R\$\s*[\d.,]+',
        r'pessoal[:\s]*R\$\s*[\d.,]+',
        r'salário[:\s]*R\$\s*[\d.,]+'
    ]
    
    for pattern in mao_obra_patterns:
        match = re.search(pattern, proposal_text, re.IGNORECASE)
        if match:
            analysis['composicao_custo']['mao_obra']['valor'] = match.group(0)
            break
    
    # Materiais
    material_patterns = [
        r'materiais?[:\s]*R\$\s*[\d.,]+',
        r'insumos?[:\s]*R\$\s*[\d.,]+',
        r'produtos?[:\s]*R\$\s*[\d.,]+'
    ]
    
    for pattern in material_patterns:
        match = re.search(pattern, proposal_text, re.IGNORECASE)
        if match:
            analysis['composicao_custo']['materiais']['valor'] = match.group(0)
            break
    
    # Equipamentos
    equip_patterns = [
        r'equipamentos?[:\s]*R\$\s*[\d.,]+',
        r'máquinas?[:\s]*R\$\s*[\d.,]+',
        r'ferramentas?[:\s]*R\$\s*[\d.,]+'
    ]
    
    for pattern in equip_patterns:
        match = re.search(pattern, proposal_text, re.IGNORECASE)
        if match:
            analysis['composicao_custo']['equipamentos']['valor'] = match.group(0)
            break
    
    # BDI
    bdi_patterns = [
        r'bdi[:\s]*(\d+(?:[.,]\d+)?)\s*%',
        r'benefícios?\s*e\s*despesas?\s*indiretas?[:\s]*(\d+(?:[.,]\d+)?)\s*%',
        r'(\d+(?:[.,]\d+)?)\s*%\s*bdi'
    ]
    
    for pattern in bdi_patterns:
        match = re.search(pattern, proposal_text, re.IGNORECASE)
        if match:
            analysis['composicao_custo']['bdi']['percentual'] = match.group(1) + '%'
            break
    
    # Análise de Condições de Pagamento
    pagamento_keywords = ['pagamento', 'parcela', 'prazo', 'adiantamento', 'entrada']
    pagamento_sections = []
    
    for line in lines:
        if any(keyword in line.lower() for keyword in pagamento_keywords) and len(line.strip()) > 20:
            pagamento_sections.append(line.strip())
    
    if pagamento_sections:
        analysis['condicoes_pagamento']['forma'] = pagamento_sections[0][:100]
        
        # Extrair parcelas
        parcela_patterns = [r'(\d+)\s*parcelas?', r'(\d+)x', r'(\d+)\s*vezes']
        for section in pagamento_sections:
            for pattern in parcela_patterns:
                matches = re.findall(pattern, section.lower())
                if matches:
                    analysis['condicoes_pagamento']['parcelas'].extend(matches)
        
        # Extrair prazos
        prazo_patterns = [r'(\d+)\s*dias?', r'(\d+)\s*meses?']
        for section in pagamento_sections:
            for pattern in prazo_patterns:
                matches = re.findall(pattern, section.lower())
                analysis['condicoes_pagamento']['prazos'].extend(matches[:2])
        
        # Extrair adiantamento
        adiant_patterns = [r'adiantamento[:\s]*(\d+(?:[.,]\d+)?)\s*%', r'entrada[:\s]*(\d+(?:[.,]\d+)?)\s*%']
        for section in pagamento_sections:
            for pattern in adiant_patterns:
                match = re.search(pattern, section.lower())
                if match:
                    analysis['condicoes_pagamento']['adiantamento'] = match.group(1) + '%'
                    break
    
    # Análise de Garantia
    garantia_keywords = ['garantia', 'warranty', 'cobertura', 'proteção']
    garantia_sections = []
    
    for line in lines:
        if any(keyword in line.lower() for keyword in garantia_keywords) and len(line.strip()) > 20:
            garantia_sections.append(line.strip())
    
    if garantia_sections:
        # Extrair prazo de garantia
        prazo_patterns = [r'(\d+)\s*anos?', r'(\d+)\s*meses?', r'(\d+)\s*dias?']
        for section in garantia_sections:
            for pattern in prazo_patterns:
                matches = re.findall(pattern, section.lower())
                if matches:
                    analysis['garantia']['prazo'] = f"{matches[0]} {pattern.split('s')[0].split('\\')[1]}"
                    break
        
        analysis['garantia']['termos'] = garantia_sections[:3]
        analysis['garantia']['cobertura'] = 'Garantia oferecida'
    
    # Análise de Treinamento
    treinamento_keywords = ['treinamento', 'capacitação', 'curso', 'instrução', 'qualificação']
    treinamento_sections = []
    
    for line in lines:
        if any(keyword in line.lower() for keyword in treinamento_keywords) and len(line.strip()) > 20:
            treinamento_sections.append(line.strip())
    
    if treinamento_sections:
        analysis['treinamento']['oferecido'] = True
        analysis['treinamento']['detalhes'] = treinamento_sections[:3]
        
        # Extrair duração
        duracao_patterns = [r'(\d+)\s*horas?', r'(\d+)\s*dias?', r'(\d+)\s*semanas?']
        for section in treinamento_sections:
            for pattern in duracao_patterns:
                matches = re.findall(pattern, section.lower())
                if matches:
                    analysis['treinamento']['duracao'] = f"{matches[0]} {pattern.split('s')[0].split('\\')[1]}"
                    break
    
    # Análise de Seguros
    seguro_keywords = ['seguro', 'apólice', 'cobertura', 'proteção', 'responsabilidade civil']
    seguro_sections = []
    
    for line in lines:
        if any(keyword in line.lower() for keyword in seguro_keywords) and len(line.strip()) > 20:
            seguro_sections.append(line.strip())
    
    if seguro_sections:
        analysis['seguros']['tipos'] = seguro_sections[:4]
        
        # Extrair coberturas
        cobertura_keywords = ['cobertura', 'proteção', 'indenização']
        for section in seguro_sections:
            if any(keyword in section.lower() for keyword in cobertura_keywords):
                analysis['seguros']['coberturas'].append(section[:100])
        
        # Extrair valores
        valor_patterns = [r'R\$\s*[\d.,]+']
        for section in seguro_sections:
            matches = re.findall(valor_patterns[0], section)
            analysis['seguros']['valores'].extend(matches[:2])
    
    # Outras Informações
    # Validade da proposta
    validade_patterns = [
        r'válida?\s*por\s*(\d+)\s*dias?',
        r'validade[:\s]*(\d+)\s*dias?',
        r'proposta\s*válida\s*até'
    ]
    
    for pattern in validade_patterns:
        match = re.search(pattern, proposal_text, re.IGNORECASE)
        if match:
            if 'até' not in pattern:
                analysis['outras_informacoes']['validade_proposta'] = f"{match.group(1)} dias"
            else:
                analysis['outras_informacoes']['validade_proposta'] = 'Data específica mencionada'
            break
    
    # Observações
    obs_keywords = ['observação', 'nota', 'importante', 'atenção', 'obs:']
    for line in lines:
        if any(keyword in line.lower() for keyword in obs_keywords) and len(line.strip()) > 25:
            analysis['outras_informacoes']['observacoes'].append(line.strip()[:150])
    
    return analysis

def analyze_excel_commercial_detailed(excel_text, company_name, cnpj):
    """Análise detalhada de dados comerciais do Excel"""
    analysis = {
        'empresa': company_name,
        'cnpj': cnpj,
        'tabela_servicos': {
            'presente': False,
            'itens': [],
            'estrutura': '',
            'detalhamento': []
        },
        'composicao_custo': {
            'mao_obra': {'valor': '', 'percentual': '', 'detalhes': []},
            'materiais': {'valor': '', 'percentual': '', 'detalhes': []},
            'equipamentos': {'valor': '', 'percentual': '', 'detalhes': []},
            'bdi': {'valor': '', 'percentual': '', 'detalhes': []}
        },
        'preco_total': '',
        'condicoes_pagamento': {
            'forma': '',
            'parcelas': [],
            'prazos': [],
            'adiantamento': ''
        },
        'garantia': {
            'prazo': '',
            'cobertura': '',
            'termos': []
        },
        'treinamento': {
            'oferecido': False,
            'detalhes': [],
            'duracao': '',
            'local': ''
        },
        'seguros': {
            'tipos': [],
            'coberturas': [],
            'valores': []
        },
        'outras_informacoes': {
            'validade_proposta': '',
            'observacoes': [],
            'condicoes_especiais': []
        }
    }
    
    lines = excel_text.split('\n')
    
    # Processar abas específicas
    current_aba = ""
    
    for line in lines:
        if "=== ABA:" in line:
            current_aba = line.replace("=== ABA:", "").strip()
            continue
        
        if not line.strip():
            continue
        
        # Processar aba de Itens de Serviços
        if "Itens Serviços" in current_aba or "ITENS" in current_aba.upper():
            analysis['tabela_servicos']['presente'] = True
            
            # Extrair itens de serviço
            if "Linha" in line and "|" in line:
                parts = line.split("|")
                if len(parts) > 2:
                    item_text = " | ".join(parts[1:4])  # Pegar primeiras colunas
                    if any(char.isdigit() for char in item_text):
                        analysis['tabela_servicos']['itens'].append(item_text.strip())
            
            # Extrair preços da tabela de serviços
            price_matches = re.findall(r'[\d.,]+', line)
            for match in price_matches:
                try:
                    if '.' in match and ',' in match:
                        clean_value = match.replace('.', '').replace(',', '.')
                    elif ',' in match:
                        clean_value = match.replace(',', '.')
                    else:
                        clean_value = match
                    
                    float_value = float(clean_value)
                    if float_value > 100:  # Filtrar valores pequenos
                        if not analysis['preco_total'] or float_value > float(re.sub(r'[^\d.]', '', analysis['preco_total'] or '0')):
                            analysis['preco_total'] = f"R$ {match}"
                except:
                    continue
        
        # Processar aba de BDI
        elif "BDI" in current_aba:
            # Extrair percentual de BDI
            bdi_matches = re.findall(r'(\d+(?:[,.]?\d*))%?', line)
            for match in bdi_matches:
                try:
                    bdi_val = float(match.replace(',', '.'))
                    if 5 <= bdi_val <= 50:  # Range típico de BDI
                        analysis['composicao_custo']['bdi']['percentual'] = f"{bdi_val}%"
                        break
                except:
                    continue
            
            # Extrair detalhes do BDI
            if len(line.strip()) > 30 and any(char.isalpha() for char in line):
                analysis['composicao_custo']['bdi']['detalhes'].append(line.strip()[:100])
        
        # Processar aba de Composição de Custo
        elif "Comp. Custo" in current_aba or "GLOBAL" in current_aba:
            # Extrair composição por categoria
            if "mão de obra" in line.lower() or "pessoal" in line.lower():
                valores = re.findall(r'[\d.,]+', line)
                if valores:
                    analysis['composicao_custo']['mao_obra']['valor'] = f"R$ {valores[-1]}"
                    analysis['composicao_custo']['mao_obra']['detalhes'].append(line.strip()[:100])
            
            elif "material" in line.lower() or "insumo" in line.lower():
                valores = re.findall(r'[\d.,]+', line)
                if valores:
                    analysis['composicao_custo']['materiais']['valor'] = f"R$ {valores[-1]}"
                    analysis['composicao_custo']['materiais']['detalhes'].append(line.strip()[:100])
            
            elif "equipamento" in line.lower() or "máquina" in line.lower():
                valores = re.findall(r'[\d.,]+', line)
                if valores:
                    analysis['composicao_custo']['equipamentos']['valor'] = f"R$ {valores[-1]}"
                    analysis['composicao_custo']['equipamentos']['detalhes'].append(line.strip()[:100])
        
        # Processar aba CARTA (informações gerais)
        elif "CARTA" in current_aba:
            # Extrair condições de pagamento
            if "pagamento" in line.lower() or "parcela" in line.lower():
                analysis['condicoes_pagamento']['forma'] = line.strip()[:150]
                
                # Extrair número de parcelas
                parcela_matches = re.findall(r'(\d+)\s*parcelas?', line.lower())
                if parcela_matches:
                    analysis['condicoes_pagamento']['parcelas'].extend(parcela_matches)
            
            # Extrair garantia
            if "garantia" in line.lower():
                analysis['garantia']['termos'].append(line.strip()[:100])
                
                # Extrair prazo de garantia
                prazo_matches = re.findall(r'(\d+)\s*(?:anos?|meses?)', line.lower())
                if prazo_matches:
                    analysis['garantia']['prazo'] = f"{prazo_matches[0]} anos/meses"
            
            # Extrair treinamento
            if "treinamento" in line.lower() or "capacitação" in line.lower():
                analysis['treinamento']['oferecido'] = True
                analysis['treinamento']['detalhes'].append(line.strip()[:100])
            
            # Extrair seguros
            if "seguro" in line.lower() or "apólice" in line.lower():
                analysis['seguros']['tipos'].append(line.strip()[:100])
            
            # Extrair validade
            if "válida" in line.lower() or "validade" in line.lower():
                analysis['outras_informacoes']['validade_proposta'] = line.strip()[:100]
    
    # Estruturar informações da tabela de serviços
    if analysis['tabela_servicos']['presente']:
        analysis['tabela_servicos']['estrutura'] = f"Tabela com {len(analysis['tabela_servicos']['itens'])} itens identificados"
        analysis['tabela_servicos']['detalhamento'] = analysis['tabela_servicos']['itens'][:5]
    
    return analysis

def generate_comparative_report(project_name, project_description, technical_analyses, commercial_analyses):
    """Gera relatório comparativo detalhado"""
    current_time = datetime.now().strftime("%d/%m/%Y %H:%M")
    
    # Calcular rankings técnicos
    tech_ranking = []
    for analysis in technical_analyses:
        score = analysis.get('score_geral', 0)
        tech_ranking.append((analysis['empresa'], score))
    
    tech_ranking.sort(key=lambda x: x[1], reverse=True)
    
    # Calcular rankings comerciais
    comm_ranking = []
    for analysis in commercial_analyses:
        if analysis.get('preco_total'):
            valor_str = re.sub(r'[^\d,.]', '', analysis['preco_total'])
            try:
                if ',' in valor_str and '.' in valor_str:
                    valor_str = valor_str.replace('.', '').replace(',', '.')
                elif ',' in valor_str:
                    valor_str = valor_str.replace(',', '.')
                valor_num = float(valor_str)
                comm_ranking.append((analysis['empresa'], analysis['preco_total'], valor_num))
            except:
                comm_ranking.append((analysis['empresa'], analysis['preco_total'], 0))
    
    comm_ranking.sort(key=lambda x: x[2])
    
    # Gerar relatório
    report = f"""# 📊 ANÁLISE COMPARATIVA DE PROPOSTAS - {project_name.upper()}

**Data:** {current_time}
**Projeto:** {project_name}
"""
    
    if project_description:
        report += f"**Descrição:** {project_description}\n"
    
    report += f"""
**Propostas Analisadas:** {len(technical_analyses)} técnicas, {len(commercial_analyses)} comerciais

---

## 🏆 RESUMO EXECUTIVO

### Rankings Gerais
"""
    
    # Ranking técnico
    if tech_ranking:
        report += "\n**🔧 Ranking Técnico:**\n"
        for i, (empresa, score) in enumerate(tech_ranking, 1):
            emoji = "🥇" if i == 1 else "🥈" if i == 2 else "🥉" if i == 3 else "📊"
            report += f"{emoji} **{i}º:** {empresa} - {score}%\n"
    
    # Ranking comercial
    if comm_ranking:
        report += "\n**💰 Ranking Comercial (Menor Preço):**\n"
        for i, (empresa, preco, _) in enumerate(comm_ranking, 1):
            emoji = "🥇" if i == 1 else "🥈" if i == 2 else "🥉" if i == 3 else "📊"
            report += f"{emoji} **{i}º:** {empresa} - {preco}\n"
    
    report += """

---

## 🔧 ANÁLISE TÉCNICA COMPARATIVA

### Matriz de Comparação Técnica

| Empresa | Metodologia | Mão de Obra | Equipamentos | Materiais | Cronograma | Score |
|---------|-------------|-------------|--------------|-----------|------------|-------|"""
    
    for analysis in technical_analyses:
        empresa = analysis['empresa']
        metodologia = "✅" if analysis['metodologia_execucao']['descricao'] else "❌"
        mao_obra = "✅" if analysis['histograma_mao_obra']['presente'] else "❌"
        equipamentos = "✅" if analysis['histograma_equipamentos']['presente'] else "❌"
        materiais = "✅" if analysis['lista_materiais']['presente'] else "❌"
        cronograma = "✅" if analysis['prazo_cronograma']['prazo_total'] else "❌"
        score = f"{analysis['score_geral']}%"
        
        report += f"\n| {empresa} | {metodologia} | {mao_obra} | {equipamentos} | {materiais} | {cronograma} | {score} |"
    
    # Análise detalhada por empresa
    for analysis in technical_analyses:
        empresa = analysis['empresa']
        cnpj = analysis.get('cnpj', 'Não identificado')
        
        report += f"""

### 🏢 {empresa}
**CNPJ:** {cnpj}
**Score Geral:** {analysis['score_geral']}%

#### 📋 Metodologia de Execução
"""
        
        if analysis['metodologia_execucao']['descricao']:
            report += f"**Descrição:** {analysis['metodologia_execucao']['descricao'][:300]}...\n"
            
            if analysis['metodologia_execucao']['fases']:
                report += f"**Fases Identificadas:** {', '.join(analysis['metodologia_execucao']['fases'][:5])}\n"
            
            if analysis['metodologia_execucao']['ferramentas']:
                report += f"**Ferramentas/Tecnologias:** {len(analysis['metodologia_execucao']['ferramentas'])} itens identificados\n"
        else:
            report += "❌ Metodologia não apresentada ou insuficiente.\n"
        
        report += "\n#### 👥 Histograma de Mão de Obra\n"
        if analysis['histograma_mao_obra']['presente']:
            report += f"✅ **Presente** - {len(analysis['histograma_mao_obra']['detalhes'])} informações identificadas\n"
            
            if analysis['histograma_mao_obra']['quantidade_total']:
                report += f"**Quantidade Total:** {analysis['histograma_mao_obra']['quantidade_total']}\n"
            
            if analysis['histograma_mao_obra']['qualificacoes']:
                report += f"**Qualificações:** {len(analysis['histograma_mao_obra']['qualificacoes'])} perfis identificados\n"
        else:
            report += "❌ Histograma de mão de obra não apresentado.\n"
        
        report += "\n#### 🔧 Histograma de Equipamentos\n"
        if analysis['histograma_equipamentos']['presente']:
            report += f"✅ **Presente** - {len(analysis['histograma_equipamentos']['equipamentos'])} equipamentos identificados\n"
            
            if analysis['histograma_equipamentos']['tecnologias']:
                report += f"**Tecnologias:** {len(analysis['histograma_equipamentos']['tecnologias'])} tecnologias avançadas\n"
        else:
            report += "❌ Histograma de equipamentos não apresentado.\n"
        
        report += "\n#### 📦 Lista de Materiais\n"
        if analysis['lista_materiais']['presente']:
            report += f"✅ **Presente** - {len(analysis['lista_materiais']['materiais'])} materiais identificados\n"
            
            if analysis['lista_materiais']['quantidades']:
                report += f"**Quantidades:** {len(analysis['lista_materiais']['quantidades'])} especificações de quantidade\n"
            
            if analysis['lista_materiais']['especificacoes']:
                report += f"**Especificações:** {len(analysis['lista_materiais']['especificacoes'])} especificações técnicas\n"
        else:
            report += "❌ Lista de materiais não apresentada.\n"
        
        report += "\n#### 📋 Obrigações e Responsabilidades\n"
        if analysis['obrigacoes']['principais']:
            report += f"✅ **Definidas** - {len(analysis['obrigacoes']['principais'])} obrigações principais\n"
            
            if analysis['obrigacoes']['responsabilidades']:
                report += f"**Responsabilidades:** {len(analysis['obrigacoes']['responsabilidades'])} responsabilidades específicas\n"
        else:
            report += "❌ Obrigações não especificadas claramente.\n"
        
        report += "\n#### 🏗️ Informações sobre Canteiro\n"
        if analysis['canteiro']['informacoes']:
            report += f"✅ **Apresentadas** - {len(analysis['canteiro']['informacoes'])} informações sobre canteiro\n"
            
            if analysis['canteiro']['logistica']:
                report += f"**Logística:** {analysis['canteiro']['logistica']}\n"
            
            if analysis['canteiro']['infraestrutura']:
                report += f"**Infraestrutura:** {analysis['canteiro']['infraestrutura']}\n"
        else:
            report += "❌ Informações sobre canteiro não apresentadas.\n"
        
        report += "\n#### ❌ Exclusões\n"
        if analysis['exclusoes']['itens_excluidos']:
            report += f"✅ **Especificadas** - {len(analysis['exclusoes']['itens_excluidos'])} exclusões identificadas\n"
            
            if analysis['exclusoes']['limitacoes']:
                report += f"**Limitações:** {len(analysis['exclusoes']['limitacoes'])} limitações específicas\n"
        else:
            report += "⚠️ Exclusões não especificadas (pode gerar problemas futuros).\n"
        
        report += "\n#### ⏰ Prazo e Cronograma\n"
        if analysis['prazo_cronograma']['prazo_total']:
            report += f"✅ **Definido** - Prazo: {analysis['prazo_cronograma']['prazo_total']}\n"
            
            if analysis['prazo_cronograma']['marcos_principais']:
                report += f"**Marcos:** {len(analysis['prazo_cronograma']['marcos_principais'])} marcos principais\n"
            
            if analysis['prazo_cronograma']['fases_cronograma']:
                report += f"**Fases:** {len(analysis['prazo_cronograma']['fases_cronograma'])} fases do cronograma\n"
            
            report += f"**Viabilidade:** {analysis['prazo_cronograma']['viabilidade']}\n"
        else:
            report += "❌ Prazo e cronograma não definidos adequadamente.\n"
        
        report += "\n#### 👨‍💼 Equipes e Recursos\n"
        if analysis['equipes_recursos']['estrutura_equipe']:
            report += f"✅ **Estrutura Definida** - {len(analysis['equipes_recursos']['estrutura_equipe'])} informações sobre equipe\n"
            
            if analysis['equipes_recursos']['coordenador']:
                report += f"**Coordenador:** {analysis['equipes_recursos']['coordenador']}\n"
            
            if analysis['equipes_recursos']['especialistas']:
                report += f"**Especialistas:** {len(analysis['equipes_recursos']['especialistas'])} especialistas identificados\n"
        else:
            report += "❌ Estrutura de equipe não apresentada.\n"
        
        report += "\n#### ✅ Pontos Fortes\n"
        for ponto in analysis['pontos_fortes'][:5]:
            report += f"• {ponto}\n"
        
        report += "\n#### ⚠️ Pontos de Atenção\n"
        for ponto in analysis['pontos_fracos'][:5]:
            report += f"• {ponto}\n"
    
    report += """

---

## 💰 ANÁLISE COMERCIAL COMPARATIVA

### Resumo de Preços

| Empresa | Preço Total | BDI | Condições Pagamento | Garantia |
|---------|-------------|-----|-------------------|----------|"""
    
    for analysis in commercial_analyses:
        empresa = analysis['empresa']
        preco = analysis.get('preco_total', 'Não informado')
        bdi = analysis['composicao_custo']['bdi'].get('percentual', 'Não informado')
        pagamento = analysis['condicoes_pagamento'].get('forma', 'Não informado')[:30] + "..."
        garantia = analysis['garantia'].get('prazo', 'Não informado')
        
        report += f"\n| {empresa} | {preco} | {bdi} | {pagamento} | {garantia} |"
    
    # Análise comercial detalhada por empresa
    for analysis in commercial_analyses:
        empresa = analysis['empresa']
        cnpj = analysis.get('cnpj', 'Não identificado')
        
        report += f"""

### 🏢 {empresa} - Análise Comercial
**CNPJ:** {cnpj}

#### 💵 Preço e Composição de Custos
**Preço Total:** {analysis.get('preco_total', 'Não informado')}

**Composição de Custos:**
"""
        
        # Mão de obra
        mao_obra = analysis['composicao_custo']['mao_obra']
        if mao_obra['valor']:
            report += f"• **Mão de Obra:** {mao_obra['valor']}"
            if mao_obra['percentual']:
                report += f" ({mao_obra['percentual']})"
            report += "\n"
        
        # Materiais
        materiais = analysis['composicao_custo']['materiais']
        if materiais['valor']:
            report += f"• **Materiais:** {materiais['valor']}"
            if materiais['percentual']:
                report += f" ({materiais['percentual']})"
            report += "\n"
        
        # Equipamentos
        equipamentos = analysis['composicao_custo']['equipamentos']
        if equipamentos['valor']:
            report += f"• **Equipamentos:** {equipamentos['valor']}"
            if equipamentos['percentual']:
                report += f" ({equipamentos['percentual']})"
            report += "\n"
        
        # BDI
        bdi = analysis['composicao_custo']['bdi']
        if bdi['percentual']:
            report += f"• **BDI:** {bdi['percentual']}"
            if bdi['valor']:
                report += f" ({bdi['valor']})"
            report += "\n"
        
        report += "\n#### 📋 Tabela de Serviços\n"
        if analysis['tabela_servicos']['presente']:
            report += f"✅ **Presente** - {analysis['tabela_servicos']['estrutura']}\n"
            report += f"**Itens Identificados:** {len(analysis['tabela_servicos']['itens'])}\n"
        else:
            report += "❌ Tabela de serviços não identificada.\n"
        
        report += "\n#### 💳 Condições de Pagamento\n"
        pagamento = analysis['condicoes_pagamento']
        if pagamento['forma']:
            report += f"**Forma:** {pagamento['forma']}\n"
        
        if pagamento['parcelas']:
            report += f"**Parcelas:** {', '.join(pagamento['parcelas'])}\n"
        
        if pagamento['prazos']:
            report += f"**Prazos:** {', '.join(pagamento['prazos'])}\n"
        
        if pagamento['adiantamento']:
            report += f"**Adiantamento:** {pagamento['adiantamento']}\n"
        
        if not any([pagamento['forma'], pagamento['parcelas'], pagamento['prazos']]):
            report += "❌ Condições de pagamento não especificadas.\n"
        
        report += "\n#### 🛡️ Garantia\n"
        garantia = analysis['garantia']
        if garantia['prazo']:
            report += f"**Prazo:** {garantia['prazo']}\n"
        
        if garantia['cobertura']:
            report += f"**Cobertura:** {garantia['cobertura']}\n"
        
        if garantia['termos']:
            report += f"**Termos:** {len(garantia['termos'])} condições especificadas\n"
        
        if not any([garantia['prazo'], garantia['cobertura'], garantia['termos']]):
            report += "❌ Garantia não especificada.\n"
        
        report += "\n#### 🎓 Treinamento\n"
        treinamento = analysis['treinamento']
        if treinamento['oferecido']:
            report += "✅ **Oferecido**\n"
            
            if treinamento['duracao']:
                report += f"**Duração:** {treinamento['duracao']}\n"
            
            if treinamento['detalhes']:
                report += f"**Detalhes:** {len(treinamento['detalhes'])} informações sobre treinamento\n"
        else:
            report += "❌ Treinamento não oferecido ou não especificado.\n"
        
        report += "\n#### 🛡️ Seguros\n"
        seguros = analysis['seguros']
        if seguros['tipos']:
            report += f"✅ **Oferecidos** - {len(seguros['tipos'])} tipos de seguro\n"
            
            if seguros['coberturas']:
                report += f"**Coberturas:** {len(seguros['coberturas'])} coberturas especificadas\n"
            
            if seguros['valores']:
                report += f"**Valores:** {len(seguros['valores'])} valores informados\n"
        else:
            report += "❌ Seguros não especificados.\n"
        
        report += "\n#### 📄 Outras Informações\n"
        outras = analysis['outras_informacoes']
        if outras['validade_proposta']:
            report += f"**Validade da Proposta:** {outras['validade_proposta']}\n"
        
        if outras['observacoes']:
            report += f"**Observações:** {len(outras['observacoes'])} observações importantes\n"
        
        if outras['condicoes_especiais']:
            report += f"**Condições Especiais:** {len(outras['condicoes_especiais'])} condições\n"
    
    report += """

---

## 🎯 CONCLUSÕES E RECOMENDAÇÕES

### Análise Comparativa Final
"""
    
    # Melhor proposta técnica
    if tech_ranking:
        melhor_tecnica = tech_ranking[0]
        report += f"""
**🏆 Melhor Proposta Técnica:** {melhor_tecnica[0]} ({melhor_tecnica[1]}%)

**Justificativa:** Esta proposta apresentou o maior score técnico, demonstrando melhor aderência aos critérios de metodologia, recursos, cronograma e estrutura organizacional.
"""
    
    # Melhor proposta comercial
    if comm_ranking:
        melhor_comercial = comm_ranking[0]
        report += f"""
**💰 Melhor Proposta Comercial:** {melhor_comercial[0]} ({melhor_comercial[1]})

**Justificativa:** Esta proposta apresentou o menor preço total, oferecendo melhor vantagem comercial.
"""
    
    # Análise de custo-benefício
    if tech_ranking and comm_ranking:
        report += "\n### 📊 Análise de Custo-Benefício\n\n"
        
        # Criar tabela de custo-benefício
        report += "| Empresa | Posição Técnica | Posição Comercial | Custo-Benefício |\n"
        report += "|---------|-----------------|-------------------|------------------|\n"
        
        for tech_pos, (tech_empresa, tech_score) in enumerate(tech_ranking, 1):
            # Encontrar posição comercial
            comm_pos = "N/A"
            for c_pos, (comm_empresa, _, _) in enumerate(comm_ranking, 1):
                if comm_empresa == tech_empresa:
                    comm_pos = c_pos
                    break
            
            # Calcular índice de custo-benefício (quanto menor, melhor)
            if comm_pos != "N/A":
                custo_beneficio = (tech_pos + comm_pos) / 2
                if custo_beneficio <= 1.5:
                    cb_status = "🥇 Excelente"
                elif custo_beneficio <= 2.5:
                    cb_status = "🥈 Bom"
                elif custo_beneficio <= 3.5:
                    cb_status = "🥉 Regular"
                else:
                    cb_status = "📊 Inferior"
            else:
                cb_status = "❌ Sem dados comerciais"
            
            report += f"| {tech_empresa} | {tech_pos}º | {comm_pos}º | {cb_status} |\n"
    
    # Recomendações finais
    report += """

### 🎯 Recomendações Finais

#### Para Tomada de Decisão:
1. **Análise Técnica:** Considere a proposta com maior score técnico para garantir qualidade de execução.
2. **Análise Comercial:** Avalie não apenas o menor preço, mas também as condições de pagamento e garantias oferecidas.
3. **Custo-Benefício:** Busque o equilíbrio entre qualidade técnica e vantagem comercial.

#### Próximos Passos Sugeridos:
1. **Esclarecimentos:** Solicite esclarecimentos para propostas com informações incompletas.
2. **Negociação:** Considere negociar condições com as propostas melhor classificadas.
3. **Verificação:** Confirme referências e capacidade técnica das empresas.

#### Pontos de Atenção:
• Propostas com exclusões não especificadas podem gerar custos adicionais.
• Cronogramas muito agressivos podem comprometer a qualidade.
• Preços muito baixos podem indicar subdimensionamento ou qualidade inferior.
"""
    
    report += f"""

---

*Relatório gerado pelo Proposal Analyzer Pro - Análise Comparativa*
*Data: {current_time}*
*Propostas analisadas: {len(technical_analyses)} técnicas, {len(commercial_analyses)} comerciais*
"""
    
    return report

@app.route('/')
def index():
    return render_template_string(HTML_TEMPLATE)

@app.route('/analyze', methods=['POST'])
def analyze_proposals():
    try:
        # Obter dados básicos
        project_name = request.form.get('projectName')
        project_description = request.form.get('projectDescription', '')
        
        if not project_name:
            return jsonify({'success': False, 'error': 'Nome do projeto é obrigatório.'})
        
        # Processar propostas técnicas
        technical_analyses = []
        tech_companies = request.form.getlist('techCompany[]')
        tech_files = request.files.getlist('techFile[]')
        
        for i, (company, file) in enumerate(zip(tech_companies, tech_files)):
            if company and file and file.filename:
                file_path = os.path.join(app.config['UPLOAD_FOLDER'], f"tech_{i}_{file.filename}")
                file.save(file_path)
                
                content = extract_text_from_file(file_path)
                tech_analysis = analyze_technical_proposal_detailed(content, company)
                technical_analyses.append(tech_analysis)
                
                # Limpar arquivo temporário
                os.remove(file_path)
                gc.collect()
        
        # Processar propostas comerciais
        commercial_analyses = []
        comm_companies = request.form.getlist('commCompany[]')
        comm_cnpjs = request.form.getlist('commCnpj[]')
        comm_files = request.files.getlist('commFile[]')
        
        for i, (company, cnpj, file) in enumerate(zip(comm_companies, comm_cnpjs, comm_files)):
            if company and file and file.filename:
                file_path = os.path.join(app.config['UPLOAD_FOLDER'], f"comm_{i}_{file.filename}")
                file.save(file_path)
                
                content = extract_text_from_file(file_path)
                comm_analysis = analyze_commercial_proposal_detailed(content, company, cnpj)
                commercial_analyses.append(comm_analysis)
                
                # Limpar arquivo temporário
                os.remove(file_path)
                gc.collect()
        
        # Validar se há propostas suficientes
        if len(technical_analyses) < 2:
            return jsonify({'success': False, 'error': 'É necessário pelo menos 2 propostas técnicas para comparação.'})
        
        if len(commercial_analyses) < 2:
            return jsonify({'success': False, 'error': 'É necessário pelo menos 2 propostas comerciais para comparação.'})
        
        # Gerar relatório comparativo
        report = generate_comparative_report(
            project_name, project_description,
            technical_analyses, commercial_analyses
        )
        
        # Salvar relatório
        report_id = f"report_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
        report_path = os.path.join(app.config['UPLOAD_FOLDER'], f"{report_id}.md")
        
        with open(report_path, 'w', encoding='utf-8') as f:
            f.write(report)
        
        # Limpar variáveis grandes
        del report, technical_analyses, commercial_analyses
        gc.collect()
        
        return jsonify({'success': True, 'report_id': report_id})
        
    except Exception as e:
        gc.collect()
        return jsonify({'success': False, 'error': str(e)})

@app.route('/download/<report_id>/<format>')
def download_report(report_id, format):
    try:
        if format == 'markdown':
            file_path = os.path.join(app.config['UPLOAD_FOLDER'], f"{report_id}.md")
            return send_file(file_path, as_attachment=True, download_name='analise_comparativa.md')
        elif format == 'pdf':
            # Gerar PDF
            md_file_path = os.path.join(app.config['UPLOAD_FOLDER'], f"{report_id}.md")
            pdf_file_path = os.path.join(app.config['UPLOAD_FOLDER'], f"analise_comparativa.pdf")
            
            # Ler markdown
            with open(md_file_path, 'r', encoding='utf-8') as f:
                markdown_content = f.read()
            
            # Gerar PDF com reportlab
            from reportlab.lib.pagesizes import A4
            from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
            from reportlab.lib.styles import getSampleStyleSheet
            from reportlab.lib.units import inch
            
            doc = SimpleDocTemplate(pdf_file_path, pagesize=A4, topMargin=1*inch)
            styles = getSampleStyleSheet()
            story = []
            
            # Processar markdown de forma simples
            lines = markdown_content.split('\n')
            for line in lines:
                line = line.strip()
                if line.startswith('# '):
                    story.append(Paragraph(line[2:], styles['Title']))
                elif line.startswith('## '):
                    story.append(Paragraph(line[3:], styles['Heading1']))
                elif line.startswith('### '):
                    story.append(Paragraph(line[4:], styles['Heading2']))
                elif line.startswith('#### '):
                    story.append(Paragraph(line[5:], styles['Heading3']))
                elif line and not line.startswith('*') and not line.startswith('|'):
                    story.append(Paragraph(line, styles['Normal']))
                
                if len(story) % 50 == 0:  # Garbage collect periodicamente
                    gc.collect()
            
            doc.build(story)
            
            # Limpar variáveis
            del markdown_content, lines, story
            gc.collect()
            
            return send_file(pdf_file_path, as_attachment=True, download_name='analise_comparativa.pdf')
        else:
            return jsonify({'error': 'Formato não suportado'}), 400
            
    except Exception as e:
        gc.collect()
        return jsonify({'error': str(e)}), 500

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=False)
