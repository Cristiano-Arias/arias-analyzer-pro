import os
import tempfile
import zipfile
from flask import Flask, request, jsonify, send_file, render_template_string
from flask_cors import CORS
import PyPDF2
import docx
import io
import re
from datetime import datetime
from collections import defaultdict

app = Flask(__name__)
CORS(app)

# Configuração de upload
UPLOAD_FOLDER = 'uploads'
if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024  # 50MB max

# HTML da interface (mantém o mesmo)
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

def analyze_tr_content(tr_text):
    """Analisa o conteúdo do Termo de Referência e extrai informações estruturadas"""
    analysis = {
        'resumo': '',
        'objeto': '',
        'requisitos_tecnicos': [],
        'prazos': [],
        'criterios_avaliacao': [],
        'qualificacoes_exigidas': [],
        'valores_estimados': []
    }
    
    # Dividir texto em seções
    sections = tr_text.split('\n')
    current_section = ''
    
    # Padrões para identificar informações importantes
    prazo_patterns = [
        r'prazo.*?(\d+).*?(dia|mês|ano)',
        r'cronograma.*?(\d+).*?(dia|mês|ano)',
        r'entrega.*?(\d+).*?(dia|mês|ano)'
    ]
    
    valor_patterns = [
        r'R\$\s*[\d.,]+',
        r'valor.*?R\$\s*[\d.,]+',
        r'orçamento.*?R\$\s*[\d.,]+'
    ]
    
    # Extrair objeto/resumo (primeiros parágrafos significativos)
    meaningful_paragraphs = [p.strip() for p in sections if len(p.strip()) > 50]
    if meaningful_paragraphs:
        analysis['resumo'] = ' '.join(meaningful_paragraphs[:3])
        analysis['objeto'] = meaningful_paragraphs[0] if meaningful_paragraphs else ''
    
    # Buscar prazos
    for section in sections:
        for pattern in prazo_patterns:
            matches = re.findall(pattern, section.lower())
            for match in matches:
                analysis['prazos'].append(f"{match[0]} {match[1]}")
    
    # Buscar valores
    for section in sections:
        for pattern in valor_patterns:
            matches = re.findall(pattern, section)
            analysis['valores_estimados'].extend(matches)
    
    # Identificar requisitos técnicos (seções que contêm palavras-chave)
    tech_keywords = ['técnico', 'especificação', 'requisito', 'metodologia', 'equipamento', 'material']
    for section in sections:
        if any(keyword in section.lower() for keyword in tech_keywords) and len(section.strip()) > 30:
            analysis['requisitos_tecnicos'].append(section.strip())
    
    # Identificar critérios de avaliação
    eval_keywords = ['avaliação', 'critério', 'pontuação', 'peso', 'classificação']
    for section in sections:
        if any(keyword in section.lower() for keyword in eval_keywords) and len(section.strip()) > 30:
            analysis['criterios_avaliacao'].append(section.strip())
    
    # Identificar qualificações exigidas
    qual_keywords = ['qualificação', 'experiência', 'certificação', 'habilitação', 'comprovação']
    for section in sections:
        if any(keyword in section.lower() for keyword in qual_keywords) and len(section.strip()) > 30:
            analysis['qualificacoes_exigidas'].append(section.strip())
    
    return analysis

def analyze_technical_proposal(proposal_text, company_name):
    """Analisa uma proposta técnica e extrai informações estruturadas"""
    analysis = {
        'empresa': company_name,
        'metodologia': '',
        'cronograma': [],
        'equipe': [],
        'equipamentos': [],
        'materiais': [],
        'diferenciais': [],
        'pontos_fortes': [],
        'pontos_fracos': []
    }
    
    sections = proposal_text.split('\n')
    
    # Padrões para identificar diferentes seções
    metodologia_keywords = ['metodologia', 'método', 'abordagem', 'estratégia', 'processo']
    cronograma_keywords = ['cronograma', 'prazo', 'etapa', 'fase', 'período']
    equipe_keywords = ['equipe', 'profissional', 'responsável', 'coordenador', 'especialista']
    equipamento_keywords = ['equipamento', 'ferramenta', 'instrumento', 'máquina']
    material_keywords = ['material', 'insumo', 'produto', 'componente']
    
    # Extrair metodologia (parágrafos que contêm palavras-chave de metodologia)
    metodologia_sections = []
    for section in sections:
        if any(keyword in section.lower() for keyword in metodologia_keywords) and len(section.strip()) > 50:
            metodologia_sections.append(section.strip())
    
    analysis['metodologia'] = ' '.join(metodologia_sections[:2]) if metodologia_sections else 'Metodologia não claramente identificada'
    
    # Extrair cronograma
    for section in sections:
        if any(keyword in section.lower() for keyword in cronograma_keywords):
            # Buscar por padrões de tempo
            time_patterns = re.findall(r'(\d+)\s*(dia|semana|mês|ano)', section.lower())
            for time_match in time_patterns:
                analysis['cronograma'].append(f"{time_match[0]} {time_match[1]}")
    
    # Extrair equipe
    for section in sections:
        if any(keyword in section.lower() for keyword in equipe_keywords) and len(section.strip()) > 20:
            analysis['equipe'].append(section.strip())
    
    # Extrair equipamentos
    for section in sections:
        if any(keyword in section.lower() for keyword in equipamento_keywords) and len(section.strip()) > 20:
            analysis['equipamentos'].append(section.strip())
    
    # Extrair materiais
    for section in sections:
        if any(keyword in section.lower() for keyword in material_keywords) and len(section.strip()) > 20:
            analysis['materiais'].append(section.strip())
    
    # Identificar pontos fortes (seções com palavras positivas)
    positive_keywords = ['experiência', 'qualificado', 'certificado', 'especializado', 'inovador', 'eficiente']
    for section in sections:
        if any(keyword in section.lower() for keyword in positive_keywords) and len(section.strip()) > 30:
            analysis['pontos_fortes'].append(section.strip())
    
    return analysis

def analyze_commercial_proposal(proposal_text, company_name, cnpj):
    """Analisa uma proposta comercial e extrai informações estruturadas"""
    analysis = {
        'empresa': company_name,
        'cnpj': cnpj,
        'preco_total': '',
        'composicao_custos': {},
        'condicoes_pagamento': '',
        'prazos': [],
        'bdi': '',
        'observacoes': []
    }
    
    # Buscar preços
    price_patterns = [
        r'R\$\s*[\d.,]+',
        r'total.*?R\$\s*[\d.,]+',
        r'valor.*?R\$\s*[\d.,]+',
        r'preço.*?R\$\s*[\d.,]+'
    ]
    
    prices_found = []
    for pattern in price_patterns:
        matches = re.findall(pattern, proposal_text, re.IGNORECASE)
        prices_found.extend(matches)
    
    if prices_found:
        # Assumir que o maior valor é o preço total
        analysis['preco_total'] = max(prices_found, key=lambda x: float(re.sub(r'[^\d,]', '', x).replace(',', '.')))
    
    # Buscar condições de pagamento
    payment_keywords = ['pagamento', 'parcela', 'à vista', 'prazo']
    sections = proposal_text.split('\n')
    
    for section in sections:
        if any(keyword in section.lower() for keyword in payment_keywords) and len(section.strip()) > 20:
            analysis['condicoes_pagamento'] = section.strip()
            break
    
    # Buscar BDI
    bdi_patterns = [r'bdi.*?(\d+[,.]?\d*)%?', r'benefício.*?(\d+[,.]?\d*)%?']
    for pattern in bdi_patterns:
        matches = re.findall(pattern, proposal_text.lower())
        if matches:
            analysis['bdi'] = matches[0] + '%'
            break
    
    # Buscar prazos
    prazo_patterns = [
        r'prazo.*?(\d+).*?(dia|mês|ano)',
        r'entrega.*?(\d+).*?(dia|mês|ano)'
    ]
    
    for pattern in prazo_patterns:
        matches = re.findall(pattern, proposal_text.lower())
        for match in matches:
            analysis['prazos'].append(f"{match[0]} {match[1]}")
    
    return analysis

def generate_comparative_analysis(tr_analysis, technical_proposals, commercial_proposals):
    """Gera análise comparativa entre propostas e TR"""
    
    # Análise técnica comparativa
    tech_comparison = {
        'aderencia_tr': {},
        'pontos_fortes_por_empresa': {},
        'pontos_fracos_por_empresa': {},
        'recomendacoes': []
    }
    
    # Análise comercial comparativa
    comm_comparison = {
        'ranking_precos': [],
        'melhor_condicao_pagamento': '',
        'analise_bdi': {},
        'recomendacoes': []
    }
    
    # Comparar propostas técnicas
    for proposal in technical_proposals:
        empresa = proposal['empresa']
        
        # Verificar aderência aos requisitos do TR
        aderencia_score = 0
        if proposal.get('metodologia'):
            aderencia_score += 25
        if proposal.get('cronograma'):
            aderencia_score += 25
        if proposal.get('equipe'):
            aderencia_score += 25
        if proposal.get('equipamentos') or proposal.get('materiais'):
            aderencia_score += 25
        
        tech_comparison['aderencia_tr'][empresa] = f"{aderencia_score}%"
        tech_comparison['pontos_fortes_por_empresa'][empresa] = proposal.get('pontos_fortes', [])
        
        # Identificar pontos fracos baseado no que está faltando
        pontos_fracos = []
        if not proposal.get('metodologia') or 'não claramente identificada' in proposal.get('metodologia', ''):
            pontos_fracos.append("Metodologia não claramente definida")
        if not proposal.get('cronograma'):
            pontos_fracos.append("Cronograma não apresentado")
        if not proposal.get('equipe'):
            pontos_fracos.append("Equipe técnica não detalhada")
        
        tech_comparison['pontos_fracos_por_empresa'][empresa] = pontos_fracos
    
    # Comparar propostas comerciais
    precos_empresas = []
    for proposal in commercial_proposals:
        if proposal.get('preco_total'):
            # Extrair valor numérico para comparação
            valor_str = re.sub(r'[^\d,]', '', proposal['preco_total']).replace(',', '.')
            try:
                valor_num = float(valor_str)
                precos_empresas.append((proposal['empresa'], proposal['preco_total'], valor_num))
            except:
                precos_empresas.append((proposal['empresa'], proposal['preco_total'], 0))
    
    # Ordenar por preço (menor para maior)
    precos_empresas.sort(key=lambda x: x[2])
    comm_comparison['ranking_precos'] = [(empresa, preco_str) for empresa, preco_str, _ in precos_empresas]
    
    # Análise de BDI
    for proposal in commercial_proposals:
        if proposal.get('bdi'):
            comm_comparison['analise_bdi'][proposal['empresa']] = proposal['bdi']
    
    return tech_comparison, comm_comparison

def generate_enhanced_report(project_name, project_description, tr_analysis, technical_analyses, commercial_analyses, tech_comparison, comm_comparison):
    """Gera relatório aprimorado com análise de IA"""
    
    current_time = datetime.now().strftime("%d/%m/%Y %H:%M")
    
    report = f"""# 📊 RELATÓRIO DE ANÁLISE DE PROPOSTAS - {project_name.upper()}

**Data de Análise:** {current_time}
**Projeto:** {project_name}
**Descrição:** {project_description if project_description else 'Não informada'}

---

## 🎯 BLOCO 1: RESUMO DO TERMO DE REFERÊNCIA

### Objeto do Projeto
{tr_analysis.get('objeto', 'Não identificado claramente')}

### Resumo Executivo do TR
{tr_analysis.get('resumo', 'Resumo não disponível')}

### Requisitos Técnicos Principais
"""
    
    if tr_analysis.get('requisitos_tecnicos'):
        for i, req in enumerate(tr_analysis['requisitos_tecnicos'][:5], 1):
            report += f"**{i}.** {req[:200]}...\n\n"
    else:
        report += "Requisitos técnicos não claramente identificados no TR.\n\n"
    
    report += "### Prazos Estabelecidos\n"
    if tr_analysis.get('prazos'):
        for prazo in tr_analysis['prazos']:
            report += f"- {prazo}\n"
    else:
        report += "Prazos não claramente especificados no TR.\n"
    
    report += "\n### Critérios de Avaliação\n"
    if tr_analysis.get('criterios_avaliacao'):
        for criterio in tr_analysis['criterios_avaliacao'][:3]:
            report += f"- {criterio[:150]}...\n"
    else:
        report += "Critérios de avaliação não claramente definidos no TR.\n"
    
    report += f"""

---

## 🔧 BLOCO 2: EQUALIZAÇÃO DAS PROPOSTAS TÉCNICAS

### Resumo Comparativo de Aderência ao TR
"""
    
    if tech_comparison.get('aderencia_tr'):
        report += "| Empresa | Aderência ao TR | Status |\n"
        report += "|---------|-----------------|--------|\n"
        for empresa, aderencia in tech_comparison['aderencia_tr'].items():
            status = "✅ Boa" if int(aderencia.replace('%', '')) >= 75 else "⚠️ Parcial" if int(aderencia.replace('%', '')) >= 50 else "❌ Insuficiente"
            report += f"| {empresa} | {aderencia} | {status} |\n"
        report += "\n"
    
    # Análise detalhada por empresa
    for analysis in technical_analyses:
        empresa = analysis['empresa']
        report += f"""
### 📋 Análise Detalhada: {empresa}

**Metodologia Proposta:**
{analysis.get('metodologia', 'Não apresentada ou não identificada')}

**Cronograma:**
"""
        if analysis.get('cronograma'):
            for item in analysis['cronograma']:
                report += f"- {item}\n"
        else:
            report += "Cronograma não apresentado de forma clara.\n"
        
        report += "\n**Equipe Técnica:**\n"
        if analysis.get('equipe'):
            for item in analysis['equipe'][:3]:
                report += f"- {item[:100]}...\n"
        else:
            report += "Equipe técnica não detalhada.\n"
        
        report += "\n**Pontos Fortes:**\n"
        pontos_fortes = tech_comparison.get('pontos_fortes_por_empresa', {}).get(empresa, [])
        if pontos_fortes:
            for ponto in pontos_fortes[:3]:
                report += f"✅ {ponto[:100]}...\n"
        else:
            report += "Pontos fortes não claramente identificados.\n"
        
        report += "\n**Pontos de Atenção:**\n"
        pontos_fracos = tech_comparison.get('pontos_fracos_por_empresa', {}).get(empresa, [])
        if pontos_fracos:
            for ponto in pontos_fracos:
                report += f"⚠️ {ponto}\n"
        else:
            report += "Nenhum ponto de atenção identificado.\n"
        
        report += "\n---\n"
    
    report += f"""

## 💰 BLOCO 3: EQUALIZAÇÃO DAS PROPOSTAS COMERCIAIS

### Ranking de Preços
"""
    
    if comm_comparison.get('ranking_precos'):
        report += "| Posição | Empresa | Preço Total | Diferença |\n"
        report += "|---------|---------|-------------|----------|\n"
        
        precos = comm_comparison['ranking_precos']
        menor_preco = None
        
        for i, (empresa, preco) in enumerate(precos, 1):
            if i == 1:
                menor_preco = preco
                diferenca = "Menor preço"
            else:
                diferenca = "A calcular"
            
            emoji = "🥇" if i == 1 else "🥈" if i == 2 else "🥉" if i == 3 else "📊"
            report += f"| {emoji} {i}º | {empresa} | {preco} | {diferenca} |\n"
        
        report += "\n"
    
    # Análise detalhada por empresa comercial
    for analysis in commercial_analyses:
        empresa = analysis['empresa']
        report += f"""
### 💼 Análise Comercial: {empresa}

**CNPJ:** {analysis.get('cnpj', 'Não informado')}
**Preço Total:** {analysis.get('preco_total', 'Não identificado')}
**BDI:** {analysis.get('bdi', 'Não informado')}

**Condições de Pagamento:**
{analysis.get('condicoes_pagamento', 'Não especificadas claramente')}

**Prazos Comerciais:**
"""
        if analysis.get('prazos'):
            for prazo in analysis['prazos']:
                report += f"- {prazo}\n"
        else:
            report += "Prazos não especificados.\n"
        
        report += "\n---\n"
    
    report += f"""

## 🎯 BLOCO 4: CONCLUSÃO E RECOMENDAÇÕES

### Síntese da Análise

**Aspecto Técnico:**
"""
    
    # Identificar melhor proposta técnica
    melhor_tecnica = ""
    maior_aderencia = 0
    
    for empresa, aderencia in tech_comparison.get('aderencia_tr', {}).items():
        aderencia_num = int(aderencia.replace('%', ''))
        if aderencia_num > maior_aderencia:
            maior_aderencia = aderencia_num
            melhor_tecnica = empresa
    
    if melhor_tecnica:
        report += f"A empresa **{melhor_tecnica}** apresentou a melhor aderência técnica ao TR ({maior_aderencia}%).\n\n"
    
    report += "**Aspecto Comercial:**\n"
    if comm_comparison.get('ranking_precos'):
        melhor_preco = comm_comparison['ranking_precos'][0]
        report += f"A empresa **{melhor_preco[0]}** apresentou o menor preço: {melhor_preco[1]}.\n\n"
    
    report += """### Recomendações Finais

**Para a Tomada de Decisão:**

1. **Verificação de Documentação:** Confirmar se todas as empresas apresentaram documentação completa de habilitação.

2. **Esclarecimentos Técnicos:** Solicitar esclarecimentos sobre pontos não claramente apresentados nas propostas técnicas.

3. **Análise de Saúde Financeira:** Verificar a situação financeira das empresas proponentes através de consultas aos órgãos competentes.

4. **Negociação:** Considerar possibilidade de negociação com as empresas melhor classificadas.

### Considerações Importantes

- Esta análise foi realizada com base no conteúdo extraído dos documentos fornecidos.
- Recomenda-se análise detalhada adicional por especialistas da área.
- Verificar conformidade com a legislação de licitações aplicável.

---

### 📈 Resumo Executivo para Decisão

**Melhor Proposta Técnica:** {melhor_tecnica if melhor_tecnica else 'A definir'}
**Melhor Proposta Comercial:** {comm_comparison['ranking_precos'][0][0] if comm_comparison.get('ranking_precos') else 'A definir'}

**Recomendação Geral:** Realizar análise conjunta dos aspectos técnicos e comerciais, considerando o melhor custo-benefício para o projeto.

---

*Relatório gerado automaticamente pelo Proposal Analyzer Pro com Análise de IA*  
*Data: {current_time}*  
*Versão: 2.0 - Enhanced AI Analysis*
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
        
        # Analisar TR com IA
        tr_analysis = analyze_tr_content(tr_text)
        
        # Processar propostas técnicas
        technical_proposals = []
        technical_analyses = []
        tech_companies = request.form.getlist('techCompany[]')
        tech_files = request.files.getlist('techFile[]')
        
        for i, (company, file) in enumerate(zip(tech_companies, tech_files)):
            if company and file and file.filename:
                file_path = os.path.join(app.config['UPLOAD_FOLDER'], f"tech_{i}_{file.filename}")
                file.save(file_path)
                content = extract_text_from_file(file_path)
                
                # Análise com IA
                tech_analysis = analyze_technical_proposal(content, company)
                technical_analyses.append(tech_analysis)
                
                technical_proposals.append({
                    'company': company,
                    'content': content
                })
        
        # Processar propostas comerciais
        commercial_proposals = []
        commercial_analyses = []
        comm_companies = request.form.getlist('commCompany[]')
        comm_cnpjs = request.form.getlist('commCnpj[]')
        comm_files = request.files.getlist('commFile[]')
        
        for i, (company, cnpj, file) in enumerate(zip(comm_companies, comm_cnpjs, comm_files)):
            if company and file and file.filename:
                file_path = os.path.join(app.config['UPLOAD_FOLDER'], f"comm_{i}_{file.filename}")
                file.save(file_path)
                content = extract_text_from_file(file_path)
                
                # Análise com IA
                comm_analysis = analyze_commercial_proposal(content, company, cnpj)
                commercial_analyses.append(comm_analysis)
                
                commercial_proposals.append({
                    'company': company,
                    'cnpj': cnpj,
                    'content': content
                })
        
        # Gerar análise comparativa
        tech_comparison, comm_comparison = generate_comparative_analysis(
            tr_analysis, technical_analyses, commercial_analyses
        )
        
        # Gerar relatório aprimorado
        report = generate_enhanced_report(
            project_name, project_description, tr_analysis,
            technical_analyses, commercial_analyses,
            tech_comparison, comm_comparison
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
            # Gerar PDF usando reportlab
            md_file_path = os.path.join(app.config['UPLOAD_FOLDER'], f"{report_id}.md")
            pdf_file_path = os.path.join(app.config['UPLOAD_FOLDER'], f"relatorio_analise.pdf")
            
            # Ler o conteúdo markdown
            with open(md_file_path, 'r', encoding='utf-8') as f:
                markdown_content = f.read()
            
            # Converter para PDF usando reportlab
            from reportlab.lib.pagesizes import letter, A4
            from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
            from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
            from reportlab.lib.units import inch
            from reportlab.lib import colors
            
            # Criar documento PDF
            doc = SimpleDocTemplate(pdf_file_path, pagesize=A4, topMargin=1*inch)
            styles = getSampleStyleSheet()
            story = []
            
            # Estilos personalizados
            title_style = ParagraphStyle(
                'CustomTitle',
                parent=styles['Heading1'],
                fontSize=18,
                spaceAfter=30,
                textColor=colors.darkblue,
                alignment=1  # Center
            )
            
            heading_style = ParagraphStyle(
                'CustomHeading',
                parent=styles['Heading2'],
                fontSize=14,
                spaceAfter=12,
                textColor=colors.darkblue,
                spaceBefore=20
            )
            
            subheading_style = ParagraphStyle(
                'CustomSubHeading',
                parent=styles['Heading3'],
                fontSize=12,
                spaceAfter=8,
                textColor=colors.darkgreen,
                spaceBefore=15
            )
            
            normal_style = ParagraphStyle(
                'CustomNormal',
                parent=styles['Normal'],
                fontSize=10,
                spaceAfter=12,
                alignment=0  # Left
            )
            
            # Processar markdown para PDF
            lines = markdown_content.split('\n')
            table_data = []
            in_table = False
            
            for line in lines:
                line = line.strip()
                if not line:
                    if not in_table:
                        story.append(Spacer(1, 12))
                elif line.startswith('# '):
                    if in_table and table_data:
                        # Adicionar tabela antes de continuar
                        table = Table(table_data)
                        table.setStyle(TableStyle([
                            ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
                            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                            ('FONTSIZE', (0, 0), (-1, 0), 10),
                            ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
                            ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
                            ('GRID', (0, 0), (-1, -1), 1, colors.black)
                        ]))
                        story.append(table)
                        story.append(Spacer(1, 20))
                        table_data = []
                        in_table = False
                    story.append(Paragraph(line[2:], title_style))
                elif line.startswith('## '):
                    if in_table and table_data:
                        table = Table(table_data)
                        table.setStyle(TableStyle([
                            ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
                            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                            ('FONTSIZE', (0, 0), (-1, 0), 10),
                            ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
                            ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
                            ('GRID', (0, 0), (-1, -1), 1, colors.black)
                        ]))
                        story.append(table)
                        story.append(Spacer(1, 20))
                        table_data = []
                        in_table = False
                    story.append(Paragraph(line[3:], heading_style))
                elif line.startswith('### '):
                    story.append(Paragraph(line[4:], subheading_style))
                elif line.startswith('|') and '|' in line[1:]:
                    # Tabela markdown
                    if not in_table:
                        in_table = True
                        table_data = []
                    
                    # Processar linha da tabela
                    cells = [cell.strip() for cell in line.split('|')[1:-1]]
                    if not all(cell.startswith('-') for cell in cells):  # Ignorar linha separadora
                        table_data.append(cells)
                elif line.startswith('**') and line.endswith('**'):
                    if in_table and table_data:
                        table = Table(table_data)
                        table.setStyle(TableStyle([
                            ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
                            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                            ('FONTSIZE', (0, 0), (-1, 0), 10),
                            ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
                            ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
                            ('GRID', (0, 0), (-1, -1), 1, colors.black)
                        ]))
                        story.append(table)
                        story.append(Spacer(1, 20))
                        table_data = []
                        in_table = False
                    story.append(Paragraph(f"<b>{line[2:-2]}</b>", normal_style))
                else:
                    if in_table and table_data:
                        table = Table(table_data)
                        table.setStyle(TableStyle([
                            ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
                            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                            ('FONTSIZE', (0, 0), (-1, 0), 10),
                            ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
                            ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
                            ('GRID', (0, 0), (-1, -1), 1, colors.black)
                        ]))
                        story.append(table)
                        story.append(Spacer(1, 20))
                        table_data = []
                        in_table = False
                    if line:
                        story.append(Paragraph(line, normal_style))
            
            # Adicionar tabela final se existir
            if in_table and table_data:
                table = Table(table_data)
                table.setStyle(TableStyle([
                    ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
                    ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                    ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                    ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                    ('FONTSIZE', (0, 0), (-1, 0), 10),
                    ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
                    ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
                    ('GRID', (0, 0), (-1, -1), 1, colors.black)
                ]))
                story.append(table)
            
            # Construir PDF
            doc.build(story)
            
            return send_file(pdf_file_path, as_attachment=True, download_name='relatorio_analise.pdf')
        else:
            return jsonify({'error': 'Formato não suportado'}), 400
            
    except Exception as e:
        return jsonify({'error': str(e)}), 500

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=False)

