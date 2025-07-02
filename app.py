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
    """Extrai texto de diferentes tipos de arquivo, incluindo Excel"""
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
        
        elif file_extension in ['.xls', '.xlsx']:
            # Processar arquivo Excel
            return extract_excel_data(file_path)
        
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

def extract_excel_data(file_path):
    """Extrai dados estruturados de arquivos Excel"""
    try:
        # Carregar o workbook
        wb = openpyxl.load_workbook(file_path)
        extracted_data = {
            'sheets': wb.sheetnames,
            'data': {}
        }
        
        # Processar cada aba
        for sheet_name in wb.sheetnames:
            try:
                df = pd.read_excel(file_path, sheet_name=sheet_name)
                
                # Converter para texto estruturado
                sheet_text = f"\n=== ABA: {sheet_name} ===\n"
                
                # Adicionar dados da planilha
                for index, row in df.iterrows():
                    row_text = " | ".join([str(cell) if pd.notna(cell) else "" for cell in row])
                    if row_text.strip():  # Só adicionar linhas não vazias
                        sheet_text += f"Linha {index + 1}: {row_text}\n"
                
                extracted_data['data'][sheet_name] = sheet_text
                
            except Exception as e:
                extracted_data['data'][sheet_name] = f"Erro ao processar aba {sheet_name}: {str(e)}"
        
        # Combinar todos os dados em um texto
        combined_text = f"ARQUIVO EXCEL: {os.path.basename(file_path)}\n"
        combined_text += f"ABAS DISPONÍVEIS: {', '.join(extracted_data['sheets'])}\n\n"
        
        for sheet_name, sheet_data in extracted_data['data'].items():
            combined_text += sheet_data + "\n"
        
        return combined_text
        
    except Exception as e:
        return f"Erro ao processar arquivo Excel: {str(e)}"

def analyze_tr_content(tr_text):
    """Analisa o conteúdo do Termo de Referência e extrai informações estruturadas"""
    analysis = {
        'resumo': '',
        'objeto': '',
        'requisitos_tecnicos': [],
        'prazos': [],
        'criterios_avaliacao': [],
        'qualificacoes_exigidas': [],
        'valores_estimados': [],
        'especificacoes_tecnicas': [],
        'metodologia_exigida': [],
        'recursos_necessarios': []
    }
    
    # Dividir texto em seções
    sections = tr_text.split('\n')
    current_section = ''
    
    # Padrões para identificar informações importantes
    prazo_patterns = [
        r'prazo.*?(\d+).*?(dia|mês|ano|semana)',
        r'cronograma.*?(\d+).*?(dia|mês|ano|semana)',
        r'entrega.*?(\d+).*?(dia|mês|ano|semana)',
        r'execução.*?(\d+).*?(dia|mês|ano|semana)'
    ]
    
    valor_patterns = [
        r'R\$\s*[\d.,]+',
        r'valor.*?R\$\s*[\d.,]+',
        r'orçamento.*?R\$\s*[\d.,]+',
        r'estimado.*?R\$\s*[\d.,]+'
    ]
    
    # Extrair objeto/resumo (primeiros parágrafos significativos)
    meaningful_paragraphs = [p.strip() for p in sections if len(p.strip()) > 50]
    if meaningful_paragraphs:
        analysis['resumo'] = ' '.join(meaningful_paragraphs[:3])
        analysis['objeto'] = meaningful_paragraphs[0] if meaningful_paragraphs else ''
    
    # Buscar prazos com mais precisão
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
    tech_keywords = ['técnico', 'especificação', 'requisito', 'metodologia', 'equipamento', 'material', 'norma', 'padrão']
    for section in sections:
        if any(keyword in section.lower() for keyword in tech_keywords) and len(section.strip()) > 30:
            analysis['requisitos_tecnicos'].append(section.strip())
    
    # Identificar especificações técnicas detalhadas
    spec_keywords = ['especificação', 'norma', 'padrão', 'certificação', 'qualidade', 'performance']
    for section in sections:
        if any(keyword in section.lower() for keyword in spec_keywords) and len(section.strip()) > 40:
            analysis['especificacoes_tecnicas'].append(section.strip())
    
    # Identificar metodologia exigida
    method_keywords = ['metodologia', 'método', 'processo', 'procedimento', 'abordagem', 'estratégia']
    for section in sections:
        if any(keyword in section.lower() for keyword in method_keywords) and len(section.strip()) > 40:
            analysis['metodologia_exigida'].append(section.strip())
    
    # Identificar critérios de avaliação
    eval_keywords = ['avaliação', 'critério', 'pontuação', 'peso', 'classificação', 'julgamento']
    for section in sections:
        if any(keyword in section.lower() for keyword in eval_keywords) and len(section.strip()) > 30:
            analysis['criterios_avaliacao'].append(section.strip())
    
    # Identificar qualificações exigidas
    qual_keywords = ['qualificação', 'experiência', 'certificação', 'habilitação', 'comprovação', 'atestado']
    for section in sections:
        if any(keyword in section.lower() for keyword in qual_keywords) and len(section.strip()) > 30:
            analysis['qualificacoes_exigidas'].append(section.strip())
    
    # Identificar recursos necessários
    resource_keywords = ['recurso', 'equipamento', 'ferramenta', 'material', 'insumo', 'mão de obra']
    for section in sections:
        if any(keyword in section.lower() for keyword in resource_keywords) and len(section.strip()) > 30:
            analysis['recursos_necessarios'].append(section.strip())
    
    return analysis

def analyze_technical_proposal_detailed(proposal_text, company_name):
    """Análise técnica detalhada e aprofundada de uma proposta"""
    analysis = {
        'empresa': company_name,
        'cnpj': '',
        'metodologia': {
            'descricao': '',
            'fases_identificadas': [],
            'ferramentas_mencionadas': [],
            'abordagem_qualitativa': '',
            'aderencia_tr': 0
        },
        'cronograma': {
            'prazo_total': '',
            'marcos_principais': [],
            'fases_detalhadas': [],
            'recursos_por_fase': [],
            'viabilidade': ''
        },
        'equipe_tecnica': {
            'coordenador': '',
            'especialistas': [],
            'qualificacoes': [],
            'experiencia_relevante': [],
            'adequacao_projeto': ''
        },
        'recursos_tecnicos': {
            'equipamentos': [],
            'materiais': [],
            'tecnologias': [],
            'inovacoes': []
        },
        'experiencia_comprovada': {
            'projetos_similares': [],
            'referencias': [],
            'certificacoes': [],
            'cases_sucesso': []
        },
        'diferenciais_competitivos': [],
        'riscos_identificados': [],
        'pontos_fortes': [],
        'pontos_fracos': [],
        'gaps_identificados': [],
        'score_detalhado': {
            'metodologia': 0,
            'cronograma': 0,
            'equipe': 0,
            'recursos': 0,
            'experiencia': 0
        }
    }
    
    sections = proposal_text.split('\n')
    
    # Extrair CNPJ com padrões mais robustos
    cnpj_patterns = [
        r'CNPJ[:\s]*(\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2})',
        r'CNPJ[:\s]*(\d{14})',
        r'(\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2})',
        r'CNPJ.*?(\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2})'
    ]
    
    for pattern in cnpj_patterns:
        matches = re.findall(pattern, proposal_text, re.IGNORECASE)
        if matches:
            analysis['cnpj'] = matches[0]
            break
    
    # Análise de Metodologia Detalhada
    metodologia_keywords = ['metodologia', 'método', 'abordagem', 'estratégia', 'processo', 'procedimento']
    metodologia_sections = []
    
    for section in sections:
        if any(keyword in section.lower() for keyword in metodologia_keywords) and len(section.strip()) > 50:
            metodologia_sections.append(section.strip())
    
    if metodologia_sections:
        analysis['metodologia']['descricao'] = ' '.join(metodologia_sections[:2])
        
        # Identificar fases da metodologia
        fase_patterns = [
            r'fase\s*(\d+)',
            r'etapa\s*(\d+)',
            r'passo\s*(\d+)',
            r'estágio\s*(\d+)'
        ]
        
        for section in metodologia_sections:
            for pattern in fase_patterns:
                matches = re.findall(pattern, section.lower())
                for match in matches:
                    analysis['metodologia']['fases_identificadas'].append(f"Fase {match}")
        
        # Identificar ferramentas mencionadas
        ferramenta_keywords = ['ferramenta', 'software', 'sistema', 'plataforma', 'tecnologia']
        for section in metodologia_sections:
            for keyword in ferramenta_keywords:
                if keyword in section.lower():
                    # Extrair contexto da ferramenta
                    words = section.split()
                    for i, word in enumerate(words):
                        if keyword in word.lower() and i < len(words) - 1:
                            analysis['metodologia']['ferramentas_mencionadas'].append(f"{word} {words[i+1]}")
        
        # Avaliar aderência (básico)
        if len(metodologia_sections) >= 2:
            analysis['metodologia']['aderencia_tr'] = 80
        elif len(metodologia_sections) == 1:
            analysis['metodologia']['aderencia_tr'] = 60
        else:
            analysis['metodologia']['aderencia_tr'] = 20
        
        analysis['score_detalhado']['metodologia'] = analysis['metodologia']['aderencia_tr']
    else:
        analysis['metodologia']['descricao'] = 'Metodologia não claramente identificada ou apresentada de forma insuficiente'
        analysis['gaps_identificados'].append('Metodologia não detalhada adequadamente')
        analysis['score_detalhado']['metodologia'] = 10
    
    # Análise de Cronograma Detalhada
    cronograma_keywords = ['cronograma', 'prazo', 'etapa', 'fase', 'período', 'duração', 'tempo']
    cronograma_sections = []
    
    for section in sections:
        if any(keyword in section.lower() for keyword in cronograma_keywords):
            cronograma_sections.append(section.strip())
    
    # Extrair prazos específicos
    time_patterns = [
        r'(\d+)\s*(dia|semana|mês|ano)',
        r'(\d+)\s*a\s*(\d+)\s*(dia|semana|mês|ano)',
        r'prazo.*?(\d+).*?(dia|semana|mês|ano)'
    ]
    
    for section in cronograma_sections:
        for pattern in time_patterns:
            matches = re.findall(pattern, section.lower())
            for match in matches:
                if len(match) == 2:
                    analysis['cronograma']['marcos_principais'].append(f"{match[0]} {match[1]}")
                elif len(match) == 4:
                    analysis['cronograma']['marcos_principais'].append(f"{match[0]} a {match[1]} {match[2]}")
    
    # Identificar fases detalhadas do cronograma
    for section in cronograma_sections:
        if len(section) > 100:  # Seções mais detalhadas
            analysis['cronograma']['fases_detalhadas'].append(section[:200] + "...")
    
    # Avaliar viabilidade do cronograma
    if analysis['cronograma']['marcos_principais']:
        analysis['cronograma']['viabilidade'] = 'Cronograma apresentado com marcos definidos'
        analysis['score_detalhado']['cronograma'] = 75
    elif cronograma_sections:
        analysis['cronograma']['viabilidade'] = 'Cronograma mencionado mas sem detalhamento adequado'
        analysis['score_detalhado']['cronograma'] = 40
    else:
        analysis['cronograma']['viabilidade'] = 'Cronograma não apresentado ou insuficiente'
        analysis['gaps_identificados'].append('Cronograma não detalhado')
        analysis['score_detalhado']['cronograma'] = 10
    
    # Análise de Equipe Técnica Detalhada
    equipe_keywords = ['equipe', 'profissional', 'responsável', 'coordenador', 'especialista', 'técnico', 'engenheiro']
    equipe_sections = []
    
    for section in sections:
        if any(keyword in section.lower() for keyword in equipe_keywords) and len(section.strip()) > 20:
            equipe_sections.append(section.strip())
    
    # Identificar coordenador/responsável técnico
    coord_keywords = ['coordenador', 'responsável técnico', 'gerente', 'líder']
    for section in equipe_sections:
        for keyword in coord_keywords:
            if keyword in section.lower():
                analysis['equipe_tecnica']['coordenador'] = section[:150] + "..."
                break
    
    # Identificar especialistas
    espec_keywords = ['especialista', 'expert', 'consultor', 'profissional especializado']
    for section in equipe_sections:
        for keyword in espec_keywords:
            if keyword in section.lower():
                analysis['equipe_tecnica']['especialistas'].append(section[:100] + "...")
    
    # Identificar qualificações
    qual_keywords = ['qualificação', 'formação', 'certificação', 'experiência', 'graduação', 'pós-graduação']
    for section in equipe_sections:
        for keyword in qual_keywords:
            if keyword in section.lower():
                analysis['equipe_tecnica']['qualificacoes'].append(section[:120] + "...")
    
    # Identificar experiência relevante
    exp_keywords = ['experiência', 'projeto similar', 'case', 'trabalho anterior', 'histórico']
    for section in equipe_sections:
        for keyword in exp_keywords:
            if keyword in section.lower():
                analysis['equipe_tecnica']['experiencia_relevante'].append(section[:120] + "...")
    
    # Avaliar adequação da equipe
    equipe_score = 0
    if analysis['equipe_tecnica']['coordenador']:
        equipe_score += 25
    if analysis['equipe_tecnica']['especialistas']:
        equipe_score += 25
    if analysis['equipe_tecnica']['qualificacoes']:
        equipe_score += 25
    if analysis['equipe_tecnica']['experiencia_relevante']:
        equipe_score += 25
    
    analysis['score_detalhado']['equipe'] = equipe_score
    
    if equipe_score >= 75:
        analysis['equipe_tecnica']['adequacao_projeto'] = 'Equipe bem estruturada e qualificada'
    elif equipe_score >= 50:
        analysis['equipe_tecnica']['adequacao_projeto'] = 'Equipe adequada com algumas lacunas'
    else:
        analysis['equipe_tecnica']['adequacao_projeto'] = 'Equipe insuficientemente detalhada'
        analysis['gaps_identificados'].append('Detalhamento insuficiente da equipe técnica')
    
    # Análise de Recursos Técnicos
    recurso_keywords = ['equipamento', 'ferramenta', 'material', 'recurso', 'tecnologia', 'software', 'hardware']
    
    for section in sections:
        for keyword in recurso_keywords:
            if keyword in section.lower() and len(section.strip()) > 30:
                if 'equipamento' in keyword or 'ferramenta' in keyword:
                    analysis['recursos_tecnicos']['equipamentos'].append(section[:100] + "...")
                elif 'material' in keyword or 'insumo' in keyword:
                    analysis['recursos_tecnicos']['materiais'].append(section[:100] + "...")
                elif 'tecnologia' in keyword or 'software' in keyword:
                    analysis['recursos_tecnicos']['tecnologias'].append(section[:100] + "...")
    
    # Avaliar recursos
    recursos_score = 0
    if analysis['recursos_tecnicos']['equipamentos']:
        recursos_score += 35
    if analysis['recursos_tecnicos']['materiais']:
        recursos_score += 35
    if analysis['recursos_tecnicos']['tecnologias']:
        recursos_score += 30
    
    analysis['score_detalhado']['recursos'] = recursos_score
    
    # Análise de Experiência Comprovada
    exp_keywords = ['projeto similar', 'experiência', 'referência', 'atestado', 'certificação', 'case']
    
    for section in sections:
        for keyword in exp_keywords:
            if keyword in section.lower() and len(section.strip()) > 40:
                if 'projeto' in keyword or 'case' in keyword:
                    analysis['experiencia_comprovada']['projetos_similares'].append(section[:150] + "...")
                elif 'referência' in keyword or 'atestado' in keyword:
                    analysis['experiencia_comprovada']['referencias'].append(section[:150] + "...")
                elif 'certificação' in keyword:
                    analysis['experiencia_comprovada']['certificacoes'].append(section[:150] + "...")
    
    # Avaliar experiência
    exp_score = 0
    if analysis['experiencia_comprovada']['projetos_similares']:
        exp_score += 40
    if analysis['experiencia_comprovada']['referencias']:
        exp_score += 30
    if analysis['experiencia_comprovada']['certificacoes']:
        exp_score += 30
    
    analysis['score_detalhado']['experiencia'] = exp_score
    
    # Identificar Diferenciais Competitivos
    diferencial_keywords = ['diferencial', 'inovação', 'vantagem', 'exclusivo', 'único', 'pioneiro']
    for section in sections:
        for keyword in diferencial_keywords:
            if keyword in section.lower() and len(section.strip()) > 40:
                analysis['diferenciais_competitivos'].append(section[:120] + "...")
    
    # Identificar Riscos
    risco_keywords = ['risco', 'problema', 'dificuldade', 'limitação', 'restrição']
    for section in sections:
        for keyword in risco_keywords:
            if keyword in section.lower() and len(section.strip()) > 30:
                analysis['riscos_identificados'].append(section[:100] + "...")
    
    # Calcular Pontos Fortes e Fracos baseado nos scores
    if analysis['score_detalhado']['metodologia'] >= 70:
        analysis['pontos_fortes'].append('Metodologia bem estruturada e detalhada')
    else:
        analysis['pontos_fracos'].append('Metodologia insuficientemente detalhada')
    
    if analysis['score_detalhado']['cronograma'] >= 70:
        analysis['pontos_fortes'].append('Cronograma bem definido com marcos claros')
    else:
        analysis['pontos_fracos'].append('Cronograma não adequadamente apresentado')
    
    if analysis['score_detalhado']['equipe'] >= 70:
        analysis['pontos_fortes'].append('Equipe técnica qualificada e bem estruturada')
    else:
        analysis['pontos_fracos'].append('Equipe técnica insuficientemente detalhada')
    
    if analysis['score_detalhado']['recursos'] >= 70:
        analysis['pontos_fortes'].append('Recursos técnicos adequados e bem especificados')
    else:
        analysis['pontos_fracos'].append('Recursos técnicos não adequadamente especificados')
    
    if analysis['score_detalhado']['experiencia'] >= 70:
        analysis['pontos_fortes'].append('Experiência comprovada em projetos similares')
    else:
        analysis['pontos_fracos'].append('Experiência em projetos similares não comprovada')
    
    return analysis

def analyze_commercial_proposal_excel(proposal_text, company_name, cnpj):
    """Analisa uma proposta comercial incluindo dados de Excel"""
    analysis = {
        'empresa': company_name,
        'cnpj': cnpj,
        'preco_total': '',
        'composicao_custos': {},
        'condicoes_pagamento': '',
        'prazos': [],
        'bdi': '',
        'observacoes': [],
        'itens_servicos': [],
        'detalhes_bdi': {}
    }
    
    # Se o texto contém dados de Excel, processar de forma diferente
    if "ARQUIVO EXCEL:" in proposal_text:
        return analyze_excel_commercial_data(proposal_text, company_name, cnpj)
    
    # Processar como PDF normal
    return analyze_commercial_proposal_pdf(proposal_text, company_name, cnpj)

def analyze_excel_commercial_data(excel_text, company_name, cnpj):
    """Analisa dados comerciais extraídos de Excel"""
    analysis = {
        'empresa': company_name,
        'cnpj': cnpj,
        'preco_total': '',
        'composicao_custos': {},
        'condicoes_pagamento': '',
        'prazos': [],
        'bdi': '',
        'observacoes': [],
        'itens_servicos': [],
        'detalhes_bdi': {}
    }
    
    lines = excel_text.split('\n')
    
    # Buscar preços na aba "Itens Serviços"
    in_itens_servicos = False
    precos_encontrados = []
    
    for line in lines:
        if "=== ABA: Itens Serviços ===" in line:
            in_itens_servicos = True
            continue
        elif "=== ABA:" in line and in_itens_servicos:
            in_itens_servicos = False
            continue
        
        if in_itens_servicos and "Preço Total(R$)" in line:
            # Próximas linhas devem conter os preços
            continue
        
        if in_itens_servicos and line.strip():
            # Buscar valores numéricos na linha
            valores = re.findall(r'[\d.,]+', line)
            for valor in valores:
                try:
                    # Tentar converter para float
                    if '.' in valor and ',' in valor:
                        # Formato brasileiro: 1.234.567,89
                        clean_valor = valor.replace('.', '').replace(',', '.')
                    elif ',' in valor:
                        # Pode ser decimal brasileiro: 1234,89
                        clean_valor = valor.replace(',', '.')
                    else:
                        clean_valor = valor
                    
                    float_valor = float(clean_valor)
                    if float_valor > 100:  # Filtrar valores muito pequenos
                        precos_encontrados.append((valor, float_valor))
                        analysis['itens_servicos'].append(f"R$ {valor}")
                except:
                    continue
    
    # Determinar preço total (maior valor ou soma)
    if precos_encontrados:
        # Ordenar por valor
        precos_encontrados.sort(key=lambda x: x[1], reverse=True)
        
        # Se há muitos valores, somar todos; senão, pegar o maior
        if len(precos_encontrados) > 5:
            total = sum([p[1] for p in precos_encontrados])
            analysis['preco_total'] = f"R$ {total:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')
        else:
            analysis['preco_total'] = f"R$ {precos_encontrados[0][0]}"
    
    # Buscar BDI na aba específica
    in_bdi = False
    for line in lines:
        if "=== ABA: BDI ===" in line:
            in_bdi = True
            continue
        elif "=== ABA:" in line and in_bdi:
            in_bdi = False
            continue
        
        if in_bdi and line.strip():
            # Buscar percentuais de BDI
            bdi_matches = re.findall(r'(\d+[,.]?\d*)%?', line)
            if bdi_matches:
                for match in bdi_matches:
                    try:
                        bdi_val = float(match.replace(',', '.'))
                        if 5 <= bdi_val <= 50:  # BDI típico entre 5% e 50%
                            analysis['bdi'] = f"{bdi_val}%"
                            break
                    except:
                        continue
    
    # Buscar composição de custos
    in_comp_custo = False
    for line in lines:
        if "=== ABA: Comp. Custo -GLOBAL ===" in line:
            in_comp_custo = True
            continue
        elif "=== ABA:" in line and in_comp_custo:
            in_comp_custo = False
            continue
        
        if in_comp_custo and line.strip():
            # Identificar categorias de custo
            if any(keyword in line.lower() for keyword in ['mão de obra', 'material', 'equipamento']):
                analysis['observacoes'].append(line.strip()[:100] + "...")
    
    # Buscar CNPJ se não fornecido
    if not analysis['cnpj']:
        cnpj_patterns = [
            r'(\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2})',
            r'CNPJ.*?(\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2})'
        ]
        
        for pattern in cnpj_patterns:
            matches = re.findall(pattern, excel_text)
            if matches:
                analysis['cnpj'] = matches[0]
                break
    
    return analysis

def analyze_commercial_proposal_pdf(proposal_text, company_name, cnpj):
    """Analisa uma proposta comercial em PDF"""
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
    
    # Buscar preços com padrões mais abrangentes
    price_patterns = [
        r'R\$\s*[\d.,]+',
        r'total.*?R\$\s*[\d.,]+',
        r'valor.*?R\$\s*[\d.,]+',
        r'preço.*?R\$\s*[\d.,]+',
        r'global.*?R\$\s*[\d.,]+',
        r'[\d.,]+\s*reais'
    ]
    
    prices_found = []
    for pattern in price_patterns:
        matches = re.findall(pattern, proposal_text, re.IGNORECASE)
        prices_found.extend(matches)
    
    if prices_found:
        # Limpar e converter preços para comparação
        cleaned_prices = []
        for price in prices_found:
            # Extrair apenas números e vírgulas/pontos
            clean_price = re.sub(r'[^\d,.]', '', price)
            if clean_price:
                try:
                    # Converter para float para comparação
                    if ',' in clean_price and '.' in clean_price:
                        # Formato brasileiro: 1.234.567,89
                        clean_price = clean_price.replace('.', '').replace(',', '.')
                    elif ',' in clean_price:
                        # Pode ser decimal brasileiro: 1234,89
                        clean_price = clean_price.replace(',', '.')
                    
                    float_value = float(clean_price)
                    cleaned_prices.append((price, float_value))
                except:
                    continue
        
        if cleaned_prices:
            # Assumir que o maior valor é o preço total
            analysis['preco_total'] = max(cleaned_prices, key=lambda x: x[1])[0]
    
    # Buscar CNPJ com padrão mais específico
    if not analysis['cnpj']:
        cnpj_patterns = [
            r'\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2}',
            r'\d{14}',
            r'CNPJ.*?(\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2})',
            r'CNPJ.*?(\d{14})'
        ]
        
        for pattern in cnpj_patterns:
            matches = re.findall(pattern, proposal_text)
            if matches:
                analysis['cnpj'] = matches[0]
                break
    
    # Buscar condições de pagamento
    payment_keywords = ['pagamento', 'parcela', 'à vista', 'prazo', 'condição']
    sections = proposal_text.split('\n')
    
    for section in sections:
        if any(keyword in section.lower() for keyword in payment_keywords) and len(section.strip()) > 20:
            analysis['condicoes_pagamento'] = section.strip()
            break
    
    # Buscar BDI
    bdi_patterns = [
        r'bdi.*?(\d+[,.]?\d*)%?',
        r'benefício.*?(\d+[,.]?\d*)%?',
        r'despesas.*?indiretas.*?(\d+[,.]?\d*)%?'
    ]
    
    for pattern in bdi_patterns:
        matches = re.findall(pattern, proposal_text.lower())
        if matches:
            analysis['bdi'] = matches[0] + '%'
            break
    
    # Buscar prazos
    prazo_patterns = [
        r'prazo.*?(\d+).*?(dia|mês|ano)',
        r'entrega.*?(\d+).*?(dia|mês|ano)',
        r'execução.*?(\d+).*?(dia|mês|ano)'
    ]
    
    for pattern in prazo_patterns:
        matches = re.findall(pattern, proposal_text.lower())
        for match in matches:
            analysis['prazos'].append(f"{match[0]} {match[1]}")
    
    return analysis

def generate_detailed_comparative_analysis(tr_analysis, technical_analyses, commercial_analyses):
    """Gera análise comparativa detalhada entre propostas e TR"""
    
    # Análise técnica comparativa detalhada
    tech_comparison = {
        'matriz_comparacao': {},
        'ranking_tecnico': [],
        'analise_gaps': {},
        'recomendacoes_tecnicas': [],
        'riscos_por_empresa': {}
    }
    
    # Análise comercial comparativa
    comm_comparison = {
        'ranking_precos': [],
        'analise_custo_beneficio': {},
        'condicoes_comparadas': {},
        'recomendacoes_comerciais': []
    }
    
    # Criar matriz de comparação técnica
    criterios_tecnicos = ['metodologia', 'cronograma', 'equipe', 'recursos', 'experiencia']
    
    for analysis in technical_analyses:
        empresa = analysis['empresa']
        scores = analysis['score_detalhado']
        
        tech_comparison['matriz_comparacao'][empresa] = scores
        
        # Calcular score total
        score_total = sum(scores.values()) / len(scores)
        tech_comparison['ranking_tecnico'].append((empresa, score_total))
        
        # Análise de gaps
        gaps = analysis['gaps_identificados']
        tech_comparison['analise_gaps'][empresa] = gaps
        
        # Riscos identificados
        riscos = analysis['riscos_identificados']
        tech_comparison['riscos_por_empresa'][empresa] = riscos
    
    # Ordenar ranking técnico
    tech_comparison['ranking_tecnico'].sort(key=lambda x: x[1], reverse=True)
    
    # Gerar recomendações técnicas
    if tech_comparison['ranking_tecnico']:
        melhor_empresa = tech_comparison['ranking_tecnico'][0][0]
        tech_comparison['recomendacoes_tecnicas'].append(
            f"Empresa {melhor_empresa} apresentou o melhor desempenho técnico geral"
        )
        
        # Recomendações específicas por critério
        for criterio in criterios_tecnicos:
            melhor_criterio = max(technical_analyses, 
                                key=lambda x: x['score_detalhado'].get(criterio, 0))
            tech_comparison['recomendacoes_tecnicas'].append(
                f"Em {criterio}: {melhor_criterio['empresa']} se destaca"
            )
    
    # Comparar propostas comerciais
    precos_empresas = []
    for analysis in commercial_analyses:
        if analysis.get('preco_total'):
            # Extrair valor numérico para comparação
            valor_str = re.sub(r'[^\d,.]', '', analysis['preco_total'])
            try:
                if ',' in valor_str and '.' in valor_str:
                    valor_str = valor_str.replace('.', '').replace(',', '.')
                elif ',' in valor_str:
                    valor_str = valor_str.replace(',', '.')
                
                valor_num = float(valor_str)
                precos_empresas.append((analysis['empresa'], analysis['preco_total'], valor_num))
            except:
                precos_empresas.append((analysis['empresa'], analysis['preco_total'], 0))
    
    # Ordenar por preço (menor para maior)
    precos_empresas.sort(key=lambda x: x[2])
    comm_comparison['ranking_precos'] = [(empresa, preco_str) for empresa, preco_str, _ in precos_empresas]
    
    # Análise de custo-benefício
    for i, (empresa_tech, score_tech) in enumerate(tech_comparison['ranking_tecnico']):
        # Encontrar posição no ranking comercial
        pos_comercial = next((j for j, (emp_comm, _) in enumerate(comm_comparison['ranking_precos']) 
                            if emp_comm == empresa_tech), len(comm_comparison['ranking_precos']))
        
        # Calcular índice custo-benefício (quanto menor, melhor)
        indice_cb = (i + 1) + (pos_comercial + 1)  # Posição técnica + posição comercial
        comm_comparison['analise_custo_beneficio'][empresa_tech] = {
            'posicao_tecnica': i + 1,
            'posicao_comercial': pos_comercial + 1,
            'indice_custo_beneficio': indice_cb,
            'score_tecnico': score_tech
        }
    
    return tech_comparison, comm_comparison

def generate_enhanced_report(project_name, project_description, tr_analysis, technical_analyses, commercial_analyses, tech_comparison, comm_comparison):
    """Gera relatório aprimorado com análise técnica detalhada e dados comerciais de Excel"""
    
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
    
    report += "### Especificações Técnicas Exigidas\n"
    if tr_analysis.get('especificacoes_tecnicas'):
        for i, spec in enumerate(tr_analysis['especificacoes_tecnicas'][:3], 1):
            report += f"**{i}.** {spec[:200]}...\n\n"
    else:
        report += "Especificações técnicas não claramente definidas no TR.\n\n"
    
    report += "### Metodologia Exigida pelo TR\n"
    if tr_analysis.get('metodologia_exigida'):
        for method in tr_analysis['metodologia_exigida'][:2]:
            report += f"- {method[:150]}...\n"
    else:
        report += "Metodologia específica não exigida ou não claramente definida no TR.\n"
    
    report += "\n### Prazos Estabelecidos\n"
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

### Matriz de Comparação Técnica Detalhada
"""
    
    if tech_comparison.get('matriz_comparacao'):
        report += "| Empresa | Metodologia | Cronograma | Equipe | Recursos | Experiência | Score Total |\n"
        report += "|---------|-------------|------------|--------|----------|-------------|-------------|\n"
        
        for empresa, score_total in tech_comparison['ranking_tecnico']:
            scores = tech_comparison['matriz_comparacao'][empresa]
            score_medio = score_total
            
            report += f"| {empresa} | {scores.get('metodologia', 0)}% | {scores.get('cronograma', 0)}% | {scores.get('equipe', 0)}% | {scores.get('recursos', 0)}% | {scores.get('experiencia', 0)}% | **{score_medio:.1f}%** |\n"
        
        report += "\n"
    
    # Análise detalhada por empresa
    for analysis in technical_analyses:
        empresa = analysis['empresa']
        cnpj = analysis.get('cnpj', 'Não identificado')
        
        report += f"""
### 📋 Análise Técnica Detalhada: {empresa}

**CNPJ:** {cnpj}

#### 🔬 Metodologia Proposta
**Descrição:** {analysis['metodologia']['descricao']}

**Fases Identificadas:**
"""
        if analysis['metodologia']['fases_identificadas']:
            for fase in analysis['metodologia']['fases_identificadas']:
                report += f"- {fase}\n"
        else:
            report += "Fases não claramente identificadas.\n"
        
        report += "\n**Ferramentas e Tecnologias Mencionadas:**\n"
        if analysis['metodologia']['ferramentas_mencionadas']:
            for ferramenta in analysis['metodologia']['ferramentas_mencionadas']:
                report += f"- {ferramenta}\n"
        else:
            report += "Ferramentas específicas não mencionadas.\n"
        
        report += f"\n**Aderência ao TR:** {analysis['metodologia']['aderencia_tr']}%\n"
        
        report += "\n#### ⏰ Cronograma e Prazos\n"
        report += f"**Viabilidade:** {analysis['cronograma']['viabilidade']}\n\n"
        
        report += "**Marcos Principais:**\n"
        if analysis['cronograma']['marcos_principais']:
            for marco in analysis['cronograma']['marcos_principais']:
                report += f"- {marco}\n"
        else:
            report += "Marcos não claramente definidos.\n"
        
        report += "\n**Fases Detalhadas:**\n"
        if analysis['cronograma']['fases_detalhadas']:
            for fase in analysis['cronograma']['fases_detalhadas'][:2]:
                report += f"- {fase}\n"
        else:
            report += "Detalhamento de fases não apresentado.\n"
        
        report += "\n#### 👥 Equipe Técnica\n"
        report += f"**Adequação ao Projeto:** {analysis['equipe_tecnica']['adequacao_projeto']}\n\n"
        
        if analysis['equipe_tecnica']['coordenador']:
            report += f"**Coordenador/Responsável Técnico:** {analysis['equipe_tecnica']['coordenador']}\n\n"
        
        report += "**Especialistas:**\n"
        if analysis['equipe_tecnica']['especialistas']:
            for esp in analysis['equipe_tecnica']['especialistas'][:3]:
                report += f"- {esp}\n"
        else:
            report += "Especialistas não claramente identificados.\n"
        
        report += "\n**Qualificações:**\n"
        if analysis['equipe_tecnica']['qualificacoes']:
            for qual in analysis['equipe_tecnica']['qualificacoes'][:3]:
                report += f"- {qual}\n"
        else:
            report += "Qualificações não detalhadas.\n"
        
        report += "\n#### 🛠️ Recursos Técnicos\n"
        
        report += "**Equipamentos:**\n"
        if analysis['recursos_tecnicos']['equipamentos']:
            for equip in analysis['recursos_tecnicos']['equipamentos'][:3]:
                report += f"- {equip}\n"
        else:
            report += "Equipamentos não especificados.\n"
        
        report += "\n**Tecnologias:**\n"
        if analysis['recursos_tecnicos']['tecnologias']:
            for tech in analysis['recursos_tecnicos']['tecnologias'][:3]:
                report += f"- {tech}\n"
        else:
            report += "Tecnologias não especificadas.\n"
        
        report += "\n#### 🏆 Experiência Comprovada\n"
        
        report += "**Projetos Similares:**\n"
        if analysis['experiencia_comprovada']['projetos_similares']:
            for proj in analysis['experiencia_comprovada']['projetos_similares'][:2]:
                report += f"- {proj}\n"
        else:
            report += "Projetos similares não comprovados.\n"
        
        report += "\n**Certificações:**\n"
        if analysis['experiencia_comprovada']['certificacoes']:
            for cert in analysis['experiencia_comprovada']['certificacoes'][:2]:
                report += f"- {cert}\n"
        else:
            report += "Certificações não apresentadas.\n"
        
        report += "\n#### ✅ Pontos Fortes\n"
        if analysis['pontos_fortes']:
            for ponto in analysis['pontos_fortes']:
                report += f"✅ {ponto}\n"
        else:
            report += "Pontos fortes não claramente identificados.\n"
        
        report += "\n#### ⚠️ Pontos de Atenção e Gaps\n"
        if analysis['pontos_fracos']:
            for ponto in analysis['pontos_fracos']:
                report += f"⚠️ {ponto}\n"
        
        if analysis['gaps_identificados']:
            for gap in analysis['gaps_identificados']:
                report += f"❌ {gap}\n"
        
        if not analysis['pontos_fracos'] and not analysis['gaps_identificados']:
            report += "Nenhum ponto de atenção crítico identificado.\n"
        
        report += "\n#### 🎯 Diferenciais Competitivos\n"
        if analysis['diferenciais_competitivos']:
            for diff in analysis['diferenciais_competitivos']:
                report += f"🌟 {diff}\n"
        else:
            report += "Diferenciais competitivos não claramente apresentados.\n"
        
        report += "\n---\n"
    
    # Ranking técnico final
    report += "\n### 🏆 Ranking Técnico Final\n"
    if tech_comparison.get('ranking_tecnico'):
        for i, (empresa, score) in enumerate(tech_comparison['ranking_tecnico'], 1):
            emoji = "🥇" if i == 1 else "🥈" if i == 2 else "🥉" if i == 3 else "📊"
            report += f"{emoji} **{i}º lugar:** {empresa} - Score: {score:.1f}%\n"
    
    report += f"""

## 💰 BLOCO 3: EQUALIZAÇÃO DAS PROPOSTAS COMERCIAIS

### Ranking de Preços
"""
    
    if comm_comparison.get('ranking_precos'):
        report += "| Posição | Empresa | Preço Total | Status |\n"
        report += "|---------|---------|-------------|--------|\n"
        
        for i, (empresa, preco) in enumerate(comm_comparison['ranking_precos'], 1):
            emoji = "🥇" if i == 1 else "🥈" if i == 2 else "🥉" if i == 3 else "📊"
            status = "Menor preço" if i == 1 else f"{i}º menor preço"
            
            report += f"| {emoji} {i}º | {empresa} | {preco} | {status} |\n"
        
        report += "\n"
    else:
        report += "**Atenção:** Não foi possível extrair informações de preços das propostas comerciais.\n\n"
    
    # Análise detalhada por empresa comercial
    for analysis in commercial_analyses:
        empresa = analysis['empresa']
        report += f"""
### 💼 Análise Comercial Detalhada: {empresa}

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
        
        # Adicionar itens de serviços se disponível
        if analysis.get('itens_servicos'):
            report += "\n**Itens de Serviços Identificados:**\n"
            for item in analysis['itens_servicos'][:5]:
                report += f"- {item}\n"
        
        # Adicionar observações se disponível
        if analysis.get('observacoes'):
            report += "\n**Observações:**\n"
            for obs in analysis['observacoes'][:3]:
                report += f"- {obs}\n"
        
        report += "\n---\n"
    
    # Análise de custo-benefício
    report += "\n### 📊 Análise de Custo-Benefício\n"
    if comm_comparison.get('analise_custo_beneficio'):
        report += "| Empresa | Posição Técnica | Posição Comercial | Índice C/B | Recomendação |\n"
        report += "|---------|-----------------|-------------------|------------|-------------|\n"
        
        # Ordenar por índice custo-benefício
        cb_sorted = sorted(comm_comparison['analise_custo_beneficio'].items(), 
                          key=lambda x: x[1]['indice_custo_beneficio'])
        
        for empresa, dados in cb_sorted:
            indice = dados['indice_custo_beneficio']
            recomendacao = "Excelente" if indice <= 4 else "Boa" if indice <= 6 else "Regular"
            
            report += f"| {empresa} | {dados['posicao_tecnica']}º | {dados['posicao_comercial']}º | {indice} | {recomendacao} |\n"
        
        report += "\n"
    
    report += f"""

## 🎯 BLOCO 4: CONCLUSÃO E RECOMENDAÇÕES

### Síntese da Análise Técnica
"""
    
    # Identificar melhor proposta técnica
    melhor_tecnica = "A definir"
    score_tecnico = 0
    if tech_comparison.get('ranking_tecnico'):
        melhor_tecnica, score_tecnico = tech_comparison['ranking_tecnico'][0]
        report += f"**Melhor Proposta Técnica:** {melhor_tecnica} (Score: {score_tecnico:.1f}%)\n\n"
        
        # Justificativa técnica
        melhor_analysis = next((a for a in technical_analyses if a['empresa'] == melhor_tecnica), None)
        if melhor_analysis:
            report += "**Justificativa:**\n"
            for ponto in melhor_analysis['pontos_fortes'][:3]:
                report += f"- {ponto}\n"
    
    report += "\n### Síntese da Análise Comercial\n"
    melhor_comercial = "A definir"
    if comm_comparison.get('ranking_precos'):
        melhor_comercial_data = comm_comparison['ranking_precos'][0]
        melhor_comercial = melhor_comercial_data[0]
        report += f"**Melhor Proposta Comercial:** {melhor_comercial} - {melhor_comercial_data[1]}\n\n"
    
    # Recomendação de custo-benefício
    melhor_cb = "A definir"
    if comm_comparison.get('analise_custo_beneficio'):
        cb_sorted = sorted(comm_comparison['analise_custo_beneficio'].items(), 
                          key=lambda x: x[1]['indice_custo_beneficio'])
        melhor_cb_data = cb_sorted[0]
        melhor_cb = melhor_cb_data[0]
        
        report += f"**Melhor Custo-Benefício:** {melhor_cb} (Índice: {melhor_cb_data[1]['indice_custo_beneficio']})\n\n"
    
    report += """### Recomendações Específicas

**Para a Tomada de Decisão:**

1. **Verificação de Documentação:** Confirmar se todas as empresas apresentaram documentação completa de habilitação.

2. **Esclarecimentos Técnicos:** Solicitar esclarecimentos sobre pontos não claramente apresentados nas propostas técnicas.

3. **Análise de Saúde Financeira:** Verificar a situação financeira das empresas proponentes através de consultas aos órgãos competentes.

4. **Negociação:** Considerar possibilidade de negociação com as empresas melhor classificadas.

5. **Visita Técnica:** Realizar visita às instalações das empresas finalistas para verificação in loco.

### Considerações Importantes

- Esta análise foi realizada com base no conteúdo extraído dos documentos fornecidos.
- Dados comerciais foram extraídos de planilhas Excel quando disponíveis.
- Recomenda-se análise detalhada adicional por especialistas da área.
- Verificar conformidade com a legislação de licitações aplicável.
- Considerar aspectos qualitativos não capturados na análise automatizada.

---

### 📈 Resumo Executivo para Decisão
"""
    
    # Resumo final
    report += f"""
**Melhor Proposta Técnica:** {melhor_tecnica}
**Melhor Proposta Comercial:** {melhor_comercial}
**Melhor Custo-Benefício:** {melhor_cb}

**Recomendação Geral:** {'Empresa ' + melhor_cb + ' apresenta o melhor equilíbrio entre qualidade técnica e proposta comercial.' if melhor_cb != 'A definir' else 'Realizar análise conjunta dos aspectos técnicos e comerciais, considerando o melhor custo-benefício para o projeto.'}
"""
    
    report += f"""

---

*Relatório gerado automaticamente pelo Proposal Analyzer Pro com Análise de IA Avançada*  
*Data: {current_time}*  
*Versão: 4.0 - Enhanced Technical Analysis with Excel Support*
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
                
                # Análise técnica detalhada com IA
                tech_analysis = analyze_technical_proposal_detailed(content, company)
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
                
                # Análise comercial com IA (incluindo Excel)
                comm_analysis = analyze_commercial_proposal_excel(content, company, cnpj)
                commercial_analyses.append(comm_analysis)
                
                commercial_proposals.append({
                    'company': company,
                    'cnpj': cnpj,
                    'content': content
                })
        
        # Gerar análise comparativa detalhada
        tech_comparison, comm_comparison = generate_detailed_comparative_analysis(
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
                elif line.startswith('### ') or line.startswith('#### '):
                    story.append(Paragraph(line[4:] if line.startswith('### ') else line[5:], subheading_style))
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

