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
import gc  # Garbage collector para liberação de memória

app = Flask(__name__)
CORS(app)

# Configuração de upload
UPLOAD_FOLDER = 'uploads'
if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['MAX_CONTENT_LENGTH'] = 20 * 1024 * 1024  # Reduzido para 20MB

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

def extract_text_from_file_optimized(file_path):
    """Extrai texto de forma otimizada para memória"""
    try:
        file_extension = os.path.splitext(file_path)[1].lower()
        
        if file_extension == '.pdf':
            return extract_pdf_optimized(file_path)
        elif file_extension in ['.doc', '.docx']:
            return extract_docx_optimized(file_path)
        elif file_extension in ['.xls', '.xlsx']:
            return extract_excel_optimized(file_path)
        elif file_extension == '.txt':
            return extract_txt_optimized(file_path)
        elif file_extension == '.zip':
            return extract_zip_optimized(file_path)
        else:
            return "Formato não suportado"
    
    except Exception as e:
        return f"Erro na extração: {str(e)}"

def extract_pdf_optimized(file_path):
    """Extração otimizada de PDF"""
    text_chunks = []
    try:
        with open(file_path, 'rb') as file:
            pdf_reader = PyPDF2.PdfReader(file)
            total_pages = len(pdf_reader.pages)
            
            # Processar em lotes para economizar memória
            batch_size = 5  # Processar 5 páginas por vez
            
            for i in range(0, total_pages, batch_size):
                batch_text = ""
                end_page = min(i + batch_size, total_pages)
                
                for page_num in range(i, end_page):
                    try:
                        page = pdf_reader.pages[page_num]
                        page_text = page.extract_text()
                        batch_text += page_text + "\n"
                        # Limpar referência da página
                        del page
                    except:
                        continue
                
                # Adicionar texto do lote e limpar variável
                if batch_text.strip():
                    text_chunks.append(batch_text)
                del batch_text
                
                # Forçar garbage collection a cada lote
                gc.collect()
        
        # Combinar chunks e limpar lista
        full_text = "\n".join(text_chunks)
        del text_chunks
        gc.collect()
        
        return full_text
        
    except Exception as e:
        return f"Erro PDF: {str(e)}"

def extract_docx_optimized(file_path):
    """Extração otimizada de DOCX"""
    try:
        doc = docx.Document(file_path)
        text_parts = []
        
        # Processar parágrafos em lotes
        batch_size = 50
        paragraphs = doc.paragraphs
        
        for i in range(0, len(paragraphs), batch_size):
            batch_text = ""
            end_idx = min(i + batch_size, len(paragraphs))
            
            for j in range(i, end_idx):
                batch_text += paragraphs[j].text + "\n"
            
            text_parts.append(batch_text)
            del batch_text
            gc.collect()
        
        full_text = "\n".join(text_parts)
        del text_parts, doc
        gc.collect()
        
        return full_text
        
    except Exception as e:
        return f"Erro DOCX: {str(e)}"

def extract_txt_optimized(file_path):
    """Extração otimizada de TXT"""
    try:
        with open(file_path, 'r', encoding='utf-8') as file:
            # Ler em chunks para arquivos grandes
            chunk_size = 8192
            text_chunks = []
            
            while True:
                chunk = file.read(chunk_size)
                if not chunk:
                    break
                text_chunks.append(chunk)
            
            full_text = "".join(text_chunks)
            del text_chunks
            gc.collect()
            
            return full_text
            
    except Exception as e:
        return f"Erro TXT: {str(e)}"

def extract_excel_optimized(file_path):
    """Extração otimizada de Excel"""
    try:
        # Usar openpyxl com read_only para economizar memória
        wb = openpyxl.load_workbook(file_path, read_only=True)
        extracted_data = []
        
        # Processar apenas abas importantes
        important_sheets = ['CARTA', 'Itens Serviços', 'Comp. Custo -GLOBAL', 'BDI']
        sheets_to_process = [sheet for sheet in wb.sheetnames if any(imp in sheet for imp in important_sheets)]
        
        for sheet_name in sheets_to_process:
            try:
                # Usar pandas com chunksize para economizar memória
                df_chunks = pd.read_excel(file_path, sheet_name=sheet_name, chunksize=100)
                
                sheet_text = f"\n=== ABA: {sheet_name} ===\n"
                
                for chunk in df_chunks:
                    # Processar apenas linhas não vazias
                    for index, row in chunk.iterrows():
                        row_text = " | ".join([str(cell) if pd.notna(cell) else "" for cell in row])
                        if row_text.strip() and not row_text.replace(" | ", "").strip() == "":
                            sheet_text += f"Linha {index + 1}: {row_text}\n"
                    
                    # Limpar chunk da memória
                    del chunk
                    gc.collect()
                
                extracted_data.append(sheet_text)
                del sheet_text
                
            except Exception as e:
                extracted_data.append(f"Erro na aba {sheet_name}: {str(e)}")
        
        # Fechar workbook
        wb.close()
        del wb
        
        # Combinar dados
        combined_text = f"ARQUIVO EXCEL: {os.path.basename(file_path)}\n"
        combined_text += "\n".join(extracted_data)
        
        del extracted_data
        gc.collect()
        
        return combined_text
        
    except Exception as e:
        return f"Erro Excel: {str(e)}"

def extract_zip_optimized(file_path):
    """Extração otimizada de ZIP"""
    try:
        extracted_text = ""
        with zipfile.ZipFile(file_path, 'r') as zip_ref:
            temp_dir = tempfile.mkdtemp()
            
            # Extrair apenas arquivos relevantes
            relevant_extensions = ['.pdf', '.docx', '.doc', '.xlsx', '.xls', '.txt']
            
            for file_info in zip_ref.filelist:
                file_ext = os.path.splitext(file_info.filename)[1].lower()
                if file_ext in relevant_extensions:
                    try:
                        zip_ref.extract(file_info, temp_dir)
                        file_path_in_zip = os.path.join(temp_dir, file_info.filename)
                        
                        # Extrair texto e adicionar
                        file_text = extract_text_from_file_optimized(file_path_in_zip)
                        extracted_text += file_text + "\n\n"
                        
                        # Remover arquivo temporário imediatamente
                        os.remove(file_path_in_zip)
                        del file_text
                        gc.collect()
                        
                    except:
                        continue
            
            # Limpar diretório temporário
            import shutil
            shutil.rmtree(temp_dir)
        
        return extracted_text
        
    except Exception as e:
        return f"Erro ZIP: {str(e)}"

def analyze_tr_content_optimized(tr_text):
    """Análise otimizada do TR"""
    # Limitar tamanho do texto para análise
    if len(tr_text) > 50000:  # Limitar a 50k caracteres
        tr_text = tr_text[:50000] + "... [texto truncado para otimização]"
    
    analysis = {
        'resumo': '',
        'objeto': '',
        'requisitos_tecnicos': [],
        'prazos': [],
        'criterios_avaliacao': [],
        'valores_estimados': []
    }
    
    # Processar em chunks menores
    lines = tr_text.split('\n')
    chunk_size = 100
    
    for i in range(0, len(lines), chunk_size):
        chunk_lines = lines[i:i+chunk_size]
        chunk_text = '\n'.join(chunk_lines)
        
        # Análise básica do chunk
        if i == 0:  # Primeiro chunk para resumo
            meaningful_lines = [line.strip() for line in chunk_lines if len(line.strip()) > 50]
            if meaningful_lines:
                analysis['resumo'] = ' '.join(meaningful_lines[:2])
                analysis['objeto'] = meaningful_lines[0] if meaningful_lines else ''
        
        # Buscar prazos
        prazo_matches = re.findall(r'(\d+)\s*(dia|mês|ano|semana)', chunk_text.lower())
        for match in prazo_matches[:3]:  # Limitar a 3 por chunk
            analysis['prazos'].append(f"{match[0]} {match[1]}")
        
        # Buscar valores
        valor_matches = re.findall(r'R\$\s*[\d.,]+', chunk_text)
        analysis['valores_estimados'].extend(valor_matches[:3])  # Limitar a 3 por chunk
        
        # Buscar requisitos técnicos
        tech_keywords = ['técnico', 'especificação', 'requisito', 'metodologia']
        for line in chunk_lines:
            if any(keyword in line.lower() for keyword in tech_keywords) and len(line.strip()) > 30:
                analysis['requisitos_tecnicos'].append(line.strip()[:200])  # Limitar tamanho
                if len(analysis['requisitos_tecnicos']) >= 10:  # Limitar quantidade
                    break
        
        # Limpar chunk da memória
        del chunk_lines, chunk_text
        gc.collect()
    
    # Limitar tamanhos finais
    analysis['requisitos_tecnicos'] = analysis['requisitos_tecnicos'][:10]
    analysis['prazos'] = list(set(analysis['prazos']))[:5]
    analysis['valores_estimados'] = list(set(analysis['valores_estimados']))[:5]
    
    return analysis

def analyze_technical_proposal_optimized(proposal_text, company_name):
    """Análise técnica otimizada"""
    # Limitar tamanho do texto
    if len(proposal_text) > 30000:
        proposal_text = proposal_text[:30000] + "... [truncado]"
    
    analysis = {
        'empresa': company_name,
        'cnpj': '',
        'metodologia': {'descricao': '', 'aderencia_tr': 0},
        'cronograma': {'viabilidade': '', 'marcos_principais': []},
        'equipe_tecnica': {'adequacao_projeto': '', 'qualificacoes': []},
        'recursos_tecnicos': {'equipamentos': [], 'tecnologias': []},
        'experiencia_comprovada': {'projetos_similares': []},
        'pontos_fortes': [],
        'pontos_fracos': [],
        'score_detalhado': {'metodologia': 0, 'cronograma': 0, 'equipe': 0, 'recursos': 0, 'experiencia': 0}
    }
    
    # Extrair CNPJ
    cnpj_match = re.search(r'(\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2})', proposal_text)
    if cnpj_match:
        analysis['cnpj'] = cnpj_match.group(1)
    
    # Análise simplificada por seções
    sections = proposal_text.split('\n')
    
    # Metodologia
    metodologia_sections = [s for s in sections if any(k in s.lower() for k in ['metodologia', 'método']) and len(s) > 50]
    if metodologia_sections:
        analysis['metodologia']['descricao'] = metodologia_sections[0][:300]
        analysis['metodologia']['aderencia_tr'] = 75 if len(metodologia_sections) > 1 else 50
        analysis['score_detalhado']['metodologia'] = analysis['metodologia']['aderencia_tr']
        analysis['pontos_fortes'].append('Metodologia apresentada')
    else:
        analysis['pontos_fracos'].append('Metodologia não detalhada')
        analysis['score_detalhado']['metodologia'] = 20
    
    # Cronograma
    cronograma_sections = [s for s in sections if any(k in s.lower() for k in ['cronograma', 'prazo']) and len(s) > 30]
    if cronograma_sections:
        analysis['cronograma']['viabilidade'] = 'Cronograma apresentado'
        # Extrair marcos
        time_matches = re.findall(r'(\d+)\s*(dia|semana|mês)', ' '.join(cronograma_sections[:3]))
        analysis['cronograma']['marcos_principais'] = [f"{m[0]} {m[1]}" for m in time_matches[:3]]
        analysis['score_detalhado']['cronograma'] = 70
        analysis['pontos_fortes'].append('Cronograma definido')
    else:
        analysis['cronograma']['viabilidade'] = 'Cronograma não apresentado'
        analysis['score_detalhado']['cronograma'] = 20
        analysis['pontos_fracos'].append('Cronograma não detalhado')
    
    # Equipe
    equipe_sections = [s for s in sections if any(k in s.lower() for k in ['equipe', 'profissional', 'coordenador']) and len(s) > 20]
    if equipe_sections:
        analysis['equipe_tecnica']['adequacao_projeto'] = 'Equipe apresentada'
        analysis['equipe_tecnica']['qualificacoes'] = [s[:150] for s in equipe_sections[:3]]
        analysis['score_detalhado']['equipe'] = 65
        analysis['pontos_fortes'].append('Equipe técnica apresentada')
    else:
        analysis['equipe_tecnica']['adequacao_projeto'] = 'Equipe não detalhada'
        analysis['score_detalhado']['equipe'] = 25
        analysis['pontos_fracos'].append('Equipe não detalhada')
    
    # Recursos
    recurso_sections = [s for s in sections if any(k in s.lower() for k in ['equipamento', 'recurso', 'material']) and len(s) > 20]
    if recurso_sections:
        analysis['recursos_tecnicos']['equipamentos'] = [s[:100] for s in recurso_sections[:3]]
        analysis['score_detalhado']['recursos'] = 60
        analysis['pontos_fortes'].append('Recursos técnicos especificados')
    else:
        analysis['score_detalhado']['recursos'] = 30
        analysis['pontos_fracos'].append('Recursos não especificados')
    
    # Experiência
    exp_sections = [s for s in sections if any(k in s.lower() for k in ['experiência', 'projeto', 'referência']) and len(s) > 30]
    if exp_sections:
        analysis['experiencia_comprovada']['projetos_similares'] = [s[:150] for s in exp_sections[:2]]
        analysis['score_detalhado']['experiencia'] = 70
        analysis['pontos_fortes'].append('Experiência comprovada')
    else:
        analysis['score_detalhado']['experiencia'] = 25
        analysis['pontos_fracos'].append('Experiência não comprovada')
    
    # Limpar variáveis grandes
    del sections, proposal_text
    gc.collect()
    
    return analysis

def analyze_commercial_proposal_optimized(proposal_text, company_name, cnpj):
    """Análise comercial otimizada"""
    analysis = {
        'empresa': company_name,
        'cnpj': cnpj,
        'preco_total': '',
        'bdi': '',
        'condicoes_pagamento': '',
        'itens_servicos': []
    }
    
    # Se é Excel, processar diferente
    if "ARQUIVO EXCEL:" in proposal_text:
        return analyze_excel_commercial_optimized(proposal_text, company_name, cnpj)
    
    # Análise simplificada para PDF
    # Buscar preços
    price_matches = re.findall(r'R\$\s*[\d.,]+', proposal_text)
    if price_matches:
        # Converter para comparação e pegar o maior
        prices_with_values = []
        for price in price_matches:
            clean_price = re.sub(r'[^\d,.]', '', price)
            try:
                if ',' in clean_price and '.' in clean_price:
                    clean_price = clean_price.replace('.', '').replace(',', '.')
                elif ',' in clean_price:
                    clean_price = clean_price.replace(',', '.')
                float_value = float(clean_price)
                prices_with_values.append((price, float_value))
            except:
                continue
        
        if prices_with_values:
            analysis['preco_total'] = max(prices_with_values, key=lambda x: x[1])[0]
    
    # Buscar CNPJ se não fornecido
    if not analysis['cnpj']:
        cnpj_match = re.search(r'(\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2})', proposal_text)
        if cnpj_match:
            analysis['cnpj'] = cnpj_match.group(1)
    
    # Buscar BDI
    bdi_match = re.search(r'bdi.*?(\d+[,.]?\d*)%?', proposal_text.lower())
    if bdi_match:
        analysis['bdi'] = bdi_match.group(1) + '%'
    
    # Limpar texto da memória
    del proposal_text
    gc.collect()
    
    return analysis

def analyze_excel_commercial_optimized(excel_text, company_name, cnpj):
    """Análise otimizada de dados comerciais do Excel"""
    analysis = {
        'empresa': company_name,
        'cnpj': cnpj,
        'preco_total': '',
        'bdi': '',
        'condicoes_pagamento': '',
        'itens_servicos': []
    }
    
    lines = excel_text.split('\n')
    
    # Processar apenas linhas relevantes
    in_itens_servicos = False
    precos_encontrados = []
    
    for line in lines:
        if "=== ABA: Itens Serviços ===" in line:
            in_itens_servicos = True
            continue
        elif "=== ABA:" in line and in_itens_servicos:
            in_itens_servicos = False
            continue
        
        if in_itens_servicos and line.strip():
            # Buscar valores numéricos
            valores = re.findall(r'[\d.,]+', line)
            for valor in valores[:3]:  # Limitar a 3 por linha
                try:
                    if '.' in valor and ',' in valor:
                        clean_valor = valor.replace('.', '').replace(',', '.')
                    elif ',' in valor:
                        clean_valor = valor.replace(',', '.')
                    else:
                        clean_valor = valor
                    
                    float_valor = float(clean_valor)
                    if float_valor > 100:
                        precos_encontrados.append((valor, float_valor))
                        if len(precos_encontrados) >= 10:  # Limitar quantidade
                            break
                except:
                    continue
        
        # Buscar BDI
        if "=== ABA: BDI ===" in line:
            # Próximas 20 linhas podem conter BDI
            for i, next_line in enumerate(lines[lines.index(line):lines.index(line)+20]):
                if i == 0:
                    continue
                bdi_matches = re.findall(r'(\d+[,.]?\d*)%?', next_line)
                for match in bdi_matches:
                    try:
                        bdi_val = float(match.replace(',', '.'))
                        if 5 <= bdi_val <= 50:
                            analysis['bdi'] = f"{bdi_val}%"
                            break
                    except:
                        continue
                if analysis['bdi']:
                    break
    
    # Determinar preço total
    if precos_encontrados:
        precos_encontrados.sort(key=lambda x: x[1], reverse=True)
        if len(precos_encontrados) > 5:
            total = sum([p[1] for p in precos_encontrados])
            analysis['preco_total'] = f"R$ {total:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')
        else:
            analysis['preco_total'] = f"R$ {precos_encontrados[0][0]}"
    
    # Limpar variáveis
    del lines, precos_encontrados
    gc.collect()
    
    return analysis

def generate_optimized_report(project_name, project_description, tr_analysis, technical_analyses, commercial_analyses):
    """Gera relatório otimizado"""
    current_time = datetime.now().strftime("%d/%m/%Y %H:%M")
    
    # Calcular rankings
    tech_ranking = []
    for analysis in technical_analyses:
        scores = analysis['score_detalhado']
        avg_score = sum(scores.values()) / len(scores) if scores else 0
        tech_ranking.append((analysis['empresa'], avg_score))
    
    tech_ranking.sort(key=lambda x: x[1], reverse=True)
    
    # Ranking comercial
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
    
    # Gerar relatório simplificado
    report = f"""# 📊 RELATÓRIO DE ANÁLISE DE PROPOSTAS - {project_name.upper()}

**Data:** {current_time}
**Projeto:** {project_name}

---

## 🎯 BLOCO 1: RESUMO DO TERMO DE REFERÊNCIA

### Objeto
{tr_analysis.get('objeto', 'Não identificado')}

### Resumo
{tr_analysis.get('resumo', 'Não disponível')}

### Prazos Identificados
"""
    
    if tr_analysis.get('prazos'):
        for prazo in tr_analysis['prazos'][:5]:
            report += f"- {prazo}\n"
    else:
        report += "Prazos não identificados.\n"
    
    report += """

---

## 🔧 BLOCO 2: EQUALIZAÇÃO TÉCNICA

### Ranking Técnico
"""
    
    if tech_ranking:
        for i, (empresa, score) in enumerate(tech_ranking, 1):
            emoji = "🥇" if i == 1 else "🥈" if i == 2 else "🥉" if i == 3 else "📊"
            report += f"{emoji} **{i}º:** {empresa} - {score:.1f}%\n"
    
    # Análise por empresa (simplificada)
    for analysis in technical_analyses:
        empresa = analysis['empresa']
        cnpj = analysis.get('cnpj', 'Não identificado')
        
        report += f"""

### {empresa}
**CNPJ:** {cnpj}

**Metodologia:** {analysis['metodologia']['descricao'][:200] if analysis['metodologia']['descricao'] else 'Não apresentada'}...

**Cronograma:** {analysis['cronograma']['viabilidade']}

**Pontos Fortes:**
"""
        for ponto in analysis['pontos_fortes'][:3]:
            report += f"✅ {ponto}\n"
        
        report += "\n**Pontos de Atenção:**\n"
        for ponto in analysis['pontos_fracos'][:3]:
            report += f"⚠️ {ponto}\n"
    
    report += """

---

## 💰 BLOCO 3: EQUALIZAÇÃO COMERCIAL

### Ranking de Preços
"""
    
    if comm_ranking:
        for i, (empresa, preco, _) in enumerate(comm_ranking, 1):
            emoji = "🥇" if i == 1 else "🥈" if i == 2 else "🥉" if i == 3 else "📊"
            report += f"{emoji} **{i}º:** {empresa} - {preco}\n"
    else:
        report += "Preços não identificados nas propostas.\n"
    
    # Análise comercial por empresa
    for analysis in commercial_analyses:
        empresa = analysis['empresa']
        report += f"""

### {empresa}
**CNPJ:** {analysis.get('cnpj', 'Não informado')}
**Preço:** {analysis.get('preco_total', 'Não identificado')}
**BDI:** {analysis.get('bdi', 'Não informado')}
"""
    
    report += """

---

## 🎯 BLOCO 4: CONCLUSÃO

### Recomendações
"""
    
    if tech_ranking:
        melhor_tecnica = tech_ranking[0][0]
        report += f"**Melhor Técnica:** {melhor_tecnica}\n"
    
    if comm_ranking:
        melhor_comercial = comm_ranking[0][0]
        report += f"**Melhor Preço:** {melhor_comercial}\n"
    
    report += f"""

**Recomendação:** Analisar conjuntamente os aspectos técnicos e comerciais para a melhor decisão.

---

*Relatório gerado pelo Proposal Analyzer Pro - Versão Otimizada*
*Data: {current_time}*
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
        
        # Processar TR
        tr_file = request.files.get('trFile')
        if not tr_file:
            return jsonify({'success': False, 'error': 'Arquivo do TR é obrigatório.'})
        
        # Salvar e processar TR
        tr_filename = f"tr_{tr_file.filename}"
        tr_path = os.path.join(app.config['UPLOAD_FOLDER'], tr_filename)
        tr_file.save(tr_path)
        
        tr_text = extract_text_from_file_optimized(tr_path)
        tr_analysis = analyze_tr_content_optimized(tr_text)
        
        # Limpar TR da memória
        del tr_text
        os.remove(tr_path)  # Remover arquivo temporário
        gc.collect()
        
        # Processar propostas técnicas
        technical_analyses = []
        tech_companies = request.form.getlist('techCompany[]')
        tech_files = request.files.getlist('techFile[]')
        
        for i, (company, file) in enumerate(zip(tech_companies, tech_files)):
            if company and file and file.filename:
                file_path = os.path.join(app.config['UPLOAD_FOLDER'], f"tech_{i}_{file.filename}")
                file.save(file_path)
                
                content = extract_text_from_file_optimized(file_path)
                tech_analysis = analyze_technical_proposal_optimized(content, company)
                technical_analyses.append(tech_analysis)
                
                # Limpar da memória
                del content
                os.remove(file_path)  # Remover arquivo temporário
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
                
                content = extract_text_from_file_optimized(file_path)
                comm_analysis = analyze_commercial_proposal_optimized(content, company, cnpj)
                commercial_analyses.append(comm_analysis)
                
                # Limpar da memória
                del content
                os.remove(file_path)  # Remover arquivo temporário
                gc.collect()
        
        # Gerar relatório otimizado
        report = generate_optimized_report(
            project_name, project_description, tr_analysis,
            technical_analyses, commercial_analyses
        )
        
        # Salvar relatório
        report_id = f"report_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
        report_path = os.path.join(app.config['UPLOAD_FOLDER'], f"{report_id}.md")
        
        with open(report_path, 'w', encoding='utf-8') as f:
            f.write(report)
        
        # Limpar variáveis grandes
        del report, tr_analysis, technical_analyses, commercial_analyses
        gc.collect()
        
        return jsonify({'success': True, 'report_id': report_id})
        
    except Exception as e:
        # Limpar memória em caso de erro
        gc.collect()
        return jsonify({'success': False, 'error': str(e)})

@app.route('/download/<report_id>/<format>')
def download_report(report_id, format):
    try:
        if format == 'markdown':
            file_path = os.path.join(app.config['UPLOAD_FOLDER'], f"{report_id}.md")
            return send_file(file_path, as_attachment=True, download_name='relatorio_analise.md')
        elif format == 'pdf':
            # Gerar PDF simplificado
            md_file_path = os.path.join(app.config['UPLOAD_FOLDER'], f"{report_id}.md")
            pdf_file_path = os.path.join(app.config['UPLOAD_FOLDER'], f"relatorio_analise.pdf")
            
            # Ler markdown
            with open(md_file_path, 'r', encoding='utf-8') as f:
                markdown_content = f.read()
            
            # Gerar PDF simples com reportlab
            from reportlab.lib.pagesizes import A4
            from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
            from reportlab.lib.styles import getSampleStyleSheet
            from reportlab.lib.units import inch
            
            doc = SimpleDocTemplate(pdf_file_path, pagesize=A4, topMargin=1*inch)
            styles = getSampleStyleSheet()
            story = []
            
            # Processar markdown de forma simples
            lines = markdown_content.split('\n')
            for line in lines[:200]:  # Limitar linhas para economizar memória
                line = line.strip()
                if line.startswith('# '):
                    story.append(Paragraph(line[2:], styles['Title']))
                elif line.startswith('## '):
                    story.append(Paragraph(line[3:], styles['Heading1']))
                elif line.startswith('### '):
                    story.append(Paragraph(line[4:], styles['Heading2']))
                elif line and not line.startswith('*'):
                    story.append(Paragraph(line, styles['Normal']))
                
                if len(story) > 0 and len(story) % 50 == 0:  # Garbage collect periodicamente
                    gc.collect()
            
            doc.build(story)
            
            # Limpar variáveis
            del markdown_content, lines, story
            gc.collect()
            
            return send_file(pdf_file_path, as_attachment=True, download_name='relatorio_analise.pdf')
        else:
            return jsonify({'error': 'Formato não suportado'}), 400
            
    except Exception as e:
        gc.collect()
        return jsonify({'error': str(e)}), 500

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=False)

