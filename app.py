import os
import tempfile
import logging
from datetime import datetime
from flask import Flask, request, jsonify, send_file, render_template_string
from flask_cors import CORS
import re
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_LEFT
import PyPDF2

app = Flask(__name__)
CORS(app)

# Configuração de logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

class ProposalComparator:
    def __init__(self):
        pass
    
    def extract_text_from_pdf(self, file_path):
        """Extrai texto do PDF com timeout e tratamento de erro"""
        try:
            text = ""
            with open(file_path, 'rb') as file:
                pdf_reader = PyPDF2.PdfReader(file)
                
                # Limitar a 10 páginas para evitar timeout
                max_pages = min(len(pdf_reader.pages), 10)
                
                for i in range(max_pages):
                    try:
                        page_text = pdf_reader.pages[i].extract_text()
                        text += page_text + "\n"
                    except Exception as e:
                        logger.warning(f"Erro na página {i}: {e}")
                        continue
                        
            return text
        except Exception as e:
            logger.error(f"Erro na extração do PDF: {e}")
            return ""
    
    def extract_proposal_data(self, text, filename):
        """Extrai dados básicos da proposta de forma robusta"""
        data = {
            'arquivo': filename,
            'empresa': 'Não identificado',
            'cnpj': 'Não informado',
            'prazo': 'Não informado',
            'valor': 'Não informado',
            'equipe': 'Não informado',
            'garantia': 'Não informado',
            'telefone': 'Não informado',
            'email': 'Não informado',
            'endereco': 'Não informado',
            'observacoes': []
        }
        
        # Limpar texto
        text = text.replace('\n', ' ').replace('\r', ' ')
        text = ' '.join(text.split())  # Remover espaços extras
        
        # Extrair nome da empresa (múltiplos padrões)
        empresa_patterns = [
            r'(?:EMPRESA|RAZÃO SOCIAL|PROPONENTE)[:\s]*([A-Z][A-Za-z\s&\-\.]+(?:LTDA|S\.?A\.?|EIRELI|ME|EPP|ENGENHARIA|CONSTRUÇÃO|PROJETOS|CONSULTORIA)[A-Za-z\s&\-\.]*)',
            r'^([A-Z][A-Za-z\s&\-\.]+(?:LTDA|S\.?A\.?|EIRELI|ME|EPP|ENGENHARIA|CONSTRUÇÃO|PROJETOS|CONSULTORIA)[A-Za-z\s&\-\.]*)',
            r'([A-Z]{2,}[A-Za-z\s&\-\.]*(?:LTDA|S\.?A\.?|EIRELI|ME|EPP|ENGENHARIA|CONSTRUÇÃO|PROJETOS|CONSULTORIA)[A-Za-z\s&\-\.]*)'
        ]
        
        for pattern in empresa_patterns:
            matches = re.findall(pattern, text, re.IGNORECASE)
            for match in matches:
                empresa = match.strip()
                if len(empresa) > 5 and len(empresa) < 80:
                    data['empresa'] = empresa
                    break
            if data['empresa'] != 'Não identificado':
                break
        
        # Extrair CNPJ
        cnpj_patterns = [
            r'CNPJ[:\s]*(\d{2}\.?\d{3}\.?\d{3}\/?\d{4}\-?\d{2})',
            r'(\d{2}\.?\d{3}\.?\d{3}\/?\d{4}\-?\d{2})'
        ]
        
        for pattern in cnpj_patterns:
            match = re.search(pattern, text)
            if match:
                data['cnpj'] = match.group(1)
                break
        
        # Extrair prazo (múltiplos formatos)
        prazo_patterns = [
            r'(?:prazo|cronograma|execução)[^:]*?(\d+)\s*(?:dias?|meses?)',
            r'(\d+)\s*(?:dias?|meses?)\s*(?:para|de|corridos|úteis)',
            r'(?:em|dentro de|até)\s*(\d+)\s*(?:dias?|meses?)',
            r'vigência[^:]*?(\d+)\s*(?:dias?|meses?)'
        ]
        
        prazos_encontrados = []
        for pattern in prazo_patterns:
            matches = re.findall(pattern, text, re.IGNORECASE)
            for match in matches:
                try:
                    prazo_num = int(match)
                    if 1 <= prazo_num <= 730:  # Entre 1 dia e 2 anos
                        prazos_encontrados.append(prazo_num)
                except:
                    continue
        
        if prazos_encontrados:
            prazo_principal = max(prazos_encontrados)
            data['prazo'] = f"{prazo_principal} dias"
        
        # Extrair valor
        valor_patterns = [
            r'(?:valor|preço|total)[^:]*?R\$\s*(\d{1,3}(?:\.\d{3})*(?:,\d{2})?)',
            r'R\$\s*(\d{1,3}(?:\.\d{3})*(?:,\d{2})?)',
            r'(?:proposta|orçamento)[^:]*?(\d{1,3}(?:\.\d{3})*(?:,\d{2})?)'
        ]
        
        for pattern in valor_patterns:
            match = re.search(pattern, text, re.IGNORECASE)
            if match:
                valor_str = match.group(1)
                try:
                    # Converter para float para validar
                    valor_num = float(valor_str.replace('.', '').replace(',', '.'))
                    if valor_num > 1000:  # Valor mínimo realista
                        data['valor'] = f"R$ {valor_str}"
                        break
                except:
                    continue
        
        # Extrair equipe
        equipe_patterns = [
            r'(?:equipe|pessoal|profissionais)[^:]*?(\d+)\s*(?:pessoas?|profissionais?)',
            r'(\d+)\s*(?:engenheiros?|técnicos?|operários?|pessoas?)',
            r'(?:composta por|formada por)[^:]*?(\d+)'
        ]
        
        equipe_total = 0
        for pattern in equipe_patterns:
            matches = re.findall(pattern, text, re.IGNORECASE)
            for match in matches:
                try:
                    num = int(match)
                    if 1 <= num <= 100:
                        equipe_total += num
                except:
                    continue
        
        if equipe_total > 0:
            data['equipe'] = f"{equipe_total} pessoas"
        
        # Extrair garantia
        garantia_patterns = [
            r'(?:garantia|warranty)[^:]*?(\d+)\s*(?:anos?|meses?)',
            r'(\d+)\s*(?:anos?|meses?)\s*(?:de\s*)?garantia'
        ]
        
        for pattern in garantia_patterns:
            match = re.search(pattern, text, re.IGNORECASE)
            if match:
                garantia_num = match.group(1)
                try:
                    num = int(garantia_num)
                    if 1 <= num <= 10:
                        data['garantia'] = f"{num} anos"
                        break
                except:
                    continue
        
        # Extrair telefone
        telefone_patterns = [
            r'(?:tel|fone|telefone)[:\s]*\(?(\d{2})\)?\s*\d{4,5}[\-\.\s]?\d{4}',
            r'\(?(\d{2})\)?\s*\d{4,5}[\-\.\s]?\d{4}'
        ]
        
        for pattern in telefone_patterns:
            match = re.search(pattern, text, re.IGNORECASE)
            if match:
                data['telefone'] = match.group(0).strip()
                break
        
        # Extrair email
        email_match = re.search(r'([a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,})', text)
        if email_match:
            data['email'] = email_match.group(1)
        
        # Extrair endereço (básico)
        endereco_patterns = [
            r'(?:endereço|rua|av|avenida)[:\s]*([^,\n]+(?:,\s*[^,\n]+){0,2})',
            r'(?:cep|CEP)[:\s]*\d{5}[\-\.]?\d{3}[^,\n]*([^,\n]+)'
        ]
        
        for pattern in endereco_patterns:
            match = re.search(pattern, text, re.IGNORECASE)
            if match:
                endereco = match.group(1).strip()
                if len(endereco) > 10:
                    data['endereco'] = endereco[:100]
                    break
        
        # Observações importantes (palavras-chave)
        observacoes_keywords = [
            'metodologia', 'scrum', 'agile', 'kanban', 'certificação', 'iso',
            'experiência', 'portfolio', 'referências', 'atestado'
        ]
        
        for keyword in observacoes_keywords:
            if keyword.lower() in text.lower():
                # Extrair contexto da palavra-chave
                pattern = rf'.{{0,50}}{keyword}.{{0,50}}'
                match = re.search(pattern, text, re.IGNORECASE)
                if match:
                    contexto = match.group(0).strip()
                    if len(contexto) > 10:
                        data['observacoes'].append(contexto)
        
        return data
    
    def analyze_proposals(self, files):
        """Analisa múltiplas propostas"""
        results = []
        
        for file_info in files:
            file_path = file_info['path']
            filename = file_info['original_name']
            
            logger.info(f"Analisando: {filename}")
            
            try:
                # Extrair texto com timeout
                text = self.extract_text_from_pdf(file_path)
                
                if text and len(text) > 50:
                    # Extrair dados
                    data = self.extract_proposal_data(text, filename)
                    results.append(data)
                    logger.info(f"Dados extraídos para {filename}: {data['empresa']}")
                else:
                    # Criar entrada vazia se não conseguir extrair
                    data = {
                        'arquivo': filename,
                        'empresa': f'Erro na extração - {filename}',
                        'cnpj': 'Erro na leitura',
                        'prazo': 'Erro na leitura',
                        'valor': 'Erro na leitura',
                        'equipe': 'Erro na leitura',
                        'garantia': 'Erro na leitura',
                        'telefone': 'Erro na leitura',
                        'email': 'Erro na leitura',
                        'endereco': 'Erro na leitura',
                        'observacoes': ['Arquivo não pôde ser processado']
                    }
                    results.append(data)
                    logger.warning(f"Falha na extração para: {filename}")
                    
            except Exception as e:
                logger.error(f"Erro no processamento de {filename}: {e}")
                # Adicionar entrada de erro
                data = {
                    'arquivo': filename,
                    'empresa': f'Erro - {filename}',
                    'cnpj': 'Erro no processamento',
                    'prazo': 'Erro no processamento',
                    'valor': 'Erro no processamento',
                    'equipe': 'Erro no processamento',
                    'garantia': 'Erro no processamento',
                    'telefone': 'Erro no processamento',
                    'email': 'Erro no processamento',
                    'endereco': 'Erro no processamento',
                    'observacoes': [f'Erro: {str(e)}']
                }
                results.append(data)
        
        return results
    
    def generate_comparison_report(self, proposals, output_path):
        """Gera relatório de comparação simples"""
        doc = SimpleDocTemplate(output_path, pagesize=A4, topMargin=0.5*inch)
        styles = getSampleStyleSheet()
        story = []
        
        # Estilos customizados
        title_style = ParagraphStyle(
            'CustomTitle',
            parent=styles['Title'],
            fontSize=18,
            spaceAfter=20,
            textColor=colors.darkblue,
            alignment=TA_CENTER
        )
        
        heading_style = ParagraphStyle(
            'CustomHeading',
            parent=styles['Heading2'],
            fontSize=12,
            spaceAfter=10,
            textColor=colors.darkblue
        )
        
        # Título
        story.append(Paragraph("COMPARAÇÃO DETALHADA DE PROPOSTAS", title_style))
        story.append(Paragraph(f"Relatório gerado em: {datetime.now().strftime('%d/%m/%Y às %H:%M')}", styles['Normal']))
        story.append(Spacer(1, 20))
        
        if not proposals:
            story.append(Paragraph("Nenhuma proposta foi processada com sucesso.", styles['Normal']))
            doc.build(story)
            return output_path
        
        # Tabela comparativa principal
        story.append(Paragraph("TABELA COMPARATIVA", heading_style))
        
        # Preparar dados da tabela
        table_data = [['Critério'] + [f"Proposta {i+1}" for i in range(len(proposals))]]
        
        # Adicionar linhas de dados
        criterios = [
            ('Empresa', 'empresa'),
            ('CNPJ', 'cnpj'),
            ('Prazo', 'prazo'),
            ('Valor', 'valor'),
            ('Equipe', 'equipe'),
            ('Garantia', 'garantia'),
            ('Telefone', 'telefone'),
            ('Email', 'email'),
            ('Endereço', 'endereco')
        ]
        
        for criterio_nome, criterio_key in criterios:
            row = [criterio_nome]
            for prop in proposals:
                valor = prop.get(criterio_key, 'N/I')
                # Limitar tamanho do texto na tabela
                if len(str(valor)) > 25:
                    valor = str(valor)[:22] + '...'
                row.append(str(valor))
            table_data.append(row)
        
        # Criar tabela
        col_widths = [1.2*inch] + [1.8*inch] * len(proposals)
        comparison_table = Table(table_data, colWidths=col_widths)
        
        # Estilo da tabela
        table_style = [
            ('BACKGROUND', (0, 0), (-1, 0), colors.darkblue),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, -1), 8),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
            ('BACKGROUND', (0, 1), (0, -1), colors.lightblue),
            ('GRID', (0, 0), (-1, -1), 1, colors.black),
            ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ]
        
        # Alternar cores das linhas
        for i in range(1, len(table_data)):
            if i % 2 == 0:
                table_style.append(('BACKGROUND', (1, i), (-1, i), colors.lightgrey))
        
        comparison_table.setStyle(TableStyle(table_style))
        story.append(comparison_table)
        story.append(Spacer(1, 30))
        
        # Detalhes individuais
        story.append(Paragraph("DETALHES POR PROPOSTA", heading_style))
        
        for i, prop in enumerate(proposals, 1):
            # Título da proposta
            story.append(Paragraph(f"PROPOSTA {i}: {prop['empresa']}", 
                                 ParagraphStyle('PropTitle', parent=styles['Heading3'], 
                                              textColor=colors.darkgreen, fontSize=11)))
            
            # Informações básicas
            info_text = f"""
            <b>Arquivo:</b> {prop['arquivo']}<br/>
            <b>CNPJ:</b> {prop['cnpj']}<br/>
            <b>Telefone:</b> {prop['telefone']}<br/>
            <b>Email:</b> {prop['email']}<br/>
            <b>Endereço:</b> {prop['endereco']}<br/>
            <b>Prazo:</b> {prop['prazo']}<br/>
            <b>Valor:</b> {prop['valor']}<br/>
            <b>Equipe:</b> {prop['equipe']}<br/>
            <b>Garantia:</b> {prop['garantia']}
            """
            
            story.append(Paragraph(info_text, styles['Normal']))
            
            # Observações
            if prop['observacoes']:
                story.append(Paragraph("<b>Observações:</b>", styles['Heading4']))
                for obs in prop['observacoes'][:3]:  # Limitar a 3 observações
                    story.append(Paragraph(f"• {obs}", styles['Normal']))
            
            story.append(Spacer(1, 15))
        
        # Resumo final
        story.append(Paragraph("RESUMO PARA ANÁLISE", heading_style))
        
        resumo_text = f"""
        Este relatório apresenta a comparação detalhada de {len(proposals)} propostas recebidas.
        Os dados foram extraídos automaticamente dos documentos enviados.
        
        <b>Próximos passos sugeridos:</b><br/>
        1. Verificar a completude dos dados extraídos<br/>
        2. Analisar os critérios mais importantes para sua decisão<br/>
        3. Solicitar esclarecimentos às empresas, se necessário<br/>
        4. Tomar decisão baseada nos critérios estabelecidos
        """
        
        story.append(Paragraph(resumo_text, styles['Normal']))
        
        # Gerar PDF
        doc.build(story)
        return output_path

# Instanciar comparador
comparator = ProposalComparator()

@app.route('/')
def index():
    return render_template_string(HTML_TEMPLATE)

@app.route('/upload', methods=['POST'])
def upload_files():
    try:
        files = request.files.getlist('files')
        
        if not files or len(files) < 1:
            return jsonify({'error': 'É necessário enviar pelo menos 1 arquivo'}), 400
        
        logger.info(f"Processando {len(files)} arquivos")
        
        # Salvar arquivos temporariamente
        temp_files = []
        upload_dir = 'uploads'
        os.makedirs(upload_dir, exist_ok=True)
        
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        
        for file in files:
            if file.filename:
                filename = f"{timestamp}_{file.filename}"
                filepath = os.path.join(upload_dir, filename)
                file.save(filepath)
                temp_files.append({
                    'path': filepath,
                    'original_name': file.filename
                })
        
        # Analisar propostas
        proposals = comparator.analyze_proposals(temp_files)
        
        if not proposals:
            return jsonify({'error': 'Falha na análise das propostas'}), 500
        
        # Gerar relatório
        report_filename = f"comparacao_propostas_{timestamp}.pdf"
        report_path = os.path.join(upload_dir, report_filename)
        comparator.generate_comparison_report(proposals, report_path)
        
        logger.info(f"Relatório gerado: {report_path}")
        
        return jsonify({
            'success': True,
            'report_url': f'/download/{report_filename}',
            'proposals_count': len(proposals),
            'summary': {
                'empresas': [p['empresa'] for p in proposals[:3]]  # Primeiras 3 empresas
            }
        })
        
    except Exception as e:
        logger.error(f"Erro no processamento: {e}")
        return jsonify({'error': f'Erro interno: {str(e)}'}), 500

@app.route('/download/<filename>')
def download_file(filename):
    try:
        file_path = os.path.join('uploads', filename)
        if os.path.exists(file_path):
            return send_file(file_path, as_attachment=True)
        else:
            return jsonify({'error': 'Arquivo não encontrado'}), 404
    except Exception as e:
        logger.error(f"Erro no download: {e}")
        return jsonify({'error': 'Erro no download'}), 500

# Template HTML ultra-simplificado
HTML_TEMPLATE = '''
<!DOCTYPE html>
<html lang="pt-BR">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Comparador de Propostas - Simples e Funcional</title>
    <style>
        * {
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }
        
        body {
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            background: linear-gradient(135deg, #2c3e50 0%, #3498db 100%);
            min-height: 100vh;
            display: flex;
            align-items: center;
            justify-content: center;
            padding: 20px;
        }
        
        .container {
            background: white;
            border-radius: 15px;
            box-shadow: 0 15px 35px rgba(0,0,0,0.1);
            padding: 40px;
            max-width: 600px;
            width: 100%;
            text-align: center;
        }
        
        .logo {
            font-size: 2.2em;
            font-weight: bold;
            color: #2c3e50;
            margin-bottom: 10px;
        }
        
        .subtitle {
            color: #7f8c8d;
            margin-bottom: 30px;
            font-size: 1.1em;
        }
        
        .features {
            background: #ecf0f1;
            padding: 20px;
            border-radius: 10px;
            margin: 20px 0;
            text-align: left;
        }
        
        .features h3 {
            color: #2c3e50;
            margin-bottom: 10px;
            text-align: center;
        }
        
        .features ul {
            list-style: none;
            padding: 0;
        }
        
        .features li {
            padding: 5px 0;
            color: #34495e;
        }
        
        .features li:before {
            content: "✓ ";
            color: #27ae60;
            font-weight: bold;
        }
        
        .upload-area {
            border: 3px dashed #3498db;
            border-radius: 15px;
            padding: 40px 20px;
            margin: 30px 0;
            background: #f8f9fa;
            transition: all 0.3s ease;
            cursor: pointer;
        }
        
        .upload-area:hover {
            border-color: #2980b9;
            background: #e3f2fd;
        }
        
        .upload-area.dragover {
            border-color: #2980b9;
            background: #bbdefb;
            transform: scale(1.02);
        }
        
        .upload-icon {
            font-size: 3em;
            color: #3498db;
            margin-bottom: 15px;
        }
        
        .upload-text {
            color: #2c3e50;
            font-size: 1.1em;
            margin-bottom: 10px;
        }
        
        .upload-hint {
            color: #7f8c8d;
            font-size: 0.9em;
        }
        
        .file-input {
            display: none;
        }
        
        .compare-btn {
            background: linear-gradient(135deg, #3498db 0%, #2980b9 100%);
            color: white;
            border: none;
            padding: 15px 40px;
            border-radius: 50px;
            font-size: 1.1em;
            font-weight: bold;
            cursor: pointer;
            transition: all 0.3s ease;
            margin-top: 20px;
            width: 100%;
            max-width: 300px;
        }
        
        .compare-btn:hover {
            transform: translateY(-2px);
            box-shadow: 0 10px 20px rgba(52, 152, 219, 0.3);
        }
        
        .compare-btn:disabled {
            background: #bdc3c7;
            cursor: not-allowed;
            transform: none;
            box-shadow: none;
        }
        
        .file-list {
            margin: 20px 0;
            text-align: left;
        }
        
        .file-item {
            background: #ecf0f1;
            padding: 10px 15px;
            margin: 5px 0;
            border-radius: 8px;
            display: flex;
            justify-content: space-between;
            align-items: center;
        }
        
        .file-name {
            color: #2c3e50;
            font-weight: 500;
        }
        
        .file-size {
            color: #7f8c8d;
            font-size: 0.9em;
        }
        
        .progress-container {
            margin: 20px 0;
            display: none;
        }
        
        .progress-bar {
            width: 100%;
            height: 20px;
            background: #ecf0f1;
            border-radius: 10px;
            overflow: hidden;
        }
        
        .progress-fill {
            height: 100%;
            background: linear-gradient(135deg, #3498db 0%, #2980b9 100%);
            width: 0%;
            transition: width 0.3s ease;
        }
        
        .progress-text {
            margin-top: 10px;
            color: #7f8c8d;
        }
        
        .result-container {
            margin-top: 30px;
            display: none;
        }
        
        .success-message {
            background: #d5f4e6;
            color: #27ae60;
            padding: 15px;
            border-radius: 10px;
            margin-bottom: 20px;
        }
        
        .download-btn {
            background: #27ae60;
            color: white;
            border: none;
            padding: 12px 30px;
            border-radius: 25px;
            font-size: 1em;
            font-weight: bold;
            cursor: pointer;
            transition: all 0.3s ease;
            text-decoration: none;
            display: inline-block;
        }
        
        .download-btn:hover {
            background: #229954;
            transform: translateY(-2px);
            box-shadow: 0 5px 15px rgba(39, 174, 96, 0.3);
        }
        
        .error-message {
            background: #fadbd8;
            color: #e74c3c;
            padding: 15px;
            border-radius: 10px;
            margin-top: 20px;
            display: none;
        }
        
        @media (max-width: 768px) {
            .container {
                padding: 20px;
                margin: 10px;
            }
            
            .logo {
                font-size: 1.8em;
            }
        }
    </style>
</head>
<body>
    <div class="container">
        <div class="logo">📊 Comparador de Propostas</div>
        <div class="subtitle">Ferramenta Simples e Funcional</div>
        
        <div class="features">
            <h3>🎯 O que esta ferramenta faz:</h3>
            <ul>
                <li>Extrai dados básicos de cada proposta</li>
                <li>Organiza informações em tabela comparativa</li>
                <li>Apresenta detalhes de cada empresa</li>
                <li>Gera relatório PDF para análise</li>
                <li>Funciona com qualquer quantidade de propostas</li>
                <li>Sem scores automáticos - você decide</li>
            </ul>
        </div>
        
        <div class="upload-area" onclick="document.getElementById('fileInput').click()">
            <div class="upload-icon">📁</div>
            <div class="upload-text">Clique aqui ou arraste as propostas</div>
            <div class="upload-hint">Aceita apenas PDF (mínimo 1 arquivo)</div>
        </div>
        
        <input type="file" id="fileInput" class="file-input" multiple accept=".pdf">
        
        <div class="file-list" id="fileList"></div>
        
        <button class="compare-btn" id="compareBtn" onclick="compareProposals()" disabled>
            Comparar Propostas
        </button>
        
        <div class="progress-container" id="progressContainer">
            <div class="progress-bar">
                <div class="progress-fill" id="progressFill"></div>
            </div>
            <div class="progress-text" id="progressText">Processando...</div>
        </div>
        
        <div class="result-container" id="resultContainer">
            <div class="success-message" id="successMessage"></div>
            <a href="#" class="download-btn" id="downloadBtn">📥 Baixar Comparação</a>
        </div>
        
        <div class="error-message" id="errorMessage"></div>
    </div>

    <script>
        let selectedFiles = [];
        
        const fileInput = document.getElementById('fileInput');
        const fileList = document.getElementById('fileList');
        const compareBtn = document.getElementById('compareBtn');
        const uploadArea = document.querySelector('.upload-area');
        
        // Drag and drop
        uploadArea.addEventListener('dragover', (e) => {
            e.preventDefault();
            uploadArea.classList.add('dragover');
        });
        
        uploadArea.addEventListener('dragleave', () => {
            uploadArea.classList.remove('dragover');
        });
        
        uploadArea.addEventListener('drop', (e) => {
            e.preventDefault();
            uploadArea.classList.remove('dragover');
            const files = Array.from(e.dataTransfer.files);
            handleFiles(files);
        });
        
        fileInput.addEventListener('change', (e) => {
            const files = Array.from(e.target.files);
            handleFiles(files);
        });
        
        function handleFiles(files) {
            const validFiles = files.filter(file => {
                const extension = '.' + file.name.split('.').pop().toLowerCase();
                return extension === '.pdf';
            });
            
            selectedFiles = validFiles;
            updateFileList();
            updateCompareButton();
        }
        
        function updateFileList() {
            fileList.innerHTML = '';
            selectedFiles.forEach((file, index) => {
                const fileItem = document.createElement('div');
                fileItem.className = 'file-item';
                fileItem.innerHTML = `
                    <span class="file-name">${file.name}</span>
                    <span class="file-size">${(file.size / 1024 / 1024).toFixed(2)} MB</span>
                `;
                fileList.appendChild(fileItem);
            });
        }
        
        function updateCompareButton() {
            compareBtn.disabled = selectedFiles.length < 1;
        }
        
        async function compareProposals() {
            if (selectedFiles.length < 1) {
                showError('É necessário selecionar pelo menos 1 arquivo PDF.');
                return;
            }
            
            // Mostrar progresso
            document.getElementById('progressContainer').style.display = 'block';
            document.getElementById('resultContainer').style.display = 'none';
            document.getElementById('errorMessage').style.display = 'none';
            compareBtn.disabled = true;
            
            // Simular progresso
            let progress = 0;
            const progressInterval = setInterval(() => {
                progress += Math.random() * 15;
                if (progress > 85) progress = 85;
                updateProgress(progress, 'Extraindo dados das propostas...');
            }, 500);
            
            try {
                const formData = new FormData();
                selectedFiles.forEach(file => {
                    formData.append('files', file);
                });
                
                const response = await fetch('/upload', {
                    method: 'POST',
                    body: formData
                });
                
                clearInterval(progressInterval);
                updateProgress(100, 'Comparação concluída!');
                
                const result = await response.json();
                
                if (result.success) {
                    showSuccess(result);
                } else {
                    showError(result.error || 'Erro desconhecido');
                }
            } catch (error) {
                clearInterval(progressInterval);
                showError('Erro de conexão: ' + error.message);
            } finally {
                compareBtn.disabled = false;
                setTimeout(() => {
                    document.getElementById('progressContainer').style.display = 'none';
                }, 2000);
            }
        }
        
        function updateProgress(percent, text) {
            document.getElementById('progressFill').style.width = percent + '%';
            document.getElementById('progressText').textContent = text;
        }
        
        function showSuccess(result) {
            const successMessage = document.getElementById('successMessage');
            const downloadBtn = document.getElementById('downloadBtn');
            const resultContainer = document.getElementById('resultContainer');
            
            successMessage.innerHTML = `
                ✅ Comparação concluída com sucesso!<br>
                📊 ${result.proposals_count} propostas processadas<br>
                📋 Empresas: ${result.summary.empresas.join(', ')}
            `;
            
            downloadBtn.href = result.report_url;
            resultContainer.style.display = 'block';
        }
        
        function showError(message) {
            const errorMessage = document.getElementById('errorMessage');
            errorMessage.textContent = '❌ ' + message;
            errorMessage.style.display = 'block';
        }
    </script>
</body>
</html>
'''

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=10000, debug=False)

