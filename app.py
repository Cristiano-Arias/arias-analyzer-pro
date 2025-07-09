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

class SimpleProposalAnalyzer:
    def __init__(self):
        pass
    
    def extract_text_from_pdf(self, file_path):
        """Extrai texto do PDF usando PyPDF2"""
        try:
            text = ""
            with open(file_path, 'rb') as file:
                pdf_reader = PyPDF2.PdfReader(file)
                for page in pdf_reader.pages:
                    text += page.extract_text() + "\n"
            return text
        except Exception as e:
            logger.error(f"Erro na extração: {e}")
            return ""
    
    def extract_basic_data(self, text, filename):
        """Extrai dados básicos da proposta"""
        data = {
            'arquivo': filename,
            'empresa': 'Não identificado',
            'cnpj': 'Não informado',
            'prazo_dias': 0,
            'valor': 0.0,
            'equipe_total': 0,
            'garantia_anos': 0,
            'telefone': 'Não informado',
            'email': 'Não informado',
            'objeto': 'Não identificado'
        }
        
        # Extrair nome da empresa (primeira linha em maiúscula ou com palavras-chave)
        empresa_patterns = [
            r'^([A-ZÁÊÇÕ\s&-]+(?:LTDA|S\.A\.|EIRELI|ME|EPP|ENGENHARIA|CONSTRUÇÃO|PROJETOS))',
            r'([A-Z][A-Za-z\s&-]+(?:LTDA|ENGENHARIA|CONSTRUÇÃO|PROJETOS)[A-Za-z\s&-]*)',
            r'EMPRESA[:\s]*([^\n]+)',
            r'RAZÃO SOCIAL[:\s]*([^\n]+)'
        ]
        
        for pattern in empresa_patterns:
            match = re.search(pattern, text, re.MULTILINE | re.IGNORECASE)
            if match:
                empresa = match.group(1).strip()
                if len(empresa) > 5:  # Filtrar nomes muito curtos
                    data['empresa'] = empresa[:50]  # Limitar tamanho
                    break
        
        # Extrair CNPJ
        cnpj_patterns = [
            r'CNPJ[:\s]*(\d{2}\.?\d{3}\.?\d{3}/?\d{4}-?\d{2})',
            r'(\d{2}\.?\d{3}\.?\d{3}/?\d{4}-?\d{2})'
        ]
        
        for pattern in cnpj_patterns:
            match = re.search(pattern, text)
            if match:
                data['cnpj'] = match.group(1)
                break
        
        # Extrair prazo (buscar por números seguidos de "dias")
        prazo_patterns = [
            r'(\d+)\s*dias?\s*(?:para|de)?\s*(?:execução|conclusão|prazo)',
            r'prazo[^:]*?(\d+)\s*dias?',
            r'execução[^:]*?(\d+)\s*dias?',
            r'(\d+)\s*dias?\s*corridos'
        ]
        
        prazos_encontrados = []
        for pattern in prazo_patterns:
            matches = re.findall(pattern, text, re.IGNORECASE)
            for match in matches:
                prazo = int(match)
                if 10 <= prazo <= 365:  # Filtrar prazos realistas
                    prazos_encontrados.append(prazo)
        
        if prazos_encontrados:
            data['prazo_dias'] = max(prazos_encontrados)  # Pegar o maior prazo encontrado
        
        # Extrair valor (buscar por valores em reais)
        valor_patterns = [
            r'R\$\s*(\d{1,3}(?:\.\d{3})*(?:,\d{2})?)',
            r'VALOR[^:]*?R\$\s*(\d{1,3}(?:\.\d{3})*(?:,\d{2})?)',
            r'TOTAL[^:]*?R\$\s*(\d{1,3}(?:\.\d{3})*(?:,\d{2})?)'
        ]
        
        for pattern in valor_patterns:
            match = re.search(pattern, text, re.IGNORECASE)
            if match:
                valor_str = match.group(1).replace('.', '').replace(',', '.')
                try:
                    valor = float(valor_str)
                    if valor > 1000:  # Filtrar valores muito baixos
                        data['valor'] = valor
                        break
                except:
                    continue
        
        # Extrair equipe (contar pessoas mencionadas)
        equipe_patterns = [
            r'(\d+)\s*(?:pedreiros?|auxiliares?|eletricistas?|operadores?|técnicos?|engenheiros?)',
            r'equipe[^:]*?(\d+)\s*(?:pessoas?|profissionais?)',
            r'(\d+)\s*(?:pessoas?|profissionais?|colaboradores?)'
        ]
        
        total_equipe = 0
        for pattern in equipe_patterns:
            matches = re.findall(pattern, text, re.IGNORECASE)
            for match in matches:
                num = int(match)
                if 1 <= num <= 50:  # Filtrar números realistas
                    total_equipe += num
        
        data['equipe_total'] = min(total_equipe, 100)  # Limitar a 100 pessoas
        
        # Extrair garantia
        garantia_patterns = [
            r'(\d+)\s*anos?\s*(?:de\s*)?garantia',
            r'garantia[^:]*?(\d+)\s*anos?',
            r'(\d+)\s*anos?\s*(?:para\s*)?(?:obras?|serviços?)'
        ]
        
        for pattern in garantia_patterns:
            match = re.search(pattern, text, re.IGNORECASE)
            if match:
                garantia = int(match.group(1))
                if 1 <= garantia <= 10:  # Filtrar garantias realistas
                    data['garantia_anos'] = garantia
                    break
        
        # Extrair telefone
        telefone_patterns = [
            r'(?:fone|tel|telefone)[:\s]*\(?(\d{2})\)?\s*\d{4,5}[-\.\s]?\d{4}',
            r'\(?(\d{2})\)?\s*\d{4,5}[-\.\s]?\d{4}'
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
        
        # Extrair objeto (primeira frase que menciona serviço/obra)
        objeto_patterns = [
            r'(?:serviço|obra|objeto)[:\s]*([^\n]+)',
            r'contratação[^:]*?([^\n]+)',
            r'execução[^:]*?([^\n]+)'
        ]
        
        for pattern in objeto_patterns:
            match = re.search(pattern, text, re.IGNORECASE)
            if match:
                objeto = match.group(1).strip()
                if len(objeto) > 10:
                    data['objeto'] = objeto[:100]  # Limitar tamanho
                    break
        
        return data
    
    def calculate_score(self, data):
        """Calcula score simples baseado nos dados"""
        score = 0
        
        # Prazo (30 pontos) - quanto menor, melhor
        if data['prazo_dias'] > 0:
            if data['prazo_dias'] <= 30:
                score += 30
            elif data['prazo_dias'] <= 60:
                score += 25
            elif data['prazo_dias'] <= 90:
                score += 20
            elif data['prazo_dias'] <= 120:
                score += 15
            else:
                score += 10
        
        # Equipe (25 pontos)
        if data['equipe_total'] > 0:
            if data['equipe_total'] >= 15:
                score += 25
            elif data['equipe_total'] >= 10:
                score += 20
            elif data['equipe_total'] >= 5:
                score += 15
            else:
                score += 10
        
        # Garantia (20 pontos)
        if data['garantia_anos'] > 0:
            if data['garantia_anos'] >= 5:
                score += 20
            elif data['garantia_anos'] >= 3:
                score += 15
            elif data['garantia_anos'] >= 1:
                score += 10
        
        # Dados completos (25 pontos)
        completude = 0
        if data['empresa'] != 'Não identificado':
            completude += 5
        if data['cnpj'] != 'Não informado':
            completude += 5
        if data['prazo_dias'] > 0:
            completude += 5
        if data['equipe_total'] > 0:
            completude += 5
        if data['telefone'] != 'Não informado':
            completude += 2.5
        if data['email'] != 'Não informado':
            completude += 2.5
        
        score += completude
        
        return min(score, 100)
    
    def analyze_proposals(self, files):
        """Analisa múltiplas propostas"""
        results = []
        
        for file_info in files:
            file_path = file_info['path']
            filename = file_info['original_name']
            
            logger.info(f"Analisando: {filename}")
            
            # Extrair texto
            text = self.extract_text_from_pdf(file_path)
            
            if text:
                # Extrair dados básicos
                data = self.extract_basic_data(text, filename)
                
                # Calcular score
                data['score'] = self.calculate_score(data)
                
                results.append(data)
                logger.info(f"Dados extraídos para {filename}: {data['empresa']}")
            else:
                logger.error(f"Falha na extração para: {filename}")
        
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
            fontSize=20,
            spaceAfter=30,
            textColor=colors.darkblue,
            alignment=TA_CENTER
        )
        
        heading_style = ParagraphStyle(
            'CustomHeading',
            parent=styles['Heading2'],
            fontSize=14,
            spaceAfter=15,
            textColor=colors.darkblue
        )
        
        # Título
        story.append(Paragraph("ANÁLISE COMPARATIVA DE PROPOSTAS", title_style))
        story.append(Paragraph(f"Relatório gerado em: {datetime.now().strftime('%d/%m/%Y às %H:%M')}", styles['Normal']))
        story.append(Spacer(1, 30))
        
        # Ranking geral
        story.append(Paragraph("1. RANKING GERAL", heading_style))
        
        # Ordenar por score
        sorted_proposals = sorted(proposals, key=lambda x: x['score'], reverse=True)
        
        # Tabela de ranking
        ranking_data = [['Pos.', 'Empresa', 'Score', 'Prazo', 'Equipe', 'Garantia']]
        
        for i, prop in enumerate(sorted_proposals, 1):
            empresa_nome = prop['empresa'][:25] + '...' if len(prop['empresa']) > 25 else prop['empresa']
            prazo_str = f"{prop['prazo_dias']} dias" if prop['prazo_dias'] > 0 else 'N/I'
            equipe_str = f"{prop['equipe_total']} pessoas" if prop['equipe_total'] > 0 else 'N/I'
            garantia_str = f"{prop['garantia_anos']} anos" if prop['garantia_anos'] > 0 else 'N/I'
            
            ranking_data.append([
                f"{i}º",
                empresa_nome,
                f"{prop['score']:.0f}%",
                prazo_str,
                equipe_str,
                garantia_str
            ])
        
        ranking_table = Table(ranking_data, colWidths=[0.5*inch, 2.5*inch, 0.8*inch, 1*inch, 1*inch, 0.8*inch])
        ranking_table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.darkblue),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, -1), 9),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
            ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
            ('GRID', (0, 0), (-1, -1), 1, colors.black),
            ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.lightgrey])
        ]))
        
        story.append(ranking_table)
        story.append(Spacer(1, 30))
        
        # Detalhes por empresa
        story.append(Paragraph("2. DETALHES POR EMPRESA", heading_style))
        
        for i, prop in enumerate(sorted_proposals):
            # Nome da empresa
            story.append(Paragraph(f"{i+1}º LUGAR: {prop['empresa']}", 
                                 ParagraphStyle('CompanyTitle', parent=styles['Heading3'], 
                                              textColor=colors.darkgreen, fontSize=12)))
            
            # Tabela de detalhes
            details_data = [
                ['Informação', 'Valor'],
                ['CNPJ', prop['cnpj']],
                ['Telefone', prop['telefone']],
                ['Email', prop['email']],
                ['Prazo de Execução', f"{prop['prazo_dias']} dias" if prop['prazo_dias'] > 0 else 'Não informado'],
                ['Equipe Total', f"{prop['equipe_total']} pessoas" if prop['equipe_total'] > 0 else 'Não informado'],
                ['Garantia', f"{prop['garantia_anos']} anos" if prop['garantia_anos'] > 0 else 'Não informado'],
                ['Valor Proposto', f"R$ {prop['valor']:,.2f}" if prop['valor'] > 0 else 'Não informado'],
                ['Score Final', f"{prop['score']:.0f}%"]
            ]
            
            if prop['objeto'] != 'Não identificado':
                details_data.insert(-1, ['Objeto', prop['objeto'][:60] + '...' if len(prop['objeto']) > 60 else prop['objeto']])
            
            details_table = Table(details_data, colWidths=[2*inch, 3.5*inch])
            details_table.setStyle(TableStyle([
                ('BACKGROUND', (0, 0), (-1, 0), colors.lightblue),
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.black),
                ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
                ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                ('FONTSIZE', (0, 0), (-1, -1), 9),
                ('GRID', (0, 0), (-1, -1), 1, colors.black),
                ('VALIGN', (0, 0), (-1, -1), 'TOP'),
                ('BACKGROUND', (0, -1), (-1, -1), colors.lightgreen)  # Destacar score
            ]))
            
            story.append(details_table)
            story.append(Spacer(1, 20))
        
        # Recomendação
        story.append(Paragraph("3. RECOMENDAÇÃO", heading_style))
        
        if sorted_proposals:
            melhor = sorted_proposals[0]
            story.append(Paragraph(f"Recomenda-se a contratação da empresa {melhor['empresa']} que obteve o melhor score geral ({melhor['score']:.0f}%).", styles['Normal']))
            story.append(Spacer(1, 10))
            
            # Justificativa
            justificativas = []
            if melhor['prazo_dias'] > 0:
                justificativas.append(f"• Prazo adequado: {melhor['prazo_dias']} dias")
            if melhor['equipe_total'] > 0:
                justificativas.append(f"• Equipe dimensionada: {melhor['equipe_total']} pessoas")
            if melhor['garantia_anos'] > 0:
                justificativas.append(f"• Garantia oferecida: {melhor['garantia_anos']} anos")
            if melhor['valor'] > 0:
                justificativas.append(f"• Valor proposto: R$ {melhor['valor']:,.2f}")
            
            if justificativas:
                story.append(Paragraph("Principais fatores:", styles['Heading4']))
                for just in justificativas:
                    story.append(Paragraph(just, styles['Normal']))
        
        # Gerar PDF
        doc.build(story)
        return output_path

# Instanciar analisador
analyzer = SimpleProposalAnalyzer()

@app.route('/')
def index():
    return render_template_string(HTML_TEMPLATE)

@app.route('/upload', methods=['POST'])
def upload_files():
    try:
        files = request.files.getlist('files')
        
        if not files or len(files) < 2:
            return jsonify({'error': 'É necessário enviar pelo menos 2 arquivos'}), 400
        
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
        proposals = analyzer.analyze_proposals(temp_files)
        
        if not proposals:
            return jsonify({'error': 'Falha na análise das propostas'}), 500
        
        # Gerar relatório
        report_filename = f"analise_comparativa_{timestamp}.pdf"
        report_path = os.path.join(upload_dir, report_filename)
        analyzer.generate_comparison_report(proposals, report_path)
        
        logger.info(f"Relatório gerado: {report_path}")
        
        return jsonify({
            'success': True,
            'report_url': f'/download/{report_filename}',
            'proposals_count': len(proposals),
            'summary': {
                'melhor_empresa': proposals[0]['empresa'] if proposals else 'N/A',
                'melhor_score': proposals[0]['score'] if proposals else 0
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

# Template HTML simplificado
HTML_TEMPLATE = '''
<!DOCTYPE html>
<html lang="pt-BR">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Arias Analyzer Pro - Versão Simplificada</title>
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
            display: flex;
            align-items: center;
            justify-content: center;
            padding: 20px;
        }
        
        .container {
            background: white;
            border-radius: 20px;
            box-shadow: 0 20px 40px rgba(0,0,0,0.1);
            padding: 40px;
            max-width: 600px;
            width: 100%;
            text-align: center;
        }
        
        .logo {
            font-size: 2.5em;
            font-weight: bold;
            color: #667eea;
            margin-bottom: 10px;
        }
        
        .subtitle {
            color: #666;
            margin-bottom: 30px;
            font-size: 1.1em;
        }
        
        .version-badge {
            background: #28a745;
            color: white;
            padding: 5px 15px;
            border-radius: 20px;
            font-size: 0.9em;
            margin-bottom: 20px;
            display: inline-block;
        }
        
        .upload-area {
            border: 3px dashed #667eea;
            border-radius: 15px;
            padding: 40px 20px;
            margin: 30px 0;
            background: #f8f9ff;
            transition: all 0.3s ease;
            cursor: pointer;
        }
        
        .upload-area:hover {
            border-color: #764ba2;
            background: #f0f2ff;
        }
        
        .upload-area.dragover {
            border-color: #764ba2;
            background: #e8ebff;
            transform: scale(1.02);
        }
        
        .upload-icon {
            font-size: 3em;
            color: #667eea;
            margin-bottom: 15px;
        }
        
        .upload-text {
            color: #333;
            font-size: 1.1em;
            margin-bottom: 10px;
        }
        
        .upload-hint {
            color: #666;
            font-size: 0.9em;
        }
        
        .file-input {
            display: none;
        }
        
        .analyze-btn {
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
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
        
        .analyze-btn:hover {
            transform: translateY(-2px);
            box-shadow: 0 10px 20px rgba(102, 126, 234, 0.3);
        }
        
        .analyze-btn:disabled {
            background: #ccc;
            cursor: not-allowed;
            transform: none;
            box-shadow: none;
        }
        
        .file-list {
            margin: 20px 0;
            text-align: left;
        }
        
        .file-item {
            background: #f0f2ff;
            padding: 10px 15px;
            margin: 5px 0;
            border-radius: 8px;
            display: flex;
            justify-content: space-between;
            align-items: center;
        }
        
        .file-name {
            color: #333;
            font-weight: 500;
        }
        
        .file-size {
            color: #666;
            font-size: 0.9em;
        }
        
        .progress-container {
            margin: 20px 0;
            display: none;
        }
        
        .progress-bar {
            width: 100%;
            height: 20px;
            background: #f0f0f0;
            border-radius: 10px;
            overflow: hidden;
        }
        
        .progress-fill {
            height: 100%;
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            width: 0%;
            transition: width 0.3s ease;
        }
        
        .progress-text {
            margin-top: 10px;
            color: #666;
        }
        
        .result-container {
            margin-top: 30px;
            display: none;
        }
        
        .success-message {
            background: #d4edda;
            color: #155724;
            padding: 15px;
            border-radius: 10px;
            margin-bottom: 20px;
        }
        
        .download-btn {
            background: #28a745;
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
            background: #218838;
            transform: translateY(-2px);
            box-shadow: 0 5px 15px rgba(40, 167, 69, 0.3);
        }
        
        .error-message {
            background: #f8d7da;
            color: #721c24;
            padding: 15px;
            border-radius: 10px;
            margin-top: 20px;
            display: none;
        }
        
        .features {
            background: #f8f9ff;
            padding: 20px;
            border-radius: 10px;
            margin: 20px 0;
            text-align: left;
        }
        
        .features h3 {
            color: #667eea;
            margin-bottom: 10px;
        }
        
        .features ul {
            list-style: none;
            padding: 0;
        }
        
        .features li {
            padding: 5px 0;
            color: #333;
        }
        
        .features li:before {
            content: "✓ ";
            color: #28a745;
            font-weight: bold;
        }
        
        @media (max-width: 768px) {
            .container {
                padding: 20px;
                margin: 10px;
            }
            
            .logo {
                font-size: 2em;
            }
        }
    </style>
</head>
<body>
    <div class="container">
        <div class="logo">📊 Arias Analyzer Pro</div>
        <div class="version-badge">Versão Simplificada</div>
        <div class="subtitle">Sistema Funcional e Confiável</div>
        
        <div class="features">
            <h3>✨ Características desta versão:</h3>
            <ul>
                <li>Extração confiável com PyPDF2</li>
                <li>Análise de dados essenciais</li>
                <li>Relatório comparativo direto</li>
                <li>Ranking automático por score</li>
                <li>Funciona sempre, sem falhas</li>
            </ul>
        </div>
        
        <div class="upload-area" onclick="document.getElementById('fileInput').click()">
            <div class="upload-icon">📁</div>
            <div class="upload-text">Clique aqui ou arraste os arquivos</div>
            <div class="upload-hint">Aceita apenas PDF (mínimo 2 arquivos)</div>
        </div>
        
        <input type="file" id="fileInput" class="file-input" multiple accept=".pdf">
        
        <div class="file-list" id="fileList"></div>
        
        <button class="analyze-btn" id="analyzeBtn" onclick="analyzeFiles()" disabled>
            Analisar e Comparar Propostas
        </button>
        
        <div class="progress-container" id="progressContainer">
            <div class="progress-bar">
                <div class="progress-fill" id="progressFill"></div>
            </div>
            <div class="progress-text" id="progressText">Processando...</div>
        </div>
        
        <div class="result-container" id="resultContainer">
            <div class="success-message" id="successMessage"></div>
            <a href="#" class="download-btn" id="downloadBtn">📥 Baixar Relatório Comparativo</a>
        </div>
        
        <div class="error-message" id="errorMessage"></div>
    </div>

    <script>
        let selectedFiles = [];
        
        const fileInput = document.getElementById('fileInput');
        const fileList = document.getElementById('fileList');
        const analyzeBtn = document.getElementById('analyzeBtn');
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
            updateAnalyzeButton();
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
        
        function updateAnalyzeButton() {
            analyzeBtn.disabled = selectedFiles.length < 2;
        }
        
        async function analyzeFiles() {
            if (selectedFiles.length < 2) {
                showError('É necessário selecionar pelo menos 2 arquivos PDF.');
                return;
            }
            
            // Mostrar progresso
            document.getElementById('progressContainer').style.display = 'block';
            document.getElementById('resultContainer').style.display = 'none';
            document.getElementById('errorMessage').style.display = 'none';
            analyzeBtn.disabled = true;
            
            // Simular progresso
            let progress = 0;
            const progressInterval = setInterval(() => {
                progress += Math.random() * 20;
                if (progress > 90) progress = 90;
                updateProgress(progress, 'Extraindo dados das propostas...');
            }, 300);
            
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
                updateProgress(100, 'Análise concluída!');
                
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
                analyzeBtn.disabled = false;
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
                ✅ Análise comparativa concluída com sucesso!<br>
                📊 ${result.proposals_count} propostas analisadas<br>
                🏆 Melhor empresa: ${result.summary.melhor_empresa} (${result.summary.melhor_score.toFixed(0)}% score)
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

