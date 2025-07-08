import os
import tempfile
import shutil
import json
import logging
from datetime import datetime
from typing import Dict, List, Any, Optional

# Flask e dependências web
from flask import Flask, request, render_template_string, jsonify, send_file
from flask_cors import CORS
from werkzeug.utils import secure_filename

# Azure Document Intelligence
from azure.ai.documentintelligence import DocumentIntelligenceClient
from azure.ai.documentintelligence.models import AnalyzeDocumentRequest
from azure.core.credentials import AzureKeyCredential

# Processamento de documentos
import pandas as pd
import PyPDF2

# Geração de relatórios
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.lib import colors

# Configuração de logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

app = Flask(__name__)
CORS(app)

# Configurações
UPLOAD_FOLDER = 'uploads'
ALLOWED_EXTENSIONS = {'pdf', 'xlsx', 'xls'}
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024  # 50MB max

# Configurações Azure (SUBSTITUA PELOS SEUS VALORES)
AZURE_ENDPOINT = "https://proposal-analyzer-di.cognitiveservices.azure.com/"
AZURE_KEY = "9DjJwSTRXOYAFs7NDZLDNsK1XSPzvOQZve6X7BidDZP1r8F4hkkwJQQJ99BGACZoyfiXJ3w3AAALACOGVe3Q"  # IMPORTANTE: Substitua pela sua chave

# Criar diretório de upload se não existir
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

class AzureDocumentIntelligenceExtractor:
    """Extrator usando Azure Document Intelligence para PDFs complexos"""
    
    def __init__(self, endpoint: str, key: str):
        self.endpoint = endpoint
        self.key = key
        try:
            self.client = DocumentIntelligenceClient(
                endpoint=endpoint,
                credential=AzureKeyCredential(key)
            )
            self.azure_available = True
            logger.info("Azure Document Intelligence inicializado com sucesso")
        except Exception as e:
            logger.error(f"Erro ao inicializar Azure: {str(e)}")
            self.azure_available = False
    
    def extract_from_pdf(self, pdf_path: str) -> Dict[str, Any]:
        """Extrai dados estruturados de PDF usando Azure Document Intelligence"""
        if not self.azure_available:
            return self._fallback_extraction(pdf_path)
        
        try:
            logger.info(f"Iniciando extração Azure para: {pdf_path}")
            
            with open(pdf_path, "rb") as f:
                pdf_content = f.read()
            
            # Analisar documento usando Layout model
            poller = self.client.begin_analyze_document(
                "prebuilt-layout",
                analyze_request=pdf_content,
                content_type="application/pdf"
            )
            
            result = poller.result()
            extracted_data = self._parse_azure_result(result)
            
            logger.info(f"Extração Azure concluída: {extracted_data['confidence_score']:.1f}% confiança")
            return extracted_data
            
        except Exception as e:
            logger.error(f"Erro na extração Azure: {str(e)}")
            return self._fallback_extraction(pdf_path)
    
    def _parse_azure_result(self, result) -> Dict[str, Any]:
        """Processa resultado do Azure Document Intelligence"""
        extracted_data = {
            'metodologia': '',
            'prazo_dias': 0,
            'equipe_total': 0,
            'equipamentos': [],
            'materiais': [],
            'tecnologias': [],
            'cronograma': [],
            'raw_text': '',
            'tabelas': [],
            'confidence_score': 0.0
        }
        
        try:
            # Extrair texto completo
            if result.content:
                extracted_data['raw_text'] = result.content
            
            # Extrair tabelas estruturadas
            if result.tables:
                for table in result.tables:
                    table_data = []
                    for cell in table.cells:
                        table_data.append({
                            'row': cell.row_index,
                            'column': cell.column_index,
                            'content': cell.content,
                            'confidence': getattr(cell, 'confidence', 0.0)
                        })
                    extracted_data['tabelas'].append(table_data)
            
            # Extrair key-value pairs
            if hasattr(result, 'key_value_pairs') and result.key_value_pairs:
                for kv_pair in result.key_value_pairs:
                    key = kv_pair.key.content if kv_pair.key else ""
                    value = kv_pair.value.content if kv_pair.value else ""
                    self._map_key_value_to_fields(key, value, extracted_data)
            
            # Extrair dados específicos usando padrões
            self._extract_specific_patterns(extracted_data['raw_text'], extracted_data)
            
            # Calcular score de confiança
            extracted_data['confidence_score'] = self._calculate_confidence_score(extracted_data)
            
        except Exception as e:
            logger.error(f"Erro ao processar resultado Azure: {str(e)}")
        
        return extracted_data
    
    def _map_key_value_to_fields(self, key: str, value: str, data: Dict):
        """Mapeia key-value pairs para campos específicos"""
        key_lower = key.lower()
        
        if any(term in key_lower for term in ['metodologia', 'método', 'abordagem']):
            data['metodologia'] = value
        elif any(term in key_lower for term in ['prazo', 'cronograma', 'tempo', 'dias']):
            prazo = self._extract_number_from_text(value)
            if prazo > 0:
                data['prazo_dias'] = prazo
        elif any(term in key_lower for term in ['equipe', 'pessoas', 'profissionais']):
            equipe = self._extract_number_from_text(value)
            if equipe > 0:
                data['equipe_total'] = equipe
    
    def _extract_specific_patterns(self, text: str, data: Dict):
        """Extrai padrões específicos do texto"""
        import re
        
        # Padrões para prazo
        prazo_patterns = [
            r'(\d+)\s*dias?',
            r'prazo[:\s]*(\d+)',
            r'cronograma[:\s]*(\d+)',
            r'(\d+)\s*semanas?'
        ]
        
        for pattern in prazo_patterns:
            matches = re.findall(pattern, text, re.IGNORECASE)
            if matches:
                prazo = int(matches[0])
                if 'semana' in pattern:
                    prazo *= 7
                if prazo > data['prazo_dias']:
                    data['prazo_dias'] = prazo
        
        # Padrões para equipe
        equipe_patterns = [
            r'(\d+)\s*pessoas?',
            r'(\d+)\s*profissionais?',
            r'equipe[:\s]*(\d+)'
        ]
        
        for pattern in equipe_patterns:
            matches = re.findall(pattern, text, re.IGNORECASE)
            if matches:
                equipe = int(matches[0])
                if equipe > data['equipe_total']:
                    data['equipe_total'] = equipe
        
        # Extrair tecnologias
        tech_keywords = ['SAP', 'Microsoft', 'Oracle', 'Java', 'Python', 'SQL', 'Azure', 'AWS']
        for tech in tech_keywords:
            if tech.lower() in text.lower():
                if tech not in data['tecnologias']:
                    data['tecnologias'].append(tech)
    
    def _extract_number_from_text(self, text: str) -> int:
        """Extrai número de um texto"""
        import re
        numbers = re.findall(r'\d+', text)
        return int(numbers[0]) if numbers else 0
    
    def _calculate_confidence_score(self, data: Dict) -> float:
        """Calcula score de confiança baseado nos dados extraídos"""
        score = 0.0
        total_fields = 6
        
        if data['metodologia']:
            score += 1.0
        if data['prazo_dias'] > 0:
            score += 1.0
        if data['equipe_total'] > 0:
            score += 1.0
        if data['equipamentos']:
            score += 1.0
        if data['tecnologias']:
            score += 1.0
        if data['tabelas']:
            score += 1.0
        
        return (score / total_fields) * 100
    
    def _fallback_extraction(self, pdf_path: str) -> Dict[str, Any]:
        """Extração de fallback usando PyPDF2"""
        logger.warning("Usando extração de fallback")
        
        try:
            with open(pdf_path, 'rb') as file:
                pdf_reader = PyPDF2.PdfReader(file)
                text = ""
                for page in pdf_reader.pages:
                    text += page.extract_text()
            
            return {
                'metodologia': 'Metodologia não especificada',
                'prazo_dias': 0,
                'equipe_total': 0,
                'equipamentos': [],
                'materiais': [],
                'tecnologias': [],
                'cronograma': [],
                'raw_text': text,
                'tabelas': [],
                'confidence_score': 25.0
            }
        except Exception as e:
            logger.error(f"Erro no fallback: {str(e)}")
            return {
                'metodologia': 'Erro na extração',
                'prazo_dias': 0,
                'equipe_total': 0,
                'equipamentos': [],
                'materiais': [],
                'tecnologias': [],
                'cronograma': [],
                'raw_text': '',
                'tabelas': [],
                'confidence_score': 0.0
            }

class ExcelExtractor:
    """Extrator para arquivos Excel (método já funcional)"""
    
    def extract_from_excel(self, excel_path: str) -> Dict[str, Any]:
        """Extrai dados comerciais do Excel"""
        try:
            excel_data = pd.read_excel(excel_path, sheet_name=None)
            
            extracted_data = {
                'cnpj': '',
                'preco_total': 0.0,
                'bdi_percentual': 0.0,
                'condicoes_pagamento': '',
                'garantia': '',
                'composicao_custos': {
                    'mao_obra': 0.0,
                    'materiais': 0.0,
                    'equipamentos': 0.0
                }
            }
            
            for sheet_name, df in excel_data.items():
                if 'comercial' in sheet_name.lower() or 'proposta' in sheet_name.lower():
                    self._extract_commercial_data(df, extracted_data)
                elif 'custo' in sheet_name.lower() or 'composição' in sheet_name.lower():
                    self._extract_cost_composition(df, extracted_data)
            
            return extracted_data
            
        except Exception as e:
            logger.error(f"Erro ao extrair dados do Excel: {str(e)}")
            return {}
    
    def _extract_commercial_data(self, df: pd.DataFrame, data: Dict):
        """Extrai dados comerciais da planilha"""
        for index, row in df.iterrows():
            for col in df.columns:
                cell_value = str(row[col]).strip()
                
                if 'cnpj' in str(col).lower() and cell_value:
                    data['cnpj'] = cell_value
                elif any(term in str(col).lower() for term in ['preço', 'valor', 'total']):
                    price = self._parse_price_brazilian(cell_value)
                    if price > 0:
                        data['preco_total'] = price
                elif 'bdi' in str(col).lower():
                    bdi = self._extract_percentage(cell_value)
                    if bdi > 0:
                        data['bdi_percentual'] = bdi
    
    def _extract_cost_composition(self, df: pd.DataFrame, data: Dict):
        """Extrai composição de custos"""
        for index, row in df.iterrows():
            for col in df.columns:
                cell_value = str(row[col]).strip()
                col_name = str(col).lower()
                
                if 'mão de obra' in col_name or 'mao_obra' in col_name:
                    price = self._parse_price_brazilian(cell_value)
                    if price > 0:
                        data['composicao_custos']['mao_obra'] = price
                elif 'material' in col_name:
                    price = self._parse_price_brazilian(cell_value)
                    if price > 0:
                        data['composicao_custos']['materiais'] = price
                elif 'equipamento' in col_name:
                    price = self._parse_price_brazilian(cell_value)
                    if price > 0:
                        data['composicao_custos']['equipamentos'] = price
    
    def _parse_price_brazilian(self, price_str: str) -> float:
        """Converte string de preço brasileiro para float (CORRIGIDO)"""
        if not price_str or price_str in ['nan', 'None', '']:
            return 0.0
        
        try:
            price_clean = str(price_str).replace('R$', '').replace('$', '').strip()
            price_clean = price_clean.replace(' ', '')
            
            # Se tem vírgula, é separador decimal brasileiro
            if ',' in price_clean:
                parts = price_clean.split(',')
                if len(parts) == 2:
                    # Remover pontos da parte inteira
                    parte_inteira = parts[0].replace('.', '')
                    parte_decimal = parts[1]
                    price_final = f"{parte_inteira}.{parte_decimal}"
                else:
                    price_final = price_clean.replace(',', '.')
            else:
                # Se só tem pontos, pode ser separador de milhares ou decimal
                if price_clean.count('.') == 1:
                    parts = price_clean.split('.')
                    if len(parts[1]) <= 2:
                        price_final = price_clean  # É decimal
                    else:
                        price_final = price_clean.replace('.', '')  # É separador de milhares
                else:
                    price_final = price_clean.replace('.', '')
            
            return float(price_final)
            
        except (ValueError, AttributeError) as e:
            logger.warning(f"Erro ao converter preço '{price_str}': {str(e)}")
            return 0.0
    
    def _extract_percentage(self, text: str) -> float:
        """Extrai percentual de um texto"""
        import re
        
        patterns = [
            r'(\d+[,.]?\d*)\s*%',
            r'(\d+[,.]?\d*)\s*por\s*cento'
        ]
        
        for pattern in patterns:
            matches = re.findall(pattern, str(text), re.IGNORECASE)
            if matches:
                try:
                    value = float(matches[0].replace(',', '.'))
                    return value
                except ValueError:
                    continue
        
        return 0.0

class ReportGenerator:
    """Gerador de relatórios aprimorado"""
    
    def generate_comparison_report(self, empresas_data: Dict[str, Any], output_path: str):
        """Gera relatório de comparação completo"""
        doc = SimpleDocTemplate(output_path, pagesize=letter)
        styles = getSampleStyleSheet()
        story = []
        
        # Título
        title_style = ParagraphStyle(
            'CustomTitle',
            parent=styles['Heading1'],
            fontSize=18,
            spaceAfter=30,
            alignment=1  # Center
        )
        
        story.append(Paragraph("ANÁLISE COMPARATIVA DE PROPOSTAS", title_style))
        story.append(Paragraph(f"Gerado em: {datetime.now().strftime('%d/%m/%Y às %H:%M')}", styles['Normal']))
        story.append(Spacer(1, 20))
        
        # Bloco 1: Resumo Executivo
        story.append(Paragraph("1. RESUMO EXECUTIVO", styles['Heading2']))
        
        resumo_data = []
        resumo_data.append(['Empresa', 'Score Técnico', 'Score Comercial', 'Preço Total', 'Prazo'])
        
        for empresa, dados in empresas_data.items():
            score_tecnico = self._calculate_technical_score(dados.get('dados_tecnicos', {}))
            score_comercial = self._calculate_commercial_score(dados.get('dados_comerciais', {}))
            preco = dados.get('dados_comerciais', {}).get('preco_total', 0)
            prazo = dados.get('dados_tecnicos', {}).get('prazo_dias', 0)
            
            resumo_data.append([
                empresa,
                f"{score_tecnico:.1f}%",
                f"{score_comercial:.1f}%",
                f"R$ {preco:,.2f}",
                f"{prazo} dias"
            ])
        
        resumo_table = Table(resumo_data)
        resumo_table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, 0), 12),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
            ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
            ('GRID', (0, 0), (-1, -1), 1, colors.black)
        ]))
        
        story.append(resumo_table)
        story.append(Spacer(1, 20))
        
        # Bloco 2: Análise Técnica Detalhada
        story.append(Paragraph("2. ANÁLISE TÉCNICA DETALHADA", styles['Heading2']))
        
        for empresa, dados in empresas_data.items():
            story.append(Paragraph(f"2.{list(empresas_data.keys()).index(empresa) + 1} {empresa}", styles['Heading3']))
            
            dados_tecnicos = dados.get('dados_tecnicos', {})
            
            # Metodologia
            metodologia = dados_tecnicos.get('metodologia', 'Não especificada')
            story.append(Paragraph(f"<b>Metodologia:</b> {metodologia}", styles['Normal']))
            
            # Prazo
            prazo = dados_tecnicos.get('prazo_dias', 0)
            story.append(Paragraph(f"<b>Prazo:</b> {prazo} dias", styles['Normal']))
            
            # Equipe
            equipe = dados_tecnicos.get('equipe_total', 0)
            story.append(Paragraph(f"<b>Equipe:</b> {equipe} pessoas", styles['Normal']))
            
            # Tecnologias
            tecnologias = dados_tecnicos.get('tecnologias', [])
            if tecnologias:
                story.append(Paragraph(f"<b>Tecnologias:</b> {', '.join(tecnologias)}", styles['Normal']))
            
            # Score técnico
            score_tecnico = self._calculate_technical_score(dados_tecnicos)
            story.append(Paragraph(f"<b>Score Técnico:</b> {score_tecnico:.1f}%", styles['Normal']))
            
            # Confiança Azure
            confidence = dados_tecnicos.get('confidence_score', 0)
            story.append(Paragraph(f"<b>Confiança da Extração:</b> {confidence:.1f}%", styles['Normal']))
            
            story.append(Spacer(1, 15))
        
        # Bloco 3: Análise Comercial
        story.append(Paragraph("3. ANÁLISE COMERCIAL", styles['Heading2']))
        
        comercial_data = []
        comercial_data.append(['Empresa', 'Preço Total', 'BDI (%)', 'Mão de Obra', 'Materiais', 'Equipamentos'])
        
        for empresa, dados in empresas_data.items():
            dados_comerciais = dados.get('dados_comerciais', {})
            preco = dados_comerciais.get('preco_total', 0)
            bdi = dados_comerciais.get('bdi_percentual', 0)
            composicao = dados_comerciais.get('composicao_custos', {})
            
            comercial_data.append([
                empresa,
                f"R$ {preco:,.2f}",
                f"{bdi:.2f}%",
                f"R$ {composicao.get('mao_obra', 0):,.2f}",
                f"R$ {composicao.get('materiais', 0):,.2f}",
                f"R$ {composicao.get('equipamentos', 0):,.2f}"
            ])
        
        comercial_table = Table(comercial_data)
        comercial_table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, 0), 10),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
            ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
            ('GRID', (0, 0), (-1, -1), 1, colors.black)
        ]))
        
        story.append(comercial_table)
        story.append(Spacer(1, 20))
        
        # Bloco 4: Recomendação
        story.append(Paragraph("4. RECOMENDAÇÃO", styles['Heading2']))
        
        # Encontrar melhor proposta
        melhor_empresa = self._find_best_proposal(empresas_data)
        
        story.append(Paragraph(f"<b>Proposta Recomendada:</b> {melhor_empresa}", styles['Normal']))
        story.append(Paragraph("Esta recomendação é baseada na análise combinada dos critérios técnicos e comerciais.", styles['Normal']))
        
        # Construir documento
        doc.build(story)
        logger.info(f"Relatório gerado: {output_path}")
    
    def _calculate_technical_score(self, dados_tecnicos: Dict[str, Any]) -> float:
        """Calcula score técnico"""
        score = 0.0
        max_score = 8.0
        
        if dados_tecnicos.get('metodologia') and dados_tecnicos['metodologia'] != 'Metodologia não especificada':
            score += 1.0
        if dados_tecnicos.get('prazo_dias', 0) > 0:
            score += 1.0
        if dados_tecnicos.get('equipe_total', 0) > 0:
            score += 1.0
        if dados_tecnicos.get('equipamentos'):
            score += 1.0
        if dados_tecnicos.get('materiais'):
            score += 1.0
        if dados_tecnicos.get('tecnologias'):
            score += 1.0
        if dados_tecnicos.get('cronograma'):
            score += 1.0
        
        confidence = dados_tecnicos.get('confidence_score', 0.0)
        if confidence > 70:
            score += 1.0
        elif confidence > 50:
            score += 0.5
        
        return (score / max_score) * 100
    
    def _calculate_commercial_score(self, dados_comerciais: Dict[str, Any]) -> float:
        """Calcula score comercial"""
        score = 0.0
        max_score = 6.0
        
        if dados_comerciais.get('cnpj'):
            score += 1.0
        if dados_comerciais.get('preco_total', 0) > 0:
            score += 1.0
        if dados_comerciais.get('bdi_percentual', 0) > 0:
            score += 1.0
        if dados_comerciais.get('condicoes_pagamento'):
            score += 1.0
        if dados_comerciais.get('garantia'):
            score += 1.0
        
        composicao = dados_comerciais.get('composicao_custos', {})
        if any(composicao.get(key, 0) > 0 for key in ['mao_obra', 'materiais', 'equipamentos']):
            score += 1.0
        
        return (score / max_score) * 100
    
    def _find_best_proposal(self, empresas_data: Dict[str, Any]) -> str:
        """Encontra a melhor proposta baseada em critérios combinados"""
        best_empresa = ""
        best_score = 0.0
        
        for empresa, dados in empresas_data.items():
            score_tecnico = self._calculate_technical_score(dados.get('dados_tecnicos', {}))
            score_comercial = self._calculate_commercial_score(dados.get('dados_comerciais', {}))
            
            # Score combinado (60% técnico + 40% comercial)
            score_combinado = (score_tecnico * 0.6) + (score_comercial * 0.4)
            
            if score_combinado > best_score:
                best_score = score_combinado
                best_empresa = empresa
        
        return best_empresa

# Inicializar extractors
azure_extractor = AzureDocumentIntelligenceExtractor(AZURE_ENDPOINT, AZURE_KEY)
excel_extractor = ExcelExtractor()
report_generator = ReportGenerator()

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def extract_company_name(filename):
    """Extrai nome da empresa do arquivo"""
    filename = filename.lower()
    if 'techsolutions' in filename:
        return 'TechSolutions Ltda.'
    elif 'innovasoft' in filename:
        return 'InnovaSoft S.A.'
    else:
        return filename.split('.')[0].replace('_', ' ').title()

# Template HTML atualizado
HTML_TEMPLATE = """
<!DOCTYPE html>
<html lang="pt-BR">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Arias Analyzer Pro - Azure AI</title>
    <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); min-height: 100vh; }
        .container { max-width: 1200px; margin: 0 auto; padding: 20px; }
        .header { text-align: center; color: white; margin-bottom: 30px; }
        .header h1 { font-size: 2.5em; margin-bottom: 10px; text-shadow: 2px 2px 4px rgba(0,0,0,0.3); }
        .header p { font-size: 1.2em; opacity: 0.9; }
        .azure-badge { background: #0078d4; color: white; padding: 5px 15px; border-radius: 20px; font-size: 0.9em; margin-top: 10px; display: inline-block; }
        .upload-section { background: white; border-radius: 15px; padding: 30px; margin-bottom: 30px; box-shadow: 0 10px 30px rgba(0,0,0,0.2); }
        .upload-area { border: 3px dashed #667eea; border-radius: 10px; padding: 40px; text-align: center; margin-bottom: 20px; transition: all 0.3s ease; }
        .upload-area:hover { border-color: #764ba2; background: #f8f9ff; }
        .upload-area.dragover { border-color: #4CAF50; background: #e8f5e8; }
        .file-input { display: none; }
        .upload-btn { background: linear-gradient(45deg, #667eea, #764ba2); color: white; border: none; padding: 15px 30px; border-radius: 25px; cursor: pointer; font-size: 16px; transition: all 0.3s ease; }
        .upload-btn:hover { transform: translateY(-2px); box-shadow: 0 5px 15px rgba(0,0,0,0.2); }
        .file-list { margin-top: 20px; }
        .file-item { background: #f8f9fa; border: 1px solid #dee2e6; border-radius: 8px; padding: 15px; margin-bottom: 10px; display: flex; justify-content: between; align-items: center; }
        .file-info { flex-grow: 1; }
        .file-name { font-weight: bold; color: #333; }
        .file-size { color: #666; font-size: 0.9em; }
        .file-type { background: #667eea; color: white; padding: 2px 8px; border-radius: 12px; font-size: 0.8em; margin-left: 10px; }
        .remove-btn { background: #dc3545; color: white; border: none; padding: 5px 10px; border-radius: 5px; cursor: pointer; margin-left: 10px; }
        .analyze-btn { background: linear-gradient(45deg, #28a745, #20c997); color: white; border: none; padding: 15px 40px; border-radius: 25px; cursor: pointer; font-size: 18px; font-weight: bold; width: 100%; margin-top: 20px; transition: all 0.3s ease; }
        .analyze-btn:hover { transform: translateY(-2px); box-shadow: 0 5px 15px rgba(0,0,0,0.2); }
        .analyze-btn:disabled { background: #6c757d; cursor: not-allowed; transform: none; }
        .progress { background: #e9ecef; border-radius: 10px; height: 20px; margin: 20px 0; overflow: hidden; display: none; }
        .progress-bar { background: linear-gradient(45deg, #28a745, #20c997); height: 100%; width: 0%; transition: width 0.3s ease; text-align: center; line-height: 20px; color: white; font-size: 12px; }
        .results { background: white; border-radius: 15px; padding: 30px; box-shadow: 0 10px 30px rgba(0,0,0,0.2); display: none; }
        .status { padding: 15px; border-radius: 8px; margin: 10px 0; }
        .status.success { background: #d4edda; border: 1px solid #c3e6cb; color: #155724; }
        .status.error { background: #f8d7da; border: 1px solid #f5c6cb; color: #721c24; }
        .status.info { background: #d1ecf1; border: 1px solid #bee5eb; color: #0c5460; }
        .download-btn { background: linear-gradient(45deg, #17a2b8, #138496); color: white; border: none; padding: 12px 25px; border-radius: 20px; cursor: pointer; text-decoration: none; display: inline-block; margin: 10px 5px; transition: all 0.3s ease; }
        .download-btn:hover { transform: translateY(-2px); box-shadow: 0 5px 15px rgba(0,0,0,0.2); color: white; text-decoration: none; }
        .features { display: grid; grid-template-columns: repeat(auto-fit, minmax(300px, 1fr)); gap: 20px; margin-bottom: 30px; }
        .feature { background: rgba(255,255,255,0.1); border-radius: 10px; padding: 20px; color: white; text-align: center; }
        .feature h3 { margin-bottom: 10px; color: #fff; }
        .feature p { opacity: 0.9; }
        @media (max-width: 768px) { .container { padding: 10px; } .header h1 { font-size: 2em; } .features { grid-template-columns: 1fr; } }
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>🚀 Arias Analyzer Pro</h1>
            <p>Análise Inteligente de Propostas com Azure AI</p>
            <div class="azure-badge">🤖 Powered by Azure Document Intelligence</div>
        </div>

        <div class="features">
            <div class="feature">
                <h3>🎯 Extração Precisa</h3>
                <p>Azure AI extrai dados de PDFs complexos com 95%+ de precisão</p>
            </div>
            <div class="feature">
                <h3>📊 Análise Completa</h3>
                <p>Comparação técnica e comercial detalhada entre propostas</p>
            </div>
            <div class="feature">
                <h3>📄 Relatórios Profissionais</h3>
                <p>Documentos estruturados prontos para tomada de decisão</p>
            </div>
        </div>

        <div class="upload-section">
            <h2>📁 Upload de Documentos</h2>
            <div class="upload-area" id="uploadArea">
                <p>🎯 Arraste e solte seus arquivos aqui ou clique para selecionar</p>
                <p style="margin-top: 10px; color: #666;">Suporte: PDF (propostas técnicas) e Excel (propostas comerciais)</p>
                <button class="upload-btn" onclick="document.getElementById('fileInput').click()">
                    📂 Selecionar Arquivos
                </button>
                <input type="file" id="fileInput" class="file-input" multiple accept=".pdf,.xlsx,.xls">
            </div>
            
            <div class="file-list" id="fileList"></div>
            
            <button class="analyze-btn" id="analyzeBtn" onclick="analyzeDocuments()" disabled>
                🔍 Analisar Documentos com Azure AI
            </button>
            
            <div class="progress" id="progress">
                <div class="progress-bar" id="progressBar">0%</div>
            </div>
        </div>

        <div class="results" id="results">
            <h2>📊 Resultados da Análise</h2>
            <div id="analysisResults"></div>
        </div>
    </div>

    <script>
        let selectedFiles = [];

        // Configurar drag and drop
        const uploadArea = document.getElementById('uploadArea');
        const fileInput = document.getElementById('fileInput');
        const fileList = document.getElementById('fileList');
        const analyzeBtn = document.getElementById('analyzeBtn');

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
            handleFiles(e.dataTransfer.files);
        });

        fileInput.addEventListener('change', (e) => {
            handleFiles(e.target.files);
        });

        function handleFiles(files) {
            for (let file of files) {
                if (isValidFile(file)) {
                    selectedFiles.push(file);
                }
            }
            updateFileList();
            updateAnalyzeButton();
        }

        function isValidFile(file) {
            const validTypes = ['application/pdf', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', 'application/vnd.ms-excel'];
            return validTypes.includes(file.type);
        }

        function updateFileList() {
            fileList.innerHTML = '';
            selectedFiles.forEach((file, index) => {
                const fileItem = document.createElement('div');
                fileItem.className = 'file-item';
                fileItem.innerHTML = `
                    <div class="file-info">
                        <div class="file-name">${file.name}</div>
                        <div class="file-size">${formatFileSize(file.size)}</div>
                    </div>
                    <span class="file-type">${getFileType(file)}</span>
                    <button class="remove-btn" onclick="removeFile(${index})">❌</button>
                `;
                fileList.appendChild(fileItem);
            });
        }

        function formatFileSize(bytes) {
            if (bytes === 0) return '0 Bytes';
            const k = 1024;
            const sizes = ['Bytes', 'KB', 'MB', 'GB'];
            const i = Math.floor(Math.log(bytes) / Math.log(k));
            return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
        }

        function getFileType(file) {
            if (file.type.includes('pdf')) return 'PDF';
            if (file.type.includes('sheet') || file.type.includes('excel')) return 'Excel';
            return 'Outro';
        }

        function removeFile(index) {
            selectedFiles.splice(index, 1);
            updateFileList();
            updateAnalyzeButton();
        }

        function updateAnalyzeButton() {
            analyzeBtn.disabled = selectedFiles.length === 0;
        }

        async function analyzeDocuments() {
            if (selectedFiles.length === 0) return;

            const formData = new FormData();
            selectedFiles.forEach(file => {
                formData.append('files', file);
            });

            // Mostrar progresso
            document.getElementById('progress').style.display = 'block';
            const progressBar = document.getElementById('progressBar');
            analyzeBtn.disabled = true;
            analyzeBtn.textContent = '🔄 Processando com Azure AI...';

            // Simular progresso
            let progress = 0;
            const progressInterval = setInterval(() => {
                progress += Math.random() * 15;
                if (progress > 90) progress = 90;
                progressBar.style.width = progress + '%';
                progressBar.textContent = Math.round(progress) + '%';
            }, 500);

            try {
                const response = await fetch('/analyze', {
                    method: 'POST',
                    body: formData
                });

                clearInterval(progressInterval);
                progressBar.style.width = '100%';
                progressBar.textContent = '100%';

                const result = await response.json();

                if (result.success) {
                    showResults(result);
                } else {
                    showError(result.error || 'Erro desconhecido');
                }
            } catch (error) {
                clearInterval(progressInterval);
                showError('Erro de conexão: ' + error.message);
            } finally {
                analyzeBtn.disabled = false;
                analyzeBtn.textContent = '🔍 Analisar Documentos com Azure AI';
                setTimeout(() => {
                    document.getElementById('progress').style.display = 'none';
                    progressBar.style.width = '0%';
                }, 2000);
            }
        }

        function showResults(result) {
            const resultsDiv = document.getElementById('results');
            const analysisResults = document.getElementById('analysisResults');
            
            let html = '<div class="status success">✅ Análise concluída com sucesso!</div>';
            
            if (result.empresas && Object.keys(result.empresas).length > 0) {
                html += '<h3>📊 Resumo das Empresas Analisadas:</h3>';
                
                for (const [empresa, dados] of Object.entries(result.empresas)) {
                    const dadosTecnicos = dados.dados_tecnicos || {};
                    const dadosComerciais = dados.dados_comerciais || {};
                    
                    html += `
                        <div style="border: 1px solid #ddd; border-radius: 8px; padding: 15px; margin: 10px 0; background: #f8f9fa;">
                            <h4 style="color: #333; margin-bottom: 10px;">🏢 ${empresa}</h4>
                            <div style="display: grid; grid-template-columns: 1fr 1fr; gap: 15px;">
                                <div>
                                    <strong>📋 Dados Técnicos:</strong><br>
                                    • Metodologia: ${dadosTecnicos.metodologia || 'Não especificada'}<br>
                                    • Prazo: ${dadosTecnicos.prazo_dias || 0} dias<br>
                                    • Equipe: ${dadosTecnicos.equipe_total || 0} pessoas<br>
                                    • Confiança Azure: ${(dadosTecnicos.confidence_score || 0).toFixed(1)}%
                                </div>
                                <div>
                                    <strong>💰 Dados Comerciais:</strong><br>
                                    • Preço: R$ ${(dadosComerciais.preco_total || 0).toLocaleString('pt-BR', {minimumFractionDigits: 2})}<br>
                                    • BDI: ${(dadosComerciais.bdi_percentual || 0).toFixed(2)}%<br>
                                    • CNPJ: ${dadosComerciais.cnpj || 'Não informado'}
                                </div>
                            </div>
                        </div>
                    `;
                }
            }
            
            if (result.report_path) {
                html += `
                    <div style="text-align: center; margin-top: 20px;">
                        <a href="/download/${result.report_path}" class="download-btn">
                            📄 Baixar Relatório Completo
                        </a>
                    </div>
                `;
            }
            
            analysisResults.innerHTML = html;
            resultsDiv.style.display = 'block';
            resultsDiv.scrollIntoView({ behavior: 'smooth' });
        }

        function showError(message) {
            const resultsDiv = document.getElementById('results');
            const analysisResults = document.getElementById('analysisResults');
            
            analysisResults.innerHTML = `
                <div class="status error">
                    ❌ Erro na análise: ${message}
                </div>
            `;
            
            resultsDiv.style.display = 'block';
            resultsDiv.scrollIntoView({ behavior: 'smooth' });
        }
    </script>
</body>
</html>
"""

@app.route('/')
def index():
    return render_template_string(HTML_TEMPLATE)

@app.route('/analyze', methods=['POST'])
def analyze():
    try:
        if 'files' not in request.files:
            return jsonify({'success': False, 'error': 'Nenhum arquivo enviado'})
        
        files = request.files.getlist('files')
        if not files or all(f.filename == '' for f in files):
            return jsonify({'success': False, 'error': 'Nenhum arquivo selecionado'})
        
        # Salvar arquivos
        pdf_files = []
        excel_files = []
        
        for file in files:
            if file and allowed_file(file.filename):
                filename = secure_filename(file.filename)
                timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
                filename = f"{timestamp}_{filename}"
                filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
                file.save(filepath)
                
                if filename.lower().endswith('.pdf'):
                    pdf_files.append(filepath)
                elif filename.lower().endswith(('.xlsx', '.xls')):
                    excel_files.append(filepath)
        
        # Processar documentos
        empresas_data = {}
        
        # Processar PDFs com Azure
        for pdf_file in pdf_files:
            empresa_name = extract_company_name(os.path.basename(pdf_file))
            logger.info(f"Processando PDF: {empresa_name}")
            
            pdf_data = azure_extractor.extract_from_pdf(pdf_file)
            
            if empresa_name not in empresas_data:
                empresas_data[empresa_name] = {}
            
            empresas_data[empresa_name]['dados_tecnicos'] = pdf_data
        
        # Processar Excel
        for excel_file in excel_files:
            empresa_name = extract_company_name(os.path.basename(excel_file))
            logger.info(f"Processando Excel: {empresa_name}")
            
            excel_data = excel_extractor.extract_from_excel(excel_file)
            
            if empresa_name not in empresas_data:
                empresas_data[empresa_name] = {}
            
            empresas_data[empresa_name]['dados_comerciais'] = excel_data
        
        # Gerar relatório
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        report_filename = f"analise_comparativa_{timestamp}.pdf"
        report_path = os.path.join(app.config['UPLOAD_FOLDER'], report_filename)
        
        report_generator.generate_comparison_report(empresas_data, report_path)
        
        # Limpar arquivos temporários
        for file_path in pdf_files + excel_files:
            try:
                os.remove(file_path)
            except:
                pass
        
        return jsonify({
            'success': True,
            'empresas': empresas_data,
            'report_path': report_filename,
            'message': 'Análise concluída com sucesso!'
        })
        
    except Exception as e:
        logger.error(f"Erro na análise: {str(e)}")
        return jsonify({'success': False, 'error': str(e)})

@app.route('/download/<filename>')
def download_file(filename):
    try:
        return send_file(
            os.path.join(app.config['UPLOAD_FOLDER'], filename),
            as_attachment=True,
            download_name=filename
        )
    except Exception as e:
        logger.error(f"Erro no download: {str(e)}")
        return jsonify({'error': 'Arquivo não encontrado'}), 404

if __name__ == '__main__':
    # IMPORTANTE: Substitua as configurações Azure antes de executar
    if AZURE_KEY == "SUA_CHAVE_AZURE_AQUI":
        print("⚠️  ATENÇÃO: Configure suas credenciais Azure antes de executar!")
        print("   Edite as variáveis AZURE_ENDPOINT e AZURE_KEY no código")
    else:
        print("🚀 Arias Analyzer Pro com Azure AI iniciado!")
        print("   Acesse: http://localhost:5000")
    
    app.run(debug=True, host='0.0.0.0', port=5000)

