import os
import tempfile
import shutil
import json
import logging
import re
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

# Processamento de documentos (SEM PANDAS)
import PyPDF2
from openpyxl import load_workbook

# Geração de relatórios
from reportlab.lib.pagesizes import letter, A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch, cm
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT

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

# Configurações Azure
AZURE_ENDPOINT = "https://proposal-analyzer-di.cognitiveservices.azure.com/"
AZURE_KEY = "SUA_CHAVE_AZURE_AQUI"

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
            'empresa': '',
            'cnpj': '',
            'metodologia': '',
            'prazo_dias': 0,
            'equipe_total': 0,
            'equipamentos': [],
            'materiais': [],
            'tecnologias': [],
            'cronograma': [],
            'preco_total': 0.0,
            'bdi_percentual': 0.0,
            'condicoes_pagamento': '',
            'garantia': '',
            'composicao_custos': {
                'mao_obra': 0.0,
                'materiais': 0.0,
                'equipamentos': 0.0
            },
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
        
        if any(term in key_lower for term in ['empresa', 'razão social', 'nome']):
            data['empresa'] = value
        elif any(term in key_lower for term in ['cnpj', 'cpf']):
            cnpj = re.sub(r'[^\d]', '', value)
            if len(cnpj) >= 11:
                data['cnpj'] = value
        elif any(term in key_lower for term in ['metodologia', 'método', 'abordagem']):
            data['metodologia'] = value
        elif any(term in key_lower for term in ['prazo', 'cronograma', 'tempo', 'dias']):
            prazo = self._extract_number_from_text(value)
            if prazo > 0:
                data['prazo_dias'] = prazo
        elif any(term in key_lower for term in ['equipe', 'pessoas', 'profissionais']):
            equipe = self._extract_number_from_text(value)
            if equipe > 0:
                data['equipe_total'] = equipe
        elif any(term in key_lower for term in ['preço', 'valor', 'total']):
            preco = self._extract_currency_from_text(value)
            if preco > 0:
                data['preco_total'] = preco
        elif any(term in key_lower for term in ['bdi', 'benefício']):
            bdi = self._extract_percentage_from_text(value)
            if bdi > 0:
                data['bdi_percentual'] = bdi
    
    def _extract_specific_patterns(self, text: str, data: Dict):
        """Extrai padrões específicos do texto usando regex avançado"""
        
        # Padrões para empresa
        empresa_patterns = [
            r'(?:Empresa|Razão Social|Nome):\s*([A-Za-z\s]+(?:Ltda|S\.A\.|EIRELI)?)',
            r'([A-Za-z\s]+(?:Ltda|S\.A\.|EIRELI))',
        ]
        
        for pattern in empresa_patterns:
            matches = re.findall(pattern, text, re.IGNORECASE)
            if matches and not data['empresa']:
                data['empresa'] = matches[0].strip()
                break
        
        # Padrões para CNPJ
        cnpj_patterns = [
            r'CNPJ[:\s]*(\d{2}\.?\d{3}\.?\d{3}/?\d{4}-?\d{2})',
            r'(\d{2}\.?\d{3}\.?\d{3}/?\d{4}-?\d{2})'
        ]
        
        for pattern in cnpj_patterns:
            matches = re.findall(pattern, text)
            if matches and not data['cnpj']:
                data['cnpj'] = matches[0]
                break
        
        # Padrões para prazo (melhorados)
        prazo_patterns = [
            r'(?:prazo|cronograma|tempo)[:\s]*(\d+)\s*dias?',
            r'(\d+)\s*dias?\s*(?:úteis|corridos)?',
            r'(?:em|dentro de)\s*(\d+)\s*dias?',
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
        
        # Padrões para equipe (melhorados)
        equipe_patterns = [
            r'(?:equipe|pessoas|profissionais)[:\s]*(\d+)',
            r'(\d+)\s*(?:pessoas|profissionais)',
            r'(?:composta por|formada por)\s*(\d+)'
        ]
        
        for pattern in equipe_patterns:
            matches = re.findall(pattern, text, re.IGNORECASE)
            if matches:
                equipe = int(matches[0])
                if equipe > data['equipe_total']:
                    data['equipe_total'] = equipe
        
        # Padrões para preço
        preco_patterns = [
            r'R\$\s*([\d.,]+)',
            r'(?:valor|preço|total)[:\s]*R\$\s*([\d.,]+)',
            r'([\d.,]+)\s*reais?'
        ]
        
        for pattern in preco_patterns:
            matches = re.findall(pattern, text, re.IGNORECASE)
            if matches:
                preco_str = matches[0].replace('.', '').replace(',', '.')
                try:
                    preco = float(preco_str)
                    if preco > data['preco_total']:
                        data['preco_total'] = preco
                except:
                    continue
        
        # Padrões para BDI
        bdi_patterns = [
            r'BDI[:\s]*(\d+(?:,\d+)?)\s*%',
            r'(?:benefício|lucro)[:\s]*(\d+(?:,\d+)?)\s*%'
        ]
        
        for pattern in bdi_patterns:
            matches = re.findall(pattern, text, re.IGNORECASE)
            if matches:
                bdi_str = matches[0].replace(',', '.')
                try:
                    bdi = float(bdi_str)
                    if bdi > data['bdi_percentual']:
                        data['bdi_percentual'] = bdi
                except:
                    continue
        
        # Extrair tecnologias
        tech_keywords = ['SAP', 'Microsoft', 'Oracle', 'Java', 'Python', 'SQL', 'Azure', 'AWS', 'Scrum', 'Kanban']
        for tech in tech_keywords:
            if tech.lower() in text.lower():
                if tech not in data['tecnologias']:
                    data['tecnologias'].append(tech)
        
        # Extrair equipamentos e materiais das tabelas
        self._extract_items_from_tables(data)
    
    def _extract_items_from_tables(self, data: Dict):
        """Extrai equipamentos e materiais das tabelas"""
        for table in data['tabelas']:
            for cell in table:
                content = cell['content'].lower()
                if any(term in content for term in ['servidor', 'computador', 'notebook', 'equipamento']):
                    if cell['content'] not in data['equipamentos']:
                        data['equipamentos'].append(cell['content'])
                elif any(term in content for term in ['licença', 'software', 'material']):
                    if cell['content'] not in data['materiais']:
                        data['materiais'].append(cell['content'])
    
    def _extract_number_from_text(self, text: str) -> int:
        """Extrai número de um texto"""
        numbers = re.findall(r'\d+', text)
        return int(numbers[0]) if numbers else 0
    
    def _extract_currency_from_text(self, text: str) -> float:
        """Extrai valor monetário de um texto"""
        # Remove símbolos e converte para float
        clean_text = re.sub(r'[^\d.,]', '', text)
        clean_text = clean_text.replace('.', '').replace(',', '.')
        try:
            return float(clean_text)
        except:
            return 0.0
    
    def _extract_percentage_from_text(self, text: str) -> float:
        """Extrai percentual de um texto"""
        numbers = re.findall(r'(\d+(?:,\d+)?)', text)
        if numbers:
            try:
                return float(numbers[0].replace(',', '.'))
            except:
                return 0.0
        return 0.0
    
    def _calculate_confidence_score(self, data: Dict) -> float:
        """Calcula score de confiança baseado nos dados extraídos"""
        score = 0.0
        total_fields = 8
        
        if data['empresa']:
            score += 1.0
        if data['cnpj']:
            score += 1.0
        if data['metodologia']:
            score += 1.0
        if data['prazo_dias'] > 0:
            score += 1.0
        if data['equipe_total'] > 0:
            score += 1.0
        if data['preco_total'] > 0:
            score += 1.0
        if data['bdi_percentual'] > 0:
            score += 1.0
        if data['tecnologias']:
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
            
            # Aplicar padrões básicos mesmo no fallback
            data = {
                'empresa': '',
                'cnpj': '',
                'metodologia': 'Metodologia não especificada',
                'prazo_dias': 0,
                'equipe_total': 0,
                'equipamentos': [],
                'materiais': [],
                'tecnologias': [],
                'cronograma': [],
                'preco_total': 0.0,
                'bdi_percentual': 0.0,
                'condicoes_pagamento': '',
                'garantia': '',
                'composicao_custos': {
                    'mao_obra': 0.0,
                    'materiais': 0.0,
                    'equipamentos': 0.0
                },
                'raw_text': text,
                'tabelas': [],
                'confidence_score': 25.0
            }
            
            self._extract_specific_patterns(text, data)
            return data
            
        except Exception as e:
            logger.error(f"Erro no fallback: {str(e)}")
            return {
                'empresa': 'Erro na extração',
                'cnpj': '',
                'metodologia': 'Erro na extração',
                'prazo_dias': 0,
                'equipe_total': 0,
                'equipamentos': [],
                'materiais': [],
                'tecnologias': [],
                'cronograma': [],
                'preco_total': 0.0,
                'bdi_percentual': 0.0,
                'condicoes_pagamento': '',
                'garantia': '',
                'composicao_custos': {
                    'mao_obra': 0.0,
                    'materiais': 0.0,
                    'equipamentos': 0.0
                },
                'raw_text': '',
                'tabelas': [],
                'confidence_score': 0.0
            }

class ExcelExtractor:
    """Extrator para arquivos Excel SEM PANDAS - usando openpyxl"""
    
    def extract_from_excel(self, excel_path: str) -> Dict[str, Any]:
        """Extrai dados comerciais do Excel usando openpyxl"""
        try:
            workbook = load_workbook(excel_path, read_only=True)
            
            extracted_data = {
                'empresa': '',
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
            
            for sheet_name in workbook.sheetnames:
                sheet = workbook[sheet_name]
                
                if any(term in sheet_name.lower() for term in ['comercial', 'proposta']):
                    self._extract_commercial_data(sheet, extracted_data)
                elif any(term in sheet_name.lower() for term in ['custo', 'composição']):
                    self._extract_cost_composition(sheet, extracted_data)
            
            workbook.close()
            return extracted_data
            
        except Exception as e:
            logger.error(f"Erro ao extrair dados do Excel: {str(e)}")
            return {}
    
    def _extract_commercial_data(self, sheet, data: Dict):
        """Extrai dados comerciais da planilha"""
        for row in sheet.iter_rows(values_only=True):
            if not row or not any(row):
                continue
            
            row_text = ' '.join(str(cell) for cell in row if cell is not None).lower()
            
            # Buscar CNPJ
            cnpj_match = re.search(r'(\d{2}\.?\d{3}\.?\d{3}/?\d{4}-?\d{2})', row_text)
            if cnpj_match and not data['cnpj']:
                data['cnpj'] = cnpj_match.group(1)
            
            # Buscar preço total
            if any(term in row_text for term in ['total', 'preço', 'valor']):
                for cell in row:
                    if isinstance(cell, (int, float)) and cell > 1000:
                        data['preco_total'] = float(cell)
                    elif isinstance(cell, str) and 'R$' in str(cell):
                        price_match = re.search(r'R\$\s*([\d.,]+)', str(cell))
                        if price_match:
                            price_str = price_match.group(1).replace('.', '').replace(',', '.')
                            try:
                                data['preco_total'] = float(price_str)
                            except:
                                pass
            
            # Buscar BDI
            if 'bdi' in row_text:
                for cell in row:
                    if isinstance(cell, (int, float)) and 0 < cell < 100:
                        data['bdi_percentual'] = float(cell)
    
    def _extract_cost_composition(self, sheet, data: Dict):
        """Extrai composição de custos da planilha"""
        for row in sheet.iter_rows(values_only=True):
            if not row or not any(row):
                continue
            
            row_text = ' '.join(str(cell) for cell in row if cell is not None).lower()
            
            # Buscar valores por categoria
            value = None
            for cell in row:
                if isinstance(cell, (int, float)) and cell > 0:
                    value = float(cell)
                    break
            
            if value:
                if any(term in row_text for term in ['mão de obra', 'mao de obra', 'pessoal']):
                    data['composicao_custos']['mao_obra'] = value
                elif any(term in row_text for term in ['material', 'insumo']):
                    data['composicao_custos']['materiais'] = value
                elif any(term in row_text for term in ['equipamento', 'hardware']):
                    data['composicao_custos']['equipamentos'] = value



class ProposalAnalyzer:
    """Analisador de propostas SEM PANDAS - usando estruturas Python nativas"""
    
    def __init__(self):
        self.azure_extractor = AzureDocumentIntelligenceExtractor(AZURE_ENDPOINT, AZURE_KEY)
        self.excel_extractor = ExcelExtractor()
    
    def analyze_proposals(self, files: List[Dict]) -> Dict[str, Any]:
        """Analisa múltiplas propostas e gera comparação"""
        proposals = []
        
        for file_info in files:
            file_path = file_info['path']
            file_type = file_info['type']
            
            try:
                if file_type == 'pdf':
                    data = self.azure_extractor.extract_from_pdf(file_path)
                elif file_type in ['xlsx', 'xls']:
                    data = self.excel_extractor.extract_from_excel(file_path)
                else:
                    continue
                
                # Adicionar informações do arquivo
                data['filename'] = file_info['filename']
                data['file_type'] = file_type
                proposals.append(data)
                
            except Exception as e:
                logger.error(f"Erro ao processar {file_info['filename']}: {str(e)}")
                continue
        
        if not proposals:
            raise ValueError("Nenhuma proposta válida foi processada")
        
        # Consolidar dados (mesclar PDF + Excel da mesma empresa)
        consolidated_proposals = self._consolidate_proposals(proposals)
        
        # Calcular scores técnicos
        self._calculate_technical_scores(consolidated_proposals)
        
        # Ordenar por ranking
        consolidated_proposals.sort(key=lambda x: x.get('score_tecnico', 0), reverse=True)
        
        return {
            'proposals': consolidated_proposals,
            'summary': self._generate_summary(consolidated_proposals),
            'analysis_date': datetime.now().strftime('%d/%m/%Y às %H:%M')
        }
    
    def _consolidate_proposals(self, proposals: List[Dict]) -> List[Dict]:
        """Consolida dados de PDF e Excel da mesma empresa"""
        consolidated = {}
        
        for proposal in proposals:
            empresa = proposal.get('empresa', '').strip()
            
            # Tentar identificar empresa por CNPJ se nome não estiver disponível
            if not empresa:
                cnpj = proposal.get('cnpj', '')
                if cnpj:
                    empresa = f"Empresa {cnpj[:8]}"
                else:
                    empresa = f"Empresa {len(consolidated) + 1}"
            
            if empresa not in consolidated:
                consolidated[empresa] = {
                    'empresa': empresa,
                    'cnpj': '',
                    'metodologia': '',
                    'prazo_dias': 0,
                    'equipe_total': 0,
                    'equipamentos': [],
                    'materiais': [],
                    'tecnologias': [],
                    'preco_total': 0.0,
                    'bdi_percentual': 0.0,
                    'condicoes_pagamento': '',
                    'garantia': '',
                    'composicao_custos': {
                        'mao_obra': 0.0,
                        'materiais': 0.0,
                        'equipamentos': 0.0
                    },
                    'confidence_score': 0.0,
                    'score_tecnico': 0.0,
                    'files_processed': []
                }
            
            # Mesclar dados
            current = consolidated[empresa]
            
            # Atualizar campos se não estiverem preenchidos ou se o novo valor for melhor
            if proposal.get('cnpj') and not current['cnpj']:
                current['cnpj'] = proposal['cnpj']
            
            if proposal.get('metodologia') and not current['metodologia']:
                current['metodologia'] = proposal['metodologia']
            
            if proposal.get('prazo_dias', 0) > current['prazo_dias']:
                current['prazo_dias'] = proposal['prazo_dias']
            
            if proposal.get('equipe_total', 0) > current['equipe_total']:
                current['equipe_total'] = proposal['equipe_total']
            
            if proposal.get('preco_total', 0) > current['preco_total']:
                current['preco_total'] = proposal['preco_total']
            
            if proposal.get('bdi_percentual', 0) > current['bdi_percentual']:
                current['bdi_percentual'] = proposal['bdi_percentual']
            
            # Mesclar listas
            for item in proposal.get('equipamentos', []):
                if item not in current['equipamentos']:
                    current['equipamentos'].append(item)
            
            for item in proposal.get('materiais', []):
                if item not in current['materiais']:
                    current['materiais'].append(item)
            
            for item in proposal.get('tecnologias', []):
                if item not in current['tecnologias']:
                    current['tecnologias'].append(item)
            
            # Atualizar composição de custos
            for key, value in proposal.get('composicao_custos', {}).items():
                if value > current['composicao_custos'].get(key, 0):
                    current['composicao_custos'][key] = value
            
            # Atualizar confidence score (usar o maior)
            if proposal.get('confidence_score', 0) > current['confidence_score']:
                current['confidence_score'] = proposal['confidence_score']
            
            # Adicionar arquivo processado
            current['files_processed'].append(proposal.get('filename', 'unknown'))
        
        return list(consolidated.values())
    
    def _calculate_technical_scores(self, proposals: List[Dict]):
        """Calcula scores técnicos baseado nos dados extraídos"""
        for proposal in proposals:
            score = 0.0
            max_score = 100.0
            
            # Metodologia (25 pontos)
            if proposal['metodologia'] and proposal['metodologia'] != 'Metodologia não especificada':
                metodologia_score = 25.0
                # Bonus para metodologias ágeis
                if any(term in proposal['metodologia'].lower() for term in ['scrum', 'kanban', 'ágil', 'agile']):
                    metodologia_score = 25.0
                elif any(term in proposal['metodologia'].lower() for term in ['cascata', 'waterfall']):
                    metodologia_score = 15.0
                else:
                    metodologia_score = 20.0
                score += metodologia_score
            
            # Prazo (20 pontos)
            if proposal['prazo_dias'] > 0:
                if proposal['prazo_dias'] <= 90:
                    score += 20.0  # Prazo excelente
                elif proposal['prazo_dias'] <= 120:
                    score += 15.0  # Prazo bom
                elif proposal['prazo_dias'] <= 150:
                    score += 10.0  # Prazo aceitável
                else:
                    score += 5.0   # Prazo ruim
            
            # Equipe (20 pontos)
            if proposal['equipe_total'] > 0:
                if proposal['equipe_total'] >= 8:
                    score += 20.0  # Equipe robusta
                elif proposal['equipe_total'] >= 5:
                    score += 15.0  # Equipe adequada
                elif proposal['equipe_total'] >= 3:
                    score += 10.0  # Equipe mínima
                else:
                    score += 5.0   # Equipe insuficiente
            
            # Recursos técnicos (15 pontos)
            recursos_score = 0
            if proposal['equipamentos']:
                recursos_score += 7.5
            if proposal['materiais']:
                recursos_score += 7.5
            score += recursos_score
            
            # Tecnologias (10 pontos)
            if proposal['tecnologias']:
                tech_score = min(len(proposal['tecnologias']) * 2, 10)
                score += tech_score
            
            # Completude dos dados (10 pontos)
            completude_score = (proposal['confidence_score'] / 100) * 10
            score += completude_score
            
            proposal['score_tecnico'] = round(score, 1)
    
    def _generate_summary(self, proposals: List[Dict]) -> Dict[str, Any]:
        """Gera resumo da análise"""
        if not proposals:
            return {}
        
        # Encontrar melhor e pior proposta
        best_technical = max(proposals, key=lambda x: x['score_tecnico'])
        best_price = min([p for p in proposals if p['preco_total'] > 0], 
                        key=lambda x: x['preco_total'], default=None)
        
        # Calcular estatísticas
        precos = [p['preco_total'] for p in proposals if p['preco_total'] > 0]
        prazos = [p['prazo_dias'] for p in proposals if p['prazo_dias'] > 0]
        
        return {
            'total_proposals': len(proposals),
            'best_technical': best_technical['empresa'] if best_technical else '',
            'best_price': best_price['empresa'] if best_price else '',
            'price_range': {
                'min': min(precos) if precos else 0,
                'max': max(precos) if precos else 0,
                'avg': sum(precos) / len(precos) if precos else 0
            },
            'deadline_range': {
                'min': min(prazos) if prazos else 0,
                'max': max(prazos) if prazos else 0,
                'avg': sum(prazos) / len(prazos) if prazos else 0
            }
        }

class ReportGenerator:
    """Gerador de relatórios com formatação visual profissional"""
    
    def __init__(self):
        self.styles = getSampleStyleSheet()
        self._setup_custom_styles()
    
    def _setup_custom_styles(self):
        """Configura estilos personalizados para o relatório"""
        # Título principal
        self.styles.add(ParagraphStyle(
            name='CustomTitle',
            parent=self.styles['Title'],
            fontSize=24,
            spaceAfter=30,
            alignment=TA_CENTER,
            textColor=colors.darkblue,
            fontName='Helvetica-Bold'
        ))
        
        # Subtítulo
        self.styles.add(ParagraphStyle(
            name='CustomSubtitle',
            parent=self.styles['Heading1'],
            fontSize=16,
            spaceAfter=20,
            spaceBefore=20,
            textColor=colors.darkblue,
            fontName='Helvetica-Bold'
        ))
        
        # Cabeçalho de seção
        self.styles.add(ParagraphStyle(
            name='SectionHeader',
            parent=self.styles['Heading2'],
            fontSize=14,
            spaceAfter=15,
            spaceBefore=15,
            textColor=colors.darkgreen,
            fontName='Helvetica-Bold',
            borderWidth=1,
            borderColor=colors.darkgreen,
            borderPadding=5
        ))
        
        # Texto normal melhorado
        self.styles.add(ParagraphStyle(
            name='CustomNormal',
            parent=self.styles['Normal'],
            fontSize=11,
            spaceAfter=10,
            fontName='Helvetica'
        ))
    
    def generate_report(self, analysis_result: Dict[str, Any], output_path: str):
        """Gera relatório PDF com formatação profissional"""
        doc = SimpleDocTemplate(
            output_path,
            pagesize=A4,
            rightMargin=2*cm,
            leftMargin=2*cm,
            topMargin=2*cm,
            bottomMargin=2*cm
        )
        
        story = []
        
        # Cabeçalho do relatório
        self._add_header(story, analysis_result)
        
        # Bloco 1: Resumo do TR
        self._add_tr_summary(story)
        
        # Bloco 2: Equalização Técnica
        self._add_technical_analysis(story, analysis_result['proposals'])
        
        # Bloco 3: Equalização Comercial
        self._add_commercial_analysis(story, analysis_result['proposals'])
        
        # Bloco 4: Conclusão
        self._add_conclusion(story, analysis_result)
        
        doc.build(story)
        logger.info(f"Relatório gerado: {output_path}")
    
    def _add_header(self, story: List, analysis_result: Dict):
        """Adiciona cabeçalho profissional"""
        # Título principal
        title = Paragraph("ANÁLISE COMPARATIVA DE PROPOSTAS", self.styles['CustomTitle'])
        story.append(title)
        
        # Data de geração
        date_text = f"<b>Data de Geração:</b> {analysis_result['analysis_date']}"
        date_para = Paragraph(date_text, self.styles['CustomNormal'])
        story.append(date_para)
        
        # Linha separadora
        story.append(Spacer(1, 20))
        story.append(self._create_separator_line())
        story.append(Spacer(1, 20))
    
    def _add_tr_summary(self, story: List):
        """Adiciona resumo do Termo de Referência"""
        # Título da seção
        section_title = Paragraph("📋 BLOCO 1: RESUMO DO TERMO DE REFERÊNCIA", self.styles['SectionHeader'])
        story.append(section_title)
        
        # Objeto
        story.append(Paragraph("<b>Objeto</b>", self.styles['CustomSubtitle']))
        story.append(Paragraph("Sistema de Gestão Empresarial", self.styles['CustomNormal']))
        
        # Especificações técnicas
        story.append(Paragraph("<b>Especificações Técnicas Exigidas</b>", self.styles['CustomSubtitle']))
        specs = [
            "• Sistema integrado de gestão",
            "• Módulos: Financeiro, Estoque, Vendas, Compras",
            "• Interface web responsiva",
            "• Banco de dados robusto",
            "• Relatórios gerenciais"
        ]
        for spec in specs:
            story.append(Paragraph(spec, self.styles['CustomNormal']))
        
        # Metodologia exigida
        story.append(Paragraph("<b>Metodologia Exigida pelo TR</b>", self.styles['CustomSubtitle']))
        metodologia = [
            "• Metodologia ágil ou híbrida",
            "• Fases bem definidas: Análise, Desenvolvimento, Testes, Implantação",
            "• Documentação técnica completa",
            "• Treinamento da equipe"
        ]
        for item in metodologia:
            story.append(Paragraph(item, self.styles['CustomNormal']))
        
        # Prazos e critérios
        story.append(Paragraph("<b>Prazos e Critérios</b>", self.styles['CustomSubtitle']))
        story.append(Paragraph("• <b>Prazo máximo:</b> 120 dias", self.styles['CustomNormal']))
        story.append(Paragraph("• <b>Critérios de avaliação:</b> Técnica (70%) + Preço (30%)", self.styles['CustomNormal']))
        
        story.append(self._create_separator_line())
    
    def _add_technical_analysis(self, story: List, proposals: List[Dict]):
        """Adiciona análise técnica detalhada"""
        # Título da seção
        section_title = Paragraph("🔧 BLOCO 2: EQUALIZAÇÃO DAS PROPOSTAS TÉCNICAS", self.styles['SectionHeader'])
        story.append(section_title)
        
        # Matriz de comparação técnica
        story.append(Paragraph("📊 Matriz de Comparação Técnica", self.styles['CustomSubtitle']))
        
        # Criar tabela de comparação
        table_data = [['Empresa', 'Metodologia', 'Prazo', 'Equipe', 'Equipamentos', 'Materiais', 'Score Total']]
        
        for proposal in proposals:
            metodologia_check = "✓" if proposal['metodologia'] and proposal['metodologia'] != 'Metodologia não especificada' else "✗"
            prazo_check = "✓" if proposal['prazo_dias'] > 0 and proposal['prazo_dias'] <= 120 else "✗"
            equipe_check = "✓" if proposal['equipe_total'] >= 5 else "✗"
            equipamentos_check = "✓" if proposal['equipamentos'] else "✗"
            materiais_check = "✓" if proposal['materiais'] else "✗"
            
            table_data.append([
                f"<b>{proposal['empresa']}</b>",
                metodologia_check,
                prazo_check,
                equipe_check,
                equipamentos_check,
                materiais_check,
                f"<b>{proposal['score_tecnico']:.1f}%</b>"
            ])
        
        table = Table(table_data, colWidths=[3*cm, 2*cm, 1.5*cm, 1.5*cm, 2*cm, 2*cm, 2*cm])
        table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.darkblue),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, 0), 10),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
            ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
            ('GRID', (0, 0), (-1, -1), 1, colors.black),
            ('FONTSIZE', (0, 1), (-1, -1), 9),
            ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.lightgrey])
        ]))
        
        story.append(table)
        story.append(Spacer(1, 20))
        
        # Ranking técnico
        story.append(Paragraph("🏆 Ranking Técnico Final", self.styles['CustomSubtitle']))
        
        for i, proposal in enumerate(proposals, 1):
            ranking_text = f"{i}. <b>🏢 {proposal['empresa']}</b> - {proposal['score_tecnico']:.1f}%"
            story.append(Paragraph(ranking_text, self.styles['CustomNormal']))
        
        story.append(Spacer(1, 20))
        
        # Análise detalhada por empresa
        story.append(Paragraph("📋 Análise Detalhada por Empresa", self.styles['CustomSubtitle']))
        
        for proposal in proposals:
            self._add_company_technical_details(story, proposal)
        
        story.append(self._create_separator_line())
    
    def _add_company_technical_details(self, story: List, proposal: Dict):
        """Adiciona detalhes técnicos de uma empresa"""
        # Nome da empresa
        company_title = f"🏢 {proposal['empresa']}"
        story.append(Paragraph(company_title, self.styles['CustomSubtitle']))
        
        # Metodologia
        story.append(Paragraph("📋 Metodologia:", self.styles['CustomNormal']))
        metodologia = proposal['metodologia'] if proposal['metodologia'] else "Não informada"
        story.append(Paragraph(f"• <b>Descrição:</b> {metodologia}", self.styles['CustomNormal']))
        
        aderencia = "✓ Boa" if proposal['metodologia'] and proposal['metodologia'] != 'Metodologia não especificada' else "✗ Não informada"
        story.append(Paragraph(f"• <b>Aderência ao TR:</b> {aderencia}", self.styles['CustomNormal']))
        
        # Cronograma
        story.append(Paragraph("⏰ Cronograma:", self.styles['CustomNormal']))
        prazo = proposal['prazo_dias'] if proposal['prazo_dias'] > 0 else "Não informado"
        story.append(Paragraph(f"• <b>Prazo Total:</b> {prazo} dias", self.styles['CustomNormal']))
        
        viabilidade = "✓ Dentro do prazo" if proposal['prazo_dias'] > 0 and proposal['prazo_dias'] <= 120 else "✗ Fora do prazo ou não informado"
        story.append(Paragraph(f"• <b>Viabilidade:</b> {viabilidade}", self.styles['CustomNormal']))
        
        # Equipe técnica
        story.append(Paragraph("👥 Equipe Técnica:", self.styles['CustomNormal']))
        equipe = proposal['equipe_total'] if proposal['equipe_total'] > 0 else "Não informada"
        story.append(Paragraph(f"• <b>Total:</b> {equipe} pessoas", self.styles['CustomNormal']))
        
        status_equipe = "✓ Adequada" if proposal['equipe_total'] >= 5 else "✗ Insuficiente ou não informada"
        story.append(Paragraph(f"• <b>Status:</b> {status_equipe}", self.styles['CustomNormal']))
        
        # Recursos técnicos
        story.append(Paragraph("🔧 Recursos Técnicos:", self.styles['CustomNormal']))
        equipamentos_count = len(proposal['equipamentos'])
        materiais_count = len(proposal['materiais'])
        story.append(Paragraph(f"• <b>Equipamentos:</b> {equipamentos_count} itens listados", self.styles['CustomNormal']))
        story.append(Paragraph(f"• <b>Materiais:</b> {materiais_count} itens listados", self.styles['CustomNormal']))
        
        # Pontos fortes
        pontos_fortes = []
        if proposal['metodologia'] and proposal['metodologia'] != 'Metodologia não especificada':
            pontos_fortes.append("Metodologia bem definida")
        if proposal['prazo_dias'] > 0 and proposal['prazo_dias'] <= 90:
            pontos_fortes.append("Prazo otimizado")
        if proposal['equipe_total'] >= 8:
            pontos_fortes.append("Equipe robusta")
        if proposal['tecnologias']:
            pontos_fortes.append("Tecnologias modernas")
        
        if pontos_fortes:
            story.append(Paragraph("✅ Pontos Fortes:", self.styles['CustomNormal']))
            for ponto in pontos_fortes:
                story.append(Paragraph(f"• {ponto}", self.styles['CustomNormal']))
        
        # Gaps e riscos
        gaps = []
        if not proposal['metodologia'] or proposal['metodologia'] == 'Metodologia não especificada':
            gaps.append("Metodologia não especificada")
        if proposal['prazo_dias'] == 0:
            gaps.append("Prazo não informado")
        if proposal['equipe_total'] == 0:
            gaps.append("Equipe não detalhada")
        if not proposal['equipamentos']:
            gaps.append("Equipamentos não listados")
        
        if gaps:
            story.append(Paragraph("⚠️ Gaps e Riscos:", self.styles['CustomNormal']))
            for gap in gaps:
                story.append(Paragraph(f"• {gap}", self.styles['CustomNormal']))
        
        story.append(Spacer(1, 15))
    
    def _add_commercial_analysis(self, story: List, proposals: List[Dict]):
        """Adiciona análise comercial"""
        # Título da seção
        section_title = Paragraph("💰 BLOCO 3: EQUALIZAÇÃO DAS PROPOSTAS COMERCIAIS", self.styles['SectionHeader'])
        story.append(section_title)
        
        # Filtrar propostas com preço
        proposals_with_price = [p for p in proposals if p['preco_total'] > 0]
        proposals_with_price.sort(key=lambda x: x['preco_total'])
        
        if not proposals_with_price:
            story.append(Paragraph("⚠️ Nenhuma proposta com informações comerciais válidas foi encontrada.", self.styles['CustomNormal']))
            return
        
        # Ranking de preços
        story.append(Paragraph("💵 Ranking de Preços", self.styles['CustomSubtitle']))
        
        # Tabela de ranking
        table_data = [['Posição', 'Empresa', 'Preço Total', 'Diferença', 'Status']]
        
        base_price = proposals_with_price[0]['preco_total']
        
        for i, proposal in enumerate(proposals_with_price, 1):
            if i == 1:
                diferenca = "Base"
                status = "🏆 Melhor Preço"
            else:
                diferenca_valor = proposal['preco_total'] - base_price
                diferenca_perc = ((proposal['preco_total'] / base_price) - 1) * 100
                diferenca = f"+R$ {diferenca_valor:,.2f}"
                status = f"📈 {diferenca_perc:.0f}% mais caro"
            
            table_data.append([
                f"<b>{i}º</b>",
                proposal['empresa'],
                f"<b>R$ {proposal['preco_total']:,.2f}</b>",
                diferenca,
                status
            ])
        
        table = Table(table_data, colWidths=[2*cm, 4*cm, 3*cm, 3*cm, 3*cm])
        table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.darkgreen),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, 0), 10),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
            ('BACKGROUND', (0, 1), (-1, -1), colors.lightgreen),
            ('GRID', (0, 0), (-1, -1), 1, colors.black),
            ('FONTSIZE', (0, 1), (-1, -1), 9),
            ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.lightgrey])
        ]))
        
        story.append(table)
        story.append(Spacer(1, 20))
        
        # Análise comercial detalhada
        story.append(Paragraph("📊 Análise Comercial Detalhada", self.styles['CustomSubtitle']))
        
        for proposal in proposals_with_price:
            self._add_company_commercial_details(story, proposal)
        
        story.append(self._create_separator_line())
    
    def _add_company_commercial_details(self, story: List, proposal: Dict):
        """Adiciona detalhes comerciais de uma empresa"""
        # Nome da empresa
        company_title = f"🏢 {proposal['empresa']}"
        story.append(Paragraph(company_title, self.styles['CustomSubtitle']))
        
        # Informações comerciais
        story.append(Paragraph("💼 Informações Comerciais:", self.styles['CustomNormal']))
        
        cnpj = proposal['cnpj'] if proposal['cnpj'] else "Não informado"
        story.append(Paragraph(f"• <b>CNPJ:</b> {cnpj}", self.styles['CustomNormal']))
        story.append(Paragraph(f"• <b>Preço Total:</b> R$ {proposal['preco_total']:,.2f}", self.styles['CustomNormal']))
        
        if proposal['bdi_percentual'] > 0:
            story.append(Paragraph(f"• <b>BDI:</b> {proposal['bdi_percentual']:.2f}%", self.styles['CustomNormal']))
        
        # Composição de custos
        if any(proposal['composicao_custos'].values()):
            story.append(Paragraph("💰 Composição de Custos:", self.styles['CustomNormal']))
            story.append(Paragraph(f"• <b>Mão de Obra:</b> R$ {proposal['composicao_custos']['mao_obra']:,.2f}", self.styles['CustomNormal']))
            story.append(Paragraph(f"• <b>Materiais:</b> R$ {proposal['composicao_custos']['materiais']:,.2f}", self.styles['CustomNormal']))
            story.append(Paragraph(f"• <b>Equipamentos:</b> R$ {proposal['composicao_custos']['equipamentos']:,.2f}", self.styles['CustomNormal']))
        
        story.append(Spacer(1, 15))
    
    def _add_conclusion(self, story: List, analysis_result: Dict):
        """Adiciona conclusão e recomendações"""
        # Título da seção
        section_title = Paragraph("🎯 BLOCO 4: CONCLUSÃO E RECOMENDAÇÕES", self.styles['SectionHeader'])
        story.append(section_title)
        
        proposals = analysis_result['proposals']
        summary = analysis_result['summary']
        
        # Síntese técnica e comercial
        story.append(Paragraph("📋 Síntese Técnica e Comercial", self.styles['CustomSubtitle']))
        
        if proposals:
            best_technical = proposals[0]  # Já ordenado por score técnico
            proposals_with_price = [p for p in proposals if p['preco_total'] > 0]
            
            if proposals_with_price:
                best_price = min(proposals_with_price, key=lambda x: x['preco_total'])
                
                story.append(Paragraph(f"• <b>Melhor Proposta Técnica:</b> {best_technical['empresa']} ({best_technical['score_tecnico']:.1f}%)", self.styles['CustomNormal']))
                story.append(Paragraph(f"• <b>Melhor Proposta Comercial:</b> {best_price['empresa']} (R$ {best_price['preco_total']:,.2f})", self.styles['CustomNormal']))
                
                # Recomendação principal
                story.append(Paragraph("🏆 Recomendação Principal", self.styles['CustomSubtitle']))
                
                if best_technical['empresa'] == best_price['empresa']:
                    story.append(Paragraph(f"A empresa <b>{best_technical['empresa']}</b> apresenta a melhor proposta tanto técnica quanto comercial, sendo a recomendação unânime para contratação.", self.styles['CustomNormal']))
                else:
                    # Análise custo-benefício
                    if best_technical['preco_total'] > 0:
                        price_diff = ((best_technical['preco_total'] / best_price['preco_total']) - 1) * 100
                        score_diff = best_technical['score_tecnico'] - best_price.get('score_tecnico', 0)
                        
                        if price_diff <= 20 and score_diff >= 10:
                            story.append(Paragraph(f"Recomenda-se a contratação da <b>{best_technical['empresa']}</b>, pois oferece qualidade técnica superior ({score_diff:.1f} pontos a mais) com diferença de preço aceitável ({price_diff:.1f}%).", self.styles['CustomNormal']))
                        else:
                            story.append(Paragraph(f"Recomenda-se análise detalhada entre <b>{best_technical['empresa']}</b> (melhor técnica) e <b>{best_price['empresa']}</b> (melhor preço) considerando os critérios de avaliação 70% técnica + 30% preço.", self.styles['CustomNormal']))
                    else:
                        story.append(Paragraph(f"Recomenda-se a <b>{best_price['empresa']}</b> pela melhor proposta comercial, mas sugere-se negociação para melhorias técnicas.", self.styles['CustomNormal']))
        
        # Ações específicas
        story.append(Paragraph("📝 Ações Específicas Recomendadas", self.styles['CustomSubtitle']))
        actions = [
            "• Solicitar esclarecimentos sobre metodologia às empresas que não especificaram",
            "• Validar disponibilidade da equipe técnica proposta",
            "• Negociar prazos mais agressivos quando possível",
            "• Confirmar garantias e condições de pagamento",
            "• Realizar reunião técnica com as empresas finalistas"
        ]
        
        for action in actions:
            story.append(Paragraph(action, self.styles['CustomNormal']))
        
        # Resumo executivo final
        if summary:
            story.append(Paragraph("📊 Resumo Executivo", self.styles['CustomSubtitle']))
            
            # Tabela resumo
            table_data = [['Métrica', 'Valor']]
            table_data.append(['Total de Propostas Analisadas', str(summary.get('total_proposals', 0))])
            
            if summary.get('price_range', {}).get('min', 0) > 0:
                table_data.append(['Menor Preço', f"R$ {summary['price_range']['min']:,.2f}"])
                table_data.append(['Maior Preço', f"R$ {summary['price_range']['max']:,.2f}"])
                table_data.append(['Preço Médio', f"R$ {summary['price_range']['avg']:,.2f}"])
            
            if summary.get('deadline_range', {}).get('min', 0) > 0:
                table_data.append(['Menor Prazo', f"{summary['deadline_range']['min']} dias"])
                table_data.append(['Maior Prazo', f"{summary['deadline_range']['max']} dias"])
            
            table = Table(table_data, colWidths=[8*cm, 4*cm])
            table.setStyle(TableStyle([
                ('BACKGROUND', (0, 0), (-1, 0), colors.navy),
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
                ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                ('FONTSIZE', (0, 0), (-1, 0), 10),
                ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
                ('BACKGROUND', (0, 1), (-1, -1), colors.lightblue),
                ('GRID', (0, 0), (-1, -1), 1, colors.black),
                ('FONTSIZE', (0, 1), (-1, -1), 9)
            ]))
            
            story.append(table)
    
    def _create_separator_line(self):
        """Cria linha separadora"""
        return Table([['---']], colWidths=[15*cm], style=TableStyle([
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('FONTSIZE', (0, 0), (-1, -1), 14),
            ('TEXTCOLOR', (0, 0), (-1, -1), colors.grey)
        ]))

# Instanciar analisador e gerador
analyzer = ProposalAnalyzer()
report_generator = ReportGenerator()

def allowed_file(filename):
    """Verifica se o arquivo tem extensão permitida"""
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

@app.route('/')
def index():
    """Página principal"""
    return render_template_string(HTML_TEMPLATE)

@app.route('/upload', methods=['POST'])
def upload_files():
    """Endpoint para upload e processamento de arquivos"""
    try:
        if 'files' not in request.files:
            return jsonify({'error': 'Nenhum arquivo enviado'}), 400
        
        files = request.files.getlist('files')
        
        if not files or all(file.filename == '' for file in files):
            return jsonify({'error': 'Nenhum arquivo selecionado'}), 400
        
        uploaded_files = []
        
        for file in files:
            if file and allowed_file(file.filename):
                filename = secure_filename(file.filename)
                timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
                filename = f"{timestamp}_{filename}"
                filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
                
                file.save(filepath)
                
                file_extension = filename.rsplit('.', 1)[1].lower()
                uploaded_files.append({
                    'filename': filename,
                    'path': filepath,
                    'type': 'pdf' if file_extension == 'pdf' else file_extension
                })
        
        if not uploaded_files:
            return jsonify({'error': 'Nenhum arquivo válido foi enviado'}), 400
        
        # Processar arquivos
        logger.info(f"Processando {len(uploaded_files)} arquivos")
        analysis_result = analyzer.analyze_proposals(uploaded_files)
        
        # Gerar relatório
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        report_filename = f'analise_comparativa_{timestamp}.pdf'
        report_path = os.path.join(app.config['UPLOAD_FOLDER'], report_filename)
        
        report_generator.generate_report(analysis_result, report_path)
        
        # Limpar arquivos temporários
        for file_info in uploaded_files:
            try:
                os.remove(file_info['path'])
            except:
                pass
        
        return jsonify({
            'success': True,
            'message': 'Análise concluída com sucesso!',
            'report_url': f'/download/{report_filename}',
            'summary': analysis_result['summary']
        })
        
    except Exception as e:
        logger.error(f"Erro no processamento: {str(e)}")
        return jsonify({'error': f'Erro no processamento: {str(e)}'}), 500

@app.route('/download/<filename>')
def download_file(filename):
    """Endpoint para download do relatório"""
    try:
        file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        if os.path.exists(file_path):
            return send_file(file_path, as_attachment=True)
        else:
            return jsonify({'error': 'Arquivo não encontrado'}), 404
    except Exception as e:
        logger.error(f"Erro no download: {str(e)}")
        return jsonify({'error': 'Erro no download'}), 500

# Template HTML (mesmo da versão anterior)
HTML_TEMPLATE = '''
<!DOCTYPE html>
<html lang="pt-BR">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Analisador de Propostas - Azure AI</title>
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
            max-width: 800px;
            margin: 0 auto;
            background: white;
            border-radius: 20px;
            box-shadow: 0 20px 40px rgba(0,0,0,0.1);
            overflow: hidden;
        }
        
        .header {
            background: linear-gradient(135deg, #4facfe 0%, #00f2fe 100%);
            color: white;
            padding: 40px;
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
        
        .upload-area {
            border: 3px dashed #4facfe;
            border-radius: 15px;
            padding: 60px 20px;
            text-align: center;
            background: #f8f9ff;
            transition: all 0.3s ease;
            cursor: pointer;
        }
        
        .upload-area:hover {
            border-color: #00f2fe;
            background: #f0f8ff;
        }
        
        .upload-area.dragover {
            border-color: #00f2fe;
            background: #e6f3ff;
            transform: scale(1.02);
        }
        
        .upload-icon {
            font-size: 4em;
            color: #4facfe;
            margin-bottom: 20px;
        }
        
        .upload-text {
            font-size: 1.3em;
            color: #333;
            margin-bottom: 10px;
        }
        
        .upload-subtext {
            color: #666;
            font-size: 0.9em;
        }
        
        #fileInput {
            display: none;
        }
        
        .file-list {
            margin-top: 20px;
            padding: 20px;
            background: #f8f9fa;
            border-radius: 10px;
            display: none;
        }
        
        .file-item {
            display: flex;
            align-items: center;
            padding: 10px;
            background: white;
            margin-bottom: 10px;
            border-radius: 8px;
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
        }
        
        .file-icon {
            font-size: 1.5em;
            margin-right: 15px;
            color: #4facfe;
        }
        
        .file-info {
            flex: 1;
        }
        
        .file-name {
            font-weight: 500;
            color: #333;
        }
        
        .file-size {
            font-size: 0.8em;
            color: #666;
        }
        
        .analyze-btn {
            background: linear-gradient(135deg, #4facfe 0%, #00f2fe 100%);
            color: white;
            border: none;
            padding: 15px 40px;
            font-size: 1.1em;
            border-radius: 50px;
            cursor: pointer;
            transition: all 0.3s ease;
            margin-top: 20px;
            width: 100%;
        }
        
        .analyze-btn:hover {
            transform: translateY(-2px);
            box-shadow: 0 10px 20px rgba(79, 172, 254, 0.3);
        }
        
        .analyze-btn:disabled {
            background: #ccc;
            cursor: not-allowed;
            transform: none;
            box-shadow: none;
        }
        
        .progress {
            display: none;
            margin-top: 20px;
        }
        
        .progress-bar {
            width: 100%;
            height: 8px;
            background: #e0e0e0;
            border-radius: 4px;
            overflow: hidden;
        }
        
        .progress-fill {
            height: 100%;
            background: linear-gradient(90deg, #4facfe, #00f2fe);
            width: 0%;
            transition: width 0.3s ease;
            animation: pulse 2s infinite;
        }
        
        @keyframes pulse {
            0% { opacity: 1; }
            50% { opacity: 0.7; }
            100% { opacity: 1; }
        }
        
        .progress-text {
            text-align: center;
            margin-top: 10px;
            color: #666;
            font-size: 0.9em;
        }
        
        .result {
            display: none;
            margin-top: 30px;
            padding: 20px;
            background: #f8f9fa;
            border-radius: 10px;
        }
        
        .result.success {
            background: #d4edda;
            border: 1px solid #c3e6cb;
            color: #155724;
        }
        
        .result.error {
            background: #f8d7da;
            border: 1px solid #f5c6cb;
            color: #721c24;
        }
        
        .download-btn {
            background: #28a745;
            color: white;
            border: none;
            padding: 12px 30px;
            border-radius: 25px;
            cursor: pointer;
            text-decoration: none;
            display: inline-block;
            margin-top: 15px;
            transition: all 0.3s ease;
        }
        
        .download-btn:hover {
            background: #218838;
            transform: translateY(-2px);
        }
        
        .features {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(250px, 1fr));
            gap: 20px;
            margin-top: 30px;
        }
        
        .feature {
            text-align: center;
            padding: 20px;
            background: #f8f9ff;
            border-radius: 15px;
        }
        
        .feature-icon {
            font-size: 2.5em;
            color: #4facfe;
            margin-bottom: 15px;
        }
        
        .feature h3 {
            color: #333;
            margin-bottom: 10px;
        }
        
        .feature p {
            color: #666;
            font-size: 0.9em;
        }
        
        .azure-badge {
            background: #0078d4;
            color: white;
            padding: 5px 15px;
            border-radius: 20px;
            font-size: 0.8em;
            display: inline-block;
            margin-top: 10px;
        }
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>🚀 Analisador de Propostas</h1>
            <p>Análise inteligente com Azure Document Intelligence</p>
            <div class="azure-badge">Powered by Azure AI</div>
        </div>
        
        <div class="content">
            <div class="upload-area" onclick="document.getElementById('fileInput').click()">
                <div class="upload-icon">📁</div>
                <div class="upload-text">Clique aqui ou arraste seus arquivos</div>
                <div class="upload-subtext">Suporte para PDF, Excel (.xlsx, .xls) - Máximo 50MB</div>
            </div>
            
            <input type="file" id="fileInput" multiple accept=".pdf,.xlsx,.xls">
            
            <div class="file-list" id="fileList"></div>
            
            <button class="analyze-btn" id="analyzeBtn" onclick="analyzeFiles()" disabled>
                🔍 Analisar Propostas
            </button>
            
            <div class="progress" id="progress">
                <div class="progress-bar">
                    <div class="progress-fill" id="progressFill"></div>
                </div>
                <div class="progress-text" id="progressText">Processando arquivos...</div>
            </div>
            
            <div class="result" id="result"></div>
            
            <div class="features">
                <div class="feature">
                    <div class="feature-icon">🤖</div>
                    <h3>Azure AI</h3>
                    <p>Extração inteligente de dados usando Azure Document Intelligence</p>
                </div>
                <div class="feature">
                    <div class="feature-icon">📊</div>
                    <h3>Análise Completa</h3>
                    <p>Comparação técnica e comercial detalhada das propostas</p>
                </div>
                <div class="feature">
                    <div class="feature-icon">📋</div>
                    <h3>Relatório Profissional</h3>
                    <p>Documento PDF formatado pronto para apresentação</p>
                </div>
            </div>
        </div>
    </div>

    <script>
        let selectedFiles = [];
        
        // Configurar drag and drop
        const uploadArea = document.querySelector('.upload-area');
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
            selectedFiles = Array.from(files);
            displayFiles();
            analyzeBtn.disabled = selectedFiles.length === 0;
        }
        
        function displayFiles() {
            if (selectedFiles.length === 0) {
                fileList.style.display = 'none';
                return;
            }
            
            fileList.style.display = 'block';
            fileList.innerHTML = '<h3>📁 Arquivos Selecionados:</h3>';
            
            selectedFiles.forEach((file, index) => {
                const fileItem = document.createElement('div');
                fileItem.className = 'file-item';
                
                const icon = file.type.includes('pdf') ? '📄' : '📊';
                const size = (file.size / 1024 / 1024).toFixed(2);
                
                fileItem.innerHTML = `
                    <div class="file-icon">${icon}</div>
                    <div class="file-info">
                        <div class="file-name">${file.name}</div>
                        <div class="file-size">${size} MB</div>
                    </div>
                `;
                
                fileList.appendChild(fileItem);
            });
        }
        
        async function analyzeFiles() {
            if (selectedFiles.length === 0) return;
            
            const formData = new FormData();
            selectedFiles.forEach(file => {
                formData.append('files', file);
            });
            
            // Mostrar progresso
            document.getElementById('progress').style.display = 'block';
            document.getElementById('result').style.display = 'none';
            analyzeBtn.disabled = true;
            
            // Simular progresso
            let progress = 0;
            const progressInterval = setInterval(() => {
                progress += Math.random() * 15;
                if (progress > 90) progress = 90;
                document.getElementById('progressFill').style.width = progress + '%';
            }, 500);
            
            try {
                const response = await fetch('/upload', {
                    method: 'POST',
                    body: formData
                });
                
                const result = await response.json();
                
                clearInterval(progressInterval);
                document.getElementById('progressFill').style.width = '100%';
                
                setTimeout(() => {
                    document.getElementById('progress').style.display = 'none';
                    showResult(result, response.ok);
                    analyzeBtn.disabled = false;
                }, 1000);
                
            } catch (error) {
                clearInterval(progressInterval);
                document.getElementById('progress').style.display = 'none';
                showResult({error: 'Erro de conexão: ' + error.message}, false);
                analyzeBtn.disabled = false;
            }
        }
        
        function showResult(result, success) {
            const resultDiv = document.getElementById('result');
            resultDiv.style.display = 'block';
            resultDiv.className = 'result ' + (success ? 'success' : 'error');
            
            if (success) {
                resultDiv.innerHTML = `
                    <h3>✅ Análise Concluída!</h3>
                    <p>${result.message}</p>
                    <a href="${result.report_url}" class="download-btn">📥 Baixar Relatório PDF</a>
                `;
            } else {
                resultDiv.innerHTML = `
                    <h3>❌ Erro na Análise</h3>
                    <p>${result.error}</p>
                `;
            }
        }
    </script>
</body>
</html>
'''

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)

