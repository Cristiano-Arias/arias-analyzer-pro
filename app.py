import os
import tempfile
import gc
import re
import json
from datetime import datetime
from flask import Flask, request, jsonify, render_template_string, send_file
from flask_cors import CORS
from werkzeug.utils import secure_filename
import PyPDF2
from docx import Document
import pandas as pd
import zipfile
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.lib import colors

app = Flask(__name__)
CORS(app)

# Configurações
UPLOAD_FOLDER = 'uploads'
ALLOWED_EXTENSIONS = {'pdf', 'docx', 'doc', 'pptx', 'ppt', 'zip', 'xlsx', 'xls'}
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024  # 50MB max

# Criar pasta de uploads se não existir
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def extract_text_from_pdf(file_path):
    """Extração aprimorada de texto de PDF com processamento em chunks"""
    try:
        text = ""
        with open(file_path, 'rb') as file:
            reader = PyPDF2.PdfReader(file)
            total_pages = len(reader.pages)
            
            # Processar em chunks de 5 páginas para economizar memória
            for i in range(0, total_pages, 5):
                chunk_end = min(i + 5, total_pages)
                chunk_text = ""
                
                for page_num in range(i, chunk_end):
                    try:
                        page = reader.pages[page_num]
                        chunk_text += page.extract_text() + "\n"
                    except Exception as e:
                        print(f"Erro ao extrair página {page_num}: {e}")
                        continue
                
                text += chunk_text
                
                # Liberar memória do chunk
                del chunk_text
                gc.collect()
                
                # Limitar texto total para evitar problemas de memória
                if len(text) > 100000:  # 100k caracteres
                    text = text[:100000]
                    break
        
        return text
    except Exception as e:
        print(f"Erro ao extrair texto do PDF {file_path}: {e}")
        return ""

def extract_text_from_docx(file_path):
    """Extração de texto de arquivo DOCX"""
    try:
        doc = Document(file_path)
        text = ""
        for paragraph in doc.paragraphs:
            text += paragraph.text + "\n"
        return text
    except Exception as e:
        print(f"Erro ao extrair texto do DOCX {file_path}: {e}")
        return ""

def extract_data_from_excel(file_path):
    """Extração aprimorada de dados de arquivo Excel"""
    try:
        excel_data = {
            'precos': {},
            'bdi': {},
            'composicao_custos': {},
            'condicoes_comerciais': {},
            'tabela_servicos': []
        }
        
        # Ler todas as abas do Excel
        excel_file = pd.ExcelFile(file_path)
        
        for sheet_name in excel_file.sheet_names:
            try:
                # Ler aba com diferentes configurações para capturar dados
                df = pd.read_excel(file_path, sheet_name=sheet_name, header=None)
                
                # Converter DataFrame para texto para análise
                sheet_text = ""
                for index, row in df.iterrows():
                    row_text = " ".join([str(cell) for cell in row if pd.notna(cell)])
                    sheet_text += row_text + "\n"
                
                # Identificar tipo de aba e extrair dados específicos
                sheet_name_lower = sheet_name.lower()
                
                if 'servi' in sheet_name_lower or 'item' in sheet_name_lower:
                    # Aba de serviços - extrair preços
                    excel_data['tabela_servicos'] = extract_services_from_sheet(df, sheet_text)
                    excel_data['precos'] = extract_prices_from_sheet(df, sheet_text)
                
                elif 'bdi' in sheet_name_lower:
                    # Aba de BDI
                    excel_data['bdi'] = extract_bdi_from_sheet(df, sheet_text)
                
                elif 'custo' in sheet_name_lower or 'comp' in sheet_name_lower:
                    # Aba de composição de custos
                    excel_data['composicao_custos'] = extract_cost_composition_from_sheet(df, sheet_text)
                
                elif 'carta' in sheet_name_lower or 'comercial' in sheet_name_lower:
                    # Aba comercial
                    excel_data['condicoes_comerciais'] = extract_commercial_conditions_from_sheet(df, sheet_text)
                
            except Exception as e:
                print(f"Erro ao processar aba {sheet_name}: {e}")
                continue
        
        return excel_data
        
    except Exception as e:
        print(f"Erro ao extrair dados do Excel {file_path}: {e}")
        return {}

def extract_services_from_sheet(df, sheet_text):
    """Extrair tabela de serviços"""
    services = []
    try:
        # Procurar por padrões de serviços e preços
        lines = sheet_text.split('\n')
        for line in lines:
            # Procurar linhas que contenham descrição de serviço e valor
            if re.search(r'(R\$|RS|\d+[,\.]\d+)', line) and len(line.strip()) > 10:
                # Extrair descrição e valor
                price_match = re.search(r'(R\$\s*)?(\d{1,3}(?:[.,]\d{3})*(?:[.,]\d{2})?)', line)
                if price_match:
                    price = price_match.group(2).replace('.', '').replace(',', '.')
                    description = re.sub(r'(R\$\s*)?(\d{1,3}(?:[.,]\d{3})*(?:[.,]\d{2})?)', '', line).strip()
                    if description:
                        services.append({
                            'descricao': description,
                            'valor': float(price) if price.replace('.', '').isdigit() else 0
                        })
    except Exception as e:
        print(f"Erro ao extrair serviços: {e}")
    
    return services

def extract_prices_from_sheet(df, sheet_text):
    """Extrair preços da planilha"""
    prices = {}
    try:
        # Procurar por preço total
        total_patterns = [
            r'total\s*:?\s*(R\$\s*)?(\d{1,3}(?:[.,]\d{3})*(?:[.,]\d{2})?)',
            r'valor\s*total\s*:?\s*(R\$\s*)?(\d{1,3}(?:[.,]\d{3})*(?:[.,]\d{2})?)',
            r'pre[çc]o\s*total\s*:?\s*(R\$\s*)?(\d{1,3}(?:[.,]\d{3})*(?:[.,]\d{2})?)'
        ]
        
        for pattern in total_patterns:
            match = re.search(pattern, sheet_text, re.IGNORECASE)
            if match:
                price_str = match.group(2).replace('.', '').replace(',', '.')
                try:
                    prices['preco_total'] = float(price_str)
                    break
                except:
                    continue
        
        # Se não encontrou preço total, somar valores individuais
        if 'preco_total' not in prices:
            all_prices = re.findall(r'(R\$\s*)?(\d{1,3}(?:[.,]\d{3})*(?:[.,]\d{2})?)', sheet_text)
            total = 0
            for price_match in all_prices:
                try:
                    price_str = price_match[1].replace('.', '').replace(',', '.')
                    price_val = float(price_str)
                    if price_val > 100:  # Filtrar valores muito pequenos
                        total += price_val
                except:
                    continue
            if total > 0:
                prices['preco_total'] = total
        
    except Exception as e:
        print(f"Erro ao extrair preços: {e}")
    
    return prices

def extract_bdi_from_sheet(df, sheet_text):
    """Extrair BDI da planilha"""
    bdi_data = {}
    try:
        # Procurar por percentual de BDI
        bdi_patterns = [
            r'bdi\s*:?\s*(\d{1,2}(?:[.,]\d{1,2})?)\s*%',
            r'benefícios?\s*e?\s*despesas?\s*indiretas?\s*:?\s*(\d{1,2}(?:[.,]\d{1,2})?)\s*%'
        ]
        
        for pattern in bdi_patterns:
            match = re.search(pattern, sheet_text, re.IGNORECASE)
            if match:
                bdi_str = match.group(1).replace(',', '.')
                try:
                    bdi_data['percentual'] = float(bdi_str)
                    break
                except:
                    continue
        
        # Procurar por componentes do BDI
        components = ['administração', 'lucro', 'impostos', 'riscos']
        for component in components:
            pattern = f'{component}\\s*:?\\s*(\\d{{1,2}}(?:[.,]\\d{{1,2}})?)'
            match = re.search(pattern, sheet_text, re.IGNORECASE)
            if match:
                try:
                    bdi_data[component] = float(match.group(1).replace(',', '.'))
                except:
                    continue
        
    except Exception as e:
        print(f"Erro ao extrair BDI: {e}")
    
    return bdi_data

def extract_cost_composition_from_sheet(df, sheet_text):
    """Extrair composição de custos"""
    composition = {}
    try:
        # Procurar por percentuais de mão de obra, materiais e equipamentos
        categories = {
            'mao_de_obra': [r'mão\s*de\s*obra', r'm\.?o\.?', r'pessoal'],
            'materiais': [r'materiais?', r'insumos?'],
            'equipamentos': [r'equipamentos?', r'máquinas?']
        }
        
        for category, patterns in categories.items():
            for pattern in patterns:
                # Procurar percentual
                regex = f'{pattern}\\s*:?\\s*(\\d{{1,3}}(?:[.,]\\d{{3}})*(?:[.,]\\d{{2}})?)'
                match = re.search(regex, sheet_text, re.IGNORECASE)
                if match:
                    value_str = match.group(1).replace('.', '').replace(',', '.')
                    try:
                        composition[category] = float(value_str)
                        break
                    except:
                        continue
        
    except Exception as e:
        print(f"Erro ao extrair composição de custos: {e}")
    
    return composition

def extract_commercial_conditions_from_sheet(df, sheet_text):
    """Extrair condições comerciais"""
    conditions = {}
    try:
        # Procurar por condições de pagamento
        payment_patterns = [
            r'pagamento\s*:?\s*([^\n]+)',
            r'condições?\s*de\s*pagamento\s*:?\s*([^\n]+)'
        ]
        
        for pattern in payment_patterns:
            match = re.search(pattern, sheet_text, re.IGNORECASE)
            if match:
                conditions['condicoes_pagamento'] = match.group(1).strip()
                break
        
        # Procurar por garantia
        warranty_patterns = [
            r'garantia\s*:?\s*([^\n]+)',
            r'prazo\s*de\s*garantia\s*:?\s*([^\n]+)'
        ]
        
        for pattern in warranty_patterns:
            match = re.search(pattern, sheet_text, re.IGNORECASE)
            if match:
                conditions['garantia'] = match.group(1).strip()
                break
        
        # Procurar por treinamento
        if re.search(r'treinamento', sheet_text, re.IGNORECASE):
            conditions['treinamento'] = 'Oferecido'
        
        # Procurar por seguros
        insurance_match = re.search(r'seguro[s]?\s*:?\s*([^\n]+)', sheet_text, re.IGNORECASE)
        if insurance_match:
            conditions['seguros'] = insurance_match.group(1).strip()
        
    except Exception as e:
        print(f"Erro ao extrair condições comerciais: {e}")
    
    return conditions

def extract_from_zip(file_path):
    """Extrair arquivos de um ZIP"""
    extracted_files = []
    try:
        with zipfile.ZipFile(file_path, 'r') as zip_ref:
            temp_dir = tempfile.mkdtemp()
            zip_ref.extractall(temp_dir)
            
            for root, dirs, files in os.walk(temp_dir):
                for file in files:
                    if allowed_file(file):
                        extracted_files.append(os.path.join(root, file))
    except Exception as e:
        print(f"Erro ao extrair ZIP {file_path}: {e}")
    
    return extracted_files

def analyze_technical_proposal_advanced(text, company_name):
    """Análise técnica aprimorada com extração detalhada"""
    analysis = {
        'empresa': company_name,
        'metodologia': extract_methodology_details(text),
        'mao_de_obra': extract_workforce_details(text),
        'equipamentos': extract_equipment_details(text),
        'materiais': extract_materials_details(text),
        'obrigacoes': extract_obligations_details(text),
        'canteiro': extract_site_details(text),
        'exclusoes': extract_exclusions_details(text),
        'cronograma': extract_schedule_details(text),
        'equipe': extract_team_details(text),
        'score_tecnico': 0,
        'pontos_fortes': [],
        'pontos_fracos': []
    }
    
    # Calcular score técnico baseado na completude das informações
    analysis['score_tecnico'] = calculate_technical_score(analysis)
    
    # Identificar pontos fortes e fracos
    analysis['pontos_fortes'], analysis['pontos_fracos'] = identify_strengths_weaknesses(analysis)
    
    return analysis

def extract_methodology_details(text):
    """Extrair detalhes da metodologia"""
    methodology = {
        'descricao': '',
        'fases': [],
        'ferramentas': [],
        'abordagem': ''
    }
    
    try:
        # Procurar por seção de metodologia
        method_patterns = [
            r'metodologia\s*:?\s*([^.]{50,500})',
            r'método\s*de\s*execução\s*:?\s*([^.]{50,500})',
            r'abordagem\s*técnica\s*:?\s*([^.]{50,500})'
        ]
        
        for pattern in method_patterns:
            match = re.search(pattern, text, re.IGNORECASE | re.DOTALL)
            if match:
                methodology['descricao'] = match.group(1).strip()
                break
        
        # Procurar por fases
        phase_patterns = [
            r'fase\s*\d+\s*:?\s*([^\n]+)',
            r'etapa\s*\d+\s*:?\s*([^\n]+)',
            r'\d+[ªº]?\s*fase\s*:?\s*([^\n]+)'
        ]
        
        for pattern in phase_patterns:
            matches = re.findall(pattern, text, re.IGNORECASE)
            methodology['fases'].extend([match.strip() for match in matches])
        
        # Procurar por ferramentas
        tool_keywords = ['software', 'ferramenta', 'sistema', 'equipamento', 'tecnologia']
        for keyword in tool_keywords:
            pattern = f'{keyword}\\s*:?\\s*([^\\n]+)'
            matches = re.findall(pattern, text, re.IGNORECASE)
            methodology['ferramentas'].extend([match.strip() for match in matches])
        
    except Exception as e:
        print(f"Erro ao extrair metodologia: {e}")
    
    return methodology

def extract_workforce_details(text):
    """Extrair detalhes da mão de obra"""
    workforce = {
        'total_pessoas': 0,
        'perfis': [],
        'qualificacoes': [],
        'experiencia': ''
    }
    
    try:
        # Procurar por números de pessoas
        people_patterns = [
            r'(\d+)\s*pessoas?',
            r'(\d+)\s*profissionais?',
            r'equipe\s*de\s*(\d+)'
        ]
        
        for pattern in people_patterns:
            match = re.search(pattern, text, re.IGNORECASE)
            if match:
                workforce['total_pessoas'] = int(match.group(1))
                break
        
        # Procurar por perfis profissionais
        profile_keywords = ['engenheiro', 'técnico', 'operador', 'supervisor', 'coordenador', 'especialista']
        for keyword in profile_keywords:
            pattern = f'{keyword}\\s*([^\\n]+)'
            matches = re.findall(pattern, text, re.IGNORECASE)
            workforce['perfis'].extend([f"{keyword}: {match.strip()}" for match in matches])
        
        # Procurar por qualificações
        qual_patterns = [
            r'qualificação\s*:?\s*([^\n]+)',
            r'certificação\s*:?\s*([^\n]+)',
            r'experiência\s*:?\s*([^\n]+)'
        ]
        
        for pattern in qual_patterns:
            matches = re.findall(pattern, text, re.IGNORECASE)
            workforce['qualificacoes'].extend([match.strip() for match in matches])
        
    except Exception as e:
        print(f"Erro ao extrair mão de obra: {e}")
    
    return workforce

def extract_equipment_details(text):
    """Extrair detalhes dos equipamentos"""
    equipment = {
        'lista': [],
        'quantidade_total': 0,
        'tecnologias': [],
        'especificacoes': []
    }
    
    try:
        # Procurar por equipamentos específicos
        equipment_keywords = ['escavadeira', 'trator', 'caminhão', 'betoneira', 'guindaste', 'compressor']
        for keyword in equipment_keywords:
            pattern = f'{keyword}\\s*([^\\n]+)'
            matches = re.findall(pattern, text, re.IGNORECASE)
            equipment['lista'].extend([f"{keyword}: {match.strip()}" for match in matches])
        
        # Contar quantidade total
        equipment['quantidade_total'] = len(equipment['lista'])
        
        # Procurar por tecnologias
        tech_patterns = [
            r'tecnologia\s*:?\s*([^\n]+)',
            r'sistema\s*:?\s*([^\n]+)',
            r'automação\s*:?\s*([^\n]+)'
        ]
        
        for pattern in tech_patterns:
            matches = re.findall(pattern, text, re.IGNORECASE)
            equipment['tecnologias'].extend([match.strip() for match in matches])
        
    except Exception as e:
        print(f"Erro ao extrair equipamentos: {e}")
    
    return equipment

def extract_materials_details(text):
    """Extrair detalhes dos materiais"""
    materials = {
        'lista': [],
        'especificacoes': [],
        'quantidades': []
    }
    
    try:
        # Procurar por materiais específicos
        material_keywords = ['concreto', 'aço', 'ferro', 'cimento', 'areia', 'brita', 'madeira']
        for keyword in material_keywords:
            pattern = f'{keyword}\\s*([^\\n]+)'
            matches = re.findall(pattern, text, re.IGNORECASE)
            materials['lista'].extend([f"{keyword}: {match.strip()}" for match in matches])
        
        # Procurar por especificações
        spec_patterns = [
            r'especificação\s*:?\s*([^\n]+)',
            r'norma\s*:?\s*([^\n]+)',
            r'qualidade\s*:?\s*([^\n]+)'
        ]
        
        for pattern in spec_patterns:
            matches = re.findall(pattern, text, re.IGNORECASE)
            materials['especificacoes'].extend([match.strip() for match in matches])
        
        # Procurar por quantidades
        qty_patterns = [
            r'(\d+(?:[.,]\d+)?)\s*(m³|m²|kg|ton|unid)',
            r'quantidade\s*:?\s*([^\n]+)'
        ]
        
        for pattern in qty_patterns:
            matches = re.findall(pattern, text, re.IGNORECASE)
            materials['quantidades'].extend([match for match in matches])
        
    except Exception as e:
        print(f"Erro ao extrair materiais: {e}")
    
    return materials

def extract_obligations_details(text):
    """Extrair detalhes das obrigações"""
    obligations = {
        'lista': [],
        'responsabilidades': [],
        'compromissos': []
    }
    
    try:
        # Procurar por obrigações
        obligation_patterns = [
            r'obrigação\s*:?\s*([^\n]+)',
            r'responsabilidade\s*:?\s*([^\n]+)',
            r'compromisso\s*:?\s*([^\n]+)',
            r'dever\s*:?\s*([^\n]+)'
        ]
        
        for pattern in obligation_patterns:
            matches = re.findall(pattern, text, re.IGNORECASE)
            obligations['lista'].extend([match.strip() for match in matches])
        
        # Separar por tipo
        for item in obligations['lista']:
            if 'responsabilidade' in item.lower():
                obligations['responsabilidades'].append(item)
            elif 'compromisso' in item.lower():
                obligations['compromissos'].append(item)
        
    except Exception as e:
        print(f"Erro ao extrair obrigações: {e}")
    
    return obligations

def extract_site_details(text):
    """Extrair detalhes do canteiro"""
    site = {
        'organizacao': [],
        'seguranca': [],
        'logistica': [],
        'instalacoes': []
    }
    
    try:
        # Procurar por informações de canteiro
        site_patterns = [
            r'canteiro\s*:?\s*([^\n]+)',
            r'organização\s*:?\s*([^\n]+)',
            r'segurança\s*:?\s*([^\n]+)',
            r'logística\s*:?\s*([^\n]+)'
        ]
        
        for pattern in site_patterns:
            matches = re.findall(pattern, text, re.IGNORECASE)
            if 'segurança' in pattern:
                site['seguranca'].extend([match.strip() for match in matches])
            elif 'logística' in pattern:
                site['logistica'].extend([match.strip() for match in matches])
            elif 'organização' in pattern:
                site['organizacao'].extend([match.strip() for match in matches])
            else:
                site['instalacoes'].extend([match.strip() for match in matches])
        
    except Exception as e:
        print(f"Erro ao extrair canteiro: {e}")
    
    return site

def extract_exclusions_details(text):
    """Extrair detalhes das exclusões"""
    exclusions = {
        'lista': [],
        'observacoes': []
    }
    
    try:
        # Procurar por exclusões
        exclusion_patterns = [
            r'exclusão\s*:?\s*([^\n]+)',
            r'não\s*inclui\s*:?\s*([^\n]+)',
            r'excluído\s*:?\s*([^\n]+)',
            r'fora\s*do\s*escopo\s*:?\s*([^\n]+)'
        ]
        
        for pattern in exclusion_patterns:
            matches = re.findall(pattern, text, re.IGNORECASE)
            exclusions['lista'].extend([match.strip() for match in matches])
        
    except Exception as e:
        print(f"Erro ao extrair exclusões: {e}")
    
    return exclusions

def extract_schedule_details(text):
    """Extrair detalhes do cronograma"""
    schedule = {
        'prazo_total': '',
        'marcos': [],
        'fases_cronograma': [],
        'viabilidade': ''
    }
    
    try:
        # Procurar por prazo total
        deadline_patterns = [
            r'prazo\s*:?\s*(\d+)\s*(dias?|meses?|semanas?)',
            r'duração\s*:?\s*(\d+)\s*(dias?|meses?|semanas?)',
            r'cronograma\s*:?\s*(\d+)\s*(dias?|meses?|semanas?)'
        ]
        
        for pattern in deadline_patterns:
            match = re.search(pattern, text, re.IGNORECASE)
            if match:
                schedule['prazo_total'] = f"{match.group(1)} {match.group(2)}"
                break
        
        # Procurar por marcos
        milestone_patterns = [
            r'marco\s*\d+\s*:?\s*([^\n]+)',
            r'entrega\s*\d+\s*:?\s*([^\n]+)',
            r'milestone\s*\d+\s*:?\s*([^\n]+)'
        ]
        
        for pattern in milestone_patterns:
            matches = re.findall(pattern, text, re.IGNORECASE)
            schedule['marcos'].extend([match.strip() for match in matches])
        
    except Exception as e:
        print(f"Erro ao extrair cronograma: {e}")
    
    return schedule

def extract_team_details(text):
    """Extrair detalhes da equipe"""
    team = {
        'coordenador': '',
        'especialistas': [],
        'estrutura': [],
        'experiencia_equipe': ''
    }
    
    try:
        # Procurar por coordenador
        coord_patterns = [
            r'coordenador\s*:?\s*([^\n]+)',
            r'gerente\s*:?\s*([^\n]+)',
            r'responsável\s*técnico\s*:?\s*([^\n]+)'
        ]
        
        for pattern in coord_patterns:
            match = re.search(pattern, text, re.IGNORECASE)
            if match:
                team['coordenador'] = match.group(1).strip()
                break
        
        # Procurar por especialistas
        specialist_patterns = [
            r'especialista\s*em\s*([^\n]+)',
            r'expert\s*em\s*([^\n]+)',
            r'consultor\s*([^\n]+)'
        ]
        
        for pattern in specialist_patterns:
            matches = re.findall(pattern, text, re.IGNORECASE)
            team['especialistas'].extend([match.strip() for match in matches])
        
    except Exception as e:
        print(f"Erro ao extrair equipe: {e}")
    
    return team

def calculate_technical_score(analysis):
    """Calcular score técnico baseado na completude"""
    score = 0
    max_score = 100
    
    # Pontuação por seção
    sections = {
        'metodologia': 20,
        'mao_de_obra': 15,
        'equipamentos': 15,
        'materiais': 15,
        'obrigacoes': 10,
        'canteiro': 10,
        'cronograma': 10,
        'equipe': 5
    }
    
    for section, max_points in sections.items():
        if section in analysis and analysis[section]:
            # Verificar se a seção tem conteúdo
            section_data = analysis[section]
            if isinstance(section_data, dict):
                # Contar campos preenchidos
                filled_fields = sum(1 for value in section_data.values() if value)
                total_fields = len(section_data)
                if total_fields > 0:
                    section_score = (filled_fields / total_fields) * max_points
                    score += section_score
            elif section_data:
                score += max_points
    
    return round(score, 1)

def identify_strengths_weaknesses(analysis):
    """Identificar pontos fortes e fracos"""
    strengths = []
    weaknesses = []
    
    # Verificar cada seção
    if analysis['metodologia']['descricao']:
        strengths.append("Metodologia bem definida")
    else:
        weaknesses.append("Metodologia não detalhada")
    
    if analysis['mao_de_obra']['total_pessoas'] > 0:
        strengths.append(f"Equipe dimensionada ({analysis['mao_de_obra']['total_pessoas']} pessoas)")
    else:
        weaknesses.append("Dimensionamento de equipe não especificado")
    
    if analysis['equipamentos']['quantidade_total'] > 0:
        strengths.append(f"Equipamentos especificados ({analysis['equipamentos']['quantidade_total']} itens)")
    else:
        weaknesses.append("Lista de equipamentos incompleta")
    
    if analysis['cronograma']['prazo_total']:
        strengths.append("Prazo definido")
    else:
        weaknesses.append("Cronograma não especificado")
    
    if not analysis['exclusoes']['lista']:
        weaknesses.append("Exclusões não especificadas (risco de custos adicionais)")
    
    return strengths, weaknesses

def analyze_commercial_proposal_advanced(excel_data, company_name):
    """Análise comercial aprimorada"""
    analysis = {
        'empresa': company_name,
        'precos': excel_data.get('precos', {}),
        'bdi': excel_data.get('bdi', {}),
        'composicao_custos': excel_data.get('composicao_custos', {}),
        'condicoes_comerciais': excel_data.get('condicoes_comerciais', {}),
        'tabela_servicos': excel_data.get('tabela_servicos', []),
        'score_comercial': 0,
        'vantagens': [],
        'desvantagens': []
    }
    
    # Calcular score comercial
    analysis['score_comercial'] = calculate_commercial_score(analysis)
    
    # Identificar vantagens e desvantagens
    analysis['vantagens'], analysis['desvantagens'] = identify_commercial_advantages(analysis)
    
    return analysis

def calculate_commercial_score(analysis):
    """Calcular score comercial"""
    score = 0
    
    # Pontuação por completude de informações
    if analysis['precos'].get('preco_total', 0) > 0:
        score += 30
    
    if analysis['bdi'].get('percentual', 0) > 0:
        score += 20
    
    if analysis['composicao_custos']:
        score += 20
    
    if analysis['condicoes_comerciais'].get('condicoes_pagamento'):
        score += 15
    
    if analysis['condicoes_comerciais'].get('garantia'):
        score += 10
    
    if analysis['tabela_servicos']:
        score += 5
    
    return round(score, 1)

def identify_commercial_advantages(analysis):
    """Identificar vantagens e desvantagens comerciais"""
    advantages = []
    disadvantages = []
    
    # Verificar preço
    if analysis['precos'].get('preco_total', 0) > 0:
        advantages.append(f"Preço total definido: R$ {analysis['precos']['preco_total']:,.2f}")
    else:
        disadvantages.append("Preço total não especificado")
    
    # Verificar BDI
    if analysis['bdi'].get('percentual', 0) > 0:
        advantages.append(f"BDI especificado: {analysis['bdi']['percentual']}%")
    else:
        disadvantages.append("BDI não especificado")
    
    # Verificar condições
    if analysis['condicoes_comerciais'].get('condicoes_pagamento'):
        advantages.append("Condições de pagamento definidas")
    else:
        disadvantages.append("Condições de pagamento não especificadas")
    
    if analysis['condicoes_comerciais'].get('garantia'):
        advantages.append("Garantia oferecida")
    else:
        disadvantages.append("Garantia não especificada")
    
    return advantages, disadvantages

def generate_comparative_analysis_advanced(technical_analyses, commercial_analyses):
    """Gerar análise comparativa avançada"""
    
    # Ordenar por score técnico
    technical_ranking = sorted(technical_analyses, key=lambda x: x['score_tecnico'], reverse=True)
    
    # Ordenar por score comercial (e preço se disponível)
    commercial_ranking = sorted(commercial_analyses, key=lambda x: x['score_comercial'], reverse=True)
    
    # Criar ranking de preços
    price_ranking = []
    for analysis in commercial_analyses:
        price = analysis['precos'].get('preco_total', 0)
        if price > 0:
            price_ranking.append({
                'empresa': analysis['empresa'],
                'preco': price
            })
    price_ranking.sort(key=lambda x: x['preco'])
    
    # Análise de custo-benefício
    cost_benefit = []
    for tech in technical_analyses:
        for comm in commercial_analyses:
            if tech['empresa'] == comm['empresa']:
                price = comm['precos'].get('preco_total', 0)
                if price > 0:
                    # Calcular índice custo-benefício (score técnico / preço normalizado)
                    normalized_price = price / max([c['precos'].get('preco_total', 1) for c in commercial_analyses])
                    cb_index = tech['score_tecnico'] / normalized_price if normalized_price > 0 else 0
                    cost_benefit.append({
                        'empresa': tech['empresa'],
                        'score_tecnico': tech['score_tecnico'],
                        'preco': price,
                        'indice_custo_beneficio': cb_index
                    })
    
    cost_benefit.sort(key=lambda x: x['indice_custo_beneficio'], reverse=True)
    
    return {
        'ranking_tecnico': technical_ranking,
        'ranking_comercial': commercial_ranking,
        'ranking_precos': price_ranking,
        'custo_beneficio': cost_benefit,
        'melhor_tecnica': technical_ranking[0]['empresa'] if technical_ranking else 'N/A',
        'melhor_comercial': price_ranking[0]['empresa'] if price_ranking else 'N/A',
        'melhor_custo_beneficio': cost_benefit[0]['empresa'] if cost_benefit else 'N/A'
    }

def generate_detailed_report(technical_analyses, commercial_analyses, comparative_analysis):
    """Gerar relatório detalhado"""
    
    report = f"""# ANÁLISE COMPARATIVA DE PROPOSTAS
**Projeto:** Cercamento Perimetral
**Data:** {datetime.now().strftime('%d/%m/%Y')}

---

## RESUMO EXECUTIVO

### Rankings Gerais
"""
    
    # Rankings
    if comparative_analysis['ranking_tecnico']:
        report += "**Ranking Técnico:**\n"
        for i, analysis in enumerate(comparative_analysis['ranking_tecnico'], 1):
            report += f"**{i}º:** {analysis['empresa']} - {analysis['score_tecnico']}%\n"
        report += "\n"
    
    if comparative_analysis['ranking_precos']:
        report += "**Ranking de Preços:**\n"
        for i, item in enumerate(comparative_analysis['ranking_precos'], 1):
            report += f"**{i}º:** {item['empresa']} - R$ {item['preco']:,.2f}\n"
        report += "\n"
    
    if comparative_analysis['custo_beneficio']:
        report += "**Ranking Custo-Benefício:**\n"
        for i, item in enumerate(comparative_analysis['custo_beneficio'], 1):
            report += f"**{i}º:** {item['empresa']} - Índice: {item['indice_custo_beneficio']:.2f}\n"
        report += "\n"
    
    # Análise Técnica Comparativa
    report += """---

## ANÁLISE TÉCNICA COMPARATIVA

### Matriz de Comparação Técnica

| Empresa | Metodologia | Mão de Obra | Equipamentos | Materiais | Cronograma | Score Total |
|---------|-------------|-------------|--------------|-----------|------------|-------------|
"""
    
    for analysis in technical_analyses:
        method_score = "✅" if analysis['metodologia']['descricao'] else "❌"
        workforce_score = "✅" if analysis['mao_de_obra']['total_pessoas'] > 0 else "❌"
        equipment_score = "✅" if analysis['equipamentos']['quantidade_total'] > 0 else "❌"
        materials_score = "✅" if analysis['materiais']['lista'] else "❌"
        schedule_score = "✅" if analysis['cronograma']['prazo_total'] else "❌"
        
        report += f"| {analysis['empresa']} | {method_score} | {workforce_score} | {equipment_score} | {materials_score} | {schedule_score} | {analysis['score_tecnico']}% |\n"
    
    # Análise detalhada por empresa
    for analysis in technical_analyses:
        report += f"""
### {analysis['empresa']} - Análise Técnica Detalhada

**Metodologia de Execução:**
{analysis['metodologia']['descricao'] if analysis['metodologia']['descricao'] else 'Não especificada'}

**Mão de Obra:**
- Total de pessoas: {analysis['mao_de_obra']['total_pessoas']}
- Perfis identificados: {len(analysis['mao_de_obra']['perfis'])}

**Equipamentos:**
- Quantidade total: {analysis['equipamentos']['quantidade_total']}
- Lista: {', '.join(analysis['equipamentos']['lista'][:3]) if analysis['equipamentos']['lista'] else 'Não especificada'}

**Materiais:**
- Itens identificados: {len(analysis['materiais']['lista'])}
- Especificações: {len(analysis['materiais']['especificacoes'])} itens

**Cronograma:**
- Prazo: {analysis['cronograma']['prazo_total'] if analysis['cronograma']['prazo_total'] else 'Não especificado'}
- Marcos: {len(analysis['cronograma']['marcos'])} identificados

**Pontos Fortes:**
{chr(10).join([f"✅ {ponto}" for ponto in analysis['pontos_fortes']])}

**Pontos de Atenção:**
{chr(10).join([f"⚠️ {ponto}" for ponto in analysis['pontos_fracos']])}
"""
    
    # Análise Comercial Comparativa
    report += """
---

## ANÁLISE COMERCIAL COMPARATIVA

### Resumo de Preços e Condições

| Empresa | Preço Total | BDI | Condições Pagamento | Garantia | Score Comercial |
|---------|-------------|-----|-------------------|----------|-----------------|
"""
    
    for analysis in commercial_analyses:
        price = analysis['precos'].get('preco_total', 0)
        price_str = f"R$ {price:,.2f}" if price > 0 else "Não informado"
        bdi = analysis['bdi'].get('percentual', 0)
        bdi_str = f"{bdi}%" if bdi > 0 else "Não informado"
        payment = analysis['condicoes_comerciais'].get('condicoes_pagamento', 'Não informado')
        warranty = analysis['condicoes_comerciais'].get('garantia', 'Não informado')
        
        report += f"| {analysis['empresa']} | {price_str} | {bdi_str} | {payment[:20]}... | {warranty[:20]}... | {analysis['score_comercial']}% |\n"
    
    # Análise detalhada comercial por empresa
    for analysis in commercial_analyses:
        report += f"""
### {analysis['empresa']} - Análise Comercial Detalhada

**Preços e Composição:**
- Preço Total: {f"R$ {analysis['precos']['preco_total']:,.2f}" if analysis['precos'].get('preco_total', 0) > 0 else 'Não informado'}
- BDI: {f"{analysis['bdi']['percentual']}%" if analysis['bdi'].get('percentual', 0) > 0 else 'Não informado'}

**Composição de Custos:**
- Mão de Obra: {f"R$ {analysis['composicao_custos']['mao_de_obra']:,.2f}" if analysis['composicao_custos'].get('mao_de_obra', 0) > 0 else 'Não informado'}
- Materiais: {f"R$ {analysis['composicao_custos']['materiais']:,.2f}" if analysis['composicao_custos'].get('materiais', 0) > 0 else 'Não informado'}
- Equipamentos: {f"R$ {analysis['composicao_custos']['equipamentos']:,.2f}" if analysis['composicao_custos'].get('equipamentos', 0) > 0 else 'Não informado'}

**Condições Comerciais:**
- Pagamento: {analysis['condicoes_comerciais'].get('condicoes_pagamento', 'Não especificado')}
- Garantia: {analysis['condicoes_comerciais'].get('garantia', 'Não especificada')}
- Treinamento: {analysis['condicoes_comerciais'].get('treinamento', 'Não especificado')}
- Seguros: {analysis['condicoes_comerciais'].get('seguros', 'Não especificados')}

**Vantagens Comerciais:**
{chr(10).join([f"✅ {vantagem}" for vantagem in analysis['vantagens']])}

**Pontos de Atenção:**
{chr(10).join([f"⚠️ {desvantagem}" for desvantagem in analysis['desvantagens']])}
"""
    
    # Conclusões e Recomendações
    report += f"""
---

## CONCLUSÕES E RECOMENDAÇÕES

### Análise Comparativa Final

**Melhor Proposta Técnica:** {comparative_analysis['melhor_tecnica']}
**Melhor Proposta Comercial:** {comparative_analysis['melhor_comercial']}
**Melhor Custo-Benefício:** {comparative_analysis['melhor_custo_beneficio']}

### Recomendações Finais

**Para Tomada de Decisão:**
1. **Análise Técnica:** Considere a proposta com maior score técnico para garantir qualidade de execução.
2. **Análise Comercial:** Avalie não apenas o menor preço, mas também as condições de pagamento e garantias oferecidas.
3. **Custo-Benefício:** Busque o equilíbrio entre qualidade técnica e vantagem comercial.

**Próximos Passos Sugeridos:**
1. **Esclarecimentos:** Solicite esclarecimentos para propostas com informações incompletas.
2. **Negociação:** Considere negociar condições com as propostas melhor classificadas.
3. **Verificação:** Confirme referências e capacidade técnica das empresas.

**Pontos de Atenção:**
⚠️ Propostas com exclusões não especificadas podem gerar custos adicionais.
⚠️ Cronogramas muito agressivos podem comprometer a qualidade.
⚠️ Preços muito baixos podem indicar subdimensionamento ou qualidade inferior.

---
"""
    
    return report

def create_pdf_report(markdown_content, output_path):
    """Criar relatório em PDF usando ReportLab"""
    try:
        doc = SimpleDocTemplate(output_path, pagesize=letter)
        styles = getSampleStyleSheet()
        story = []
        
        # Estilo personalizado para títulos
        title_style = ParagraphStyle(
            'CustomTitle',
            parent=styles['Heading1'],
            fontSize=16,
            spaceAfter=12,
            textColor=colors.darkblue
        )
        
        subtitle_style = ParagraphStyle(
            'CustomSubtitle',
            parent=styles['Heading2'],
            fontSize=14,
            spaceAfter=10,
            textColor=colors.darkgreen
        )
        
        # Processar markdown para PDF
        lines = markdown_content.split('\n')
        current_table_data = []
        in_table = False
        
        for line in lines:
            line = line.strip()
            
            if not line:
                if current_table_data and in_table:
                    # Criar tabela
                    table = Table(current_table_data)
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
                    story.append(Spacer(1, 12))
                    current_table_data = []
                    in_table = False
                continue
            
            if line.startswith('# '):
                # Título principal
                title_text = line[2:].strip()
                story.append(Paragraph(title_text, title_style))
                story.append(Spacer(1, 12))
            elif line.startswith('## '):
                # Subtítulo
                subtitle_text = line[3:].strip()
                story.append(Paragraph(subtitle_text, subtitle_style))
                story.append(Spacer(1, 8))
            elif line.startswith('### '):
                # Subtítulo menor
                subsubtitle_text = line[4:].strip()
                story.append(Paragraph(subsubtitle_text, styles['Heading3']))
                story.append(Spacer(1, 6))
            elif line.startswith('|') and '|' in line[1:]:
                # Linha de tabela
                cells = [cell.strip() for cell in line.split('|')[1:-1]]
                current_table_data.append(cells)
                in_table = True
            elif line.startswith('**') and line.endswith('**'):
                # Texto em negrito
                bold_text = line[2:-2]
                story.append(Paragraph(f"<b>{bold_text}</b>", styles['Normal']))
                story.append(Spacer(1, 6))
            else:
                # Texto normal
                if line.startswith('- ') or line.startswith('✅ ') or line.startswith('⚠️ ') or line.startswith('❌ '):
                    # Lista
                    story.append(Paragraph(line, styles['Normal']))
                else:
                    # Parágrafo normal
                    story.append(Paragraph(line, styles['Normal']))
                story.append(Spacer(1, 4))
        
        # Adicionar tabela final se existir
        if current_table_data:
            table = Table(current_table_data)
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
        
        doc.build(story)
        return True
        
    except Exception as e:
        print(f"Erro ao criar PDF: {e}")
        return False

# Template HTML simplificado
HTML_TEMPLATE = """
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
            border-radius: 15px;
            box-shadow: 0 20px 40px rgba(0,0,0,0.1);
            overflow: hidden;
        }
        
        .header {
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
            padding: 30px;
            text-align: center;
        }
        
        .header h1 {
            font-size: 2.5em;
            margin-bottom: 10px;
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
            border: 2px solid #f0f0f0;
            border-radius: 10px;
            background: #fafafa;
        }
        
        .section h2 {
            color: #333;
            margin-bottom: 20px;
            font-size: 1.8em;
            border-bottom: 3px solid #667eea;
            padding-bottom: 10px;
        }
        
        .upload-area {
            border: 3px dashed #667eea;
            border-radius: 10px;
            padding: 30px;
            text-align: center;
            background: white;
            margin-bottom: 20px;
            transition: all 0.3s ease;
        }
        
        .upload-area:hover {
            border-color: #764ba2;
            background: #f8f9ff;
        }
        
        .upload-area.dragover {
            border-color: #764ba2;
            background: #f0f4ff;
        }
        
        .file-input {
            display: none;
        }
        
        .upload-btn {
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
            padding: 12px 30px;
            border: none;
            border-radius: 25px;
            cursor: pointer;
            font-size: 16px;
            transition: all 0.3s ease;
        }
        
        .upload-btn:hover {
            transform: translateY(-2px);
            box-shadow: 0 5px 15px rgba(102, 126, 234, 0.4);
        }
        
        .file-list {
            margin-top: 15px;
        }
        
        .file-item {
            background: white;
            padding: 15px;
            margin: 10px 0;
            border-radius: 8px;
            border-left: 4px solid #667eea;
            display: flex;
            justify-content: space-between;
            align-items: center;
            box-shadow: 0 2px 5px rgba(0,0,0,0.1);
        }
        
        .file-item.success {
            border-left-color: #28a745;
            background: #f8fff9;
        }
        
        .company-input {
            width: 100%;
            padding: 12px;
            border: 2px solid #ddd;
            border-radius: 8px;
            font-size: 16px;
            margin-bottom: 10px;
            transition: border-color 0.3s ease;
        }
        
        .company-input:focus {
            outline: none;
            border-color: #667eea;
        }
        
        .cnpj-input {
            width: 100%;
            padding: 12px;
            border: 2px solid #ddd;
            border-radius: 8px;
            font-size: 16px;
            margin-bottom: 20px;
            transition: border-color 0.3s ease;
        }
        
        .cnpj-input:focus {
            outline: none;
            border-color: #667eea;
        }
        
        .add-proposal-btn {
            background: #28a745;
            color: white;
            padding: 10px 20px;
            border: none;
            border-radius: 20px;
            cursor: pointer;
            font-size: 14px;
            margin-top: 10px;
        }
        
        .add-proposal-btn:hover {
            background: #218838;
        }
        
        .generate-btn {
            background: linear-gradient(135deg, #dc3545 0%, #c82333 100%);
            color: white;
            padding: 15px 40px;
            border: none;
            border-radius: 30px;
            cursor: pointer;
            font-size: 18px;
            font-weight: bold;
            display: block;
            margin: 30px auto;
            transition: all 0.3s ease;
        }
        
        .generate-btn:hover {
            transform: translateY(-2px);
            box-shadow: 0 8px 20px rgba(220, 53, 69, 0.4);
        }
        
        .generate-btn:disabled {
            background: #ccc;
            cursor: not-allowed;
            transform: none;
            box-shadow: none;
        }
        
        .loading {
            display: none;
            text-align: center;
            padding: 40px;
        }
        
        .spinner {
            border: 4px solid #f3f3f3;
            border-top: 4px solid #667eea;
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
            background: #f8f9fa;
            border-radius: 10px;
            border: 2px solid #28a745;
        }
        
        .download-btn {
            background: #28a745;
            color: white;
            padding: 12px 25px;
            border: none;
            border-radius: 25px;
            cursor: pointer;
            font-size: 16px;
            margin: 10px;
            text-decoration: none;
            display: inline-block;
            transition: all 0.3s ease;
        }
        
        .download-btn:hover {
            background: #218838;
            transform: translateY(-2px);
        }
        
        .error {
            background: #f8d7da;
            color: #721c24;
            padding: 15px;
            border-radius: 8px;
            margin: 15px 0;
            border: 1px solid #f5c6cb;
        }
        
        .success {
            background: #d4edda;
            color: #155724;
            padding: 15px;
            border-radius: 8px;
            margin: 15px 0;
            border: 1px solid #c3e6cb;
        }
        
        @media (max-width: 768px) {
            .container {
                margin: 10px;
                border-radius: 10px;
            }
            
            .content {
                padding: 20px;
            }
            
            .header h1 {
                font-size: 2em;
            }
            
            .section {
                padding: 20px;
            }
        }
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>🚀 Proposal Analyzer Pro</h1>
            <p>Análise Inteligente e Comparação de Propostas Técnicas e Comerciais</p>
        </div>
        
        <div class="content">
            <div class="section">
                <h2>📋 Propostas Técnicas</h2>
                <p>Faça upload das propostas técnicas (PDF, DOCX, PPT) para análise comparativa detalhada.</p>
                
                <div id="technical-proposals">
                    <div class="proposal-group">
                        <h3>Proposta Técnica 1</h3>
                        <input type="text" class="company-input" placeholder="Nome da Empresa" required>
                        <div class="upload-area" onclick="document.getElementById('tech-file-1').click()">
                            <p>📁 Clique aqui ou arraste o arquivo da proposta técnica</p>
                            <button type="button" class="upload-btn">Selecionar Arquivo</button>
                            <input type="file" id="tech-file-1" class="file-input" accept=".pdf,.docx,.doc,.pptx,.ppt">
                        </div>
                        <div class="file-list" id="tech-files-1"></div>
                    </div>
                </div>
                
                <button type="button" class="add-proposal-btn" onclick="addTechnicalProposal()">+ Adicionar Proposta Técnica</button>
            </div>
            
            <div class="section">
                <h2>💰 Propostas Comerciais</h2>
                <p>Faça upload das propostas comerciais (Excel, PDF) para análise de preços e condições.</p>
                
                <div id="commercial-proposals">
                    <div class="proposal-group">
                        <h3>Proposta Comercial 1</h3>
                        <input type="text" class="company-input" placeholder="Nome da Empresa" required>
                        <input type="text" class="cnpj-input" placeholder="CNPJ (Opcional)" pattern="[0-9]{2}[.][0-9]{3}[.][0-9]{3}[/][0-9]{4}[-][0-9]{2}">
                        <div class="upload-area" onclick="document.getElementById('comm-file-1').click()">
                            <p>📊 Clique aqui ou arraste o arquivo da proposta comercial</p>
                            <button type="button" class="upload-btn">Selecionar Arquivo</button>
                            <input type="file" id="comm-file-1" class="file-input" accept=".xlsx,.xls,.pdf,.docx,.doc">
                        </div>
                        <div class="file-list" id="comm-files-1"></div>
                    </div>
                </div>
                
                <button type="button" class="add-proposal-btn" onclick="addCommercialProposal()">+ Adicionar Proposta Comercial</button>
            </div>
            
            <button type="button" class="generate-btn" onclick="generateReport()" id="generate-btn">
                🧠 Gerar Relatório com Análise IA
            </button>
            
            <div class="loading" id="loading">
                <div class="spinner"></div>
                <h3>Processando documentos e gerando análise...</h3>
                <p>Isso pode levar alguns minutos. Por favor, aguarde.</p>
            </div>
            
            <div class="result" id="result">
                <h3>✅ Relatório Gerado com Sucesso!</h3>
                <p>Seu relatório de análise comparativa está pronto para download.</p>
                <div>
                    <a href="#" class="download-btn" id="download-pdf">📄 Download PDF</a>
                    <a href="#" class="download-btn" id="download-md">📝 Download Markdown</a>
                </div>
            </div>
        </div>
    </div>

    <script>
        let technicalProposalCount = 1;
        let commercialProposalCount = 1;

        function addTechnicalProposal() {
            if (technicalProposalCount >= 4) {
                alert('Máximo de 4 propostas técnicas permitidas.');
                return;
            }
            
            technicalProposalCount++;
            const container = document.getElementById('technical-proposals');
            const newProposal = document.createElement('div');
            newProposal.className = 'proposal-group';
            newProposal.innerHTML = `
                <h3>Proposta Técnica ${technicalProposalCount}</h3>
                <input type="text" class="company-input" placeholder="Nome da Empresa" required>
                <div class="upload-area" onclick="document.getElementById('tech-file-${technicalProposalCount}').click()">
                    <p>📁 Clique aqui ou arraste o arquivo da proposta técnica</p>
                    <button type="button" class="upload-btn">Selecionar Arquivo</button>
                    <input type="file" id="tech-file-${technicalProposalCount}" class="file-input" accept=".pdf,.docx,.doc,.pptx,.ppt">
                </div>
                <div class="file-list" id="tech-files-${technicalProposalCount}"></div>
            `;
            container.appendChild(newProposal);
            setupFileHandlers(`tech-file-${technicalProposalCount}`, `tech-files-${technicalProposalCount}`);
        }

        function addCommercialProposal() {
            if (commercialProposalCount >= 4) {
                alert('Máximo de 4 propostas comerciais permitidas.');
                return;
            }
            
            commercialProposalCount++;
            const container = document.getElementById('commercial-proposals');
            const newProposal = document.createElement('div');
            newProposal.className = 'proposal-group';
            newProposal.innerHTML = `
                <h3>Proposta Comercial ${commercialProposalCount}</h3>
                <input type="text" class="company-input" placeholder="Nome da Empresa" required>
                <input type="text" class="cnpj-input" placeholder="CNPJ (Opcional)" pattern="[0-9]{2}[.][0-9]{3}[.][0-9]{3}[/][0-9]{4}[-][0-9]{2}">
                <div class="upload-area" onclick="document.getElementById('comm-file-${commercialProposalCount}').click()">
                    <p>📊 Clique aqui ou arraste o arquivo da proposta comercial</p>
                    <button type="button" class="upload-btn">Selecionar Arquivo</button>
                    <input type="file" id="comm-file-${commercialProposalCount}" class="file-input" accept=".xlsx,.xls,.pdf,.docx,.doc">
                </div>
                <div class="file-list" id="comm-files-${commercialProposalCount}"></div>
            `;
            container.appendChild(newProposal);
            setupFileHandlers(`comm-file-${commercialProposalCount}`, `comm-files-${commercialProposalCount}`);
        }

        function setupFileHandlers(inputId, listId) {
            const input = document.getElementById(inputId);
            const fileList = document.getElementById(listId);
            
            input.addEventListener('change', function(e) {
                displayFiles(e.target.files, fileList);
            });
        }

        function displayFiles(files, container) {
            container.innerHTML = '';
            for (let file of files) {
                const fileItem = document.createElement('div');
                fileItem.className = 'file-item success';
                fileItem.innerHTML = `
                    <span>✅ ${file.name}</span>
                    <span>${(file.size / 1024 / 1024).toFixed(2)} MB</span>
                `;
                container.appendChild(fileItem);
            }
        }

        // Setup inicial para os primeiros campos
        setupFileHandlers('tech-file-1', 'tech-files-1');
        setupFileHandlers('comm-file-1', 'comm-files-1');

        // Drag and drop
        document.addEventListener('DOMContentLoaded', function() {
            const uploadAreas = document.querySelectorAll('.upload-area');
            
            uploadAreas.forEach(area => {
                area.addEventListener('dragover', function(e) {
                    e.preventDefault();
                    this.classList.add('dragover');
                });
                
                area.addEventListener('dragleave', function(e) {
                    e.preventDefault();
                    this.classList.remove('dragover');
                });
                
                area.addEventListener('drop', function(e) {
                    e.preventDefault();
                    this.classList.remove('dragover');
                    
                    const input = this.querySelector('input[type="file"]');
                    const fileList = this.parentNode.querySelector('.file-list');
                    
                    input.files = e.dataTransfer.files;
                    displayFiles(e.dataTransfer.files, fileList);
                });
            });
        });

        async function generateReport() {
            const formData = new FormData();
            
            // Coletar propostas técnicas
            const technicalProposals = [];
            for (let i = 1; i <= technicalProposalCount; i++) {
                const companyInput = document.querySelector(`#technical-proposals .proposal-group:nth-child(${i}) .company-input`);
                const fileInput = document.getElementById(`tech-file-${i}`);
                
                if (companyInput && companyInput.value && fileInput && fileInput.files.length > 0) {
                    technicalProposals.push({
                        company: companyInput.value,
                        file: fileInput.files[0]
                    });
                }
            }
            
            // Coletar propostas comerciais
            const commercialProposals = [];
            for (let i = 1; i <= commercialProposalCount; i++) {
                const companyInput = document.querySelector(`#commercial-proposals .proposal-group:nth-child(${i}) .company-input`);
                const cnpjInput = document.querySelector(`#commercial-proposals .proposal-group:nth-child(${i}) .cnpj-input`);
                const fileInput = document.getElementById(`comm-file-${i}`);
                
                if (companyInput && companyInput.value && fileInput && fileInput.files.length > 0) {
                    commercialProposals.push({
                        company: companyInput.value,
                        cnpj: cnpjInput ? cnpjInput.value : '',
                        file: fileInput.files[0]
                    });
                }
            }
            
            // Validações
            if (technicalProposals.length === 0) {
                alert('Por favor, adicione pelo menos uma proposta técnica.');
                return;
            }
            
            if (commercialProposals.length === 0) {
                alert('Por favor, adicione pelo menos uma proposta comercial.');
                return;
            }
            
            // Adicionar arquivos ao FormData
            technicalProposals.forEach((proposal, index) => {
                formData.append(`technical_${index}`, proposal.file);
                formData.append(`technical_${index}_company`, proposal.company);
            });
            
            commercialProposals.forEach((proposal, index) => {
                formData.append(`commercial_${index}`, proposal.file);
                formData.append(`commercial_${index}_company`, proposal.company);
                formData.append(`commercial_${index}_cnpj`, proposal.cnpj);
            });
            
            // Mostrar loading
            document.getElementById('loading').style.display = 'block';
            document.getElementById('generate-btn').disabled = true;
            document.getElementById('result').style.display = 'none';
            
            try {
                const response = await fetch('/analyze', {
                    method: 'POST',
                    body: formData
                });
                
                if (!response.ok) {
                    throw new Error(`Erro HTTP: ${response.status}`);
                }
                
                const result = await response.json();
                
                if (result.success) {
                    // Mostrar resultado
                    document.getElementById('loading').style.display = 'none';
                    document.getElementById('result').style.display = 'block';
                    
                    // Configurar links de download
                    document.getElementById('download-pdf').href = `/download/${result.report_id}/pdf`;
                    document.getElementById('download-md').href = `/download/${result.report_id}/markdown`;
                } else {
                    throw new Error(result.error || 'Erro desconhecido');
                }
                
            } catch (error) {
                console.error('Erro:', error);
                document.getElementById('loading').style.display = 'none';
                alert('Erro ao gerar relatório: ' + error.message);
            } finally {
                document.getElementById('generate-btn').disabled = false;
            }
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
        # Coletar arquivos e informações
        technical_files = []
        commercial_files = []
        
        # Processar arquivos técnicos
        for key in request.files:
            if key.startswith('technical_') and not key.endswith('_company'):
                file = request.files[key]
                company_key = f"{key}_company"
                company_name = request.form.get(company_key, 'Empresa Desconhecida')
                
                if file and allowed_file(file.filename):
                    filename = secure_filename(file.filename)
                    filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
                    file.save(filepath)
                    technical_files.append({
                        'path': filepath,
                        'company': company_name,
                        'filename': filename
                    })
        
        # Processar arquivos comerciais
        for key in request.files:
            if key.startswith('commercial_') and not key.endswith('_company') and not key.endswith('_cnpj'):
                file = request.files[key]
                company_key = f"{key}_company"
                cnpj_key = f"{key}_cnpj"
                company_name = request.form.get(company_key, 'Empresa Desconhecida')
                cnpj = request.form.get(cnpj_key, '')
                
                if file and allowed_file(file.filename):
                    filename = secure_filename(file.filename)
                    filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
                    file.save(filepath)
                    commercial_files.append({
                        'path': filepath,
                        'company': company_name,
                        'cnpj': cnpj,
                        'filename': filename
                    })
        
        if not technical_files or not commercial_files:
            return jsonify({'success': False, 'error': 'Pelo menos uma proposta técnica e uma comercial são necessárias'})
        
        # Processar propostas técnicas
        technical_analyses = []
        for file_info in technical_files:
            try:
                # Extrair texto baseado no tipo de arquivo
                if file_info['filename'].lower().endswith('.pdf'):
                    text = extract_text_from_pdf(file_info['path'])
                elif file_info['filename'].lower().endswith(('.docx', '.doc')):
                    text = extract_text_from_docx(file_info['path'])
                elif file_info['filename'].lower().endswith('.zip'):
                    # Processar ZIP
                    extracted_files = extract_from_zip(file_info['path'])
                    text = ""
                    for extracted_file in extracted_files:
                        if extracted_file.lower().endswith('.pdf'):
                            text += extract_text_from_pdf(extracted_file) + "\n"
                        elif extracted_file.lower().endswith(('.docx', '.doc')):
                            text += extract_text_from_docx(extracted_file) + "\n"
                else:
                    text = ""
                
                if text:
                    analysis = analyze_technical_proposal_advanced(text, file_info['company'])
                    technical_analyses.append(analysis)
                
                # Limpar arquivo temporário
                os.remove(file_info['path'])
                
            except Exception as e:
                print(f"Erro ao processar arquivo técnico {file_info['filename']}: {e}")
                continue
        
        # Processar propostas comerciais
        commercial_analyses = []
        for file_info in commercial_files:
            try:
                excel_data = {}
                
                # Extrair dados baseado no tipo de arquivo
                if file_info['filename'].lower().endswith(('.xlsx', '.xls')):
                    excel_data = extract_data_from_excel(file_info['path'])
                elif file_info['filename'].lower().endswith('.pdf'):
                    # Para PDFs comerciais, tentar extrair dados básicos
                    text = extract_text_from_pdf(file_info['path'])
                    # Converter texto em estrutura similar ao Excel
                    excel_data = {
                        'precos': extract_prices_from_text(text),
                        'bdi': extract_bdi_from_text(text),
                        'composicao_custos': {},
                        'condicoes_comerciais': extract_commercial_from_text(text),
                        'tabela_servicos': []
                    }
                
                # Adicionar CNPJ se fornecido
                if file_info['cnpj']:
                    excel_data['cnpj'] = file_info['cnpj']
                
                analysis = analyze_commercial_proposal_advanced(excel_data, file_info['company'])
                commercial_analyses.append(analysis)
                
                # Limpar arquivo temporário
                os.remove(file_info['path'])
                
            except Exception as e:
                print(f"Erro ao processar arquivo comercial {file_info['filename']}: {e}")
                continue
        
        if not technical_analyses or not commercial_analyses:
            return jsonify({'success': False, 'error': 'Erro ao processar os arquivos. Verifique os formatos e tente novamente.'})
        
        # Gerar análise comparativa
        comparative_analysis = generate_comparative_analysis_advanced(technical_analyses, commercial_analyses)
        
        # Gerar relatório
        report_content = generate_detailed_report(technical_analyses, commercial_analyses, comparative_analysis)
        
        # Salvar relatório
        report_id = datetime.now().strftime('%Y%m%d_%H%M%S')
        md_path = os.path.join(app.config['UPLOAD_FOLDER'], f"{report_id}.md")
        
        with open(md_path, 'w', encoding='utf-8') as f:
            f.write(report_content)
        
        # Liberar memória
        del technical_analyses, commercial_analyses, comparative_analysis
        gc.collect()
        
        return jsonify({
            'success': True,
            'report_id': report_id,
            'message': 'Relatório gerado com sucesso!'
        })
        
    except Exception as e:
        print(f"Erro geral na análise: {e}")
        return jsonify({'success': False, 'error': f'Erro interno: {str(e)}'})

def extract_prices_from_text(text):
    """Extrair preços de texto PDF"""
    prices = {}
    try:
        # Procurar por preço total
        total_patterns = [
            r'total\s*:?\s*(R\$\s*)?(\d{1,3}(?:[.,]\d{3})*(?:[.,]\d{2})?)',
            r'valor\s*total\s*:?\s*(R\$\s*)?(\d{1,3}(?:[.,]\d{3})*(?:[.,]\d{2})?)'
        ]
        
        for pattern in total_patterns:
            match = re.search(pattern, text, re.IGNORECASE)
            if match:
                price_str = match.group(2).replace('.', '').replace(',', '.')
                try:
                    prices['preco_total'] = float(price_str)
                    break
                except:
                    continue
    except Exception as e:
        print(f"Erro ao extrair preços do texto: {e}")
    
    return prices

def extract_bdi_from_text(text):
    """Extrair BDI de texto PDF"""
    bdi_data = {}
    try:
        bdi_patterns = [
            r'bdi\s*:?\s*(\d{1,2}(?:[.,]\d{1,2})?)\s*%'
        ]
        
        for pattern in bdi_patterns:
            match = re.search(pattern, text, re.IGNORECASE)
            if match:
                bdi_str = match.group(1).replace(',', '.')
                try:
                    bdi_data['percentual'] = float(bdi_str)
                    break
                except:
                    continue
    except Exception as e:
        print(f"Erro ao extrair BDI do texto: {e}")
    
    return bdi_data

def extract_commercial_from_text(text):
    """Extrair condições comerciais de texto PDF"""
    conditions = {}
    try:
        # Procurar por condições de pagamento
        payment_patterns = [
            r'pagamento\s*:?\s*([^\n]+)'
        ]
        
        for pattern in payment_patterns:
            match = re.search(pattern, text, re.IGNORECASE)
            if match:
                conditions['condicoes_pagamento'] = match.group(1).strip()
                break
        
        # Procurar por garantia
        warranty_patterns = [
            r'garantia\s*:?\s*([^\n]+)'
        ]
        
        for pattern in warranty_patterns:
            match = re.search(pattern, text, re.IGNORECASE)
            if match:
                conditions['garantia'] = match.group(1).strip()
                break
    except Exception as e:
        print(f"Erro ao extrair condições comerciais: {e}")
    
    return conditions

@app.route('/download/<report_id>/<format>')
def download_report(report_id, format):
    try:
        if format == 'markdown':
            file_path = os.path.join(app.config['UPLOAD_FOLDER'], f"{report_id}.md")
            return send_file(file_path, as_attachment=True, download_name=f"analise_comparativa_{report_id}.md")
        elif format == 'pdf':
            md_file_path = os.path.join(app.config['UPLOAD_FOLDER'], f"{report_id}.md")
            pdf_file_path = os.path.join(app.config['UPLOAD_FOLDER'], f"analise_comparativa_{report_id}.pdf")

            # Verificar se o arquivo markdown existe
            if not os.path.exists(md_file_path):
                return jsonify({'error': 'Arquivo de relatório não encontrado.'}), 404

            # Ler conteúdo do markdown
            with open(md_file_path, 'r', encoding='utf-8') as f:
                markdown_content = f.read()

            # Criar PDF
            if create_pdf_report(markdown_content, pdf_file_path):
                return send_file(pdf_file_path, as_attachment=True, download_name=f"analise_comparativa_{report_id}.pdf")
            else:
                return jsonify({'error': 'Erro ao gerar PDF.'}), 500

        else:
            return jsonify({'error': 'Formato não suportado.'}), 400

    except Exception as e:
        return jsonify({'error': str(e)}), 500

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=10000, debug=False)
