import os
import tempfile
import zipfile
import re
import gc
import json
from datetime import datetime
from flask import Flask, request, jsonify, send_file, render_template_string
from flask_cors import CORS
from werkzeug.utils import secure_filename
import PyPDF2
from docx import Document
import pandas as pd
from reportlab.lib.pagesizes import letter, A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.lib import colors
from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_JUSTIFY
import subprocess

app = Flask(__name__)
CORS(app)

# Configurações
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024  # 50MB max

# Criar diretório de uploads se não existir
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)

def extract_text_from_pdf(file_path):
    """Extrai texto de PDF com tratamento robusto"""
    try:
        text = ""
        with open(file_path, 'rb') as file:
            reader = PyPDF2.PdfReader(file)
            for page in reader.pages:
                page_text = page.extract_text()
                if page_text:
                    text += page_text + "\n"
        return text.strip()
    except Exception as e:
        print(f"Erro ao extrair PDF {file_path}: {e}")
        return ""

def extract_text_from_docx(file_path):
    """Extrai texto de DOCX"""
    try:
        doc = Document(file_path)
        text = ""
        for paragraph in doc.paragraphs:
            text += paragraph.text + "\n"
        return text.strip()
    except Exception as e:
        print(f"Erro ao extrair DOCX {file_path}: {e}")
        return ""

def extract_data_from_excel(file_path):
    """Extração robusta de dados de Excel"""
    try:
        # Ler todas as abas do Excel
        excel_data = pd.read_excel(file_path, sheet_name=None, engine='openpyxl')
        
        extracted_data = {
            'empresa': '',
            'cnpj': '',
            'preco_total': 0.0,
            'bdi_total': 0.0,
            'condicoes_pagamento': '',
            'garantia': '',
            'treinamento': '',
            'seguros': '',
            'composicao_custos': {},
            'tabela_servicos': [],
            'outras_informacoes': ''
        }
        
        # Processar cada aba
        for sheet_name, df in excel_data.items():
            sheet_lower = sheet_name.lower()
            
            # ABA CARTA - Informações gerais
            if 'carta' in sheet_lower or 'resumo' in sheet_lower or 'geral' in sheet_lower:
                for idx, row in df.iterrows():
                    if pd.notna(row.iloc[0]):
                        cell_text = str(row.iloc[0]).lower()
                        cell_value = str(row.iloc[1]) if len(row) > 1 and pd.notna(row.iloc[1]) else ''
                        
                        # Extrair empresa
                        if 'empresa' in cell_text and not extracted_data['empresa']:
                            extracted_data['empresa'] = cell_value
                        
                        # Extrair CNPJ
                        if 'cnpj' in cell_text:
                            cnpj_match = re.search(r'\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2}', cell_value)
                            if cnpj_match:
                                extracted_data['cnpj'] = cnpj_match.group()
                        
                        # Extrair preço total
                        if 'preço total' in cell_text or 'valor total' in cell_text:
                            price_match = re.search(r'R\$\s*([\d.,]+)', cell_value.replace('.', '').replace(',', '.'))
                            if price_match:
                                try:
                                    # Tratar formatação brasileira de números
                                    price_str = price_match.group(1)
                                    if '.' in price_str:
                                        # Se tem ponto, assumir que é separador de milhares
                                        price_str = price_str.replace('.', '')
                                    extracted_data['preco_total'] = float(price_str)
                                except:
                                    pass
                        
                        # Extrair condições comerciais
                        if 'pagamento' in cell_text:
                            extracted_data['condicoes_pagamento'] = cell_value
                        if 'garantia' in cell_text:
                            extracted_data['garantia'] = cell_value
                        if 'treinamento' in cell_text:
                            extracted_data['treinamento'] = cell_value
                        if 'seguro' in cell_text:
                            extracted_data['seguros'] = cell_value
                        if 'outras informações' in cell_text or 'observações' in cell_text:
                            extracted_data['outras_informacoes'] = cell_value
            
            # ABA ITENS SERVIÇOS - Tabela de serviços
            elif 'serviço' in sheet_lower or 'item' in sheet_lower:
                # Procurar por colunas de preço
                for col in df.columns:
                    if 'preço total' in str(col).lower() or 'valor total' in str(col).lower():
                        # Somar todos os valores da coluna
                        total = 0
                        for val in df[col]:
                            if pd.notna(val) and isinstance(val, (int, float)):
                                total += val
                            elif pd.notna(val):
                                # Tentar extrair número do texto
                                val_str = str(val).replace('R$', '').replace('.', '').replace(',', '.')
                                try:
                                    num_val = float(re.sub(r'[^\d.]', '', val_str))
                                    total += num_val
                                except:
                                    pass
                        
                        if total > extracted_data['preco_total']:
                            extracted_data['preco_total'] = total
                
                # Extrair itens da tabela
                if len(df) > 0:
                    for idx, row in df.iterrows():
                        if pd.notna(row.iloc[0]) and str(row.iloc[0]).strip():
                            item = {
                                'item': str(row.iloc[0]) if pd.notna(row.iloc[0]) else '',
                                'descricao': str(row.iloc[1]) if len(row) > 1 and pd.notna(row.iloc[1]) else '',
                                'quantidade': str(row.iloc[3]) if len(row) > 3 and pd.notna(row.iloc[3]) else '',
                                'preco_unitario': str(row.iloc[4]) if len(row) > 4 and pd.notna(row.iloc[4]) else '',
                                'preco_total': str(row.iloc[5]) if len(row) > 5 and pd.notna(row.iloc[5]) else ''
                            }
                            extracted_data['tabela_servicos'].append(item)
            
            # ABA BDI
            elif 'bdi' in sheet_lower:
                for idx, row in df.iterrows():
                    if pd.notna(row.iloc[0]):
                        cell_text = str(row.iloc[0]).lower()
                        if 'total' in cell_text and len(row) > 1 and pd.notna(row.iloc[1]):
                            try:
                                bdi_val = float(row.iloc[1])
                                extracted_data['bdi_total'] = bdi_val
                            except:
                                pass
            
            # ABA COMPOSIÇÃO DE CUSTOS
            elif 'custo' in sheet_lower or 'comp' in sheet_lower:
                for idx, row in df.iterrows():
                    if pd.notna(row.iloc[0]) and len(row) > 1 and pd.notna(row.iloc[1]):
                        categoria = str(row.iloc[0]).lower()
                        try:
                            valor = float(row.iloc[1])
                            if 'mão de obra' in categoria or 'mao de obra' in categoria:
                                extracted_data['composicao_custos']['mao_obra'] = valor
                            elif 'material' in categoria:
                                extracted_data['composicao_custos']['materiais'] = valor
                            elif 'equipamento' in categoria:
                                extracted_data['composicao_custos']['equipamentos'] = valor
                        except:
                            pass
        
        return extracted_data
        
    except Exception as e:
        print(f"Erro ao extrair Excel {file_path}: {e}")
        return {
            'empresa': '',
            'cnpj': '',
            'preco_total': 0.0,
            'bdi_total': 0.0,
            'condicoes_pagamento': '',
            'garantia': '',
            'treinamento': '',
            'seguros': '',
            'composicao_custos': {},
            'tabela_servicos': [],
            'outras_informacoes': ''
        }

def extract_technical_data_from_pdf(text):
    """Extração robusta de dados técnicos de PDF"""
    try:
        data = {
            'empresa': '',
            'cnpj': '',
            'metodologia': '',
            'prazo_total': '',
            'equipe_total': 0,
            'perfis_equipe': [],
            'equipamentos': [],
            'materiais': [],
            'cronograma': [],
            'obrigacoes': [],
            'exclusoes': [],
            'canteiro': '',
            'experiencia': []
        }
        
        # Extrair empresa
        empresa_match = re.search(r'Empresa:\s*([^\n]+)', text, re.IGNORECASE)
        if empresa_match:
            data['empresa'] = empresa_match.group(1).strip()
        
        # Extrair CNPJ
        cnpj_match = re.search(r'CNPJ:\s*(\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2})', text)
        if cnpj_match:
            data['cnpj'] = cnpj_match.group(1)
        
        # Extrair metodologia
        metodologia_patterns = [
            r'Metodologia[^:]*:\s*([^\.]+\.)',
            r'Abordagem Geral[^:]*:\s*([^\.]+\.)',
            r'metodologia[^\.]*([^\.]+\.)'
        ]
        for pattern in metodologia_patterns:
            match = re.search(pattern, text, re.IGNORECASE | re.DOTALL)
            if match:
                data['metodologia'] = match.group(1).strip()
                break
        
        # Extrair prazo total
        prazo_patterns = [
            r'Prazo:\s*(\d+)\s*dias',
            r'Prazo Total[^:]*:\s*(\d+)\s*dias',
            r'(\d+)\s*dias'
        ]
        for pattern in prazo_patterns:
            match = re.search(pattern, text, re.IGNORECASE)
            if match:
                data['prazo_total'] = f"{match.group(1)} dias"
                break
        
        # Extrair total de pessoas da equipe
        equipe_patterns = [
            r'Total de pessoas:\s*(\d+)',
            r'(\d+)\s*profissionais',
            r'equipe.*?(\d+).*?pessoas'
        ]
        for pattern in equipe_patterns:
            match = re.search(pattern, text, re.IGNORECASE)
            if match:
                data['equipe_total'] = int(match.group(1))
                break
        
        # Extrair perfis da equipe
        perfis_section = re.search(r'Perfis Profissionais:(.*?)(?=###|\n\n|\Z)', text, re.IGNORECASE | re.DOTALL)
        if perfis_section:
            perfis_text = perfis_section.group(1)
            perfis = re.findall(r'([A-Za-z\s]+):\s*([^\n]+)', perfis_text)
            data['perfis_equipe'] = [{'cargo': p[0].strip(), 'descricao': p[1].strip()} for p in perfis]
        
        # Extrair equipamentos
        equipamentos_section = re.search(r'Lista de Equipamentos:(.*?)(?=###|\n\n|\Z)', text, re.IGNORECASE | re.DOTALL)
        if equipamentos_section:
            equip_text = equipamentos_section.group(1)
            equipamentos = re.findall(r'\|\s*\d+\s*\|\s*([^|]+)\s*\|\s*(\d+)\s*\|\s*([^|]+)\s*\|', equip_text)
            data['equipamentos'] = [{'descricao': e[0].strip(), 'quantidade': e[1].strip(), 'tecnologia': e[2].strip()} for e in equipamentos]
        
        # Extrair materiais
        materiais_section = re.search(r'Lista de Materiais:(.*?)(?=###|\n\n|\Z)', text, re.IGNORECASE | re.DOTALL)
        if materiais_section:
            mat_text = materiais_section.group(1)
            materiais = re.findall(r'\|\s*\d+\s*\|\s*([^|]+)\s*\|\s*([^|]+)\s*\|\s*(\d+)\s*\|\s*([^|]+)\s*\|', mat_text)
            data['materiais'] = [{'descricao': m[0].strip(), 'unidade': m[1].strip(), 'quantidade': m[2].strip(), 'especificacao': m[3].strip()} for m in materiais]
        
        # Extrair marcos do cronograma
        marcos_section = re.search(r'Marcos Principais:(.*?)(?=###|\n\n|\Z)', text, re.IGNORECASE | re.DOTALL)
        if marcos_section:
            marcos_text = marcos_section.group(1)
            marcos = re.findall(r'Marco \d+:\s*([^-]+)-\s*(\d+\s*dias)', marcos_text)
            data['cronograma'] = [{'marco': m[0].strip(), 'prazo': m[1].strip()} for m in marcos]
        
        # Extrair obrigações
        obrigacoes_section = re.search(r'Obrigações da Proponente:(.*?)(?=###|\n\n|Obrigações do Contratante|\Z)', text, re.IGNORECASE | re.DOTALL)
        if obrigacoes_section:
            obrig_text = obrigacoes_section.group(1)
            obrigacoes = re.findall(r'Obrigação:\s*([^\n]+)', obrig_text)
            data['obrigacoes'] = [o.strip() for o in obrigacoes]
        
        # Extrair exclusões
        exclusoes_section = re.search(r'Itens Não Incluídos:(.*?)(?=###|\n\n|\Z)', text, re.IGNORECASE | re.DOTALL)
        if exclusoes_section:
            excl_text = exclusoes_section.group(1)
            exclusoes = re.findall(r'Exclusão:\s*([^\n]+)', excl_text)
            data['exclusoes'] = [e.strip() for e in exclusoes]
        
        # Extrair informações sobre canteiro
        canteiro_section = re.search(r'Informações sobre Canteiro:(.*?)(?=###|\n\n|\Z)', text, re.IGNORECASE | re.DOTALL)
        if canteiro_section:
            data['canteiro'] = canteiro_section.group(1).strip()
        
        # Extrair experiência comprovada
        exp_section = re.search(r'Projetos Similares:(.*?)(?=###|\n\n|\Z)', text, re.IGNORECASE | re.DOTALL)
        if exp_section:
            exp_text = exp_section.group(1)
            projetos = re.findall(r'Projeto:\s*([^\n]+)', exp_text)
            data['experiencia'] = [p.strip() for p in projetos]
        
        return data
        
    except Exception as e:
        print(f"Erro ao extrair dados técnicos: {e}")
        return {
            'empresa': '',
            'cnpj': '',
            'metodologia': '',
            'prazo_total': '',
            'equipe_total': 0,
            'perfis_equipe': [],
            'equipamentos': [],
            'materiais': [],
            'cronograma': [],
            'obrigacoes': [],
            'exclusoes': [],
            'canteiro': '',
            'experiencia': []
        }

def calculate_technical_score(tech_data):
    """Calcula score técnico baseado na completude dos dados"""
    score = 0
    max_score = 8
    
    # Metodologia (peso 2)
    if tech_data.get('metodologia'):
        score += 2
    
    # Equipe (peso 1)
    if tech_data.get('equipe_total', 0) > 0:
        score += 1
    
    # Equipamentos (peso 1)
    if tech_data.get('equipamentos'):
        score += 1
    
    # Materiais (peso 1)
    if tech_data.get('materiais'):
        score += 1
    
    # Cronograma (peso 1)
    if tech_data.get('prazo_total') or tech_data.get('cronograma'):
        score += 1
    
    # Obrigações (peso 1)
    if tech_data.get('obrigacoes'):
        score += 1
    
    # Experiência (peso 1)
    if tech_data.get('experiencia'):
        score += 1
    
    return (score / max_score) * 100

def calculate_commercial_score(comm_data):
    """Calcula score comercial baseado na completude dos dados"""
    score = 0
    max_score = 6
    
    # Preço total (peso 2)
    if comm_data.get('preco_total', 0) > 0:
        score += 2
    
    # BDI (peso 1)
    if comm_data.get('bdi_total', 0) > 0:
        score += 1
    
    # Condições de pagamento (peso 1)
    if comm_data.get('condicoes_pagamento'):
        score += 1
    
    # Garantia (peso 1)
    if comm_data.get('garantia'):
        score += 1
    
    # Composição de custos (peso 1)
    if comm_data.get('composicao_custos'):
        score += 1
    
    return (score / max_score) * 100

def generate_comparative_report(proposals_data):
    """Gera relatório comparativo detalhado"""
    try:
        # Separar dados técnicos e comerciais
        technical_data = []
        commercial_data = []
        
        for proposal in proposals_data:
            if proposal['type'] == 'technical':
                technical_data.append(proposal)
            elif proposal['type'] == 'commercial':
                commercial_data.append(proposal)
        
        # Combinar dados por empresa
        combined_data = {}
        
        # Processar dados técnicos
        for tech in technical_data:
            empresa = tech['data'].get('empresa', 'Empresa Desconhecida')
            if empresa not in combined_data:
                combined_data[empresa] = {'technical': {}, 'commercial': {}}
            combined_data[empresa]['technical'] = tech['data']
            combined_data[empresa]['technical']['score'] = calculate_technical_score(tech['data'])
        
        # Processar dados comerciais
        for comm in commercial_data:
            empresa = comm['data'].get('empresa', 'Empresa Desconhecida')
            if empresa not in combined_data:
                combined_data[empresa] = {'technical': {}, 'commercial': {}}
            combined_data[empresa]['commercial'] = comm['data']
            combined_data[empresa]['commercial']['score'] = calculate_commercial_score(comm['data'])
        
        # Gerar rankings
        tech_ranking = sorted(
            [(empresa, data['technical'].get('score', 0)) for empresa, data in combined_data.items() if data['technical']],
            key=lambda x: x[1], reverse=True
        )
        
        price_ranking = sorted(
            [(empresa, data['commercial'].get('preco_total', 0)) for empresa, data in combined_data.items() if data['commercial'].get('preco_total', 0) > 0],
            key=lambda x: x[1]
        )
        
        # Gerar relatório
        report = f"""# ANÁLISE COMPARATIVA DE PROPOSTAS

**Data:** {datetime.now().strftime('%d/%m/%Y')}

---

## RESUMO EXECUTIVO

### Rankings Gerais

**Ranking Técnico:**
"""
        
        for i, (empresa, score) in enumerate(tech_ranking, 1):
            report += f"**{i}º:** {empresa} - {score:.1f}%\n"
        
        report += "\n**Ranking de Preços:**\n"
        for i, (empresa, preco) in enumerate(price_ranking, 1):
            report += f"**{i}º:** {empresa} - R$ {preco:,.2f}\n"
        
        report += "\n---\n\n## ANÁLISE TÉCNICA COMPARATIVA\n\n"
        
        # Matriz de comparação técnica
        report += "### Matriz de Comparação Técnica\n\n"
        report += "| Empresa | Metodologia | Equipe | Equipamentos | Materiais | Cronograma | Score Total |\n"
        report += "|---------|-------------|--------|--------------|-----------|------------|-------------|\n"
        
        for empresa, data in combined_data.items():
            if data['technical']:
                tech = data['technical']
                metodologia = "✅" if tech.get('metodologia') else "❌"
                equipe = "✅" if tech.get('equipe_total', 0) > 0 else "❌"
                equipamentos = "✅" if tech.get('equipamentos') else "❌"
                materiais = "✅" if tech.get('materiais') else "❌"
                cronograma = "✅" if tech.get('prazo_total') or tech.get('cronograma') else "❌"
                score = tech.get('score', 0)
                
                report += f"| {empresa} | {metodologia} | {equipe} | {equipamentos} | {materiais} | {cronograma} | {score:.1f}% |\n"
        
        # Análise detalhada por empresa
        for empresa, data in combined_data.items():
            if data['technical']:
                tech = data['technical']
                report += f"\n### {empresa} - Análise Técnica Detalhada\n\n"
                
                if tech.get('metodologia'):
                    report += f"**Metodologia de Execução:**\n{tech['metodologia']}\n\n"
                
                report += f"**Equipe e Recursos:**\n"
                if tech.get('equipe_total', 0) > 0:
                    report += f"- Total de pessoas: {tech['equipe_total']}\n"
                if tech.get('perfis_equipe'):
                    report += f"- Perfis identificados: {len(tech['perfis_equipe'])}\n"
                
                if tech.get('equipamentos'):
                    report += f"- Equipamentos: {len(tech['equipamentos'])} itens\n"
                if tech.get('materiais'):
                    report += f"- Materiais: {len(tech['materiais'])} itens\n"
                
                if tech.get('prazo_total'):
                    report += f"**Cronograma:**\n- Prazo: {tech['prazo_total']}\n"
                if tech.get('cronograma'):
                    report += f"- Marcos: {len(tech['cronograma'])} identificados\n"
                
                report += "\n"
        
        report += "\n---\n\n## ANÁLISE COMERCIAL COMPARATIVA\n\n"
        
        # Tabela de preços e condições
        report += "### Resumo de Preços e Condições\n\n"
        report += "| Empresa | Preço Total | BDI | Condições Pagamento | Garantia | Score Comercial |\n"
        report += "|---------|-------------|-----|---------------------|----------|----------------|\n"
        
        for empresa, data in combined_data.items():
            if data['commercial']:
                comm = data['commercial']
                preco = f"R$ {comm.get('preco_total', 0):,.2f}" if comm.get('preco_total', 0) > 0 else "Não informado"
                bdi = f"{comm.get('bdi_total', 0):.2f}%" if comm.get('bdi_total', 0) > 0 else "Não informado"
                pagamento = comm.get('condicoes_pagamento', 'Não especificado')[:20] + "..." if len(comm.get('condicoes_pagamento', '')) > 20 else comm.get('condicoes_pagamento', 'Não especificado')
                garantia = comm.get('garantia', 'Não especificada')[:20] + "..." if len(comm.get('garantia', '')) > 20 else comm.get('garantia', 'Não especificada')
                score = comm.get('score', 0)
                
                report += f"| {empresa} | {preco} | {bdi} | {pagamento} | {garantia} | {score:.0f}% |\n"
        
        # Análise comercial detalhada
        for empresa, data in combined_data.items():
            if data['commercial']:
                comm = data['commercial']
                report += f"\n### {empresa} - Análise Comercial Detalhada\n\n"
                
                report += f"**Preços e Composição:**\n"
                if comm.get('preco_total', 0) > 0:
                    report += f"- Preço Total: R$ {comm['preco_total']:,.2f}\n"
                if comm.get('bdi_total', 0) > 0:
                    report += f"- BDI: {comm['bdi_total']:.2f}%\n"
                
                if comm.get('composicao_custos'):
                    report += f"\n**Composição de Custos:**\n"
                    custos = comm['composicao_custos']
                    if custos.get('mao_obra'):
                        report += f"- Mão de Obra: R$ {custos['mao_obra']:,.2f}\n"
                    if custos.get('materiais'):
                        report += f"- Materiais: R$ {custos['materiais']:,.2f}\n"
                    if custos.get('equipamentos'):
                        report += f"- Equipamentos: R$ {custos['equipamentos']:,.2f}\n"
                
                report += f"\n**Condições Comerciais:**\n"
                if comm.get('condicoes_pagamento'):
                    report += f"- Pagamento: {comm['condicoes_pagamento']}\n"
                if comm.get('garantia'):
                    report += f"- Garantia: {comm['garantia']}\n"
                if comm.get('treinamento'):
                    report += f"- Treinamento: {comm['treinamento']}\n"
                if comm.get('seguros'):
                    report += f"- Seguros: {comm['seguros']}\n"
                
                report += "\n"
        
        report += "\n---\n\n## CONCLUSÕES E RECOMENDAÇÕES\n\n"
        
        # Determinar melhores propostas
        melhor_tecnica = tech_ranking[0][0] if tech_ranking else "A definir"
        melhor_comercial = price_ranking[0][0] if price_ranking else "A definir"
        
        report += f"### Análise Comparativa Final\n\n"
        report += f"**Melhor Proposta Técnica:** {melhor_tecnica}\n"
        report += f"**Melhor Proposta Comercial:** {melhor_comercial}\n\n"
        
        # Recomendações
        report += f"### Recomendações Finais\n\n"
        report += f"**Para Tomada de Decisão:**\n"
        report += f"1. **Análise Técnica:** Considere a proposta com maior score técnico para garantir qualidade de execução.\n"
        report += f"2. **Análise Comercial:** Avalie não apenas o menor preço, mas também as condições de pagamento e garantias oferecidas.\n"
        report += f"3. **Custo-Benefício:** Busque o equilíbrio entre qualidade técnica e vantagem comercial.\n\n"
        
        report += f"**Próximos Passos Sugeridos:**\n"
        report += f"1. **Esclarecimentos:** Solicite esclarecimentos para propostas com informações incompletas.\n"
        report += f"2. **Negociação:** Considere negociar condições com as propostas melhor classificadas.\n"
        report += f"3. **Verificação:** Confirme referências e capacidade técnica das empresas.\n\n"
        
        return report
        
    except Exception as e:
        print(f"Erro ao gerar relatório: {e}")
        return "Erro ao gerar relatório comparativo."

def create_pdf_from_markdown(markdown_content, output_path):
    """Converte markdown para PDF usando ReportLab"""
    try:
        doc = SimpleDocTemplate(output_path, pagesize=A4)
        styles = getSampleStyleSheet()
        story = []
        
        # Estilo personalizado
        title_style = ParagraphStyle(
            'CustomTitle',
            parent=styles['Heading1'],
            fontSize=16,
            spaceAfter=12,
            textColor=colors.darkblue
        )
        
        heading_style = ParagraphStyle(
            'CustomHeading',
            parent=styles['Heading2'],
            fontSize=14,
            spaceAfter=10,
            textColor=colors.darkblue
        )
        
        # Processar markdown linha por linha
        lines = markdown_content.split('\n')
        for line in lines:
            line = line.strip()
            if not line:
                story.append(Spacer(1, 6))
            elif line.startswith('# '):
                story.append(Paragraph(line[2:], title_style))
            elif line.startswith('## '):
                story.append(Paragraph(line[3:], heading_style))
            elif line.startswith('### '):
                story.append(Paragraph(line[4:], styles['Heading3']))
            elif line.startswith('**') and line.endswith('**'):
                story.append(Paragraph(f"<b>{line[2:-2]}</b>", styles['Normal']))
            elif line.startswith('|'):
                # Processar tabela (simplificado)
                story.append(Paragraph(line.replace('|', ' | '), styles['Normal']))
            else:
                story.append(Paragraph(line, styles['Normal']))
        
        doc.build(story)
        return True
        
    except Exception as e:
        print(f"Erro ao criar PDF: {e}")
        return False

@app.route('/')
def index():
    return render_template_string('''
<!DOCTYPE html>
<html lang="pt-BR">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Proposal Analyzer Pro</title>
    <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); min-height: 100vh; }
        .container { max-width: 1200px; margin: 0 auto; padding: 20px; }
        .header { text-align: center; color: white; margin-bottom: 40px; }
        .header h1 { font-size: 3rem; margin-bottom: 10px; text-shadow: 2px 2px 4px rgba(0,0,0,0.3); }
        .header p { font-size: 1.2rem; opacity: 0.9; }
        .main-content { background: white; border-radius: 20px; padding: 40px; box-shadow: 0 20px 40px rgba(0,0,0,0.1); }
        .upload-section { margin-bottom: 40px; }
        .upload-section h2 { color: #333; margin-bottom: 20px; font-size: 1.5rem; }
        .proposal-group { background: #f8f9fa; border-radius: 15px; padding: 25px; margin-bottom: 25px; border-left: 5px solid #667eea; }
        .proposal-group h3 { color: #667eea; margin-bottom: 15px; }
        .file-input-group { margin-bottom: 15px; }
        .file-input-group label { display: block; margin-bottom: 5px; font-weight: 600; color: #555; }
        .file-input { width: 100%; padding: 12px; border: 2px dashed #ddd; border-radius: 10px; background: white; transition: all 0.3s; }
        .file-input:hover { border-color: #667eea; }
        .company-input { width: 100%; padding: 12px; border: 2px solid #ddd; border-radius: 10px; margin-bottom: 10px; }
        .cnpj-input { width: 100%; padding: 12px; border: 2px solid #ddd; border-radius: 10px; }
        .add-proposal-btn { background: #28a745; color: white; border: none; padding: 12px 24px; border-radius: 10px; cursor: pointer; margin-top: 15px; }
        .add-proposal-btn:hover { background: #218838; }
        .analyze-btn { background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); color: white; border: none; padding: 15px 40px; border-radius: 50px; font-size: 1.1rem; cursor: pointer; display: block; margin: 30px auto; transition: transform 0.3s; }
        .analyze-btn:hover { transform: translateY(-2px); }
        .loading { display: none; text-align: center; margin: 20px 0; }
        .loading-spinner { border: 4px solid #f3f3f3; border-top: 4px solid #667eea; border-radius: 50%; width: 50px; height: 50px; animation: spin 1s linear infinite; margin: 0 auto 20px; }
        @keyframes spin { 0% { transform: rotate(0deg); } 100% { transform: rotate(360deg); } }
        .results { display: none; margin-top: 30px; }
        .download-btn { background: #17a2b8; color: white; border: none; padding: 12px 24px; border-radius: 10px; margin: 10px; cursor: pointer; text-decoration: none; display: inline-block; }
        .download-btn:hover { background: #138496; }
        .file-selected { border-color: #28a745 !important; background-color: #f8fff9 !important; }
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>🚀 Proposal Analyzer Pro</h1>
            <p>Análise Inteligente e Comparação de Propostas Técnicas e Comerciais</p>
        </div>
        
        <div class="main-content">
            <form id="analysisForm" enctype="multipart/form-data">
                <div class="upload-section">
                    <h2>📋 Upload de Propostas</h2>
                    <div id="proposalsContainer">
                        <div class="proposal-group">
                            <h3>Proposta Técnica 1</h3>
                            <div class="file-input-group">
                                <label>Nome da Empresa</label>
                                <input type="text" name="company_name_tech_1" class="company-input" placeholder="Digite o nome da empresa">
                            </div>
                            <div class="file-input-group">
                                <label>CNPJ (Opcional)</label>
                                <input type="text" name="cnpj_tech_1" class="cnpj-input" placeholder="00.000.000/0000-00">
                            </div>
                            <div class="file-input-group">
                                <label>Arquivo da Proposta Técnica</label>
                                <input type="file" name="technical_1" class="file-input" accept=".pdf,.docx,.doc" onchange="handleFileSelect(this)">
                            </div>
                        </div>
                        
                        <div class="proposal-group">
                            <h3>Proposta Comercial 1</h3>
                            <div class="file-input-group">
                                <label>Nome da Empresa</label>
                                <input type="text" name="company_name_comm_1" class="company-input" placeholder="Digite o nome da empresa">
                            </div>
                            <div class="file-input-group">
                                <label>CNPJ (Opcional)</label>
                                <input type="text" name="cnpj_comm_1" class="cnpj-input" placeholder="00.000.000/0000-00">
                            </div>
                            <div class="file-input-group">
                                <label>Arquivo da Proposta Comercial</label>
                                <input type="file" name="commercial_1" class="file-input" accept=".xlsx,.xls" onchange="handleFileSelect(this)">
                            </div>
                        </div>
                    </div>
                    
                    <button type="button" class="add-proposal-btn" onclick="addProposal()">+ Adicionar Proposta Técnica</button>
                    <button type="button" class="add-proposal-btn" onclick="addCommercialProposal()">+ Adicionar Proposta Comercial</button>
                </div>
                
                <button type="submit" class="analyze-btn">🧠 Gerar Relatório com Análise IA</button>
            </form>
            
            <div class="loading" id="loading">
                <div class="loading-spinner"></div>
                <p>Processando documentos e gerando análise...</p>
                <p>Isso pode levar alguns minutos. Por favor, aguarde.</p>
            </div>
            
            <div class="results" id="results">
                <h2>📊 Relatório Gerado</h2>
                <p>Seu relatório de análise comparativa foi gerado com sucesso!</p>
                <a href="#" id="downloadMarkdown" class="download-btn">📄 Download Markdown</a>
                <a href="#" id="downloadPDF" class="download-btn">📑 Download PDF</a>
            </div>
        </div>
    </div>

    <script>
        let proposalCount = 1;
        let commercialCount = 1;

        function handleFileSelect(input) {
            if (input.files.length > 0) {
                input.classList.add('file-selected');
            } else {
                input.classList.remove('file-selected');
            }
        }

        function addProposal() {
            proposalCount++;
            const container = document.getElementById('proposalsContainer');
            const newProposal = document.createElement('div');
            newProposal.className = 'proposal-group';
            newProposal.innerHTML = `
                <h3>Proposta Técnica ${proposalCount}</h3>
                <div class="file-input-group">
                    <label>Nome da Empresa</label>
                    <input type="text" name="company_name_tech_${proposalCount}" class="company-input" placeholder="Digite o nome da empresa">
                </div>
                <div class="file-input-group">
                    <label>CNPJ (Opcional)</label>
                    <input type="text" name="cnpj_tech_${proposalCount}" class="cnpj-input" placeholder="00.000.000/0000-00">
                </div>
                <div class="file-input-group">
                    <label>Arquivo da Proposta Técnica</label>
                    <input type="file" name="technical_${proposalCount}" class="file-input" accept=".pdf,.docx,.doc" onchange="handleFileSelect(this)">
                </div>
            `;
            container.appendChild(newProposal);
        }

        function addCommercialProposal() {
            commercialCount++;
            const container = document.getElementById('proposalsContainer');
            const newProposal = document.createElement('div');
            newProposal.className = 'proposal-group';
            newProposal.innerHTML = `
                <h3>Proposta Comercial ${commercialCount}</h3>
                <div class="file-input-group">
                    <label>Nome da Empresa</label>
                    <input type="text" name="company_name_comm_${commercialCount}" class="company-input" placeholder="Digite o nome da empresa">
                </div>
                <div class="file-input-group">
                    <label>CNPJ (Opcional)</label>
                    <input type="text" name="cnpj_comm_${commercialCount}" class="cnpj-input" placeholder="00.000.000/0000-00">
                </div>
                <div class="file-input-group">
                    <label>Arquivo da Proposta Comercial</label>
                    <input type="file" name="commercial_${commercialCount}" class="file-input" accept=".xlsx,.xls" onchange="handleFileSelect(this)">
                </div>
            `;
            container.appendChild(newProposal);
        }

        document.getElementById('analysisForm').addEventListener('submit', async function(e) {
            e.preventDefault();
            
            const formData = new FormData(this);
            const loading = document.getElementById('loading');
            const results = document.getElementById('results');
            
            loading.style.display = 'block';
            results.style.display = 'none';
            
            try {
                const response = await fetch('/analyze', {
                    method: 'POST',
                    body: formData
                });
                
                const data = await response.json();
                
                if (data.success) {
                    document.getElementById('downloadMarkdown').href = `/download/${data.report_id}/markdown`;
                    document.getElementById('downloadPDF').href = `/download/${data.report_id}/pdf`;
                    results.style.display = 'block';
                } else {
                    alert('Erro: ' + data.error);
                }
            } catch (error) {
                alert('Erro na comunicação com o servidor: ' + error.message);
            } finally {
                loading.style.display = 'none';
            }
        });
    </script>
</body>
</html>
    ''')

@app.route('/analyze', methods=['POST'])
def analyze_proposals():
    try:
        # Processar arquivos enviados
        proposals_data = []
        
        # Processar propostas técnicas
        for key in request.files:
            if key.startswith('technical_'):
                file = request.files[key]
                if file and file.filename:
                    # Salvar arquivo
                    filename = secure_filename(file.filename)
                    file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
                    file.save(file_path)
                    
                    # Extrair texto
                    if filename.lower().endswith('.pdf'):
                        text = extract_text_from_pdf(file_path)
                    elif filename.lower().endswith(('.docx', '.doc')):
                        text = extract_text_from_docx(file_path)
                    else:
                        continue
                    
                    # Extrair dados técnicos
                    tech_data = extract_technical_data_from_pdf(text)
                    
                    # Adicionar nome da empresa do formulário se não extraído
                    proposal_num = key.split('_')[1]
                    company_name = request.form.get(f'company_name_tech_{proposal_num}', '')
                    if company_name and not tech_data.get('empresa'):
                        tech_data['empresa'] = company_name
                    
                    proposals_data.append({
                        'type': 'technical',
                        'data': tech_data,
                        'filename': filename
                    })
                    
                    # Limpar arquivo
                    os.remove(file_path)
        
        # Processar propostas comerciais
        for key in request.files:
            if key.startswith('commercial_'):
                file = request.files[key]
                if file and file.filename:
                    # Salvar arquivo
                    filename = secure_filename(file.filename)
                    file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
                    file.save(file_path)
                    
                    # Extrair dados comerciais
                    comm_data = extract_data_from_excel(file_path)
                    
                    # Adicionar nome da empresa do formulário se não extraído
                    proposal_num = key.split('_')[1]
                    company_name = request.form.get(f'company_name_comm_{proposal_num}', '')
                    if company_name and not comm_data.get('empresa'):
                        comm_data['empresa'] = company_name
                    
                    proposals_data.append({
                        'type': 'commercial',
                        'data': comm_data,
                        'filename': filename
                    })
                    
                    # Limpar arquivo
                    os.remove(file_path)
        
        if not proposals_data:
            return jsonify({'success': False, 'error': 'Nenhum arquivo válido foi enviado.'})
        
        # Gerar relatório
        report_content = generate_comparative_report(proposals_data)
        
        # Salvar relatório
        report_id = datetime.now().strftime('%Y%m%d_%H%M%S')
        report_path = os.path.join(app.config['UPLOAD_FOLDER'], f"analise_comparativa_{report_id}.md")
        
        with open(report_path, 'w', encoding='utf-8') as f:
            f.write(report_content)
        
        # Limpar memória
        del proposals_data
        gc.collect()
        
        return jsonify({
            'success': True,
            'report_id': report_id,
            'message': 'Análise concluída com sucesso!'
        })
        
    except Exception as e:
        print(f"Erro na análise: {e}")
        return jsonify({'success': False, 'error': str(e)})

@app.route('/download/<report_id>/<format>')
def download_report(report_id, format):
    try:
        if format == 'markdown':
            file_path = os.path.join(app.config['UPLOAD_FOLDER'], f"analise_comparativa_{report_id}.md")
            return send_file(file_path, as_attachment=True, download_name=f"analise_comparativa_{report_id}.md")
        elif format == 'pdf':
            md_file_path = os.path.join(app.config['UPLOAD_FOLDER'], f"analise_comparativa_{report_id}.md")
            pdf_file_path = os.path.join(app.config['UPLOAD_FOLDER'], f"analise_comparativa_{report_id}.pdf")

            # Verificar se o arquivo markdown existe
            if not os.path.exists(md_file_path):
                return jsonify({'error': 'Arquivo de relatório não encontrado.'}), 404

            # Ler conteúdo markdown
            with open(md_file_path, 'r', encoding='utf-8') as f:
                markdown_content = f.read()

            # Criar PDF
            if create_pdf_from_markdown(markdown_content, pdf_file_path):
                return send_file(pdf_file_path, as_attachment=True, download_name=f"analise_comparativa_{report_id}.pdf")
            else:
                return jsonify({'error': 'Erro ao gerar PDF.'}), 500

        else:
            return jsonify({'error': 'Formato não suportado.'}), 400

    except Exception as e:
        return jsonify({'error': str(e)}), 500

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=10000, debug=False)
