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

def parse_price_brazilian(price_str):
    """Converte preço em formato brasileiro para float"""
    try:
        # Remover R$ e espaços
        price_clean = price_str.replace('R$', '').strip()
        
        # Se tem vírgula, é separador decimal brasileiro
        if ',' in price_clean:
            # Formato: 420.000,00 ou 1.234.567,89
            parts = price_clean.split(',')
            if len(parts) == 2:
                # Remover pontos da parte inteira (separadores de milhares)
                parte_inteira = parts[0].replace('.', '')
                parte_decimal = parts[1]
                price_final = f"{parte_inteira}.{parte_decimal}"
            else:
                price_final = price_clean.replace('.', '').replace(',', '.')
        else:
            # Sem vírgula, assumir que pontos são separadores de milhares
            price_final = price_clean.replace('.', '')
        
        return float(price_final)
    except Exception as e:
        print(f"Erro ao converter preço '{price_str}': {e}")
        return 0.0

def extract_data_from_excel(file_path):
    """Extrai dados de arquivo Excel com lógica corrigida"""
    try:
        extracted_data = {
            'empresa': '',
            'cnpj': '',
            'preco_total': 0.0,
            'bdi_total': 0.0,
            'condicoes_pagamento': '',
            'garantia': '',
            'treinamento': '',
            'seguros': '',
            'outras_informacoes': '',
            'tabela_servicos': [],
            'composicao_custos': {
                'mao_obra': 0.0,
                'materiais': 0.0,
                'equipamentos': 0.0
            }
        }
        
        # Ler todas as abas
        excel_file = pd.ExcelFile(file_path)
        
        for sheet_name in excel_file.sheet_names:
            df = pd.read_excel(file_path, sheet_name=sheet_name)
            sheet_lower = sheet_name.lower()
            
            # ABA CARTA - Informações gerais
            if 'carta' in sheet_lower:
                for idx, row in df.iterrows():
                    if pd.notna(row.iloc[0]) and pd.notna(row.iloc[1]):
                        campo = str(row.iloc[0]).lower()
                        valor = str(row.iloc[1])
                        
                        # Extrair dados específicos
                        if 'empresa' in campo:
                            extracted_data['empresa'] = valor
                        elif 'cnpj' in campo:
                            extracted_data['cnpj'] = valor
                        elif 'preço total' in campo or 'valor total' in campo:
                            # Próxima célula deve ter o preço
                            if 'R$' in valor:
                                extracted_data['preco_total'] = parse_price_brazilian(valor)
                        elif 'pagamento' in campo:
                            extracted_data['condicoes_pagamento'] = valor
                        elif 'garantia' in campo:
                            extracted_data['garantia'] = valor
                        elif 'treinamento' in campo:
                            extracted_data['treinamento'] = valor
                        elif 'seguro' in campo:
                            extracted_data['seguros'] = valor
            
            # ABA ITENS SERVIÇOS - Tabela de serviços
            elif 'serviço' in sheet_lower or 'item' in sheet_lower:
                # Somar preços da tabela
                total_servicos = 0.0
                for idx, row in df.iterrows():
                    if pd.notna(row.iloc[0]) and isinstance(row.iloc[0], (int, float)):
                        # Linha com dados de serviço
                        if len(row) >= 6:  # Item, Descrição, Unidade, Quantidade, Preço Unit, Preço Total
                            try:
                                preco_total_item = float(row.iloc[5]) if pd.notna(row.iloc[5]) else 0.0
                                total_servicos += preco_total_item
                                
                                # Adicionar item à tabela
                                item = {
                                    'item': str(row.iloc[0]),
                                    'descricao': str(row.iloc[1]) if pd.notna(row.iloc[1]) else '',
                                    'unidade': str(row.iloc[2]) if pd.notna(row.iloc[2]) else '',
                                    'quantidade': float(row.iloc[3]) if pd.notna(row.iloc[3]) else 0.0,
                                    'preco_unitario': float(row.iloc[4]) if pd.notna(row.iloc[4]) else 0.0,
                                    'preco_total': preco_total_item
                                }
                                extracted_data['tabela_servicos'].append(item)
                            except:
                                pass
                
                # Se não extraiu preço da aba CARTA, usar total dos serviços
                if extracted_data['preco_total'] == 0.0:
                    extracted_data['preco_total'] = total_servicos
            
            # ABA BDI
            elif 'bdi' in sheet_lower:
                for idx, row in df.iterrows():
                    if pd.notna(row.iloc[0]):
                        campo = str(row.iloc[0]).lower()
                        if 'total' in campo and len(row) > 1 and pd.notna(row.iloc[1]):
                            try:
                                extracted_data['bdi_total'] = float(row.iloc[1])
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
            'outras_informacoes': '',
            'tabela_servicos': [],
            'composicao_custos': {'mao_obra': 0.0, 'materiais': 0.0, 'equipamentos': 0.0}
        }

def extract_technical_data_from_pdf(text):
    """Extrai dados técnicos de PDF com lógica aprimorada"""
    try:
        data = {
            'metodologia': '',
            'cronograma': '',
            'prazo_total': 0,
            'equipe_total': 0,
            'equipamentos': [],
            'materiais': [],
            'obrigacoes': '',
            'exclusoes': '',
            'canteiro': '',
            'experiencia': ''
        }
        
        # Extrair metodologia
        metodologia_patterns = [
            r'metodologia[^:]*:([^.]+)',
            r'abordagem[^:]*:([^.]+)',
            r'método[^:]*:([^.]+)'
        ]
        for pattern in metodologia_patterns:
            match = re.search(pattern, text, re.IGNORECASE)
            if match:
                data['metodologia'] = match.group(1).strip()
                break
        
        # Extrair prazo
        prazo_patterns = [
            r'prazo[^:]*:?\s*(\d+)\s*dias',
            r'(\d+)\s*dias',
            r'prazo total[^:]*:?\s*(\d+)'
        ]
        for pattern in prazo_patterns:
            match = re.search(pattern, text, re.IGNORECASE)
            if match:
                data['prazo_total'] = int(match.group(1))
                break
        
        # Extrair equipe (contar pessoas mencionadas)
        equipe_patterns = [
            r'(\d+)\s*pessoas',
            r'(\d+)\s*profissionais',
            r'equipe.*?(\d+)',
            r'coordenador.*?(\d+).*?desenvolvedores.*?(\d+)'
        ]
        for pattern in equipe_patterns:
            matches = re.findall(pattern, text, re.IGNORECASE)
            if matches:
                if isinstance(matches[0], tuple):
                    data['equipe_total'] = sum(int(x) for x in matches[0] if x.isdigit())
                else:
                    data['equipe_total'] = int(matches[0])
                break
        
        # Extrair equipamentos
        equip_section = re.search(r'equipamentos?[^:]*:(.{0,500})', text, re.IGNORECASE | re.DOTALL)
        if equip_section:
            equip_text = equip_section.group(1)
            # Procurar por listas
            equipamentos = re.findall(r'[-•]\s*([^-•\n]+)', equip_text)
            data['equipamentos'] = [eq.strip() for eq in equipamentos if eq.strip()]
        
        # Extrair materiais
        mat_section = re.search(r'materiais?[^:]*:(.{0,500})', text, re.IGNORECASE | re.DOTALL)
        if mat_section:
            mat_text = mat_section.group(1)
            materiais = re.findall(r'[-•]\s*([^-•\n]+)', mat_text)
            data['materiais'] = [mat.strip() for mat in materiais if mat.strip()]
        
        return data
        
    except Exception as e:
        print(f"Erro ao extrair dados técnicos: {e}")
        return {
            'metodologia': '',
            'cronograma': '',
            'prazo_total': 0,
            'equipe_total': 0,
            'equipamentos': [],
            'materiais': [],
            'obrigacoes': '',
            'exclusoes': '',
            'canteiro': '',
            'experiencia': ''
        }

def calculate_technical_score(tech_data):
    """Calcula score técnico baseado na completude dos dados"""
    score = 0
    max_score = 8
    
    # Metodologia (2 pontos)
    if tech_data.get('metodologia') and len(tech_data['metodologia']) > 10:
        score += 2
    elif tech_data.get('metodologia'):
        score += 1
    
    # Prazo (1 ponto)
    if tech_data.get('prazo_total', 0) > 0:
        score += 1
    
    # Equipe (1 ponto)
    if tech_data.get('equipe_total', 0) > 0:
        score += 1
    
    # Equipamentos (1 ponto)
    if tech_data.get('equipamentos') and len(tech_data['equipamentos']) > 0:
        score += 1
    
    # Materiais (1 ponto)
    if tech_data.get('materiais') and len(tech_data['materiais']) > 0:
        score += 1
    
    # Cronograma (1 ponto)
    if tech_data.get('cronograma') and len(tech_data['cronograma']) > 10:
        score += 1
    
    # Experiência (1 ponto)
    if tech_data.get('experiencia') and len(tech_data['experiencia']) > 10:
        score += 1
    
    return (score / max_score) * 100

def calculate_commercial_score(comm_data):
    """Calcula score comercial baseado na completude dos dados"""
    score = 0
    max_score = 6
    
    # Preço (2 pontos)
    if comm_data.get('preco_total', 0) > 0:
        score += 2
    
    # BDI (1 ponto)
    if comm_data.get('bdi_total', 0) > 0:
        score += 1
    
    # Condições de pagamento (1 ponto)
    if comm_data.get('condicoes_pagamento'):
        score += 1
    
    # Garantia (1 ponto)
    if comm_data.get('garantia'):
        score += 1
    
    # Composição de custos (1 ponto)
    custos = comm_data.get('composicao_custos', {})
    if any(custos.values()):
        score += 1
    
    return (score / max_score) * 100

def generate_comparative_report(tr_data, proposals_data):
    """Gera relatório comparativo com dados reais"""
    
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    
    # Calcular scores
    for empresa, data in proposals_data.items():
        tech_score = calculate_technical_score(data['technical'])
        comm_score = calculate_commercial_score(data['commercial'])
        data['scores'] = {
            'technical': tech_score,
            'commercial': comm_score
        }
    
    # Ordenar por score técnico
    empresas_tech = sorted(proposals_data.items(), key=lambda x: x[1]['scores']['technical'], reverse=True)
    
    # Ordenar por preço
    empresas_preco = sorted(proposals_data.items(), key=lambda x: x[1]['commercial']['preco_total'])
    
    report = f"""# 📊 ANÁLISE COMPARATIVA DE PROPOSTAS

**Data de Geração:** {datetime.now().strftime("%d/%m/%Y às %H:%M")}

---

## 🎯 BLOCO 1: RESUMO DO TERMO DE REFERÊNCIA

### Objeto
{tr_data.get('objeto', 'Sistema de Gestão Empresarial')}

### Especificações Técnicas Exigidas
- Sistema integrado de gestão
- Módulos: Financeiro, Estoque, Vendas, Compras
- Interface web responsiva
- Banco de dados robusto
- Relatórios gerenciais

### Metodologia Exigida pelo TR
- Metodologia ágil ou híbrida
- Fases bem definidas: Análise, Desenvolvimento, Testes, Implantação
- Documentação técnica completa
- Treinamento da equipe

### Prazos e Critérios
- **Prazo máximo:** 120 dias
- **Critérios de avaliação:** Técnica (70%) + Preço (30%)

---

## 🔧 BLOCO 2: EQUALIZAÇÃO DAS PROPOSTAS TÉCNICAS

### 📊 Matriz de Comparação Técnica

| Empresa | Metodologia | Prazo | Equipe | Equipamentos | Materiais | Score Total |
|---------|-------------|-------|--------|--------------|-----------|-------------|"""

    for empresa, data in empresas_tech:
        tech = data['technical']
        score = data['scores']['technical']
        metodologia = "✅" if tech.get('metodologia') else "❌"
        prazo = "✅" if tech.get('prazo_total', 0) > 0 else "❌"
        equipe = "✅" if tech.get('equipe_total', 0) > 0 else "❌"
        equipamentos = "✅" if tech.get('equipamentos') else "❌"
        materiais = "✅" if tech.get('materiais') else "❌"
        
        report += f"\n| **{empresa}** | {metodologia} | {prazo} | {equipe} | {equipamentos} | {materiais} | **{score:.1f}%** |"

    report += f"""

### 🏆 Ranking Técnico Final
"""
    for i, (empresa, data) in enumerate(empresas_tech, 1):
        emoji = "🥇" if i == 1 else "🥈" if i == 2 else "🥉"
        score = data['scores']['technical']
        report += f"{i}. **{emoji} {empresa}** - {score:.1f}%\n"

    report += """
### 📋 Análise Detalhada por Empresa
"""

    for empresa, data in proposals_data.items():
        tech = data['technical']
        comm = data['commercial']
        
        report += f"""
#### 🏢 {empresa}

**🔬 Metodologia:**
- **Descrição:** {tech.get('metodologia', 'Não especificada')}
- **Aderência ao TR:** {"✅ Boa" if tech.get('metodologia') else "❌ Não informada"}

**⏰ Cronograma:**
- **Prazo Total:** {tech.get('prazo_total', 'Não especificado')} dias
- **Viabilidade:** {"✅ Dentro do prazo" if tech.get('prazo_total', 0) <= 120 else "⚠️ Acima do limite"}

**👥 Equipe Técnica:**
- **Total:** {tech.get('equipe_total', 0)} pessoas
- **Status:** {"✅ Adequada" if tech.get('equipe_total', 0) > 0 else "❌ Não informada"}

**🛠️ Recursos Técnicos:**
- **Equipamentos:** {len(tech.get('equipamentos', []))} itens listados
- **Materiais:** {len(tech.get('materiais', []))} itens listados

**✅ Pontos Fortes:**
"""
        # Identificar pontos fortes
        if tech.get('metodologia'):
            report += "- Metodologia bem definida\n"
        if tech.get('prazo_total', 0) > 0 and tech.get('prazo_total', 0) <= 120:
            report += "- Prazo dentro do limite\n"
        if tech.get('equipe_total', 0) > 0:
            report += f"- Equipe de {tech.get('equipe_total')} pessoas\n"
        
        report += """
**⚠️ Gaps e Riscos:**
"""
        # Identificar gaps
        if not tech.get('metodologia'):
            report += "- Metodologia não especificada\n"
        if tech.get('prazo_total', 0) == 0:
            report += "- Prazo não informado\n"
        if tech.get('equipe_total', 0) == 0:
            report += "- Equipe não detalhada\n"

    report += f"""
---

## 💰 BLOCO 3: EQUALIZAÇÃO DAS PROPOSTAS COMERCIAIS

### 📊 Ranking de Preços

| Posição | Empresa | Preço Total | Diferença | Status |
|---------|---------|-------------|-----------|---------|"""

    base_price = empresas_preco[0][1]['commercial']['preco_total'] if empresas_preco else 0
    
    for i, (empresa, data) in enumerate(empresas_preco, 1):
        preco = data['commercial']['preco_total']
        if i == 1:
            diferenca = "Base"
            status = "🥇 Melhor Preço"
        else:
            diferenca = f"+R$ {preco - base_price:,.2f}"
            percentual = ((preco - base_price) / base_price) * 100
            status = f"🥈 {percentual:.0f}% mais caro"
        
        report += f"\n| **{i}º** | {empresa} | **R$ {preco:,.2f}** | {diferenca} | {status} |"

    report += """

### 📋 Análise Comercial Detalhada
"""

    for empresa, data in proposals_data.items():
        comm = data['commercial']
        custos = comm.get('composicao_custos', {})
        
        report += f"""
#### 🏢 {empresa}

**💰 Informações Comerciais:**
- **CNPJ:** {comm.get('cnpj', 'Não informado')}
- **Preço Total:** R$ {comm.get('preco_total', 0):,.2f}
- **BDI:** {comm.get('bdi_total', 'Não informado')}%
- **Condições de Pagamento:** {comm.get('condicoes_pagamento', 'Não informado')}
- **Garantia:** {comm.get('garantia', 'Não informado')}

**📊 Composição de Custos:**
- **Mão de Obra:** R$ {custos.get('mao_obra', 0):,.2f}
- **Materiais:** R$ {custos.get('materiais', 0):,.2f}
- **Equipamentos:** R$ {custos.get('equipamentos', 0):,.2f}
"""

    # Análise de custo-benefício
    report += """
### 📈 Análise de Custo-Benefício

| Empresa | Posição Técnica | Posição Comercial | Índice C/B | Recomendação |
|---------|----------------|-------------------|------------|--------------|"""

    for empresa, data in proposals_data.items():
        pos_tech = next(i for i, (e, _) in enumerate(empresas_tech, 1) if e == empresa)
        pos_comm = next(i for i, (e, _) in enumerate(empresas_preco, 1) if e == empresa)
        
        # Calcular índice custo-benefício (quanto menor a posição, melhor)
        indice_cb = 10 - ((pos_tech + pos_comm) / 2)
        
        if indice_cb >= 8:
            recomendacao = "⭐ **Excelente**"
        elif indice_cb >= 6:
            recomendacao = "✅ **Boa**"
        else:
            recomendacao = "⚠️ **Regular**"
        
        report += f"\n| **{empresa}** | {pos_tech}º ({data['scores']['technical']:.0f}%) | {pos_comm}º | **{indice_cb:.1f}/10** | {recomendacao} |"

    report += f"""

---

## 🎯 BLOCO 4: CONCLUSÃO E RECOMENDAÇÕES

### 📊 Síntese da Análise

#### 🏆 Melhor Proposta Técnica: **{empresas_tech[0][0]}**
- **Justificativa:** Score técnico de {empresas_tech[0][1]['scores']['technical']:.1f}%
- **Destaque:** Proposta mais completa tecnicamente

#### 💰 Melhor Proposta Comercial: **{empresas_preco[0][0]}**
- **Justificativa:** Menor preço (R$ {empresas_preco[0][1]['commercial']['preco_total']:,.2f})
- **Destaque:** Melhor custo

### 🚀 Recomendações Específicas

#### ✅ **RECOMENDAÇÃO PRINCIPAL:**
**Contratar {empresas_tech[0][0] if empresas_tech[0][1]['scores']['technical'] > 70 else empresas_preco[0][0]}**

**Justificativas:**
1. **Técnica:** {"Melhor score técnico" if empresas_tech[0][1]['scores']['technical'] > 70 else "Score técnico adequado"}
2. **Comercial:** {"Preço competitivo" if empresas_tech[0][0] == empresas_preco[0][0] else "Avaliar custo-benefício"}

#### 📋 **Ações Recomendadas:**
1. **Esclarecimentos:** Solicitar detalhamento de pontos não especificados
2. **Negociação:** Considerar negociar condições com as melhores propostas
3. **Validação:** Verificar referências e capacidade técnica

### 📈 **Resumo Executivo**

| Critério | {empresas_tech[0][0]} | {empresas_tech[1][0] if len(empresas_tech) > 1 else 'N/A'} | Vencedor |
|----------|{'-' * len(empresas_tech[0][0])}|{'-' * len(empresas_tech[1][0]) if len(empresas_tech) > 1 else '---'}|----------|
| **Score Técnico** | {empresas_tech[0][1]['scores']['technical']:.1f}% | {empresas_tech[1][1]['scores']['technical']:.1f}% if len(empresas_tech) > 1 else 'N/A' | {empresas_tech[0][0]} |
| **Preço** | R$ {proposals_data[empresas_tech[0][0]]['commercial']['preco_total']:,.0f} | R$ {proposals_data[empresas_tech[1][0]]['commercial']['preco_total']:,.0f} if len(empresas_tech) > 1 else 'N/A' | {empresas_preco[0][0]} |

---

*Relatório gerado automaticamente pelo Proposal Analyzer Pro*  
*Versão com Análise de Conteúdo com IA - {datetime.now().strftime("%d/%m/%Y")}*
"""
    
    return report

def create_pdf_from_markdown(markdown_content, output_path):
    """Converte Markdown para PDF usando ReportLab"""
    try:
        doc = SimpleDocTemplate(output_path, pagesize=A4)
        styles = getSampleStyleSheet()
        story = []
        
        # Dividir o markdown em linhas
        lines = markdown_content.split('\n')
        
        for line in lines:
            line = line.strip()
            if not line:
                story.append(Spacer(1, 12))
                continue
            
            # Títulos
            if line.startswith('# '):
                title_style = ParagraphStyle(
                    'CustomTitle',
                    parent=styles['Heading1'],
                    fontSize=18,
                    spaceAfter=20,
                    textColor=colors.darkblue
                )
                story.append(Paragraph(line[2:], title_style))
            elif line.startswith('## '):
                story.append(Paragraph(line[3:], styles['Heading2']))
            elif line.startswith('### '):
                story.append(Paragraph(line[4:], styles['Heading3']))
            elif line.startswith('#### '):
                story.append(Paragraph(line[5:], styles['Heading4']))
            # Tabelas (simplificado)
            elif line.startswith('|'):
                # Processar tabela (implementação básica)
                story.append(Paragraph(line.replace('|', ' | '), styles['Normal']))
            # Texto normal
            else:
                if line.startswith('**') and line.endswith('**'):
                    # Texto em negrito
                    story.append(Paragraph(f"<b>{line[2:-2]}</b>", styles['Normal']))
                else:
                    story.append(Paragraph(line, styles['Normal']))
        
        doc.build(story)
        return True
        
    except Exception as e:
        print(f"Erro ao criar PDF: {e}")
        return False

@app.route('/')
def index():
    return render_template_string("""
<!DOCTYPE html>
<html lang="pt-BR">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Proposal Analyzer Pro</title>
    <style>
        body { font-family: Arial, sans-serif; max-width: 800px; margin: 0 auto; padding: 20px; }
        .upload-area { border: 2px dashed #ccc; padding: 40px; text-align: center; margin: 20px 0; }
        .file-list { margin: 20px 0; }
        .file-item { padding: 10px; border: 1px solid #ddd; margin: 5px 0; }
        button { background: #007bff; color: white; padding: 10px 20px; border: none; cursor: pointer; }
        button:hover { background: #0056b3; }
        .progress { display: none; margin: 20px 0; }
        .result { margin: 20px 0; padding: 20px; background: #f8f9fa; border-radius: 5px; }
    </style>
</head>
<body>
    <h1>📊 Proposal Analyzer Pro</h1>
    <p>Sistema de Análise Comparativa de Propostas com IA</p>
    
    <div class="upload-area" onclick="document.getElementById('fileInput').click()">
        <p>📁 Clique aqui ou arraste arquivos para fazer upload</p>
        <p><small>Aceita: PDF, DOCX, XLSX (máx. 50MB)</small></p>
        <input type="file" id="fileInput" multiple accept=".pdf,.docx,.xlsx" style="display: none;">
    </div>
    
    <div class="file-list" id="fileList"></div>
    
    <button onclick="analyzeProposals()" id="analyzeBtn" disabled>🔍 Analisar Propostas</button>
    
    <div class="progress" id="progress">
        <p>⏳ Processando documentos...</p>
    </div>
    
    <div class="result" id="result" style="display: none;"></div>

    <script>
        let uploadedFiles = [];
        
        document.getElementById('fileInput').addEventListener('change', function(e) {
            handleFiles(e.target.files);
        });
        
        function handleFiles(files) {
            for (let file of files) {
                uploadedFiles.push(file);
                addFileToList(file);
            }
            updateAnalyzeButton();
        }
        
        function addFileToList(file) {
            const fileList = document.getElementById('fileList');
            const fileItem = document.createElement('div');
            fileItem.className = 'file-item';
            fileItem.innerHTML = `
                <span>📄 ${file.name} (${(file.size/1024/1024).toFixed(2)} MB)</span>
                <button onclick="removeFile('${file.name}')" style="float: right; background: #dc3545;">❌</button>
            `;
            fileList.appendChild(fileItem);
        }
        
        function removeFile(fileName) {
            uploadedFiles = uploadedFiles.filter(f => f.name !== fileName);
            updateFileList();
            updateAnalyzeButton();
        }
        
        function updateFileList() {
            const fileList = document.getElementById('fileList');
            fileList.innerHTML = '';
            uploadedFiles.forEach(addFileToList);
        }
        
        function updateAnalyzeButton() {
            document.getElementById('analyzeBtn').disabled = uploadedFiles.length === 0;
        }
        
        async function analyzeProposals() {
            if (uploadedFiles.length === 0) return;
            
            const formData = new FormData();
            uploadedFiles.forEach(file => {
                formData.append('files', file);
            });
            
            document.getElementById('progress').style.display = 'block';
            document.getElementById('result').style.display = 'none';
            
            try {
                const response = await fetch('/analyze', {
                    method: 'POST',
                    body: formData
                });
                
                const result = await response.json();
                
                if (result.success) {
                    document.getElementById('result').innerHTML = `
                        <h3>✅ Análise Concluída!</h3>
                        <p><strong>Relatório ID:</strong> ${result.report_id}</p>
                        <p><strong>Empresas Analisadas:</strong> ${result.companies_count}</p>
                        <div style="margin: 20px 0;">
                            <a href="/download/${result.report_id}/pdf" target="_blank">
                                <button>📄 Download PDF</button>
                            </a>
                            <a href="/download/${result.report_id}/markdown" target="_blank">
                                <button>📝 Download Markdown</button>
                            </a>
                        </div>
                    `;
                    document.getElementById('result').style.display = 'block';
                } else {
                    throw new Error(result.error || 'Erro desconhecido');
                }
            } catch (error) {
                document.getElementById('result').innerHTML = `
                    <h3>❌ Erro na Análise</h3>
                    <p>${error.message}</p>
                `;
                document.getElementById('result').style.display = 'block';
            }
            
            document.getElementById('progress').style.display = 'none';
        }
    </script>
</body>
</html>
    """)

@app.route('/analyze', methods=['POST'])
def analyze():
    try:
        files = request.files.getlist('files')
        if not files:
            return jsonify({'success': False, 'error': 'Nenhum arquivo enviado'})
        
        # Salvar arquivos
        uploaded_files = []
        for file in files:
            if file.filename:
                filename = secure_filename(file.filename)
                filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
                file.save(filepath)
                uploaded_files.append(filepath)
        
        # Processar arquivos
        proposals_data = {}
        tr_data = {'objeto': 'Sistema de Gestão Empresarial'}
        
        # Agrupar arquivos por empresa
        companies = {}
        for filepath in uploaded_files:
            filename = os.path.basename(filepath).lower()
            
            # Identificar empresa pelo nome do arquivo
            if 'techsolutions' in filename:
                company = 'TechSolutions Ltda.'
            elif 'innovasoft' in filename:
                company = 'InnovaSoft S.A.'
            else:
                # Tentar extrair nome da empresa do arquivo
                company = filename.split('_')[0].title()
            
            if company not in companies:
                companies[company] = {'technical': {}, 'commercial': {}}
            
            # Processar arquivo
            if filepath.endswith('.xlsx'):
                companies[company]['commercial'] = extract_data_from_excel(filepath)
            elif filepath.endswith('.pdf') or filepath.endswith('.docx'):
                if filepath.endswith('.pdf'):
                    text = extract_text_from_pdf(filepath)
                else:
                    text = extract_text_from_docx(filepath)
                companies[company]['technical'] = extract_technical_data_from_pdf(text)
        
        # Gerar relatório
        report_content = generate_comparative_report(tr_data, companies)
        
        # Salvar relatório
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        report_id = f"analise_comparativa_{timestamp}"
        
        # Salvar Markdown
        md_path = os.path.join(app.config['UPLOAD_FOLDER'], f"{report_id}.md")
        with open(md_path, 'w', encoding='utf-8') as f:
            f.write(report_content)
        
        # Gerar PDF
        pdf_path = os.path.join(app.config['UPLOAD_FOLDER'], f"{report_id}.pdf")
        create_pdf_from_markdown(report_content, pdf_path)
        
        # Limpar arquivos temporários
        for filepath in uploaded_files:
            try:
                os.remove(filepath)
            except:
                pass
        
        return jsonify({
            'success': True,
            'report_id': report_id,
            'companies_count': len(companies)
        })
        
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)})

@app.route('/download/<report_id>/<format>')
def download_report(report_id, format):
    try:
        if format == 'markdown':
            file_path = os.path.join(app.config['UPLOAD_FOLDER'], f"{report_id}.md")
            return send_file(file_path, as_attachment=True, download_name=f"{report_id}.md")
        elif format == 'pdf':
            file_path = os.path.join(app.config['UPLOAD_FOLDER'], f"{report_id}.pdf")
            return send_file(file_path, as_attachment=True, download_name=f"{report_id}.pdf")
        else:
            return jsonify({'error': 'Formato não suportado.'}), 400
    except Exception as e:
        return jsonify({'error': str(e)}), 500

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)

