import os
import tempfile
import logging
from datetime import datetime
from flask import Flask, request, jsonify, send_file, render_template_string
from flask_cors import CORS
import re
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_JUSTIFY
import PyPDF2
from openpyxl import Workbook
import json

# Configuração do Azure Document Intelligence
from azure.ai.documentintelligence import DocumentIntelligenceClient
from azure.core.credentials import AzureKeyCredential

# Configurações Azure
AZURE_ENDPOINT = "https://proposal-analyzer-eastus.cognitiveservices.azure.com/"
AZURE_KEY = "2WSbc2H8NbocAvetZtpuqx6fhkHULpBgLyTQg2tD8BKG2E74Pm1wJQQJ99BGACYeBjFXJ3w3AAALACOGu7AE"

app = Flask(__name__)
CORS(app)

# Configuração de logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# Inicializar cliente Azure
try:
    azure_client = DocumentIntelligenceClient(
        endpoint=AZURE_ENDPOINT,
        credential=AzureKeyCredential(AZURE_KEY)
    )
    logger.info("Azure Document Intelligence inicializado com sucesso")
except Exception as e:
    logger.error(f"Erro ao inicializar Azure: {e}")
    azure_client = None

class ProposalAnalyzer:
    def __init__(self):
        self.azure_client = azure_client
        
    def extract_with_azure(self, file_path):
        """Extrai dados usando Azure Document Intelligence"""
        try:
            logger.info(f"Iniciando extração Azure para: {file_path}")
            
            with open(file_path, "rb") as f:
                poller = self.azure_client.begin_analyze_document(
                    "prebuilt-layout", 
                    analyze_request=f,
                    content_type="application/pdf"
                )
                result = poller.result()
            
            # Extrair texto completo
            full_text = ""
            if result.content:
                full_text = result.content
            
            confidence = getattr(result, 'confidence', 0.7) * 100
            logger.info(f"Extração Azure concluída: {confidence:.1f}% confiança")
            
            return full_text, confidence
            
        except Exception as e:
            logger.error(f"Erro na extração Azure: {e}")
            return None, 0
    
    def extract_with_pypdf2(self, file_path):
        """Extrai dados usando PyPDF2 como fallback"""
        try:
            logger.warning("Usando extração de fallback")
            text = ""
            with open(file_path, 'rb') as file:
                pdf_reader = PyPDF2.PdfReader(file)
                for page in pdf_reader.pages:
                    text += page.extract_text() + "\n"
            return text, 50.0  # Confiança menor para fallback
        except Exception as e:
            logger.error(f"Erro na extração PyPDF2: {e}")
            return "", 0
    
    def extract_company_data(self, text):
        """Extrai dados específicos da empresa da proposta"""
        data = {
            'nome_empresa': '',
            'cnpj': '',
            'endereco': '',
            'telefone': '',
            'email': '',
            'objeto': '',
            'prazo_total': 0,
            'prazo_mobilizacao': 0,
            'prazo_execucao': 0,
            'equipe_total': 0,
            'engenheiros': [],
            'metodologia': '',
            'garantia_civil': 0,
            'garantia_outros': 0,
            'equipamentos': [],
            'experiencia': [],
            'valor_total': 0.0,
            'bdi': 0.0
        }
        
        # Extrair nome da empresa (primeira linha em maiúscula)
        nome_match = re.search(r'^([A-ZÁÊÇÕ\s&-]+(?:LTDA|S\.A\.|EIRELI|ME|EPP)?)', text, re.MULTILINE)
        if nome_match:
            data['nome_empresa'] = nome_match.group(1).strip()
        
        # Extrair CNPJ
        cnpj_match = re.search(r'CNPJ[:\s]*(\d{2}\.?\d{3}\.?\d{3}/?\d{4}-?\d{2})', text, re.IGNORECASE)
        if cnpj_match:
            data['cnpj'] = cnpj_match.group(1)
        
        # Extrair endereço
        endereco_match = re.search(r'(?:Avenida|Rua|Av\.|R\.)\s+([^,\n]+(?:,\s*[^,\n]+)*)', text, re.IGNORECASE)
        if endereco_match:
            data['endereco'] = endereco_match.group(0).strip()
        
        # Extrair telefone
        telefone_match = re.search(r'(?:Fone|Tel|Telefone)[:\s]*\(?(\d{2})\)?\s*\d{4,5}[-\.\s]?\d{4}', text, re.IGNORECASE)
        if telefone_match:
            data['telefone'] = telefone_match.group(0).split(':')[-1].strip()
        
        # Extrair email
        email_match = re.search(r'Email[:\s]*([a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,})', text, re.IGNORECASE)
        if email_match:
            data['email'] = email_match.group(1)
        
        # Extrair objeto/serviço
        objeto_patterns = [
            r'SERVIÇO[:\s]*([^\n]+(?:\n[^\n]+)*?)(?=\n\n|\nPROPOSTA|\nAPRESENTAÇÃO)',
            r'OBRA[:\s]*([^\n]+(?:\n[^\n]+)*?)(?=\n\n|\nLOCAL)',
            r'OBJETO[:\s]*([^\n]+(?:\n[^\n]+)*?)(?=\n\n|\nESCOPO)'
        ]
        for pattern in objeto_patterns:
            match = re.search(pattern, text, re.IGNORECASE | re.DOTALL)
            if match:
                data['objeto'] = match.group(1).strip()
                break
        
        # Extrair prazos
        prazo_patterns = [
            r'prazo[^:]*?(\d+)\s*dias?\s*(?:para\s*)?(?:execução|total)', 
            r'execução[^:]*?(\d+)\s*dias?',
            r'(\d+)\s*dias?\s*(?:para\s*)?(?:execução|conclusão)'
        ]
        for pattern in prazo_patterns:
            match = re.search(pattern, text, re.IGNORECASE)
            if match:
                data['prazo_execucao'] = int(match.group(1))
                break
        
        # Extrair prazo de mobilização
        mob_match = re.search(r'mobilização[^:]*?(\d+)\s*dias?', text, re.IGNORECASE)
        if mob_match:
            data['prazo_mobilizacao'] = int(mob_match.group(1))
        
        # Calcular prazo total
        data['prazo_total'] = data['prazo_execucao'] + data['prazo_mobilizacao']
        
        # Extrair equipe técnica
        engenheiros = re.findall(r'([A-Z][a-z]+\s+[A-Z][a-z]+(?:\s+[A-Z][a-z]+)*)\s*[-–]\s*Engenheiro\s+(\w+)', text)
        data['engenheiros'] = [{'nome': nome, 'especialidade': esp} for nome, esp in engenheiros]
        
        # Contar equipe total (buscar por números de pessoas)
        equipe_patterns = [
            r'(\d+)\s*(?:Pedreiros?|pedreiros?)',
            r'(\d+)\s*(?:Auxiliares?|auxiliares?)',
            r'(\d+)\s*(?:Eletricistas?|eletricistas?)',
            r'(\d+)\s*(?:Operadores?|operadores?)',
            r'(\d+)\s*(?:Técnicos?|técnicos?)'
        ]
        total_equipe = len(data['engenheiros'])  # Começar com engenheiros
        for pattern in equipe_patterns:
            matches = re.findall(pattern, text, re.IGNORECASE)
            for match in matches:
                total_equipe += int(match)
        data['equipe_total'] = total_equipe
        
        # Extrair metodologia
        metodologia_match = re.search(r'(?:PLANO DE EXECUÇÃO|METODOLOGIA|EXECUÇÃO)[:\s]*([^§]+?)(?=\n\d+\.\d+|\nGARANTIAS|\nPRAZO)', text, re.IGNORECASE | re.DOTALL)
        if metodologia_match:
            data['metodologia'] = metodologia_match.group(1).strip()[:500]  # Limitar tamanho
        
        # Extrair garantias
        garantia_civil = re.search(r'(\d+)\s*(?:anos?|meses?)\s*(?:para\s*)?(?:obras?\s*)?civis?', text, re.IGNORECASE)
        if garantia_civil:
            data['garantia_civil'] = int(garantia_civil.group(1))
        
        garantia_outros = re.search(r'(\d+)\s*(?:anos?|meses?)\s*(?:para\s*)?(?:demais|outros|serviços)', text, re.IGNORECASE)
        if garantia_outros:
            data['garantia_outros'] = int(garantia_outros.group(1))
        
        # Extrair equipamentos
        equipamentos = re.findall(r'(?:Locação|Fornecimento)\s+de\s+([^,\n]+)', text, re.IGNORECASE)
        data['equipamentos'] = equipamentos[:10]  # Limitar a 10 itens
        
        # Extrair experiência/referências
        experiencia = re.findall(r'Cliente[:\s]*([^\n]+)', text, re.IGNORECASE)
        data['experiencia'] = experiencia[:5]  # Limitar a 5 referências
        
        return data
    
    def calculate_technical_score(self, data):
        """Calcula score técnico baseado nos dados extraídos"""
        score = 0
        max_score = 100
        
        # Prazo (25 pontos) - Quanto menor, melhor
        if data['prazo_total'] > 0:
            if data['prazo_total'] <= 60:
                score += 25
            elif data['prazo_total'] <= 90:
                score += 20
            elif data['prazo_total'] <= 120:
                score += 15
            else:
                score += 10
        
        # Equipe técnica (25 pontos)
        if data['equipe_total'] > 0:
            if data['equipe_total'] >= 15:
                score += 25
            elif data['equipe_total'] >= 10:
                score += 20
            elif data['equipe_total'] >= 5:
                score += 15
            else:
                score += 10
        
        # Engenheiros (20 pontos)
        num_engenheiros = len(data['engenheiros'])
        if num_engenheiros >= 3:
            score += 20
        elif num_engenheiros >= 2:
            score += 15
        elif num_engenheiros >= 1:
            score += 10
        
        # Metodologia (15 pontos)
        if data['metodologia']:
            if len(data['metodologia']) > 200:
                score += 15
            elif len(data['metodologia']) > 100:
                score += 10
            else:
                score += 5
        
        # Experiência (10 pontos)
        num_experiencia = len(data['experiencia'])
        if num_experiencia >= 5:
            score += 10
        elif num_experiencia >= 3:
            score += 7
        elif num_experiencia >= 1:
            score += 5
        
        # Garantias (5 pontos)
        if data['garantia_civil'] >= 5:
            score += 5
        elif data['garantia_civil'] >= 3:
            score += 3
        elif data['garantia_civil'] >= 1:
            score += 2
        
        return min(score, max_score)
    
    def analyze_proposals(self, files):
        """Analisa múltiplas propostas"""
        results = []
        
        for file_info in files:
            file_path = file_info['path']
            logger.info(f"Analisando: {file_path}")
            
            # Tentar extração com Azure primeiro
            text, confidence = None, 0
            if self.azure_client:
                text, confidence = self.extract_with_azure(file_path)
            
            # Fallback para PyPDF2 se Azure falhar
            if not text:
                text, confidence = self.extract_with_pypdf2(file_path)
            
            if text:
                # Extrair dados da empresa
                company_data = self.extract_company_data(text)
                
                # Calcular score técnico
                technical_score = self.calculate_technical_score(company_data)
                
                # Adicionar informações extras
                company_data['confidence'] = confidence
                company_data['technical_score'] = technical_score
                company_data['file_name'] = file_info['original_name']
                
                results.append(company_data)
            else:
                logger.error(f"Falha na extração para: {file_path}")
        
        return results
    
    def generate_technical_report(self, proposals, output_path):
        """Gera relatório técnico especializado"""
        doc = SimpleDocTemplate(output_path, pagesize=A4, topMargin=0.5*inch)
        styles = getSampleStyleSheet()
        story = []
        
        # Estilos customizados
        title_style = ParagraphStyle(
            'CustomTitle',
            parent=styles['Title'],
            fontSize=18,
            spaceAfter=30,
            textColor=colors.darkblue,
            alignment=TA_CENTER
        )
        
        heading_style = ParagraphStyle(
            'CustomHeading',
            parent=styles['Heading2'],
            fontSize=14,
            spaceAfter=12,
            textColor=colors.darkblue,
            leftIndent=0
        )
        
        # Título
        story.append(Paragraph("ANÁLISE E EQUALIZAÇÃO TÉCNICA DE PROPOSTAS", title_style))
        story.append(Paragraph("Avaliação Técnica Especializada", styles['Normal']))
        story.append(Paragraph(f"Data de Geração: {datetime.now().strftime('%d/%m/%Y às %H:%M')}", styles['Normal']))
        story.append(Spacer(1, 20))
        
        # Seção 1: Ranking Técnico
        story.append(Paragraph("SEÇÃO 1: RANKING TÉCNICO GERAL", heading_style))
        
        # Ordenar por score técnico
        sorted_proposals = sorted(proposals, key=lambda x: x['technical_score'], reverse=True)
        
        # Tabela de ranking
        ranking_data = [['Posição', 'Empresa', 'Score Técnico', 'Prazo (dias)', 'Equipe']]
        for i, prop in enumerate(sorted_proposals, 1):
            ranking_data.append([
                str(i),
                prop['nome_empresa'][:30] if prop['nome_empresa'] else 'N/I',
                f"{prop['technical_score']:.1f}%",
                str(prop['prazo_total']) if prop['prazo_total'] > 0 else 'N/I',
                str(prop['equipe_total']) if prop['equipe_total'] > 0 else 'N/I'
            ])
        
        ranking_table = Table(ranking_data, colWidths=[0.8*inch, 2.5*inch, 1.2*inch, 1.2*inch, 1*inch])
        ranking_table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.darkblue),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, 0), 10),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
            ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
            ('GRID', (0, 0), (-1, -1), 1, colors.black)
        ]))
        story.append(ranking_table)
        story.append(Spacer(1, 20))
        
        # Seção 2: Análise Detalhada por Empresa
        story.append(Paragraph("SEÇÃO 2: ANÁLISE TÉCNICA DETALHADA", heading_style))
        
        for i, prop in enumerate(sorted_proposals):
            if i > 0:
                story.append(PageBreak())
            
            # Nome da empresa
            story.append(Paragraph(f"{prop['nome_empresa']} - Score: {prop['technical_score']:.1f}%", 
                                 ParagraphStyle('CompanyTitle', parent=styles['Heading3'], 
                                              textColor=colors.darkgreen, fontSize=12)))
            story.append(Spacer(1, 10))
            
            # Dados básicos
            basic_data = [
                ['Informação', 'Detalhes'],
                ['CNPJ', prop['cnpj'] if prop['cnpj'] else 'Não informado'],
                ['Endereço', prop['endereco'][:50] + '...' if len(prop['endereco']) > 50 else prop['endereco'] if prop['endereco'] else 'Não informado'],
                ['Telefone', prop['telefone'] if prop['telefone'] else 'Não informado'],
                ['Email', prop['email'] if prop['email'] else 'Não informado']
            ]
            
            basic_table = Table(basic_data, colWidths=[1.5*inch, 4*inch])
            basic_table.setStyle(TableStyle([
                ('BACKGROUND', (0, 0), (-1, 0), colors.lightblue),
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.black),
                ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
                ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                ('FONTSIZE', (0, 0), (-1, -1), 9),
                ('GRID', (0, 0), (-1, -1), 1, colors.black),
                ('VALIGN', (0, 0), (-1, -1), 'TOP')
            ]))
            story.append(basic_table)
            story.append(Spacer(1, 15))
            
            # Objeto do serviço
            if prop['objeto']:
                story.append(Paragraph("Objeto do Serviço:", styles['Heading4']))
                story.append(Paragraph(prop['objeto'], styles['Normal']))
                story.append(Spacer(1, 10))
            
            # Cronograma
            story.append(Paragraph("Cronograma:", styles['Heading4']))
            cronograma_text = f"• Prazo de Execução: {prop['prazo_execucao']} dias<br/>"
            if prop['prazo_mobilizacao'] > 0:
                cronograma_text += f"• Prazo de Mobilização: {prop['prazo_mobilizacao']} dias<br/>"
            cronograma_text += f"• Prazo Total: {prop['prazo_total']} dias"
            story.append(Paragraph(cronograma_text, styles['Normal']))
            story.append(Spacer(1, 10))
            
            # Equipe técnica
            story.append(Paragraph("Equipe Técnica:", styles['Heading4']))
            equipe_text = f"• Total da Equipe: {prop['equipe_total']} pessoas<br/>"
            if prop['engenheiros']:
                equipe_text += "• Engenheiros:<br/>"
                for eng in prop['engenheiros']:
                    equipe_text += f"  - {eng['nome']} (Engenheiro {eng['especialidade']})<br/>"
            story.append(Paragraph(equipe_text, styles['Normal']))
            story.append(Spacer(1, 10))
            
            # Metodologia
            if prop['metodologia']:
                story.append(Paragraph("Metodologia de Execução:", styles['Heading4']))
                metodologia_resumo = prop['metodologia'][:300] + "..." if len(prop['metodologia']) > 300 else prop['metodologia']
                story.append(Paragraph(metodologia_resumo, styles['Normal']))
                story.append(Spacer(1, 10))
            
            # Garantias
            if prop['garantia_civil'] > 0 or prop['garantia_outros'] > 0:
                story.append(Paragraph("Garantias:", styles['Heading4']))
                garantias_text = ""
                if prop['garantia_civil'] > 0:
                    garantias_text += f"• Obras Civis: {prop['garantia_civil']} anos<br/>"
                if prop['garantia_outros'] > 0:
                    garantias_text += f"• Demais Serviços: {prop['garantia_outros']} anos<br/>"
                story.append(Paragraph(garantias_text, styles['Normal']))
                story.append(Spacer(1, 10))
            
            # Experiência
            if prop['experiencia']:
                story.append(Paragraph("Experiência/Referências:", styles['Heading4']))
                exp_text = ""
                for exp in prop['experiencia'][:3]:  # Mostrar apenas 3 principais
                    exp_text += f"• {exp}<br/>"
                story.append(Paragraph(exp_text, styles['Normal']))
                story.append(Spacer(1, 10))
        
        # Seção 3: Recomendações
        story.append(PageBreak())
        story.append(Paragraph("SEÇÃO 3: RECOMENDAÇÕES TÉCNICAS", heading_style))
        
        if sorted_proposals:
            melhor_proposta = sorted_proposals[0]
            story.append(Paragraph("Recomendação Técnica Principal:", styles['Heading4']))
            story.append(Paragraph(f"Com base na análise técnica detalhada, recomenda-se a empresa {melhor_proposta['nome_empresa']} que obteve o melhor score técnico ({melhor_proposta['technical_score']:.1f}%).", styles['Normal']))
            story.append(Spacer(1, 10))
            
            story.append(Paragraph("Justificativa Técnica:", styles['Heading4']))
            justificativas = []
            if melhor_proposta['prazo_total'] > 0:
                justificativas.append(f"• Cronograma viável: {melhor_proposta['prazo_total']} dias")
            if melhor_proposta['equipe_total'] > 0:
                justificativas.append(f"• Equipe robusta: {melhor_proposta['equipe_total']} pessoas")
            if melhor_proposta['engenheiros']:
                justificativas.append(f"• Corpo técnico qualificado: {len(melhor_proposta['engenheiros'])} engenheiros")
            
            for just in justificativas:
                story.append(Paragraph(just, styles['Normal']))
        
        # Gerar PDF
        doc.build(story)
        return output_path
    
    def generate_commercial_report(self, proposals, output_path):
        """Gera relatório comercial especializado"""
        doc = SimpleDocTemplate(output_path, pagesize=A4, topMargin=0.5*inch)
        styles = getSampleStyleSheet()
        story = []
        
        # Estilos customizados
        title_style = ParagraphStyle(
            'CustomTitle',
            parent=styles['Title'],
            fontSize=18,
            spaceAfter=30,
            textColor=colors.darkgreen,
            alignment=TA_CENTER
        )
        
        heading_style = ParagraphStyle(
            'CustomHeading',
            parent=styles['Heading2'],
            fontSize=14,
            spaceAfter=12,
            textColor=colors.darkgreen,
            leftIndent=0
        )
        
        # Título
        story.append(Paragraph("ANÁLISE E EQUALIZAÇÃO COMERCIAL DE PROPOSTAS", title_style))
        story.append(Paragraph("Avaliação Comercial Especializada", styles['Normal']))
        story.append(Paragraph(f"Data de Geração: {datetime.now().strftime('%d/%m/%Y às %H:%M')}", styles['Normal']))
        story.append(Spacer(1, 20))
        
        # Seção 1: Ranking Comercial
        story.append(Paragraph("SEÇÃO 1: RANKING COMERCIAL", heading_style))
        
        # Para relatório comercial, ordenar por menor preço (quando disponível)
        # Por enquanto, usar score técnico como proxy
        sorted_proposals = sorted(proposals, key=lambda x: x['technical_score'], reverse=True)
        
        # Tabela de ranking comercial
        ranking_data = [['Posição', 'Empresa', 'Valor Proposto', 'Prazo', 'Garantias']]
        for i, prop in enumerate(sorted_proposals, 1):
            valor_str = f"R$ {prop['valor_total']:,.2f}" if prop['valor_total'] > 0 else 'A definir'
            garantia_str = f"{prop['garantia_civil']}a/{prop['garantia_outros']}a" if prop['garantia_civil'] > 0 else 'N/I'
            
            ranking_data.append([
                str(i),
                prop['nome_empresa'][:25] if prop['nome_empresa'] else 'N/I',
                valor_str,
                f"{prop['prazo_total']} dias" if prop['prazo_total'] > 0 else 'N/I',
                garantia_str
            ])
        
        ranking_table = Table(ranking_data, colWidths=[0.8*inch, 2.2*inch, 1.5*inch, 1.2*inch, 1*inch])
        ranking_table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.darkgreen),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, 0), 10),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
            ('BACKGROUND', (0, 1), (-1, -1), colors.lightgreen),
            ('GRID', (0, 0), (-1, -1), 1, colors.black)
        ]))
        story.append(ranking_table)
        story.append(Spacer(1, 20))
        
        # Seção 2: Análise Comercial Detalhada
        story.append(Paragraph("SEÇÃO 2: ANÁLISE COMERCIAL DETALHADA", heading_style))
        
        for i, prop in enumerate(sorted_proposals):
            if i > 0:
                story.append(PageBreak())
            
            # Nome da empresa
            story.append(Paragraph(f"{prop['nome_empresa']}", 
                                 ParagraphStyle('CompanyTitle', parent=styles['Heading3'], 
                                              textColor=colors.darkgreen, fontSize=12)))
            story.append(Spacer(1, 10))
            
            # Dados comerciais
            commercial_data = [
                ['Aspecto Comercial', 'Detalhes'],
                ['Empresa', prop['nome_empresa'] if prop['nome_empresa'] else 'Não informado'],
                ['CNPJ', prop['cnpj'] if prop['cnpj'] else 'Não informado'],
                ['Endereço', prop['endereco'][:60] + '...' if len(prop['endereco']) > 60 else prop['endereco'] if prop['endereco'] else 'Não informado'],
                ['Contato', f"{prop['telefone']} / {prop['email']}" if prop['telefone'] or prop['email'] else 'Não informado'],
                ['Prazo Total', f"{prop['prazo_total']} dias" if prop['prazo_total'] > 0 else 'Não informado'],
                ['Garantia Civil', f"{prop['garantia_civil']} anos" if prop['garantia_civil'] > 0 else 'Não informado'],
                ['Garantia Serviços', f"{prop['garantia_outros']} anos" if prop['garantia_outros'] > 0 else 'Não informado']
            ]
            
            commercial_table = Table(commercial_data, colWidths=[2*inch, 3.5*inch])
            commercial_table.setStyle(TableStyle([
                ('BACKGROUND', (0, 0), (-1, 0), colors.lightgreen),
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.black),
                ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
                ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                ('FONTSIZE', (0, 0), (-1, -1), 9),
                ('GRID', (0, 0), (-1, -1), 1, colors.black),
                ('VALIGN', (0, 0), (-1, -1), 'TOP')
            ]))
            story.append(commercial_table)
            story.append(Spacer(1, 15))
            
            # Condições comerciais
            story.append(Paragraph("Condições Comerciais:", styles['Heading4']))
            condicoes_text = f"• Prazo de Execução: {prop['prazo_execucao']} dias<br/>"
            if prop['prazo_mobilizacao'] > 0:
                condicoes_text += f"• Prazo de Mobilização: {prop['prazo_mobilizacao']} dias<br/>"
            if prop['garantia_civil'] > 0:
                condicoes_text += f"• Garantia para Obras Civis: {prop['garantia_civil']} anos<br/>"
            if prop['garantia_outros'] > 0:
                condicoes_text += f"• Garantia para Demais Serviços: {prop['garantia_outros']} anos<br/>"
            story.append(Paragraph(condicoes_text, styles['Normal']))
            story.append(Spacer(1, 10))
            
            # Capacidade técnica (relevante para avaliação comercial)
            story.append(Paragraph("Capacidade Técnica:", styles['Heading4']))
            capacidade_text = f"• Equipe Proposta: {prop['equipe_total']} pessoas<br/>"
            if prop['engenheiros']:
                capacidade_text += f"• Engenheiros: {len(prop['engenheiros'])} profissionais<br/>"
            if prop['experiencia']:
                capacidade_text += f"• Referências: {len(prop['experiencia'])} clientes<br/>"
            story.append(Paragraph(capacidade_text, styles['Normal']))
            story.append(Spacer(1, 10))
        
        # Seção 3: Recomendações Comerciais
        story.append(PageBreak())
        story.append(Paragraph("SEÇÃO 3: RECOMENDAÇÕES COMERCIAIS", heading_style))
        
        if sorted_proposals:
            melhor_proposta = sorted_proposals[0]
            story.append(Paragraph("Recomendação Comercial Principal:", styles['Heading4']))
            story.append(Paragraph(f"Com base na análise comercial, recomenda-se a empresa {melhor_proposta['nome_empresa']} considerando o conjunto de fatores comerciais e técnicos.", styles['Normal']))
            story.append(Spacer(1, 10))
            
            story.append(Paragraph("Justificativa Comercial:", styles['Heading4']))
            justificativas = []
            if melhor_proposta['prazo_total'] > 0:
                justificativas.append(f"• Prazo competitivo: {melhor_proposta['prazo_total']} dias")
            if melhor_proposta['garantia_civil'] > 0:
                justificativas.append(f"• Garantia adequada: {melhor_proposta['garantia_civil']} anos para obras civis")
            if melhor_proposta['experiencia']:
                justificativas.append(f"• Experiência comprovada: {len(melhor_proposta['experiencia'])} referências")
            
            for just in justificativas:
                story.append(Paragraph(just, styles['Normal']))
            
            story.append(Spacer(1, 15))
            story.append(Paragraph("Próximos Passos Comerciais:", styles['Heading4']))
            proximos_passos = [
                "• Solicitar detalhamento da composição de custos",
                "• Validar condições de pagamento propostas", 
                "• Confirmar prazos de execução e garantias",
                "• Verificar documentação fiscal e jurídica",
                "• Negociar condições finais do contrato"
            ]
            
            for passo in proximos_passos:
                story.append(Paragraph(passo, styles['Normal']))
        
        # Gerar PDF
        doc.build(story)
        return output_path

# Instanciar analisador
analyzer = ProposalAnalyzer()

@app.route('/')
def index():
    return render_template_string(HTML_TEMPLATE)

@app.route('/upload', methods=['POST'])
def upload_files():
    try:
        files = request.files.getlist('files')
        report_type = request.form.get('report_type', 'technical')
        
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
        
        # Gerar relatório baseado no tipo selecionado
        if report_type == 'technical':
            report_filename = f"analise_tecnica_{timestamp}.pdf"
            report_path = os.path.join(upload_dir, report_filename)
            analyzer.generate_technical_report(proposals, report_path)
            logger.info(f"Relatório técnico gerado: {report_path}")
        else:  # commercial
            report_filename = f"analise_comercial_{timestamp}.pdf"
            report_path = os.path.join(upload_dir, report_filename)
            analyzer.generate_commercial_report(proposals, report_path)
            logger.info(f"Relatório comercial gerado: {report_path}")
        
        return jsonify({
            'success': True,
            'report_url': f'/download/{report_filename}',
            'report_type': 'Técnico' if report_type == 'technical' else 'Comercial',
            'proposals_count': len(proposals)
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

# Template HTML atualizado para 2 relatórios
HTML_TEMPLATE = '''
<!DOCTYPE html>
<html lang="pt-BR">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Arias Analyzer Pro - Análise de Propostas</title>
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
        
        .report-type {
            margin: 20px 0;
            text-align: left;
        }
        
        .report-type label {
            font-weight: bold;
            color: #333;
            margin-bottom: 10px;
            display: block;
        }
        
        .radio-group {
            display: flex;
            gap: 20px;
            justify-content: center;
            flex-wrap: wrap;
        }
        
        .radio-option {
            display: flex;
            align-items: center;
            gap: 8px;
            padding: 12px 20px;
            border: 2px solid #ddd;
            border-radius: 10px;
            cursor: pointer;
            transition: all 0.3s ease;
            background: white;
        }
        
        .radio-option:hover {
            border-color: #667eea;
            background: #f8f9ff;
        }
        
        .radio-option input[type="radio"]:checked + .radio-label {
            color: #667eea;
            font-weight: bold;
        }
        
        .radio-option:has(input[type="radio"]:checked) {
            border-color: #667eea;
            background: #f8f9ff;
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
        
        @media (max-width: 768px) {
            .container {
                padding: 20px;
                margin: 10px;
            }
            
            .logo {
                font-size: 2em;
            }
            
            .radio-group {
                flex-direction: column;
                align-items: center;
            }
            
            .radio-option {
                width: 100%;
                max-width: 250px;
                justify-content: center;
            }
        }
    </style>
</head>
<body>
    <div class="container">
        <div class="logo">📊 Arias Analyzer Pro</div>
        <div class="subtitle">Análise Inteligente de Propostas com IA</div>
        
        <div class="upload-area" onclick="document.getElementById('fileInput').click()">
            <div class="upload-icon">📁</div>
            <div class="upload-text">Clique aqui ou arraste os arquivos</div>
            <div class="upload-hint">Aceita PDF e Excel (mínimo 2 arquivos)</div>
        </div>
        
        <input type="file" id="fileInput" class="file-input" multiple accept=".pdf,.xlsx,.xls">
        
        <div class="file-list" id="fileList"></div>
        
        <div class="report-type">
            <label>Tipo de Relatório:</label>
            <div class="radio-group">
                <div class="radio-option">
                    <input type="radio" id="technical" name="reportType" value="technical" checked>
                    <label for="technical" class="radio-label">📋 Análise Técnica</label>
                </div>
                <div class="radio-option">
                    <input type="radio" id="commercial" name="reportType" value="commercial">
                    <label for="commercial" class="radio-label">💰 Análise Comercial</label>
                </div>
            </div>
        </div>
        
        <button class="analyze-btn" id="analyzeBtn" onclick="analyzeFiles()" disabled>
            Analisar Propostas
        </button>
        
        <div class="progress-container" id="progressContainer">
            <div class="progress-bar">
                <div class="progress-fill" id="progressFill"></div>
            </div>
            <div class="progress-text" id="progressText">Processando...</div>
        </div>
        
        <div class="result-container" id="resultContainer">
            <div class="success-message" id="successMessage"></div>
            <a href="#" class="download-btn" id="downloadBtn">📥 Baixar Relatório</a>
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
                const validTypes = ['.pdf', '.xlsx', '.xls'];
                const extension = '.' + file.name.split('.').pop().toLowerCase();
                return validTypes.includes(extension);
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
                showError('É necessário selecionar pelo menos 2 arquivos.');
                return;
            }
            
            const reportType = document.querySelector('input[name="reportType"]:checked').value;
            
            // Mostrar progresso
            document.getElementById('progressContainer').style.display = 'block';
            document.getElementById('resultContainer').style.display = 'none';
            document.getElementById('errorMessage').style.display = 'none';
            analyzeBtn.disabled = true;
            
            // Simular progresso
            let progress = 0;
            const progressInterval = setInterval(() => {
                progress += Math.random() * 15;
                if (progress > 90) progress = 90;
                updateProgress(progress, 'Processando arquivos...');
            }, 500);
            
            try {
                const formData = new FormData();
                selectedFiles.forEach(file => {
                    formData.append('files', file);
                });
                formData.append('report_type', reportType);
                
                const response = await fetch('/upload', {
                    method: 'POST',
                    body: formData
                });
                
                clearInterval(progressInterval);
                updateProgress(100, 'Concluído!');
                
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
                ✅ Relatório ${result.report_type} gerado com sucesso!<br>
                📊 ${result.proposals_count} propostas analisadas
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

