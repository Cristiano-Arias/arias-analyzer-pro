from docx import Document
from docx.shared import Pt
from docx.oxml.ns import qn
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT

def gerar_relatorio_word(nome_arquivo_analisado, resultado_analise, nome_arquivo_saida):
    document = Document()

    # Título principal
    titulo = document.add_heading('Relatório de Análise - Proposal Analyzer by Arias', level=1)
    titulo.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER

    # Arquivo analisado
    par_arquivo = document.add_paragraph()
    par_arquivo.add_run('Arquivo analisado: ').bold = True
    par_arquivo.add_run(nome_arquivo_analisado)

    document.add_paragraph('')

    # Subtítulo
    subtitulo = document.add_paragraph()
    subtitulo.add_run('Resultado da Análise:').bold = True

    # Conteúdo da análise (formatado em parágrafos separados)
    for item in resultado_analise.split('\n'):
        document.add_paragraph(item, style='Normal')

    # Salvar
    document.save(nome_arquivo_saida)
