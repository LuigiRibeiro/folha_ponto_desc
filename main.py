from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.platypus import Table, TableStyle, Image
from calendar import monthrange, weekday
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.platypus import Paragraph
import PyPDF2
import io
from reportlab.lib.pagesizes import letter
from PyPDF2 import PdfReader, PdfWriter
from datetime import datetime


def criar_pdf_controle_ponto(nome_mes, mes, ano, feriados, facultativos):
    # Configurações do PDF
    nome_arquivo = f"controle_ponto_{mes}_{ano}.pdf"
    pdf = canvas.Canvas(nome_arquivo, pagesize=A4)
    largura_pagina, altura_pagina = A4


    # Adicionar imagem no topo da página
    pdf.drawImage("cabeçalho.png", 27, 560, width=540, height=260)

    # Cabeçalho
    estilo = getSampleStyleSheet()["Normal"]
    estilo.fontName = "Times-Roman"
    estilo.fontSize = 16
    estilo.leading = 14

    # Texto com estilo
    texto = f"Mês de {nome_mes} / {ano}"
    paragrafo = Paragraph(texto, estilo)
    paragrafo.wrapOn(pdf, 400, 100)
    paragrafo.drawOn(pdf, largura_pagina / 2 - len(nome_mes) * 4 - 55, altura_pagina - 100)

    pdf.drawString(30, 30, f"*De acordo com a CI UERJ/GR Nº 246 /2023.")

    # Tabela de dias do mês
    dias_no_mes = monthrange(ano, mes)[1]
    dias_ate_15 = min(dias_no_mes, 15)

    # Adicionar retângulo cinza ao redor das tabelas
    pdf.setStrokeColor(colors.black)
    pdf.setFillColor('#c0c0c0')
    altura_linhas = [20] + [30] * max(dias_ate_15, dias_no_mes - 15)
    pdf.rect(30, altura_pagina - 290, largura_pagina - 61, -sum(altura_linhas), fill=1)

    # Dividindo os dias em duas colunas
    dados_tabela_esquerda = [['DIA']]
    dados_tabela_direita = [['DIA']]
    for dia in range(1, dias_ate_15 + 1):
        dia_format = str(dia).zfill(2)
        dados_tabela_esquerda.append([str(dia_format)])
    for dia in range(dias_ate_15 + 1, dias_no_mes + 1):
        dia_format = str(dia).zfill(2)
        dados_tabela_direita.append([str(dia_format)])

    larg_util = largura_pagina / 2 - 28
    largura_colunas = [int(larg_util*0.08)]
    altura_linhas_E = [20] + [30] * dias_ate_15
    altura_linhas_D = [20] + [30] * (dias_no_mes - 15)
    tabela_esquerda = Table(dados_tabela_esquerda, colWidths=largura_colunas, rowHeights=altura_linhas_E)
    tabela_direita = Table(dados_tabela_direita, colWidths=largura_colunas, rowHeights=altura_linhas_D)

    # Estilo da tabela
    estilo = TableStyle([
        ('GRID', (0, 0), (-1, -1), 1, colors.black),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('BACKGROUND', (0, 0), (-1, -1), colors.white),
        ('FONTNAME', (0, 0), (-1, -1), 'Times-Roman'),  # Define a fonte como Helvetica
        ('FONTSIZE', (0, 0), (-1, -1), 7)
    ])
    tabela_esquerda.setStyle(estilo)
    tabela_direita.setStyle(estilo)

    # Posicionamento da tabela
    tabela_esquerda.wrapOn(pdf, largura_pagina / 2 - 30, altura_pagina)
    tabela_esquerda.drawOn(pdf, 30, altura_pagina - sum(altura_linhas_E) - 290)
    tabela_direita.wrapOn(pdf, largura_pagina / 2 - 30, altura_pagina)
    tabela_direita.drawOn(pdf, largura_pagina / 2 - 0.5, altura_pagina - sum(altura_linhas_D) - 290)

    # Dividindo os dias em duas colunas
    dados_tabela_esquerda = [['INÍCIO', 'TÉRMINO']]
    dados_tabela_direita = [['INÍCIO', 'TÉRMINO']]

    largura_colunas = [int(larg_util * 0.457), int(larg_util * 0.457)]
    altura_linhas_E = [12]
    altura_linhas_D = [12]
    tabela_esquerda = Table(dados_tabela_esquerda, colWidths=largura_colunas, rowHeights=altura_linhas_E)
    tabela_direita = Table(dados_tabela_direita, colWidths=largura_colunas, rowHeights=altura_linhas_D)

    # Estilo da tabela
    estilo = TableStyle([
        ('GRID', (0, 0), (-1, -1), 1, colors.black),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('BACKGROUND', (0, 0), (-1, -1), colors.white),
        ('FONTNAME', (0, 0), (-1, -1), 'Times-Roman'),
        ('FONTSIZE', (0, 0), (-1, -1), 7),
        ('TOPPADDING', (0, 0), (-1, -1), 6)
    ])

    tabela_esquerda.setStyle(estilo)
    tabela_direita.setStyle(estilo)

    # Posicionamento da tabela
    tabela_esquerda.wrapOn(pdf, largura_pagina / 2 - 30, altura_pagina)
    tabela_esquerda.drawOn(pdf, 30 + int(larg_util * 0.08), altura_pagina - sum(altura_linhas_E) - 290)
    tabela_direita.wrapOn(pdf, largura_pagina / 2 - 30, altura_pagina)
    tabela_direita.drawOn(pdf, largura_pagina / 2 - 0.5 + int(larg_util * 0.08), altura_pagina - sum(altura_linhas_D) - 290)

    # Dividindo os dias em duas colunas
    dados_tabela_esquerda = [['HORA', 'RUBRICA', 'HORA', 'RUBRICA']]
    dados_tabela_direita = [['HORA', 'RUBRICA', 'HORA', 'RUBRICA']]

    largura_colunas = [int(larg_util * 0.15), int(larg_util * 0.308), int(larg_util * 0.15), int(larg_util * 0.308)]
    altura_linhas_E = [8]
    altura_linhas_D = [8]
    tabela_esquerda = Table(dados_tabela_esquerda, colWidths=largura_colunas, rowHeights=altura_linhas_E)
    tabela_direita = Table(dados_tabela_direita, colWidths=largura_colunas, rowHeights=altura_linhas_D)

    # Estilo da tabela
    estilo = TableStyle([
        ('GRID', (0, 0), (-1, -1), 1, colors.black),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('BACKGROUND', (0, 0), (-1, -1), colors.white),
        ('FONTNAME', (0, 0), (-1, -1), 'Times-Roman'),
        ('FONTSIZE', (0, 0), (-1, -1), 6),
        ('TOPPADDING', (0, 0), (-1, -1), 7)
    ])

    tabela_esquerda.setStyle(estilo)
    tabela_direita.setStyle(estilo)

    # Posicionamento da tabela
    tabela_esquerda.wrapOn(pdf, largura_pagina / 2 - 30, altura_pagina)
    tabela_esquerda.drawOn(pdf, 30 + int(larg_util * 0.08), altura_pagina - sum(altura_linhas_E) - 290 - 12)
    tabela_direita.wrapOn(pdf, largura_pagina / 2 - 30, altura_pagina)
    tabela_direita.drawOn(pdf, largura_pagina / 2 - 0.5 + int(larg_util * 0.08),
                          altura_pagina - sum(altura_linhas_D) - 290 - 12)

    a = Image("rubrica.png", 20, 45)

    # Dividindo os dias em duas colunas
    dados_tabela_esquerda = []
    dados_tabela_direita = []
    for dia in range(1, dias_ate_15 + 1):
        dia_semana = weekday(ano, mes, dia).value
        if dia_semana == 5 or dia_semana == 6 or dia in feriados or dia in facultativos:
            dados_tabela_esquerda.append(['', '', '', ''])
            dados_tabela_esquerda.append(['', '', '', ''])
        else:
            dados_tabela_esquerda.append(['08H', a, '12H', a])
            dados_tabela_esquerda.append(['13H', '', '17H', ''])

    for dia in range(dias_ate_15 + 1, dias_no_mes + 1):
        dia_semana = weekday(ano, mes, dia).value
        if dia_semana == 5 or dia_semana == 6 or dia in feriados or dia in facultativos:
            dados_tabela_direita.append(['', '', '', ''])
            dados_tabela_direita.append(['', '', '', ''])
        else:
            dados_tabela_direita.append(['08H', a, '12H', a])
            dados_tabela_direita.append(['13H', '', '17H', ''])

    larg_util = largura_pagina / 2 - 28
    largura_colunas = [int(larg_util * 0.15), int(larg_util * 0.308), int(larg_util * 0.15), int(larg_util * 0.308)]
    altura_linhas_E = [15] * dias_ate_15 * 2
    altura_linhas_D = [15] * (dias_no_mes - 15) * 2
    tabela_esquerda = Table(dados_tabela_esquerda, colWidths=largura_colunas, rowHeights=altura_linhas_E)
    tabela_direita = Table(dados_tabela_direita, colWidths=largura_colunas, rowHeights=altura_linhas_D)

    # Estilo da tabela
    estilo = TableStyle([
        ('GRID', (0, 0), (-1, -1), 1, colors.black),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('BACKGROUND', (0, 0), (-1, -1), colors.white),
        ('FONTNAME', (0, 0), (-1, -1), 'Helvetica'),
        ('FONTSIZE', (0, 0), (-1, -1), 10),
        ('FONTWEIGHT', (0, 0), (-1, -1), 'BOLD')
    ])
    tabela_esquerda.setStyle(estilo)
    tabela_direita.setStyle(estilo)

    # Posicionamento da tabela
    tabela_esquerda.wrapOn(pdf, largura_pagina / 2 - 30, altura_pagina)
    tabela_esquerda.drawOn(pdf, 30 + int(larg_util * 0.08), altura_pagina - sum(altura_linhas_E) - 290 - 20)
    tabela_direita.wrapOn(pdf, largura_pagina / 2 - 30, altura_pagina)
    tabela_direita.drawOn(pdf, largura_pagina / 2 - 0.5 + int(larg_util * 0.08), altura_pagina - sum(altura_linhas_D) - 290 - 20)

    # Dividindo os dias em duas colunas
    dados_tabela_esquerda = []
    dados_tabela_direita = []
    for dia in range(1, dias_ate_15 + 1):
        dia_semana = weekday(ano, mes, dia).value
        if dia_semana == 5:
            dados_tabela_esquerda.append(['SÁBADO', 'SÁBADO', 'SÁBADO', 'SÁBADO'])
            dados_tabela_esquerda.append(['SÁBADO', 'SÁBADO', 'SÁBADO', 'SÁBADO'])
        elif dia_semana == 6:
            dados_tabela_esquerda.append(['DOMINGO', 'DOMINGO', 'DOMINGO', 'DOMINGO'])
            dados_tabela_esquerda.append(['DOMINGO', 'DOMINGO', 'DOMINGO', 'DOMINGO'])
        elif dia in feriados:
            dados_tabela_esquerda.append(['FERIADO', 'FERIADO', 'FERIADO', 'FERIADO'])
            dados_tabela_esquerda.append(['FERIADO', 'FERIADO', 'FERIADO', 'FERIADO'])
        elif dia in facultativos:
            dados_tabela_esquerda.append(['PONTO', 'FACULTATIVO', 'PONTO', 'FACULTATIVO'])
            dados_tabela_esquerda.append(['PONTO', 'FACULTATIVO', 'PONTO', 'FACULTATIVO'])
        else:
            dados_tabela_esquerda.append(['', '', '', ''])
            dados_tabela_esquerda.append(['', '', '', ''])

    for dia in range(dias_ate_15 + 1, dias_no_mes + 1):
        dia_semana = weekday(ano, mes, dia).value
        if dia_semana == 5:
            dados_tabela_direita.append(['SÁBADO', 'SÁBADO', 'SÁBADO', 'SÁBADO'])
            dados_tabela_direita.append(['SÁBADO', 'SÁBADO', 'SÁBADO', 'SÁBADO'])
        elif dia_semana == 6:
            dados_tabela_direita.append(['DOMINGO', 'DOMINGO', 'DOMINGO', 'DOMINGO'])
            dados_tabela_direita.append(['DOMINGO', 'DOMINGO', 'DOMINGO', 'DOMINGO'])
        elif dia in feriados:
            dados_tabela_direita.append(['FERIADO', 'FERIADO', 'FERIADO', 'FERIADO'])
            dados_tabela_direita.append(['FERIADO', 'FERIADO', 'FERIADO', 'FERIADO'])
        elif dia in facultativos:
            dados_tabela_direita.append(['PONTO', 'FACULTATIVO', 'PONTO', 'FACULTATIVO'])
            dados_tabela_direita.append(['PONTO', 'FACULTATIVO', 'PONTO', 'FACULTATIVO'])
        else:
            dados_tabela_direita.append(['', '', '', ''])
            dados_tabela_direita.append(['', '', '', ''])

    larg_util = largura_pagina / 2 - 28
    largura_colunas = [int(larg_util * 0.15), int(larg_util * 0.308), int(larg_util * 0.15), int(larg_util * 0.308)]
    altura_linhas_E = [15] * dias_ate_15 * 2
    altura_linhas_D = [15] * (dias_no_mes - 15) * 2
    tabela_esquerda = Table(dados_tabela_esquerda, colWidths=largura_colunas, rowHeights=altura_linhas_E)
    tabela_direita = Table(dados_tabela_direita, colWidths=largura_colunas, rowHeights=altura_linhas_D)

    # Estilo da tabela
    estilo = TableStyle([
        ('GRID', (0, 0), (-1, -1), 1, colors.black),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('FONTNAME', (0, 0), (-1, -1), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, -1), 8),
        ('FONTWEIGHT', (0, 0), (-1, -1), 'BOLD'),
        ('TEXTCOLOR', (0, 0), (-1, -1), colors.red),
        ('TOPPADDING', (0, 0), (-1, -1), 5)
    ])
    tabela_esquerda.setStyle(estilo)
    tabela_direita.setStyle(estilo)

    # Posicionamento da tabela
    tabela_esquerda.wrapOn(pdf, largura_pagina / 2 - 30, altura_pagina)
    tabela_esquerda.drawOn(pdf, 30 + int(larg_util * 0.08), altura_pagina - sum(altura_linhas_E) - 290 - 20)
    tabela_direita.wrapOn(pdf, largura_pagina / 2 - 30, altura_pagina)
    tabela_direita.drawOn(pdf, largura_pagina / 2 - 0.5 + int(larg_util * 0.08),
                          altura_pagina - sum(altura_linhas_D) - 290 - 20)

    # Fechar PDF
    pdf.save()

    print(f"PDF gerado com sucesso: {nome_arquivo}")


def juntar_pdf(mes, ano):
    # Abrir o PDF existente
    pdf_existente = PyPDF2.PdfReader(open(f"controle_ponto_{mes}_{ano}.pdf", "rb"))

    # Abrir o PDF com a imagem
    pdf_imagem = PyPDF2.PdfReader(open("FOLHA FREQ_VERSO ASSINADA.pdf", "rb"))

    # Criar um novo PDF
    pdf_novo = PyPDF2.PdfWriter()

    # Adicionar páginas do PDF existente ao novo PDF
    for pagina_num in range(len(pdf_existente.pages)):
        pagina = pdf_existente.pages[pagina_num]
        pdf_novo.add_page(pagina)

    # Adicionar páginas do PDF com a imagem ao novo PDF (assumindo que a imagem está na primeira página)
    pagina_imagem = pdf_imagem.pages[0]
    pdf_novo.add_page(pagina_imagem)

    # Salvar o novo PDF
    with open(f"FOLHA FREQ {mes}_{ano} - LUIGI.pdf", "wb") as f:
        pdf_novo.write(f)


def ajustar_verso():
    packet = io.BytesIO()
    can = canvas.Canvas(packet, pagesize=A4)
    data_atual = datetime.now().strftime("%d   %m   %Y")
    can.drawString(66, 86, data_atual)
    can.drawImage("assinatura_com_fundo.png", 40, 95, width=88, height=40)
    can.save()

    # move to the beginning of the StringIO buffer
    packet.seek(0)

    # create a new PDF with Reportlab
    new_pdf = PdfReader(packet)
    # read your existing PDF
    existing_pdf = PdfReader(open("FOLHA FREQ_VERSO.pdf", "rb"))
    output = PdfWriter()
    # add the "watermark" (which is the new pdf) on the existing page
    page = existing_pdf.pages[0]
    page.merge_page(new_pdf.pages[0])
    output.add_page(page)
    # finally, write "output" to a real file
    output_stream = open("FOLHA FREQ_VERSO ASSINADA.pdf", "wb")
    output.write(output_stream)
    output_stream.close()


if __name__ == '__main__':
    # Exemplo de uso
    mes = 4
    nomes_mes = 'ABRIL'
    ano = 2024
    feriados = [23]
    falcultativos = [22]
    criar_pdf_controle_ponto(nomes_mes, mes, ano, feriados, falcultativos)
    ajustar_verso()
    juntar_pdf(mes, ano)

