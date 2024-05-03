from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.platypus import Table, TableStyle, Image
from calendar import monthrange, weekday


def criar_pdf_controle_ponto(mes, ano, feriados, facultativos):
    # Configurações do PDF
    nome_arquivo = f"controle_ponto_{mes}_{ano}.pdf"
    pdf = canvas.Canvas(nome_arquivo, pagesize=A4)
    largura_pagina, altura_pagina = A4

    # Cabeçalho
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

if __name__ == '__main__':
    # Exemplo de uso
    mes = 5
    ano = 2024
    feriados = [27]
    falcultativos = [28]
    criar_pdf_controle_ponto(mes, ano, feriados, falcultativos)

