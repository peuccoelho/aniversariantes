from dotenv import load_dotenv
import os
import pandas as pd
from datetime import datetime
import calendar
from openpyxl import load_workbook
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors

load_dotenv()

arquivo_excel = os.getenv("ARQUIVO_EXCEL")

df = pd.read_excel(arquivo_excel)

if not arquivo_excel:
    print("Variável ARQUIVO_EXCEL não encontrada no .env.")
    exit()

df.columns = df.columns.str.strip().str.upper()

# converter múltiplos formatos de data
def converter_data(data):
    formatos = ['%d/%m/%Y', '%d.%m.%Y', '%d%m%Y']
    for f in formatos:
        try:
            return datetime.strptime(str(data), f)
        except:
            continue
    return pd.NaT

df["DATA DE NASCIMENTO"] = df["DATA DE NASCIMENTO"].apply(converter_data)
df = df[pd.to_datetime(df["DATA DE NASCIMENTO"], errors="coerce").notna()]  
df["DATA DE NASCIMENTO"] = pd.to_datetime(df["DATA DE NASCIMENTO"], dayfirst=True)  


mes_escolhido = int(input("Digite o número do mês (1 a 12): "))

meses_pt = [
    "", "Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho",
    "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"
]
nome_mes = meses_pt[mes_escolhido]

# Filtrar aniversariantes
aniversariantes = df[df['DATA DE NASCIMENTO'].dt.month == mes_escolhido].copy()


if aniversariantes.empty:
    print(f"Nenhuma pessoa faz aniversário em {nome_mes}.")
else:
    print(f"{len(aniversariantes)} aniversariante(s) encontrado(s) para {nome_mes}.")

    aniversariantes['DIA'] = aniversariantes['DATA DE NASCIMENTO'].dt.day
    aniversariantes = aniversariantes.sort_values(by='DIA')
    aniversariantes.drop(columns=['DIA'], inplace=True)

    colunas_saida = ['DATA DE NASCIMENTO', 'GH', 'NOME COMPLETO', 'ETOR']
    aniversariantes_saida = aniversariantes[colunas_saida].copy()

    # formatar a data como dd/mm/yyyy para excel
    aniversariantes_saida['DATA DE NASCIMENTO'] = aniversariantes_saida['DATA DE NASCIMENTO'].dt.strftime('%d/%m/%Y')

    nome_arquivo_excel = f'aniversariantes_mes_{mes_escolhido:02}.xlsx'
    nome_aba = f"Aniversariantes_{nome_mes}"

    with pd.ExcelWriter(nome_arquivo_excel, engine='openpyxl') as writer:
        aniversariantes_saida.to_excel(writer, sheet_name=nome_aba, index=False, startrow=2)

    wb = load_workbook(nome_arquivo_excel)
    ws = wb[nome_aba]
    ws.merge_cells('A1:D1')
    ws['A1'] = f"Aniversariantes do mês {nome_mes}"

    from openpyxl.styles import Font, Alignment
    ws['A1'].font = Font(bold=True, size=14)
    ws['A1'].alignment = Alignment(horizontal='center')
    wb.save(nome_arquivo_excel)

    # gerar o PDF
    nome_arquivo_pdf = f'aniversariantes_mes_{mes_escolhido:02}.pdf'
    doc = SimpleDocTemplate(nome_arquivo_pdf, pagesize=A4)
    styles = getSampleStyleSheet()

    titulo_style = ParagraphStyle(
        name='CenterTitle',
        parent=styles['Title'],
        alignment=1,  
        fontSize=16,
        spaceAfter=12
    )

    elementos = []

    titulo = f"Aniversariantes do mês {nome_mes}"
    elementos.append(Paragraph(titulo, titulo_style))
    elementos.append(Spacer(1, 12))

    # preparar dados para o PDF
    dados_pdf = aniversariantes_saida.copy()
    dados_pdf['DATA DE NASCIMENTO'] = pd.to_datetime(
        dados_pdf['DATA DE NASCIMENTO'], format='%d/%m/%Y', errors='coerce'
    ).dt.strftime('%d/%m')


    dados_pdf.rename(columns={'ETOR': 'SETOR'}, inplace=True)

    colunas_pdf = ['DATA DE NASCIMENTO', 'GH', 'NOME COMPLETO', 'SETOR']
    dados = [colunas_pdf] + dados_pdf.values.tolist()

    tabela = Table(dados, hAlign='CENTER')
    tabela.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.darkblue),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, -1), 10),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 8),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
    ]))

    elementos.append(tabela)
    doc.build(elementos)

    print(f"Arquivos gerados com sucesso: {nome_arquivo_excel} e {nome_arquivo_pdf}")
