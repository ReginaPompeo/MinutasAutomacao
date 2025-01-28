from docx import Document
import pandas as pd

# Carregar a planilha do Excel
df = pd.read_excel('PLANILHA_TESTE.xlsx')

# Carregar o documento Word
doc = Document('MODELO MINUTA DE ACORDO.docx')

# Iterar sobre as linhas da planilha
for index, row in df.iterrows():
    vara = row['VARA']  # Coluna "VARA" na planilha
    uf = row['UF']      # Coluna "UF" na planilha
    comarca = row['COMARCA']  # Coluna "COMARCA" na planilha
    cnj = row['CNJ'] 
    nome = row['NOME']
    valor_acordo = row['VALOR ACORDO']

    # Iterar sobre os parágrafos do documento
    for para in doc.paragraphs:
        # Iterar sobre os runs para preservar a formatação
        for run in para.runs:
            if '<ÓRGÃO DE TRAMITAÇÃO>' in run.text:
                run.text = run.text.replace('<ÓRGÃO DE TRAMITAÇÃO>', vara)
            if '<UF>' in run.text:
                run.text = run.text.replace('<UF>', uf)
            if '<COMARCA>' in run.text:
                run.text = run.text.replace('<COMARCA>', comarca)                
            if '<NR. DO PROCESSO>' in run.text:
                run.text = run.text.replace('<NR. DO PROCESSO>', cnj)
            if '<POUPADOR/SUCESSORES>' in run.text:
                run.text = run.text.replace('<POUPADOR/SUCESSORES>', nome)    
            if '<VALOR DO ACORDO>' in run.text:
                run.text = run.text.replace('<VALOR ACORDO>', valor_acordo)

# Salvar o documento atualizado
doc.save('documento_atualizado.docx')
