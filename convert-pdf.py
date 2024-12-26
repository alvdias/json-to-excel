import pdfplumber
import pandas as pd
import re
from typing import List, Dict

def extrair_texto_pdf(caminho_pdf: str) -> List[str]:

    textos = []
    
    with pdfplumber.open(caminho_pdf) as pdf:
        for pagina in pdf.pages:
            texto = pagina.extract_text()
            if texto:
                textos.append(texto)
    
    return textos

def processar_texto(textos: List[str]) -> List[Dict]:
    dados = []
    
    for texto in textos:

        linhas = texto.split('\n')
        
        for linha in linhas:

            linha = re.sub(r'\s+', ' ', linha.strip())
            
            if linha:
                campos = re.split(r'\s{2,}', linha)
                
                registro = {}
                for i, campo in enumerate(campos):
                    registro[f'campo_{i+1}'] = campo
                
                dados.append(registro)
    
    return dados

def salvar_excel(dados: List[Dict], caminho_saida: str):
    df = pd.DataFrame(dados)
    
    if len(dados) > 0:
        primeira_linha = list(dados[0].values())
        if all(isinstance(x, str) for x in primeira_linha):
            df.columns = primeira_linha
            df = df.iloc[1:]
    
    df.to_excel(caminho_saida, index=False)

def pdf_para_excel(caminho_pdf: str, caminho_excel: str):

    try:
        textos = extrair_texto_pdf(caminho_pdf)
        dados = processar_texto(textos)
        salvar_excel(dados, caminho_excel)
        
        print(f"Arquivo Excel criado com sucesso: {caminho_excel}")
        
    except Exception as e:
        print(f"Erro ao processar o PDF: {str(e)}")

if __name__ == "__main__":
    caminho_pdf = "diego.pdf"
    caminho_excel = "saida-diego.xlsx"
    pdf_para_excel(caminho_pdf, caminho_excel)