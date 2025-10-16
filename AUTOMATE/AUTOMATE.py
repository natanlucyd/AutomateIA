import sys
import pandas as pd
import requests
from fpdf import FPDF


print("Argumentos recebidos:", sys.argv)

if len(sys.argv) != 3:
    print("Uso correto: python AUTOMATE.py <arquivo_excel> <arquivo_pdf>")
    sys.exit(1)

arquivo_excel = sys.argv[1]
arquivo_pdf = sys.argv[2]

print("Excel recebido:", arquivo_excel)
print("PDF recebido:", arquivo_pdf)


API_KEY = ""  # 🔑 coloque aqui sua chave 
BASE_URL = ""  # endpoint 


try:
    df = pd.read_excel(arquivo_excel, sheet_name="Relatorio_PIPE", engine="openpyxl")
except Exception as e:
    print("Erro ao abrir o Excel:", e)
    sys.exit(1)


tabela_texto = df.to_string(index=False)

prompt = f"""
Você é um analista de negócios. Gere um resumo executivo e estruturado com base nesta tabela PIPE( que seria produtos que estão em "Reunião" "Cotação" "Em andamneto" ou seja não foram concluidos ainda):

{tabela_texto}

O resumo deve conter:
- Principais destaques (em tópicos)
- Maiores volumes e filiais relevantes
- Produtos de maior impacto
- Insights e recomendações estratégicas
- Comente sobre as reuniões e sobre os gráficos e de insights sobre isso
- De um resumo detalhado e Principais insights sobre a planilha no geral
- Sugerir fazer alguma coisa para diminuir a quantidade desses pipes e concluir eles( ja que quanto mais PIPE seria mais dinheiro na mesa que não foi recebido ainda)
"""


headers = {
    "x-api-key": API_KEY,
    "Content-Type": "application/json"
}

payload = {
    "message": prompt,
    "message_type": "user",
    "model": "gpt-5-mini",  
    "stream": False,
    "ephemeral": True
}

try:
    response = requests.post(BASE_URL, headers=headers, json=payload)
    response.raise_for_status()
    data = response.json()


    resumo = ""
    if "messages" in data:
        for msg in reversed(data["messages"]):
            if msg.get("message_type") == "agent" and msg.get("message"):
                resumo = msg["message"].strip()
                break
    elif "message" in data:
        resumo = data["message"]
    else:
        resumo = str(data)

    if not resumo.strip():
        resumo = "Não foi possível gerar o resumo automaticamente."

except Exception as e:
    print("Erro ao chamar a API Jumpad:", e)
    resumo = "Erro ao obter resposta da IA."

class PDF(FPDF):
    def header(self):
        self.set_font("ArialUnicode", size=10)
        self.cell(0, 10, "Relatório PIPE - Resumo Gerado por IA (BeSmart Jumpad)", align="C", ln=1)
        self.ln(2)

    def footer(self):
        self.set_y(-15)
        self.set_font("ArialUnicode", size=8)
        self.cell(0, 10, f"Página {self.page_no()}", align="C")


import re

def limpar_markdown(texto):
    texto = re.sub(r"\*\*(.*?)\*\*", r"\1", texto)  
    texto = re.sub(r"#+\s*", "", texto)             
    texto = re.sub(r"[_`]", "", texto)              

resumo = limpar_markdown(resumo)


pdf = PDF()
pdf.add_font("ArialUnicode", fname="C:\\Windows\\Fonts\\arial.ttf")
pdf.set_font("ArialUnicode", size=12)

pdf.add_page()
pdf.multi_cell(0, 8, resumo)
pdf.output(arquivo_pdf)

print(f"✅ PDF gerado com sucesso: {arquivo_pdf}")
