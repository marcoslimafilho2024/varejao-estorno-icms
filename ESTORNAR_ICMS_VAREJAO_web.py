# ESTORNAR_ICMS_VAREJAO_web.py

from flask import Flask, render_template, request, redirect, url_for, send_file
import pandas as pd
import locale
from io import BytesIO
from fpdf import FPDF
from dados_biblioteca import dados_biblioteca  # Importa os dados embutidos

# Configuração de formatação de moeda brasileira
locale.setlocale(locale.LC_ALL, 'pt_BR.UTF-8')

app = Flask(__name__)

# Função para carregar a biblioteca de produtos estornados diretamente dos dados JSON incorporados
def carregar_bd_geral():
    return pd.DataFrame(dados_biblioteca)

# Função para validar o layout do arquivo Excel
def validar_layout(df):
    required_columns = ['IDPRODUTO', 'VALICM']
    for col in required_columns:
        if col not in df.columns:
            raise ValueError(f"O arquivo não contém a coluna obrigatória: {col}")

# Função para realizar o cálculo de estorno usando os dados da biblioteca
def calcular_estorno(df):
    produtos_estornados = carregar_bd_geral()
    valicm_total = df['VALICM'].sum()
    produtos_estorno_ids = produtos_estornados[produtos_estornados['ESTORNAR'] == 'SIM']['IDPRODUTO']
    df_filtrado = df[df['IDPRODUTO'].isin(produtos_estorno_ids)]
    valor_total_estornado = df_filtrado['VALICM'].sum()
    valor_liquido = valicm_total - valor_total_estornado
    return valicm_total, valor_total_estornado, valor_liquido

# Rota principal para exibir o formulário de upload - agora acessível tanto em "/" quanto em "/estorno_icms"
@app.route("/", methods=["GET", "POST"])
@app.route("/estorno_icms", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        # Receber o arquivo enviado pelo usuário
        file = request.files["file"]
        if file:
            try:
                df = pd.read_excel(file)
                validar_layout(df)
                valicm_total, valor_estornado, valor_liquido = calcular_estorno(df)
                
                # Formatando os valores para exibição
                valicm_total_formatado = locale.currency(valicm_total, grouping=True, symbol=True)
                valor_estornado_formatado = locale.currency(valor_estornado, grouping=True, symbol=True)
                valor_liquido_formatado = locale.currency(valor_liquido, grouping=True, symbol=True)

                # Renderizando os resultados na página
                return render_template("index.html", 
                                       valicm_total=valicm_total_formatado,
                                       valor_estornado=valor_estornado_formatado,
                                       valor_liquido=valor_liquido_formatado)
            except ValueError as ve:
                return render_template("index.html", error=str(ve))
            except Exception as e:
                return render_template("index.html", error=f"Erro ao processar o arquivo: {str(e)}")
    return render_template("index.html")

# Rota para gerar o PDF da memória de cálculo com caminho específico "/estorno_icms/download_pdf"
@app.route("/estorno_icms/download_pdf")
def download_pdf():
    try:
        valicm_total = request.args.get("valicm_total")
        valor_estornado = request.args.get("valor_estornado")
        valor_liquido = request.args.get("valor_liquido")

        pdf = FPDF()
        pdf.add_page()
        pdf.set_font("Arial", "B", 16)
        pdf.cell(200, 10, "Memória de Cálculo do Estorno do ICMS", ln=True, align="C")
        pdf.set_font("Arial", size=12)
        pdf.cell(200, 10, f"VALICM TOTAL: {valicm_total}", ln=True)
        pdf.cell(200, 10, f"VALOR ESTORNADO: {valor_estornado}", ln=True)
        pdf.cell(200, 10, f"VALOR LÍQUIDO: {valor_liquido}", ln=True)

        # Criar PDF em memória
        pdf_buffer = BytesIO()
        pdf.output(pdf_buffer)
        pdf_buffer.seek(0)

        return send_file(pdf_buffer, as_attachment=True, download_name="memoria_calculo_estorno_icms.pdf")
    except Exception as e:
        return redirect(url_for("index", error=f"Erro ao gerar PDF: {str(e)}"))

if __name__ == "__main__":
    app.run(debug=True)
