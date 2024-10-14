import PyPDF2

# Caminho para o arquivo PDF
caminho_pdf = 'C:/Users/estevao.freixo/Desktop/Python/calc.pdf'

# Abrindo o arquivo PDF
with open(caminho_pdf, 'rb') as arquivo:
    leitor_pdf = PyPDF2.PdfReader(arquivo)
    
    # Loop através das páginas e extrair texto
    for pagina in leitor_pdf.pages:
        texto = pagina.extract_text()
        print(texto)
