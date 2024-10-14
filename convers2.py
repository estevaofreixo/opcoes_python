import re
import win32com.client as win32
from pypdf import PdfReader
from pathlib import Path
from datetime import datetime

# Definir o caminho da pasta Downloads do usuário
pasta_downloads = Path.home() / 'Downloads'

# Função para encontrar o último arquivo PDF baixado
def encontrar_ultimo_pdf(pasta):
    arquivos_pdf = list(pasta.glob('*.pdf'))  # Encontra todos os arquivos PDF na pasta
    if arquivos_pdf:
        return max(arquivos_pdf, key=lambda arquivo: arquivo.stat().st_mtime)  # Retorna o mais recente
    else:
        return None

# Encontrar o último arquivo PDF baixado
caminho_pdf = encontrar_ultimo_pdf(pasta_downloads)

if caminho_pdf is None:
    print("Nenhum arquivo PDF encontrado na pasta Downloads.")
else:
    # Função para extrair a combinação alfanumérica entre o underline e o hífen do nome do arquivo
    def extrair_combinacao_alfanumerica(nome_arquivo):
        padrao = r'_(\w+)'  # Expressão regular para encontrar a combinação entre "_" e "-"
        match = re.search(padrao, nome_arquivo)
        if match:
            return match.group(1)
        return None

    # Extraindo a combinação alfanumérica do nome do arquivo
    combinacao_alfanumerica = extrair_combinacao_alfanumerica(caminho_pdf.name)
    
    if combinacao_alfanumerica:
        print(f"Combinação alfanumérica extraída: {combinacao_alfanumerica}")
    else:
        print("Nenhuma combinação alfanumérica encontrada no nome do arquivo.")

    # Leitura do PDF baixado
    leitor_pdf = PdfReader(caminho_pdf)

    # Variáveis para armazenar os valores capturados
    valor_data_liquidacao = None
    valor_liquido_reclamante = None
    valor_deposito_fgts = None
    valor_irpf_reclamante = None
    valor_custas_judiciais = None
    valor_honorarios_primeira = None
    valor_honorarios_segunda = None
    valor_honorarios_terceira = None
    valor_c_primeira = None
    valor_c_segunda = None
    valor_d_segunda = None
    valor_contribuicao_social = None  # Nova variável para a contribuição social

    # Função para corrigir possíveis inversões de datas
    def corrigir_data_invertida(data_str):
        try:
            return datetime.strptime(data_str, '%d/%m/%Y')
        except ValueError:
            try:
                return datetime.strptime(data_str, '%m/%d/%Y')
            except ValueError:
                return None

    # Função para capturar a segunda data no formato dd/mm/yyyy, caractere por caractere
    def encontrar_segunda_data_caractere_por_caractere(texto):
        datas = []
        buffer = ""
        for char in texto:
            if char.isdigit() or char == '/':
                buffer += char
            else:
                if len(buffer) == 10 and buffer.count('/') == 2:  # Verifica se é uma data
                    datas.append(buffer)
                buffer = ""  # Reiniciar o buffer para a próxima data

        if len(datas) >= 2:
            return corrigir_data_invertida(datas[1])  # Retorna a segunda data
        return None

    # Função para encontrar o maior número em uma linha (considerando formatação com vírgula e ponto)
    def encontrar_maior_numero(linha):
        numeros = re.findall(r'\d{1,3}(?:\.\d{3})*,\d{2}|\d{1,3}(?:\.\d{3})*', linha)
        numeros_convertidos = [float(num.replace('.', '').replace(',', '.')) for num in numeros if num]
        if numeros_convertidos:
            return max(numeros_convertidos)
        return None

    # Função para encontrar o número posterior a um texto (com flexibilidade para texto intermediário)
    def encontrar_numero_posterior(texto, padrao):
        ocorrencia = re.search(rf'{padrao}.*?(\d{{1,3}}(?:\.\d{{3}})*,\d{{2}})', texto, re.DOTALL)
        if ocorrencia:
            return float(ocorrencia.group(1).replace('.', '').replace(',', '.'))
        return None

    # Função para encontrar o número anterior a um texto
    def encontrar_numero_anterior(texto, padrao):
        ocorrencia = re.search(rf'(\d{{1,3}}(?:\.\d{{3}})*,\d{{2}})\s*{padrao}', texto)
        if ocorrencia:
            return float(ocorrencia.group(1).replace('.', '').replace(',', '.'))
        return None

    # Função para dividir o texto em duas partes: antes e depois de "CUSTAS JUDICIAIS DEVIDAS PELO RECLAMADO"
    def dividir_texto_custas(texto):
        partes = texto.split("CUSTAS JUDICIAIS DEVIDAS PELO RECLAMADO")
        if len(partes) == 2:
            return partes[0], partes[1]  # Parte antes de "CUSTAS" e parte depois de "CUSTAS"
        return texto, None  # Se "CUSTAS" não for encontrado, retornamos o texto inteiro como primeira parte

    # Itera sobre todas as páginas do PDF
    contador_d = 0  # Inicializa a variável contador_d
    contador_c = 0  # Inicializa a variável contador_c
    contador_honorarios = 0  # Variável auxiliar para controlar as ocorrências de honorários

    for pagina_num, pagina in enumerate(leitor_pdf.pages, 1):
        texto = pagina.extract_text()

        # Dividir o texto em duas partes: antes e depois de "CUSTAS JUDICIAIS DEVIDAS PELO RECLAMADO"
        texto_antes_custas, texto_depois_custas = dividir_texto_custas(texto)

        # Processar honorários antes de "CUSTAS JUDICIAIS DEVIDAS PELO RECLAMADO"
        if texto_antes_custas:
            linhas_antes = texto_antes_custas.split('\n')
            for linha in linhas_antes:
                if linha.startswith("HONORÁRIOS LÍQUIDOS PARA"):
                    contador_honorarios += 1
                    valor_honorarios = encontrar_numero_posterior(linha, "HONORÁRIOS LÍQUIDOS PARA")
                    if contador_honorarios == 1:
                        valor_honorarios_primeira = valor_honorarios
                    elif contador_honorarios == 2:
                        valor_honorarios_segunda = valor_honorarios

                # Encontrar o valor da "CONTRIBUIÇÃO SOCIAL SOBRE SALÁRIOS DEVIDOS"
                if "CONTRIBUIÇÃO SOCIAL SOBRE SALÁRIOS DEVIDOS" in linha:
                    valor_contribuicao_social = encontrar_numero_posterior(linha, "CONTRIBUIÇÃO SOCIAL SOBRE SALÁRIOS DEVIDOS")

                # Encontrar o valor de "LÍQUIDO DEVIDO AO RECLAMANTE" na mesma linha
                if "LÍQUIDO DEVIDO AO RECLAMANTE" in linha:
                    valor_liquido_reclamante = encontrar_numero_posterior(linha, "LÍQUIDO DEVIDO AO RECLAMANTE")

        if texto_depois_custas:
            linhas_depois = texto_depois_custas.split('\n')
            for linha in linhas_depois:
                if linha.startswith("HONORÁRIOS LÍQUIDOS PARA"):
                    valor_honorarios_terceira = encontrar_numero_posterior(linha, "HONORÁRIOS LÍQUIDOS PARA")

        linhas = texto.split('\n')

        if valor_data_liquidacao is None:
            valor_data_liquidacao = encontrar_segunda_data_caractere_por_caractere(texto)  # Alterado para segunda data

        if valor_deposito_fgts is None:
            valor_deposito_fgts = encontrar_numero_posterior(texto, "DEPÓSITO FGTS")

        if valor_irpf_reclamante is None:
            valor_irpf_reclamante = encontrar_numero_posterior(texto, "IRPF DEVIDO PELO RECLAMANTE")

        if valor_custas_judiciais is None:
            valor_custas_judiciais = encontrar_numero_posterior(texto, "CUSTAS JUDICIAIS DEVIDAS PELO RECLAMADO")

        for linha in linhas:
            if "C = A x B" in linha:
                contador_c += 1
                maior_numero = encontrar_maior_numero(linha)
                if contador_c == 1 and maior_numero is not None:
                    valor_c_primeira = maior_numero
                elif contador_c == 2 and maior_numero is not None:
                    valor_c_segunda = maior_numero

            elif "D = A x B" in linha:
                contador_d += 1
                if contador_d == 2:
                    maior_numero = encontrar_maior_numero(linha)
                    if maior_numero is not None:
                        valor_d_segunda = maior_numero

    # Manipulação do Excel via COM (pywin32)
    excel = win32.Dispatch("Excel.Application")

    # Tenta acessar o arquivo do Excel já aberto, caso contrário, o abre
    try:
        workbook = excel.Workbooks("CONVERSAO EM TR 3.xlsm")  # Tenta acessar o arquivo já aberto
    except Exception as e:
        caminho_arquivo_excel = r"C:\Users\estevao.freixo\Desktop\TRT\HOMOLOGACAO\CONVERSAO EM TR.xlsm"  # Ajuste o caminho para o correto
        workbook = excel.Workbooks.Open(caminho_arquivo_excel)

    worksheet = workbook.Sheets(1)

    # Função para forçar o cálculo de uma célula
    def forcar_calculo_celula(celula):
        worksheet.Range(celula).Value = worksheet.Range(celula).Value

    # Limpar células antes de preencher com novos valores
    def limpar_celulas():
        celulas_para_limpar = ["D9", "C34", "C35", "C13", "C20", "C24", "C21", "C22", "C23", "C28", "C29", "C30", "C19", "C18", "C14", "C15", "C16"]
        for celula in celulas_para_limpar:
            worksheet.Range(celula).Value = None

    # Limpar as células antes de preencher com novos valores
    limpar_celulas()

    # Preencher as células com novos valores
    if valor_data_liquidacao is not None:
        worksheet.Range("C34").Value = valor_data_liquidacao
        forcar_calculo_celula("C34")

    if combinacao_alfanumerica is not None:
        worksheet.Range("C35").Value = combinacao_alfanumerica
        forcar_calculo_celula("C35")

    if valor_liquido_reclamante is not None:
        worksheet.Range("C13").Value = valor_liquido_reclamante  # Valor inserido na célula C13
        forcar_calculo_celula("C13")

    if valor_deposito_fgts is not None:
        worksheet.Range("C20").Value = valor_deposito_fgts
        forcar_calculo_celula("C20")

    if valor_honorarios_primeira is not None:
        worksheet.Range("C24").Value = valor_honorarios_primeira
        forcar_calculo_celula("C24")

    if valor_honorarios_segunda is not None:
        worksheet.Range("C21").Value = valor_honorarios_segunda
        forcar_calculo_celula("C21")

    if valor_honorarios_terceira is not None:
        worksheet.Range("C23").Value = valor_honorarios_terceira
        forcar_calculo_celula("C23")

    if valor_irpf_reclamante is not None:
        worksheet.Range("C19").Value = valor_irpf_reclamante
        forcar_calculo_celula("C19")

    if valor_custas_judiciais is not None:
        worksheet.Range("C18").Value = valor_custas_judiciais
        forcar_calculo_celula("C18")

    if valor_d_segunda is not None:
        worksheet.Range("C14").Value = valor_d_segunda
        forcar_calculo_celula("C14")

    if valor_c_primeira is not None:
        worksheet.Range("C15").Value = valor_c_primeira
        forcar_calculo_celula("C15")

    if valor_contribuicao_social is not None:  # Preencher a célula C15 com o valor da contribuição social
        worksheet.Range("C15").Value = valor_contribuicao_social
        forcar_calculo_celula("C15")

    if valor_c_segunda is not None:
        worksheet.Range("C16").Value = valor_c_segunda
        forcar_calculo_celula("C16")

    excel.Visible = True

    print(f"Valores numéricos e datas inseridos corretamente nas células do arquivo: {caminho_pdf.name}.")
