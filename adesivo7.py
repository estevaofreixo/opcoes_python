
import fitz  # PyMuPDF
import re
import os
import subprocess
import platform
from pathlib import Path

# Função para encontrar o arquivo PDF mais recente na pasta Downloads
def encontrar_pdf_mais_recente(diretorio):
    caminho_diretorio = Path(diretorio)
    arquivos_pdf = list(caminho_diretorio.glob('*.pdf'))  # Buscar apenas arquivos .pdf
    if not arquivos_pdf:
        raise FileNotFoundError("Nenhum arquivo PDF encontrado no diretório Downloads.")
    
    # Encontrar o arquivo PDF mais recente com base na data de modificação
    arquivo_pdf_mais_recente = max(arquivos_pdf, key=lambda f: f.stat().st_mtime)
    return arquivo_pdf_mais_recente

# Lista de expressões a serem procuradas (sensível a maiúsculas e minúsculas)
expressoes = [
    r'TST - Acórdão',
    r'TST - Decisão/Despacho',
    r'Acórdão',
    r'Sentença'
]

# Expressão regular para encontrar horários no formato hh:mm (também sensível à caixa)
expressao_horario = r'\b\d{2}:\d{2}\b'

# Função para realçar o texto encontrado
def realcar_texto(pagina, areas_texto):
    for area in areas_texto:
        # Adicionar realce ao redor da área onde o texto foi encontrado
        pagina.add_highlight_annot(area)

# Função para adicionar uma nota adesiva (sticky note) na página
def adicionar_nota_adesiva(pagina, area, texto_nota):
    # Obter o centro da área para posicionar a nota adesiva
    x0, y0, x1, y1 = area
    centro_x = (x0 + x1) / 2
    centro_y = (y0 + y1) / 2
    # Adicionar a nota adesiva com o texto fornecido
    pagina.add_text_annot((centro_x, centro_y), texto_nota)

# Função para verificar se a expressão está ao lado de um horário e com exatamente 1 espaço entre eles
def expressao_ao_lado_de_horario_sem_texto_visivel(texto, posicao_expressao):
    horarios = list(re.finditer(expressao_horario, texto))
    
    # Procurar se existe um horário imediatamente antes da expressão
    for horario in horarios:
        inicio_horario, fim_horario = horario.span()

        # Verificar se há apenas **espaços ou caracteres invisíveis** entre o horário e a expressão
        if fim_horario < posicao_expressao:
            texto_entre = texto[fim_horario:posicao_expressao]
            # Verificar se o texto entre o horário e a expressão contém **apenas espaços ou caracteres invisíveis**
            if texto_entre.isspace():  # Se o texto entre contém apenas espaços ou caracteres invisíveis
                return True
    return False

# Função para verificar se já processamos a expressão na mesma linha
def na_mesma_linha(area_atual, areas_processadas):
    # Considera que estão na mesma linha se a diferença na coordenada y0 for pequena
    for area in areas_processadas:
        _, y0_area, _, _ = area
        _, y0_atual, _, _ = area_atual
        if abs(y0_atual - y0_area) < 2:  # Tolerância para considerar que estão na mesma linha
            return True
    return False

# Função para processar as ocorrências da expressão e aplicar a numeração apenas nas realçadas
def processar_ultima_ocorrencia_por_expressao(pagina, ultima_ocorrencia_expressao, contadores, multiplas_ocorrencias):
    for expressao_completa, area in ultima_ocorrencia_expressao.items():
        # Obter apenas a primeira palavra da expressão
        primeira_palavra = expressao_completa.split()[0]

        # Verificar se a primeira palavra tem mais de uma ocorrência realçada
        if multiplas_ocorrencias[primeira_palavra]:
            # Incrementar o contador para essa expressão, começando em 1
            contadores[primeira_palavra] += 1
            texto_nota = f"{primeira_palavra} {contadores[primeira_palavra]}"  # Apenas a primeira palavra com numeração
        else:
            texto_nota = primeira_palavra  # Não numerar se houver apenas uma ocorrência

        # Realçar o texto encontrado
        realcar_texto(pagina, [area])
        # Adicionar uma nota adesiva com a numeração correspondente (ou sem numeração)
        adicionar_nota_adesiva(pagina, area, texto_nota)
        print(f"Expressão '{texto_nota}' processada e anotada.")

# Função principal para processar o PDF já aberto
def processar_pdf_aberto(pdf_documento):
    if not pdf_documento:
        print("Erro: O PDF não foi carregado corretamente.")
        return
    
    # Verificação adicional para garantir que o PDF foi carregado
    if pdf_documento.page_count == 0:
        print("Erro: O PDF parece estar vazio ou não foi carregado corretamente.")
        return

    # Contadores para numerar as expressões encontradas baseadas na primeira palavra
    contadores = {
        'TST': 0,
        'Acórdão': 0,
        'Sentença': 0
    }

    # Para armazenar se há mais de uma ocorrência de cada expressão realçada
    multiplas_ocorrencias = {
        'TST': False,
        'Acórdão': False,
        'Sentença': False
    }

    # Armazenar o número de ocorrências realçadas por expressão (baseado na primeira palavra)
    ocorrencias_realcadas = {
        'TST': 0,
        'Acórdão': 0,
        'Sentença': 0
    }

    # Processar cada página do PDF
    for pagina_numero in range(pdf_documento.page_count):
        pagina = pdf_documento.load_page(pagina_numero)
        texto_completo = pagina.get_text("text")
        print(f"Processando a página {pagina_numero + 1}/{pdf_documento.page_count}.")

        # Armazenar as áreas de expressões processadas por linha e as últimas ocorrências
        areas_processadas = []
        ultima_ocorrencia_expressao = {}

        # Procurar as expressões na página (sensível à caixa)
        for expressao in expressoes:
            # Procurar a expressão no texto para verificar a posição
            for ocorrencia in re.finditer(expressao, texto_completo):  # Sensível à caixa
                inicio_expressao, fim_expressao = ocorrencia.span()

                # Verificar se a expressão está ao lado de um horário com exatamente 1 espaço entre eles
                if expressao_ao_lado_de_horario_sem_texto_visivel(texto_completo, inicio_expressao):
                    areas_expressao = pagina.search_for(expressao)  # Realçar a área
                    for area in areas_expressao:
                        # Verificar se a área já foi processada nesta linha
                        if na_mesma_linha(area, areas_processadas):
                            continue  # Pula se já processamos uma expressão na mesma linha
                        
                        # Marcar a área como processada
                        areas_processadas.append(area)

                        # Atualizar a última ocorrência da expressão
                        ultima_ocorrencia_expressao[expressao] = area

                        # Incrementar o contador de ocorrências realçadas para a primeira palavra
                        primeira_palavra = expressao.split()[0]
                        ocorrencias_realcadas[primeira_palavra] += 1

        # Verificar se a primeira palavra da expressão tem múltiplas ocorrências realçadas
        for expressao in ocorrencias_realcadas:
            if ocorrencias_realcadas[expressao] > 1:
                multiplas_ocorrencias[expressao] = True

        # Processar a última ocorrência de cada expressão após verificar todas as linhas
        processar_ultima_ocorrencia_por_expressao(pagina, ultima_ocorrencia_expressao, contadores, multiplas_ocorrencias)

# Para salvar o arquivo após o processamento
def salvar_pdf(pdf_documento, caminho_saida):
    try:
        # Verificar se o arquivo já existe, se sim, remover o arquivo antes de salvar
        if os.path.exists(caminho_saida):
            os.remove(caminho_saida)  # Remove o arquivo existente

        pdf_documento.save(caminho_saida)
        pdf_documento.close()
        print(f"PDF salvo com sucesso em {caminho_saida}")
        abrir_pdf(caminho_saida)  # Abrir o arquivo após o salvamento
    except Exception as e:
        print(f"Erro ao salvar o PDF: {e}")

# Função para abrir o PDF após o salvamento
def abrir_pdf(caminho_saida):
    try:
        if platform.system() == 'Windows':
            os.startfile(caminho_saida)  # Windows
        elif platform.system() == 'Darwin':  # macOS
            subprocess.call(['open', caminho_saida])
        else:  # Linux (assumindo que xdg-open esteja disponível)
            subprocess.call(['xdg-open', caminho_saida])
        print(f"PDF aberto: {caminho_saida}")
    except Exception as e:
        print(f"Erro ao abrir o PDF: {e}")

# Exemplo de como usar o código
try:
    # Caminho para a pasta Downloads do usuário
    pasta_downloads = str(Path.home() / 'Downloads')

    # Encontrar o PDF mais recente na pasta Downloads
    arquivo_pdf_mais_recente = encontrar_pdf_mais_recente(pasta_downloads)
    print(f"Processando o arquivo: {arquivo_pdf_mais_recente}")

    # Abrir o PDF mais recente encontrado
    pdf_documento = fitz.open(str(arquivo_pdf_mais_recente))
    print("PDF carregado com sucesso.")
    
    # Extrair o nome base do arquivo PDF
    nome_arquivo_original = arquivo_pdf_mais_recente.stem  # Obtem o nome do arquivo sem a extensão
    caminho_saida = f'C:/Users/estevao.freixo/Downloads/{nome_arquivo_original}_editado.pdf'  # Adiciona '_editado'
                      
except Exception as e:
    print(f"Erro ao abrir o PDF: {e}")
    pdf_documento = None

if pdf_documento:
    processar_pdf_aberto(pdf_documento)
    salvar_pdf(pdf_documento, caminho_saida)  # Usa o caminho de saída gerado anteriormente
