import pyperclip
import time
import os
import subprocess

def fechar_outros_blocos_de_notas():
    # Lista todos os processos em execução
    processos = subprocess.check_output(["tasklist"]).decode("latin-1")  # Use 'latin-1' em vez de 'utf-8'
    # Nome do processo do Bloco de Notas
    nome_processo = "notepad.exe"

    # Fecha todos os processos do Bloco de Notas
    for linha in processos.splitlines():
        if nome_processo in linha:
            # Tenta fechar o processo
            try:
                pid = int(linha.split()[1])  # Obtém o PID do processo
                subprocess.run(["taskkill", "/F", "/PID", str(pid)])  # Força o fechamento
            except Exception as e:
                print(f"Erro ao fechar o Bloco de Notas: {e}")

def monitorar_area_de_transferencia():
    # Limpa o conteúdo inicial da área de transferência
    pyperclip.copy("")

    # Nome do arquivo onde o conteúdo será salvo
    nome_arquivo = "copiado.txt"

    # Esvazia o arquivo para garantir que comece em branco
    with open(nome_arquivo, "w", encoding="utf-8") as arquivo:
        pass  # Apenas abre e fecha o arquivo para limpá-lo

    print("Monitorando a área de transferência. Copie algum texto para adicionar ao arquivo.")

    # Inicializa a variável para o conteúdo anterior da área de transferência
    conteudo_anterior = ""

    while True:
        try:
            # Obtém o conteúdo atual da área de transferência
            conteudo_atual = pyperclip.paste()

            # Se o conteúdo atual for diferente do anterior e não estiver vazio
            if conteudo_atual != conteudo_anterior and conteudo_atual:
                # Atualiza o conteúdo anterior
                conteudo_anterior = conteudo_atual

                # Escreve uma linha em branco antes do novo conteúdo
                with open(nome_arquivo, "a", encoding="utf-8") as arquivo:
                    arquivo.write("\n")
                    arquivo.write(conteudo_atual + "\n")
                    arquivo.flush()

                # Fecha outros arquivos de bloco de notas
                fechar_outros_blocos_de_notas()

                # Reabre o arquivo para o usuário ver o conteúdo atualizado
                os.startfile(nome_arquivo)

                print(f"Conteúdo copiado: {conteudo_atual}")

        except Exception as e:
            print(f"Ocorreu um erro: {e}")

        # Espera um segundo antes de verificar novamente
        time.sleep(1)

if __name__ == "__main__":
    monitorar_area_de_transferencia()
