import pyperclip
import time
import os

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

    # Abre o arquivo no modo append para mantê-lo aberto durante a execução
    with open(nome_arquivo, "a", encoding="utf-8") as arquivo:
        while True:
            try:
                # Obtém o conteúdo atual da área de transferência
                conteudo_atual = pyperclip.paste()

                # Se o conteúdo atual for diferente do anterior e não estiver vazio
                if conteudo_atual != conteudo_anterior and conteudo_atual:
                    # Atualiza o conteúdo anterior
                    conteudo_anterior = conteudo_atual

                    # Escreve uma linha em branco antes do novo conteúdo
                    arquivo.write("\n")
                    # Escreve o novo conteúdo no arquivo
                    arquivo.write(conteudo_atual + "\n")
                    # Descarrega o buffer para garantir que o conteúdo seja salvo
                    arquivo.flush()

                    # Reabre o arquivo para o usuário ver o conteúdo atualizado
                    os.startfile(nome_arquivo)

                    print(f"Conteúdo copiado: {conteudo_atual}")

            except Exception as e:
                print(f"Ocorreu um erro: {e}")

            # Espera um segundo antes de verificar novamente
            time.sleep(1.5)

if __name__ == "__main__":
    monitorar_area_de_transferencia()
