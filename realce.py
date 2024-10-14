import pyperclip
import time
import subprocess
import os

def limpar_area_de_transferencia():
    # Limpa o conteúdo da área de transferência definindo como uma string vazia
    pyperclip.copy("")  # Define a área de transferência como vazia
    print("Área de transferência limpa. Copie um novo conteúdo agora.")

def esperar_por_copiar():
    print("Copie o conteúdo que deseja colar no Bloco de Notas.")
    print("O programa vai esperar até que você copie algo novo.")

    conteudo_anterior = ""

    while True:
        # Obtém o conteúdo atual da área de transferência
        conteudo_atual = pyperclip.paste()

        # Verifica se o conteúdo da área de transferência mudou e se não é igual ao anterior
        if conteudo_atual != conteudo_anterior and conteudo_atual.strip() != "":
            print("Conteúdo copiado!")
            return conteudo_atual  # Retorna o novo conteúdo
        time.sleep(1)  # Espera 1 segundo antes de verificar novamente

def salvar_no_bloco_de_notas(conteudo, primeiro_conteudo=False):
    # Salva o conteúdo em um arquivo temporário
    caminho_arquivo = "conteudo_temp.txt"
    
    # Se for o primeiro conteúdo, cria o arquivo e abre o Bloco de Notas
    if primeiro_conteudo:
        with open(caminho_arquivo, 'w', encoding='utf-8') as arquivo:  # 'w' para criar o arquivo
            arquivo.write(conteudo + "\n")  # Adiciona o conteúdo
        subprocess.Popen(['notepad.exe', caminho_arquivo])  # Abre o Bloco de Notas
    else:
        # Adiciona novo conteúdo ao arquivo
        with open(caminho_arquivo, 'a', encoding='utf-8') as arquivo:  # 'a' para anexar ao arquivo
            arquivo.write(conteudo + "\n")  # Adiciona uma nova linha após o conteúdo

if __name__ == "__main__":
    limpar_area_de_transferencia()  # Limpa a área de transferência
    
    primeiro_copiado = True  # Variável para controlar se é o primeiro conteúdo copiado

    while True:
        conteudo_copiado = esperar_por_copiar()  # Aguarda o usuário copiar o conteúdo
        salvar_no_bloco_de_notas(conteudo_copiado, primeiro_copiado)  # Salva no arquivo
        
        if primeiro_copiado:
            primeiro_copiado = False  # Depois da primeira cópia, muda para False
        
        print("Conteúdo salvo. Aguardando 3 segundos antes de esperar novo conteúdo...")
        time.sleep(3)  # Aguarda 3 segundos antes de permitir nova cópia
