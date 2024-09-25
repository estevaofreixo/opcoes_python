from flask import Flask, render_template, redirect, url_for
import subprocess

app = Flask(__name__)

# Função para executar o script Python
def run_script(script_name):
    try:
        # Chama o script Python usando subprocess
        subprocess.run(["python", script_name], check=True)
        return True
    except Exception as e:
        print(f"Erro ao executar {script_name}: {e}")
        return False

# Página inicial com os botões
@app.route('/')
def index():
    return render_template('index.html')

# Rota para executar o primeiro script
@app.route('/run-script1')
def run_script1():
    run_script("convers.py")
    return redirect(url_for('index'))

# Rota para executar o segundo script
@app.route('/run-script2')
def run_script2():
    run_script("adesivo7.py")
    return redirect(url_for('index'))


if __name__ == '__main__':
    app.run(debug=True)
