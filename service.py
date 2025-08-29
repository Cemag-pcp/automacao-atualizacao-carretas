from flask import Flask

import os
import time
from datetime import date, timedelta, datetime
from zoneinfo import ZoneInfo

# 2. Bibliotecas externas
from flask import Flask, render_template, jsonify
import schedule
import threading
from api import main

app = Flask(__name__)

@app.route("/")
def home():
    return render_template('home.html')

@app.route("/executar/",methods=['POST'])
def executar_auto():
    try:
        main()
    except Exception as e:
        print('Erro: ',e)
        return jsonify({'status':'erro'},400)
    return jsonify({'status':'ok'},200)

def agendar_atualizacao():
    print('agendar_atualizacao()')
    schedule.every().day.at("17:16").do(main)
    while True:
        jobs = schedule.get_jobs()  # Retorna a lista de jobs pendentes
        schedule.run_pending()
        time.sleep(5)
        if datetime.now().hour == 16 and datetime.now().minute >= 5:
            print(jobs)

# Inicia o agendamento em uma thread separada
def start_thread():
    print('start_thread()')
    agendamento_thread = threading.Thread(target=agendar_atualizacao)
    agendamento_thread.daemon = True # Isso permite que a thread seja finalizada quando o programa principal for finalizado
    agendamento_thread.start()

# Função para iniciar a thread de agendamento após o app estar configurado
def init_agendamento():
    print('init_agendamento()')
    with app.app_context():
        start_thread()  # Inicia a thread de agendamento

# Executa o agendamento quando o app for iniciado
init_agendamento()

# Executando a aplicação
if __name__ == '__main__':
    app.run(debug=True)