import os
from threading import Thread
from flask import Flask, render_template, request, redirect, url_for, flash, session, send_from_directory
from openpyxl import load_workbook, Workbook
from datetime import datetime
from kivy.app import App
from kivy.utils import platform

# --- CONFIGURAÇÃO DE CAMINHOS PARA ANDROID ---
if platform == 'android':
    from android.storage import app_storage_path
    BASE_DIR = app_storage_path()  # Pasta segura para salvar o Excel e Fotos
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

ARQUIVO = os.path.join(BASE_DIR, "base_dados.xlsx")
UPLOAD = os.path.join(BASE_DIR, "uploads")

if not os.path.exists(UPLOAD):
    os.makedirs(UPLOAD)

# --- SEU CÓDIGO FLASK (RESUMIDO) ---
app = Flask(__name__)
app.secret_key = "chave_seguranca_inspecao_pro"

# ... (Mantenha todas as suas rotas: @app.route("/") etc.) ...
# Certifique-se de que todas as funções usem a variável ARQUIVO e UPLOAD globais.

# --- PARTE DO KIVY PARA ABRIR A WEBVIEW ---
# Usaremos a biblioteca jnius para abrir o navegador nativo do Android dentro do App
from kivy.uix.modalview import ModalView
from kivy.clock import Clock

class FlaskApp(App):
    def build(self):
        # 1. Inicia o Flask em uma thread separada
        flask_thread = Thread(target=self.run_flask)
        flask_thread.daemon = True
        flask_thread.start()

        # 2. Aguarda um segundo e abre a WebView
        # No Android real, usamos o componente nativo via pyjnius ou kivy-garden.webview
        from webbrowser import open as open_browser
        Clock.schedule_once(lambda dt: open_browser("http://127.0.0.1:5000"), 2)
        
        return ModalView() # Retorna uma tela vazia (o navegador ficará por cima)

    def run_flask(self):
        # Garante que o arquivo excel existe antes de iniciar
        # (Chame sua função verificar_arquivo_excel aqui)
        app.run(host="127.0.0.1", port=5000, debug=False)

if __name__ == "__main__":
    FlaskApp().run()