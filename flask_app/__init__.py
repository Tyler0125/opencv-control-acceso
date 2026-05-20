# flask_app/__init__.py

from flask import Flask

# Inicializa la aplicación Flask
app = Flask(__name__)

# --- AÑADE ESTA LÍNEA ---
# Es CRUCIAL para las sesiones y los mensajes flash.
# En producción, usa una clave muy larga y generada aleatoriamente.
app.config['SECRET_KEY'] = 'una_clave_secreta_generada_aleatoriamente_y_unica_para_tu_aplicacion'
# Puedes generar una con: import os; os.urandom(24)

# Importa los controladores para registrar las rutas
from flask_app.controllers import root_controller

# --- AÑADE ESTA LÍNEA CRUCIAL PARA REGISTRAR EL BLUEPRINT ---
app.register_blueprint(root_controller.root_bp)

# Si tienes otros controladores (ej. auth_controller.py), impórtalos también:
# from flask_app.controllers import auth_controller
