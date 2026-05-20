# Importar la clase Flask del módulo flask
from flask import Flask
# Importar el Blueprint 'root_bp' desde el módulo root_controller
from flask_app.controllers.root_controller import root_bp
# Importar el módulo os para manejar rutas de archivos
import os

# Crear una instancia de la aplicación Flask
# Se especifica la ruta a la carpeta de plantillas para que Jinja2 pueda encontrarlas.
app = Flask(__name__, template_folder='flask_app/templates')

# Configurar una clave secreta para la aplicación.
# ¡IMPORTANTE!: En un entorno de producción, esta clave debe ser una cadena aleatoria y compleja.
# Es utilizada por Flask para proteger las sesiones del lado del cliente y los mensajes flash.
app.secret_key = 'super_secret_key_para_proyectoradios' 

# Registrar el Blueprint 'root_bp' con la aplicación Flask.
# Esto asocia las rutas definidas en root_controller.py con la aplicación principal.
app.register_blueprint(root_bp)

# Bloque para ejecutar la aplicación solo cuando el script se ejecuta directamente
if __name__ == '__main__':
    # Ejecutar la aplicación Flask.
    # debug=True activa el modo de depuración.
    # host='0.0.0.0' hace que el servidor sea accesible desde cualquier IP en tu red local.
    port = int(os.getenv('PORT', '5050'))
    app.run(debug=True, use_reloader=False, threaded=True, host='0.0.0.0', port=port)