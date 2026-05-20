# Sistema de Control de Acceso — Reconocimiento Facial

Aplicación web desarrollada con **Flask** y **OpenCV** para controlar el acceso de personas mediante reconocimiento facial. Permite registrar rostros, escanearlos en tiempo real desde el navegador, y registrar automáticamente entradas y salidas con exportación a Excel.

## Demo en vivo

**[Ver demo → opencv-flask-xn43a3ze3a-uc.a.run.app/escanear-rostro](https://opencv-flask-xn43a3ze3a-uc.a.run.app/escanear-rostro)**

> Desplegado en Google Cloud Run — proyecto `tylerrv25@gmail.com`

## Funcionalidades

- **Registro de rostros** — Captura y almacena el rostro de una persona usando la cámara del navegador
- **Escaneo y reconocimiento** — Detecta el rostro en tiempo real y lo compara con los registrados
- **Control INGRESO / SALIDA** — Alterna automáticamente entre entrada y salida por persona
- **Historial de accesos** — Vista web con todos los eventos ordenados por fecha
- **Exportación a Excel** — Registro persistente en `.xlsx` con tablas formateadas y backup en `.txt`
- **Seguridad por IP** — Las rutas de escaneo y registro solo son accesibles desde la red local

## Tecnologías

| Área | Tecnología |
|---|---|
| Backend | Python, Flask |
| Visión computacional | OpenCV (Haar Cascade + comparación de histogramas) |
| Manejo de datos | openpyxl, pandas, numpy |
| Frontend | HTML, Jinja2 |
| Autenticación | Flask-Bcrypt |
| Variables de entorno | python-dotenv |

## Estructura del proyecto

```
OpenCV/
├── flask_app/
│   ├── controllers/
│   │   └── root_controller.py   # Rutas y lógica de control
│   ├── utils/
│   │   └── file_manager.py      # Reconocimiento facial, Excel, registro
│   ├── templates/root/
│   │   ├── escanear_rostro.html     # Vista de escaneo (cámara)
│   │   ├── registrar_rostro.html    # Vista de registro de persona
│   │   └── historial_accesos.html   # Vista de historial
│   ├── data/                    # Generado en runtime (ignorado por git)
│   │   ├── faces/               # Imágenes de rostros registrados
│   │   └── registro_accesos.xlsx
│   └── __init__.py
├── server.py                    # Punto de entrada
├── Pipfile                      # Dependencias
└── .env.example                 # Plantilla de variables de entorno
```

## Cómo funciona el reconocimiento

1. **Detección**: Se usa el clasificador **Haar Cascade** de OpenCV (`haarcascade_frontalface_default.xml`) para detectar rostros en el frame
2. **Normalización**: El rostro detectado se convierte a escala de grises y se redimensiona a 200×200 px
3. **Comparación**: Se compara el histograma del rostro escaneado contra todos los registrados usando `cv2.compareHist` con correlación
4. **Reconocimiento**: Si el score supera **0.72**, se considera una coincidencia válida

## Rutas disponibles

| Ruta | Acceso | Descripción |
|---|---|---|
| `GET /` | Local → escáner / Red → historial | Redirige según origen de la petición |
| `GET /escanear-rostro` | Solo local | Vista con cámara para escanear |
| `GET /registro-rostro` | Solo local | Formulario para registrar persona |
| `GET /registros-rostro` | Todos | Historial de accesos |
| `POST /api/registrar-rostro` | Solo local | API para guardar rostro |
| `POST /api/escanear-rostro` | Solo local | API para reconocer y registrar acceso |

## Instalación

```bash
# 1. Clonar el repositorio
git clone https://github.com/Tyler0125/opencv-control-acceso.git
cd opencv-control-acceso

# 2. Instalar dependencias
pipenv install

# 3. Configurar variables de entorno
cp .env.example .env
# Editar .env con tus valores

# 4. Ejecutar
pipenv run python server.py
```

La aplicación estará disponible en `http://localhost:5000`

## Requisitos

- Python 3.13
- Cámara web (para escaneo desde el navegador)
- Pipenv

## Variables de entorno

```env
APP_SECRET_KEY=clave-secreta-larga
DB_NAME=nombre_base_de_datos
DB_HOSTNAME=localhost
DB_USERNAME=root
DB_PASSWORD=contraseña
```

## Autor

**Tyler Rojas** — [GitHub](https://github.com/Tyler0125)
