from flask import Blueprint, jsonify, render_template, request

from flask_app.utils.file_manager import (
    decode_base64_image,
    list_face_access_events,
    register_face_from_frame,
    scan_face_and_register_access,
)

root_bp = Blueprint('root', __name__)


def _is_local_request():
    return request.remote_addr in ('127.0.0.1', '::1')


@root_bp.route('/', methods=['GET'])
def index():
    if _is_local_request():
        return render_template('root/escanear_rostro.html')
    return registros_rostro()


@root_bp.route('/escanear-rostro', methods=['GET'])
def escanear_rostro():
    if not _is_local_request():
        return 'Acceso no permitido desde la red. Esta vista solo está habilitada localmente.', 403
    return render_template('root/escanear_rostro.html')


@root_bp.route('/registro-rostro', methods=['GET'])
def registro_rostro():
    if not _is_local_request():
        return 'Acceso no permitido desde la red. Esta vista solo está habilitada localmente.', 403
    return render_template('root/registrar_rostro.html')


@root_bp.route('/registros-rostro', methods=['GET'])
def registros_rostro():
    can_manage = _is_local_request()
    try:
        events = list_face_access_events()
        return render_template('root/historial_accesos.html', events=events, error_message=None, can_manage=can_manage)
    except PermissionError as e:
        return render_template('root/historial_accesos.html', events=[], error_message=str(e), can_manage=can_manage)


@root_bp.route('/api/registrar-rostro', methods=['POST'])
def api_registrar_rostro():
    if not _is_local_request():
        return jsonify({'ok': False, 'message': 'Acceso no permitido desde la red.'}), 403

    payload = request.get_json(silent=True) or {}
    nombre = (payload.get('nombre') or '').strip()
    image_data = payload.get('image_data')

    if not nombre:
        return jsonify({'ok': False, 'message': 'Debes ingresar un nombre.'}), 400

    frame = decode_base64_image(image_data)
    if frame is None:
        return jsonify({'ok': False, 'message': 'Imagen inválida o vacía.'}), 400

    try:
        ok, message = register_face_from_frame(nombre, frame)
        status_code = 200 if ok else 400
        return jsonify({'ok': ok, 'message': message}), status_code
    except PermissionError as e:
        return jsonify({'ok': False, 'message': str(e)}), 409


@root_bp.route('/api/escanear-rostro', methods=['POST'])
def api_escanear_rostro():
    if not _is_local_request():
        return jsonify({'ok': False, 'message': 'Acceso no permitido desde la red.'}), 403

    payload = request.get_json(silent=True) or {}
    image_data = payload.get('image_data')

    frame = decode_base64_image(image_data)
    if frame is None:
        return jsonify({'ok': False, 'message': 'Imagen inválida o vacía.'}), 400

    try:
        ok, message, event_data = scan_face_and_register_access(frame)
        status_code = 200 if ok else 404
        return jsonify({'ok': ok, 'message': message, 'event': event_data}), status_code
    except PermissionError as e:
        return jsonify({'ok': False, 'message': str(e), 'event': None}), 409
