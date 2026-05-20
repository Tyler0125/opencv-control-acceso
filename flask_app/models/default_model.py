from dataclasses import dataclass, fields
from datetime import datetime
from typing import ClassVar
from flask_app.config.mysqlconnection import connectToMySQL

from dotenv import load_dotenv
import os

load_dotenv()

@dataclass
class default_model:
    table_name: ClassVar[str] = None

    id: int
    created_at: datetime
    updated_at: datetime

    def __init__(self,data):
    
        for field in fields(self):
            setattr(self, field.name, data.get(field.name))

    # Método para obtener todos los registros de la tabla
    @classmethod
    def all(cls):
        query = f'SELECT * FROM {cls.table_name}'
        results = connectToMySQL(os.getenv('DB_NAME')).query_db(query)
        instances = []
        for result in results:
            instances.append(cls(result))
        return instances
    
    # Método para obtener todos los registros de la tabla con un query parametrizado
    @classmethod
    def all_by_query(cls,query):
        results = connectToMySQL(os.getenv('DB_NAME')).query_db(query)
        instances = []
        for result in results:
            instances.append(cls(result))
        return instances

    # Método para obtener el registros de la tabla segun el id
    @classmethod
    def find_by_id(cls, id:int):
        try:
            query = f'SELECT * from {cls.table_name} WHERE id=%(id)s'
            result = connectToMySQL(os.getenv('DB_NAME')).query_db(query, {'id': id})
            if not result:
                return None
            return cls(result[0])
        except Exception as e:
            print(e)
            return None
        
    # Método para insertar registros en la tabla
    @classmethod
    def save(cls, data):
        query = f'INSERT INTO {cls.table_name} ({cls.get_fields()}) VALUES ({cls.get_data_values()})'
        return connectToMySQL(os.getenv('DB_NAME')).query_db(query, data)
    
    # Método para actualizar registros en la tabla por id
    @classmethod
    def update_by_id(cls,data):
        query = f'UPDATE {cls.table_name} SET {cls.string_update()} WHERE id = %(id)s'
        return connectToMySQL(os.getenv('DB_NAME')).query_db(query, data)
    
    # Método para actualizar registros en la tabla 
    @classmethod
    def update(cls, data):
        query = f'UPDATE {cls.table_name} ({cls.get_fields()}) VALUES ({cls.get_data_values()})'
        return connectToMySQL(os.getenv('DB_NAME')).query_db(query, data)

    # Método para eliminar un registro de la tabla
    @classmethod
    def delete_by_id(cls, id:int):
        # Crea el query
        query = f'DELETE FROM {cls.table_name} WHERE id = %(id)s'
        return connectToMySQL(os.getenv('DB_NAME')).query_db(query, {'id': id})

    # Método para buscar registros por columna con query parametrizado
    @classmethod
    def find_one_by_field(cls, field_name: str, value):
        if not field_name:
            return None
        query = f'SELECT * FROM {cls.table_name} WHERE {field_name} = %(value)s LIMIT 1'
        result = connectToMySQL(os.getenv('DB_NAME')).query_db(query, {'value': value})
        if not result:
            return None
        return cls(result[0])

    # Método para registrar eventos de acceso (requiere tabla compatible en DB)
    @classmethod
    def save_face_access_event(cls, person_name: str, access_type: str, event_time: datetime = None):
        if event_time is None:
            event_time = datetime.now()
        query = (
            f'INSERT INTO {cls.table_name} (person_name, access_type, event_time) '
            f'VALUES (%(person_name)s, %(access_type)s, %(event_time)s)'
        )
        data = {
            'person_name': person_name,
            'access_type': access_type,
            'event_time': event_time.strftime('%Y-%m-%d %H:%M:%S')
        }
        return connectToMySQL(os.getenv('DB_NAME')).query_db(query, data)

    # Método para ver todos los campos
    @classmethod
    def data_fields(cls):
        class_data_fields = []
        for field in fields(cls):
            if (field.name != 'id' and field.name != 'created_at' and field.name != 'updated_at'):
                class_data_fields.append(field.name)
        return class_data_fields
    # Método para obtener todos los campos separados por coma
    @classmethod
    def get_fields(cls):
        return ', '.join(cls.data_fields())
    # Método para ver los valores de  todos los campos separados por coma
    @classmethod
    def get_data_values(cls):
        return ', '.join([f'%({field})s' for field in cls.data_fields()])
    # Método para ver los campos y sus valores en formato 'campo=valor' separados por coma
    @classmethod
    def string_update(cls):
        return ', '.join([f'{field} = %({field})s' for field in cls.data_fields()])
    
    # Método para correr y retornar una query parametrizada
    @staticmethod
    def run_query(query, data):
        return connectToMySQL(os.getenv('DB_NAME')).query_db(query, data)