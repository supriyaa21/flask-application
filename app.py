from flask import Flask, request, jsonify
from flask_cors import CORS
import os
import pandas as pd
import pypyodbc
from sqlalchemy import create_engine
from sqlalchemy.engine import URL
from mangum import Mangum

app = Flask(__name__, static_folder='static', static_url_path='/')
CORS(app, resources={r"/uploads": {"origins": "*"}}, methods=["POST"])

app.config['UPLOAD_FOLDER'] = 'uploads'
SERVER_NAME = '192.168.125.250\\SQL2017'
DATABASE_NAME = 'pallet1db'
TABLE_NAME = 'demo_table'
SQL_USERNAME = 'sa'
SQL_PASSWORD = 'Triune@123'

connection_string = f"""
    DRIVER={{SQL Server}};
    SERVER={SERVER_NAME};
    DATABASE={DATABASE_NAME};
    UID={SQL_USERNAME};
    PWD={SQL_PASSWORD};
"""
connection_url = URL.create('mssql+pyodbc', query={'odbc_connect': connection_string})
engine = create_engine(connection_url, module=pypyodbc)

if not os.path.exists(app.config['UPLOAD_FOLDER']):
    os.makedirs(app.config['UPLOAD_FOLDER'])

@app.route('/')
def index():
    return app.send_static_file('index.html')

@app.route('/uploads', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return jsonify({'message': 'No file part'}), 400

    file = request.files['file']
    if file.filename == '':
        return jsonify({'message': 'No selected file'}), 400

    file_path = os.path.join(app.config['UPLOAD_FOLDER'], file.filename)
    file.save(file_path)

    try:
        excel_data = pd.read_excel(file_path, sheet_name=None)
        for sheet_name, df_data in excel_data.items():
            df_data.to_sql(TABLE_NAME, engine, if_exists='append', index=False)
        return jsonify({'message': 'Data loaded successfully into SQL Server.'}), 200
    except Exception as e:
        return jsonify({'message': str(e)}), 500

# This part is needed to wrap the Flask app as a serverless function


if __name__ == '__main__':
    from waitress import serve
    serve(app,host='0.0.0.0',port=5000)