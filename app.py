from flask import Flask, render_template, request, jsonify, send_file
from dotenv import load_dotenv
from flask_cors import CORS
import cv2
import numpy as np
import face_recognition
import os
import base64
from io import BytesIO
from PIL import Image
import pickle
import openpyxl
import pandas as pd
import io
from datetime import datetime
import pytz
load_dotenv()
app = Flask(__name__)

CORS(app)
path = 'persons'
classNames = os.listdir(path)
def loadFaceEncodings():
    with open('face_encodings.pickle', 'rb') as file:
        encodeListKnown = pickle.load(file)
    return encodeListKnown
encodeListKnown = loadFaceEncodings()
def load_excel_file(file_name):
    try:
        return openpyxl.load_workbook(file_name)
    except FileNotFoundError:
        print(f"Error: File '{file_name}' not found.")
        return None
def find_name(sheet, search_name):
    column_to_search = sheet['A']
    for cell in column_to_search:
        if cell.value == search_name:
            return cell.row
    return None
def generate_search_name():
    try:
        with open('received_name.txt', 'r') as file:
            received_name = file.read().strip()
        return received_name
    except FileNotFoundError:
        print("Error: File 'received_name.txt' not found.")
        return None
def update_bonus(file_name):
    search_name= generate_search_name()
    try:
        workbook=load_excel_file(file_name)
        sheet=workbook.active
        row = find_name(sheet, search_name)
        if row is not None:
            current_value = sheet[f'B{row}'].value
            current_value = 0 if current_value is None else current_value
            sheet[f'B{row}'] = current_value + 1
            update_time(sheet, row, "Bonus")
            workbook.save(file_name)
        else:
            print(f"Name '{search_name}' not found in Column A.")
    except Exception as e:
        print(f"Error updating bonus: {e}")
def update_minus(file_name):
    search_name=generate_search_name()
    try:
        workbook=load_excel_file(file_name)
        sheet=workbook.active
        row = find_name(sheet, search_name)
        if row is not None:
            current_value = sheet[f'B{row}'].value
            current_value = 0 if current_value is None else current_value

            sheet[f'B{row}'] = current_value - 1
            update_time(sheet, row, "Minus")
            workbook.save(file_name)
        else:
            print(f"Name '{search_name}' not found in Column A.")
    except Exception as e:
        print(f"Error updating minus: {e}")
def update_alert(file_name):
    search_name = generate_search_name()
    try:
        workbook = load_excel_file(file_name)
        sheet = workbook.active
        row = find_name(sheet, search_name)
        if row is not None:
            sheet[f'C{row}'] = (sheet[f'C{row}'].value or 0) + 1
            update_time(sheet, row, "Alert")
            workbook.save(file_name)
        else:
            print(f"Name '{search_name}' not found in Column A.")
    except Exception as e:
        print(f"Error updating alert: {e}")
def add_comment(file_name, comment):
    search_name=generate_search_name()
    try:
        workbook=load_excel_file(file_name)
        sheet=workbook.active
        row = find_name(sheet, search_name)
        if row is not None:
            existing_comment = sheet[f'D{row}'].value
            sheet[f'D{row}'] = f"{existing_comment} (({comment})),"
            update_time(sheet, row, "Comment")
            workbook.save(file_name)
        else:
            print(f"Name '{search_name}' not found in Column A.")
    except Exception as e:
        print(f"Error adding comment: {e}")
def update_time(sheet, row, action):
    cairo_timezone = pytz.timezone('Africa/Cairo')
    cairo_time = datetime.now(cairo_timezone).strftime("%Y-%m-%d %I:%M %p")
    existing_date = sheet[f'E{row}'].value
    sheet[f'E{row}'] = f"{existing_date} (({cairo_time}){action}),"
@app.route('/')
def index():
    return render_template('index.html')
@app.route('/indexar')
def indexar():
    return render_template('indexar.html')
@app.route('/teacher')
def apps():
    return render_template('teacher.html')
@app.route('/recognize', methods=['POST'])
def recognize():
    data = request.json
    image_data = data.get('image')
    if image_data:
        image = Image.open(BytesIO(base64.b64decode(image_data.split(',')[1])))
        img_np = np.array(image)
        img_rgb = cv2.cvtColor(img_np, cv2.COLOR_BGR2RGB)
        face_locations = face_recognition.face_locations(img_rgb)
        face_encodings = face_recognition.face_encodings(img_rgb, face_locations)
        result = []

        for encodeface, faceLoc in zip(face_encodings, face_locations):
            matches = face_recognition.compare_faces(encodeListKnown, encodeface)
            faceDis = face_recognition.face_distance(encodeListKnown, encodeface)
            matchIndex = np.argmin(faceDis)

            if matches[matchIndex]:
                name = classNames[matchIndex]
                result.append({'name': name})

        return jsonify(result)
@app.route('/send-pop', methods=['POST'])
def receive_name():
    try:
        data = request.get_json()
        received_name = data.get('name', '')
        new_file_name = f'received_name.txt'
        with open(new_file_name, 'w') as file:
            file.write(received_name)
        return jsonify({'status': 'success', 'message': 'Name received and processed'})
    except Exception as e:
        return jsonify({'status': 'error', 'message': f"Error: {e}"})
@app.route('/download_original_excel')
def download_original_excel():
    return send_file('dgrees_sheet.xlsx', as_attachment=True, download_name='dgrees_sheet.xlsx')
@app.route('/get_excel_data')
def get_excel_data():
    df = pd.read_excel('dgrees_sheet.xlsx')
    excel_data = io.BytesIO()
    df.to_excel(excel_data, index=False, engine='openpyxl')
    excel_data.seek(0)
    return send_file(excel_data, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', as_attachment=True, download_name='output.xlsx')
@app.route('/update_bonus', methods=['POST'])
def update_bonus_route():
    try:
        update_bonus('dgrees_sheet.xlsx')
        return jsonify({'status': 'success', 'message': 'Bonus updated'})
    except Exception as e:
        return jsonify({'status': 'error', 'message': f"Error: {e}"})
@app.route('/update_minus', methods=['POST'])
def update_minus_route():
    try:
        update_minus('dgrees_sheet.xlsx')
        return jsonify({'status': 'success', 'message': 'Minus updated'})
    except Exception as e:
        return jsonify({'status': 'error', 'message': f"Error: {e}"})
@app.route('/update_alert', methods=['POST'])
def update_alert_route():
    try:
        update_alert('dgrees_sheet.xlsx')
        return jsonify({'status': 'success', 'message': 'Alert updated'})
    except Exception as e:
        return jsonify({'status': 'error', 'message': f"Error: {e}"})
@app.route('/add_comment', methods=['POST'])
def add_comment_route():

    try:
        data = request.get_json()
        comment = data.get('comment', '')
        add_comment('dgrees_sheet.xlsx', comment)
        return jsonify({'status': 'success', 'message': 'Comment added'})
    except Exception as e:
        return jsonify({'status': 'error', 'message': f"Error: {e}"})
if __name__ == '__main__':
    app.run()