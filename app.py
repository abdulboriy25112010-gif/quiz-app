from flask import Flask, request, jsonify, send_from_directory, session
import json
import os
import openpyxl
from openpyxl import load_workbook
from datetime import datetime

app = Flask(__name__, static_folder='static')
app.secret_key = 'your-secret-key-change-this-in-production'

RESULTS_FILE = 'results.xlsx'
ADMIN_PASSWORD = 'Behzod5664@'  # Change this!

QUESTIONS = [
    {"id": 1, "question": "What is the capital of France?", "options": ["Berlin", "Madrid", "Paris", "Rome"], "answer": "Paris"},
    {"id": 2, "question": "What is 5 + 7?", "options": ["10", "11", "12", "13"], "answer": "12"},
    {"id": 3, "question": "Which planet is closest to the Sun?", "options": ["Venus", "Mercury", "Earth", "Mars"], "answer": "Mercury"},
    {"id": 4, "question": "What color is the sky on a clear day?", "options": ["Green", "Red", "Blue", "Yellow"], "answer": "Blue"},
    {"id": 5, "question": "How many days are in a week?", "options": ["5", "6", "7", "8"], "answer": "7"},
    {"id": 6, "question": "What is the largest ocean?", "options": ["Atlantic", "Indian", "Arctic", "Pacific"], "answer": "Pacific"},
    {"id": 7, "question": "Which animal is known as the 'King of the Jungle'?", "options": ["Tiger", "Lion", "Elephant", "Bear"], "answer": "Lion"},
    {"id": 8, "question": "How many sides does a triangle have?", "options": ["2", "3", "4", "5"], "answer": "3"},
    {"id": 9, "question": "What is the boiling point of water in Celsius?", "options": ["90°C", "95°C", "100°C", "110°C"], "answer": "100°C"},
    {"id": 10, "question": "Which country has the largest population?", "options": ["USA", "India", "Russia", "China"], "answer": "India"},
    {"id": 11, "question": "What is H2O commonly known as?", "options": ["Salt", "Sugar", "Water", "Air"], "answer": "Water"},
    {"id": 12, "question": "How many months are in a year?", "options": ["10", "11", "12", "13"], "answer": "12"},
    {"id": 13, "question": "What is the fastest land animal?", "options": ["Lion", "Horse", "Cheetah", "Leopard"], "answer": "Cheetah"},
    {"id": 14, "question": "Which gas do plants absorb from the air?", "options": ["Oxygen", "Nitrogen", "CO2", "Hydrogen"], "answer": "CO2"},
    {"id": 15, "question": "What is the currency of the USA?", "options": ["Euro", "Pound", "Dollar", "Yen"], "answer": "Dollar"},
    {"id": 16, "question": "How many hours are in a day?", "options": ["20", "22", "24", "26"], "answer": "24"},
    {"id": 17, "question": "What is the smallest continent?", "options": ["Europe", "Australia", "Antarctica", "South America"], "answer": "Australia"},
    {"id": 18, "question": "Which organ pumps blood in the human body?", "options": ["Liver", "Kidney", "Lung", "Heart"], "answer": "Heart"},
    {"id": 19, "question": "What is 10 x 10?", "options": ["10", "100", "1000", "110"], "answer": "100"},
    {"id": 20, "question": "How many letters are in the English alphabet?", "options": ["24", "25", "26", "27"], "answer": "26"},
    {"id": 21, "question": "What is the opposite of 'hot'?", "options": ["Warm", "Cold", "Cool", "Mild"], "answer": "Cold"},
    {"id": 22, "question": "Which is the largest planet in our solar system?", "options": ["Saturn", "Neptune", "Jupiter", "Uranus"], "answer": "Jupiter"},
    {"id": 23, "question": "What do bees produce?", "options": ["Milk", "Honey", "Wax only", "Nectar"], "answer": "Honey"},
    {"id": 24, "question": "How many continents are there on Earth?", "options": ["5", "6", "7", "8"], "answer": "7"},
    {"id": 25, "question": "What is the chemical symbol for Gold?", "options": ["Go", "Gd", "Au", "Ag"], "answer": "Au"},
    {"id": 26, "question": "Which sense is used to hear?", "options": ["Sight", "Touch", "Taste", "Hearing"], "answer": "Hearing"},
    {"id": 27, "question": "What is the square root of 64?", "options": ["6", "7", "8", "9"], "answer": "8"},
    {"id": 28, "question": "Which country invented pizza?", "options": ["France", "Spain", "Italy", "Greece"], "answer": "Italy"},
    {"id": 29, "question": "What is the main language spoken in Brazil?", "options": ["Spanish", "English", "Portuguese", "French"], "answer": "Portuguese"},
    {"id": 30, "question": "How many zeros are in one million?", "options": ["5", "6", "7", "8"], "answer": "6"},
]

def get_completed_users():
    if not os.path.exists(RESULTS_FILE):
        return set()
    wb = load_workbook(RESULTS_FILE)
    ws = wb.active
    completed = set()
    for row in ws.iter_rows(min_row=2, values_only=True):
        if row[0]:
            completed.add(row[0].lower().strip())
    return completed

def save_result(email, name, answers, score, total_time):
    if not os.path.exists(RESULTS_FILE):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Quiz Results"
        ws.append(["Email", "Full Name", "Score", "Total Questions", "Percentage", "Time Taken (s)", "Date", "Answers"])
    else:
        wb = load_workbook(RESULTS_FILE)
        ws = wb.active

    percentage = round((score / 30) * 100, 1)
    date_str = datetime.now().strftime("%Y-%m-%d %H:%M")
    answers_str = json.dumps(answers)
    ws.append([email, name, score, 30, f"{percentage}%", total_time, date_str, answers_str])
    wb.save(RESULTS_FILE)

@app.route('/')
def index():
    return send_from_directory('static', 'index.html')

@app.route('/api/start', methods=['POST'])
def start_quiz():
    data = request.json
    email = data.get('email', '').lower().strip()
    name = data.get('name', '').strip()
    if not email or not name:
        return jsonify({'error': 'Email and name are required'}), 400
    completed = get_completed_users()
    if email in completed:
        return jsonify({'error': 'You have already completed this quiz. It can only be taken once.'}), 403
    session['email'] = email
    session['name'] = name
    return jsonify({'success': True, 'questions': QUESTIONS})

@app.route('/api/submit', methods=['POST'])
def submit_quiz():
    data = request.json
    email = session.get('email')
    name = session.get('name')
    if not email:
        return jsonify({'error': 'Session expired. Please restart.'}), 401
    completed = get_completed_users()
    if email in completed:
        return jsonify({'error': 'Already submitted.'}), 403
    answers = data.get('answers', {})
    total_time = data.get('totalTime', 0)
    score = 0
    result_details = {}
    for q in QUESTIONS:
        qid = str(q['id'])
        user_ans = answers.get(qid, '')
        correct = q['answer']
        is_correct = user_ans == correct
        if is_correct:
            score += 1
        result_details[qid] = {'userAnswer': user_ans, 'correct': correct, 'isCorrect': is_correct}
    save_result(email, name, result_details, score, total_time)
    session.clear()
    return jsonify({'success': True, 'score': score, 'total': 30, 'details': result_details})

@app.route('/api/admin/results')
def admin_results():
    password = request.args.get('password', '')
    if password != ADMIN_PASSWORD:
        return jsonify({'error': 'Unauthorized'}), 401
    if not os.path.exists(RESULTS_FILE):
        return jsonify({'results': []})
    wb = load_workbook(RESULTS_FILE)
    ws = wb.active
    headers = [cell.value for cell in ws[1]]
    results = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        results.append(dict(zip(headers, row)))
    return jsonify({'results': results})

@app.route('/api/admin/download')
def download_excel():
    password = request.args.get('password', '')
    if password != ADMIN_PASSWORD:
        return jsonify({'error': 'Unauthorized'}), 401
    if not os.path.exists(RESULTS_FILE):
        return jsonify({'error': 'No results yet'}), 404
    return send_from_directory('.', RESULTS_FILE, as_attachment=True)

if __name__ == '__main__':
    os.makedirs('static', exist_ok=True)
    app.run(debug=True)
