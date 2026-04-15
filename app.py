from flask import Flask, request, jsonify, send_from_directory, session
import json
import os
import openpyxl
from openpyxl import load_workbook
from datetime import datetime

app = Flask(__name__, static_folder='static')
app.secret_key = 'your-secret-key-change-this-in-production'

RESULTS_FILE = 'results.xlsx'
ADMIN_PASSWORD = 'admin123'  # Change this!

QUESTIONS = [
    {"id": 1, "question": "Falsafa so'zi qaysi tildan olingan va uning ma'nosi nima?", "options": ["Lotin tilidan, «ilm-fan» ma'nosi", "Yunoncha «philos» va «sophia», «donishmandlikni sevish»", "Arabcha tilidan, «hikmat» ma'nosi", "Fors tilidan, «aql» ma'nosi"], "answer": "Yunoncha «philos» va «sophia», «donishmandlikni sevish»"},
    {"id": 2, "question": "Qadimgi Yunonistonda falsafaning asosiy muammosi nima edi?", "options": ["Davlat boshqaruvi muammosi", "Borliq va mavjudlikning mohiyati", "Iqtisodiy munosabatlar", "Harbiy strategiya"], "answer": "Borliq va mavjudlikning mohiyati"},
    {"id": 3, "question": "Ontologiya nimani o'rganadi?", "options": ["Bilish jarayonini", "Borliq va mavjudlikning asosiy tamoyillarini", "Axloq va odobni", "Jamiyat tuzilishini"], "answer": "Borliq va mavjudlikning asosiy tamoyillarini"},
    {"id": 4, "question": "Gnoseologiya falsafaning qaysi sohasiga kiradi?", "options": ["Borliq haqidagi ta'limotga", "Bilish nazariyasiga", "Axloq falsafasiga", "Estetikaga"], "answer": "Bilish nazariyasiga"},
    {"id": 5, "question": "Quyidagi faylasuflardan qaysi biri antik davr faylasufi hisoblanadi?", "options": ["Immanuel Kant", "René Descartes", "Sokrat", "Karl Marks"], "answer": "Sokrat"},
    {"id": 6, "question": "«Men faqat bitta narsani bilaman — hech narsa bilmasligimni» — bu fikr kimga tegishli?", "options": ["Platonga", "Aristotelga", "Sokratga", "Epikurga"], "answer": "Sokratga"},
    {"id": 7, "question": "Platon qaysi falsafiy tushunchani asosiy deb hisoblagan?", "options": ["Materiya", "G'oyalar (ideya) olami", "Atom", "Energiya"], "answer": "G'oyalar (ideya) olami"},
    {"id": 8, "question": "Aristotelning mantiq haqidagi asarlari to'plami qanday nomlanadi?", "options": ["«Metafizika»", "«Siyosat»", "«Organon»", "«Etika»"], "answer": "«Organon»"},
    {"id": 9, "question": "Materializm va idealizm o'rtasidagi asosiy farq nimada?", "options": ["Metodologiyada", "Birlamchi — materiya yoki ong ekanligini belgilashda", "Axloqiy qarashlarda", "Siyosiy pozitsiyalarda"], "answer": "Birlamchi — materiya yoki ong ekanligini belgilashda"},
    {"id": 10, "question": "Dialektika nima?", "options": ["Fizika qonunlari tizimi", "Qarama-qarshiliklar birligi orqali taraqqiyotni tushuntiruvchi ta'limot", "Mantiqiy xatolarni aniqlash usuli", "Tarixiy faktlarni o'rganish metodi"], "answer": "Qarama-qarshiliklar birligi orqali taraqqiyotni tushuntiruvchi ta'limot"},
    {"id": 11, "question": "Ibn Sino (Avitsenna) falsafasining asosiy yo'nalishi qaysi?", "options": ["Skeptitsizm", "Neoplatonizm va aristotelizm sintezi", "Nihilizm", "Pragmatizm"], "answer": "Neoplatonizm va aristotelizm sintezi"},
    {"id": 12, "question": "Al-Forobiy qaysi unvon bilan tanilgan?", "options": ["«Sharq faylasufi»", "«Ikkinchi muallim» (Muallim us-soniy)", "«Donishmandlar donishmandi»", "«Birinchi tabib»"], "answer": "«Ikkinchi muallim» (Muallim us-soniy)"},
    {"id": 13, "question": "Dekart falsafasining asosiy tamoyili qaysi?", "options": ["«Borliq — harakat»", "«Cogito ergo sum» — «Men o'ylayman, demak mavjudman»", "«Bilim — kuch»", "«Hamma narsa oqadi»"], "answer": "«Cogito ergo sum» — «Men o'ylayman, demak mavjudman»"},
    {"id": 14, "question": "Kantning bosh falsafiy asari qanday nomlanadi?", "options": ["«Aql haqida nutq»", "«Sof aqlning tanqidi»", "«Ijtimoiy shartnoma»", "«Leviafan»"], "answer": "«Sof aqlning tanqidi»"},
    {"id": 15, "question": "Gegelning dialektik metodidagi uchlik (triada) qanday tushunchalardan iborat?", "options": ["Materiya, harakat, fazo", "Tezis, antitezis, sintez", "Sabab, natija, maqsad", "Idea, ruh, tabiat"], "answer": "Tezis, antitezis, sintez"},
    {"id": 16, "question": "Marksizm falsafasining asosiy tamoyili nima?", "options": ["Idealistik dialektika", "Dialektik va tarixiy materializm", "Pragmatizm", "Ekzistentsializm"], "answer": "Dialektik va tarixiy materializm"},
    {"id": 17, "question": "Ekzistentsializm falsafasining asosiy muammosi nima?", "options": ["Tabiat qonunlari", "Insonning erkinligi, tanlovi va mas'uliyati", "Ijtimoiy adolat", "Bilishning chegaralari"], "answer": "Insonning erkinligi, tanlovi va mas'uliyati"},
    {"id": 18, "question": "«Inson — o'z mohiyatini o'zi yaratadi» — bu fikr qaysi yo'nalishga xos?", "options": ["Stoitsizm", "Ekzistentsializm (J.-P. Sartr)", "Pozitivizm", "Empirizm"], "answer": "Ekzistentsializm (J.-P. Sartr)"},
    {"id": 19, "question": "Pragmatizm falsafasiga ko'ra haqiqat nima?", "options": ["Ilohiy vahiy", "Amalda foydali va samarali bo'lgan narsa", "Mantiqiy xulosalarning to'g'riligi", "Sezgi organlarimiz bergan ma'lumot"], "answer": "Amalda foydali va samarali bo'lgan narsa"},
    {"id": 20, "question": "Empirizm ta'limotiga ko'ra bilimning asosiy manbai nima?", "options": ["Tug'ma g'oyalar", "Tajriba va sezgi", "Ilohiy aql", "Mantiqiy deduksiya"], "answer": "Tajriba va sezgi"},
    {"id": 21, "question": "Falsafada «agnostitsizm» nimani anglatadi?", "options": ["Xudoning mavjudligini inkor etish", "Dunyoni to'liq bilish mumkin emasligini ta'kidlash", "Barcha narsalarni bilish mumkinligiga ishonch", "Faqat moddiy narsalar mavjudligini ta'kidlash"], "answer": "Dunyoni to'liq bilish mumkin emasligini ta'kidlash"},
    {"id": 22, "question": "Axloq falsafasida «kategorik imperativ» tushunchasi kimga tegishli?", "options": ["Gegelga", "Kantga", "Nitsshega", "Shopengauerga"], "answer": "Kantga"},
    {"id": 23, "question": "Nitsshening «Xudo o'ldi» iborasi nimani anglatadi?", "options": ["Xristianlikning qulashi", "An'anaviy qadriyatlar va mutlaq haqiqatning inqirozi", "Ateizmning g'alabasini", "Insoniyatning halokati"], "answer": "An'anaviy qadriyatlar va mutlaq haqiqatning inqirozi"},
    {"id": 24, "question": "Falsafada «hermenevtika» nima?", "options": ["Arxeologik qazishmalar usuli", "Matnlarni tushunish va talqin qilish nazariyasi", "Riyoziyot tarmog'i", "Tabiiy hodisalarni o'rganish metodi"], "answer": "Matnlarni tushunish va talqin qilish nazariyasi"},
    {"id": 25, "question": "Zamonaviy falsafada «postmodernizm» qaysi g'oyani markazga qo'yadi?", "options": ["Yagona mutlaq haqiqat mavjudligini", "Haqiqat ko'pqirrali va nisbiy, «katta rivoyatlar»ni rad etish", "Ilm-fan cheksiz taraqqiyotini", "Insonning tabiat ustidan hukmronligini"], "answer": "Haqiqat ko'pqirrali va nisbiy, «katta rivoyatlar»ni rad etish"},
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
        ws.title = "Test natijalari"
        ws.append(["Email", "To'liq ism", "Ball", "Jami savollar", "Foiz", "Ketgan vaqt (s)", "Sana", "Javoblar"])
    else:
        wb = load_workbook(RESULTS_FILE)
        ws = wb.active
    percentage = round((score / 25) * 100, 1)
    date_str = datetime.now().strftime("%Y-%m-%d %H:%M")
    answers_str = json.dumps(answers, ensure_ascii=False)
    ws.append([email, name, score, 25, f"{percentage}%", total_time, date_str, answers_str])
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
        return jsonify({'error': 'Email va ism kiritilishi shart'}), 400
    completed = get_completed_users()
    if email in completed:
        return jsonify({'error': 'Siz bu testni allaqachon topshirgansiz. Test faqat bir marta topshirilishi mumkin.'}), 403
    session['email'] = email
    session['name'] = name
    return jsonify({'success': True, 'questions': QUESTIONS})

@app.route('/api/submit', methods=['POST'])
def submit_quiz():
    data = request.json
    email = session.get('email')
    name = session.get('name')
    if not email:
        return jsonify({'error': 'Sessiya tugadi. Qaytadan boshlang.'}), 401
    completed = get_completed_users()
    if email in completed:
        return jsonify({'error': 'Allaqachon topshirilgan.'}), 403
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
    return jsonify({'success': True, 'score': score, 'total': 25, 'details': result_details})

@app.route('/api/admin/results')
def admin_results():
    password = request.args.get('password', '')
    if password != ADMIN_PASSWORD:
        return jsonify({'error': 'Ruxsat yoq'}), 401
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
        return jsonify({'error': 'Ruxsat yoq'}), 401
    if not os.path.exists(RESULTS_FILE):
        return jsonify({'error': 'Hali natijalar yoq'}), 404
    return send_from_directory('.', RESULTS_FILE, as_attachment=True)

if __name__ == '__main__':
    os.makedirs('static', exist_ok=True)
    app.run(debug=True)
