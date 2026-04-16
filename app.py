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
    {
        "id": 1,
        "question": "Islom dini qaysi yili va qayerda vujudga kelgan?",
        "options": [
            "610-yil, Madina shahrida",
            "610-yil, Makka shahrida",
            "622-yil, Makka shahrida",
            "570-yil, Taif shahrida"
        ],
        "answer": "610-yil, Makka shahrida"
    },
    {
        "id": 2,
        "question": "Islomning muqaddas kitobi — Qur'oni Karim qaysi tilda nozil bo'lgan?",
        "options": [
            "Fors tilida",
            "Ibroniy tilida",
            "Arab tilida",
            "Suryon tilida"
        ],
        "answer": "Arab tilida"
    },
    {
        "id": 3,
        "question": "Muhammad (s.a.v.) qaysi yili tug'ilgan va qaysi yili vafot etgan?",
        "options": [
            "570 — 632-yillar",
            "560 — 622-yillar",
            "571 — 640-yillar",
            "580 — 650-yillar"
        ],
        "answer": "570 — 632-yillar"
    },
    {
        "id": 4,
        "question": "Islomning besh rukni (ustuni) to'g'ri keltirilgan qatorni toping:",
        "options": [
            "Kalima, namoz, ro'za, zakot, haj",
            "Namoz, zakot, ro'za, haj, jihod",
            "Iymon, namoz, ro'za, sadaqa, haj",
            "Kalima, namoz, zakot, sadaqa, umra"
        ],
        "answer": "Kalima, namoz, ro'za, zakot, haj"
    },
    {
        "id": 5,
        "question": "Hijrat — Payg'ambarimiz (s.a.v.)ning Makkadan Madinaga ko'chishi qaysi yili sodir bo'lgan?",
        "options": [
            "610-yil",
            "615-yil",
            "622-yil",
            "630-yil"
        ],
        "answer": "622-yil"
    },
    {
        "id": 6,
        "question": "Islomda birinchi xalifa kim bo'lgan?",
        "options": [
            "Umar ibn Xattob (r.a.)",
            "Abu Bakr Siddiq (r.a.)",
            "Usmon ibn Affon (r.a.)",
            "Ali ibn Abu Tolib (r.a.)"
        ],
        "answer": "Abu Bakr Siddiq (r.a.)"
    },
    {
        "id": 7,
        "question": "To'rt xulafoi roshidin (to'g'ri yo'ldagi xalifalar) davri qaysi yillarda bo'lgan?",
        "options": [
            "622 — 661-yillar",
            "632 — 661-yillar",
            "630 — 680-yillar",
            "640 — 700-yillar"
        ],
        "answer": "632 — 661-yillar"
    },
    {
        "id": 8,
        "question": "Qur'oni Karim necha sura va necha oyatdan iborat?",
        "options": [
            "110 sura, 6000 oyat",
            "114 sura, 6236 oyat",
            "120 sura, 7000 oyat",
            "114 sura, 5000 oyat"
        ],
        "answer": "114 sura, 6236 oyat"
    },
    {
        "id": 9,
        "question": "Hadis ilmida «Sihoh as-sitta» (oltita ishonchli hadis to'plami) deb nimalar ataladi?",
        "options": [
            "Buxoriy, Muslim, Abu Dovud, Termiziy, Nasaiy, Ibn Moja asarlari",
            "Buxoriy, Muslim, Molik, Shofiy, Ahmad, Termiziy asarlari",
            "Buxoriy, Muslim, Tabari, Qurtubiy, Ibn Kasir, Baydoviy asarlari",
            "Buxoriy, Muslim, Abu Hanifa, Molik, Shofiy, Hanbal asarlari"
        ],
        "answer": "Buxoriy, Muslim, Abu Dovud, Termiziy, Nasaiy, Ibn Moja asarlari"
    },
    {
        "id": 10,
        "question": "Imom Buxoriy qayerda tug'ilgan va uning mashhur asari qanday nomlanadi?",
        "options": [
            "Samarqandda tug'ilgan, «Al-Muvatta»",
            "Buxoroda tug'ilgan, «Al-Jome' as-sahih»",
            "Termizda tug'ilgan, «As-Sunan»",
            "Toshkentda tug'ilgan, «Al-Musnad»"
        ],
        "answer": "Buxoroda tug'ilgan, «Al-Jome' as-sahih»"
    },
    {
        "id": 11,
        "question": "Islomda to'rtta asosiy fiqh (huquq) maktabi — mazhablar qaysilar?",
        "options": [
            "Hanafiy, Molikiy, Shofiy, Hanbaliy",
            "Hanafiy, Jaffariy, Zaydiy, Ismoiliy",
            "Hanafiy, Shofiy, Mutaziliy, Ashariy",
            "Molikiy, Shofiy, Hanbaliy, Vahhobiy"
        ],
        "answer": "Hanafiy, Molikiy, Shofiy, Hanbaliy"
    },
    {
        "id": 12,
        "question": "Umaviylar xalifaligi qaysi shaharda joylashgan edi?",
        "options": [
            "Bag'dod",
            "Qohira",
            "Damashq",
            "Kordova"
        ],
        "answer": "Damashq"
    },
    {
        "id": 13,
        "question": "Abbosiylar xalifaligi qaysi yillarda hukmronlik qilgan?",
        "options": [
            "632 — 750-yillar",
            "661 — 750-yillar",
            "750 — 1258-yillar",
            "800 — 1300-yillar"
        ],
        "answer": "750 — 1258-yillar"
    },
    {
        "id": 14,
        "question": "Islomda «tavhid» tushunchasi nimani anglatadi?",
        "options": [
            "Payg'ambarlarga ishonch",
            "Allohning yagonaligiga e'tiqod",
            "Oxirat kuniga ishonch",
            "Farishtalar mavjudligiga e'tiqod"
        ],
        "answer": "Allohning yagonaligiga e'tiqod"
    },
    {
        "id": 15,
        "question": "«Tasavvuf» — islomiy mistitsizm qaysi asrdan boshlab shakllanib rivojlandi?",
        "options": [
            "VII — VIII asrlardan",
            "X — XI asrlardan",
            "XII — XIII asrlardan",
            "XIV — XV asrlardan"
        ],
        "answer": "VII — VIII asrlardan"
    },
    {
        "id": 16,
        "question": "Buyuk alloma va faqih Abu Hanifa (r.a.) qaysi shaharda yashagan va vafot etgan?",
        "options": [
            "Damashqda",
            "Makkada",
            "Bag'dodda",
            "Basrada"
        ],
        "answer": "Bag'dodda"
    },
    {
        "id": 17,
        "question": "Islomda «ijmo'» tushunchasi fiqhda nimani bildiradi?",
        "options": [
            "Qiyos — mantiqiy qiyoslash usuli",
            "Olimlar va mujtahidlarning biror masalada yakdil kelishuvi",
            "Shaxsiy fikr va ijtihod",
            "Hadislarni yig'ish jarayoni"
        ],
        "answer": "Olimlar va mujtahidlarning biror masalada yakdil kelishuvi"
    },
    {
        "id": 18,
        "question": "Islomda zakot kim tomonidan va qancha miqdorda to'lanadi?",
        "options": [
            "Barcha musulmonlar tomonidan, daromadning 5%",
            "Nisob miqdoriga yetgan mol-mulk egasi, mol-mulkning 2,5%",
            "Faqat boylar tomonidan, ixtiyoriy miqdorda",
            "Davlat tomonidan, soliq sifatida yig'iladi"
        ],
        "answer": "Nisob miqdoriga yetgan mol-mulk egasi, mol-mulkning 2,5%"
    },
    {
        "id": 19,
        "question": "Ramazon oyida ro'za tutish islomda qachon farz qilindi?",
        "options": [
            "Hijratdan oldin, Makkada",
            "Hijratdan keyin, Madinada — 2-hijriy yili",
            "Muhammad (s.a.v.) vafotidan keyin",
            "Birinchi xalifalar davrida"
        ],
        "answer": "Hijratdan keyin, Madinada — 2-hijriy yili"
    },
    {
        "id": 20,
        "question": "Islomda «shari'at» tushunchasi nimani anglatadi?",
        "options": [
            "Faqat namoz va ibodatlar majmuasi",
            "Qur'on va Hadisga asoslangan ilohiy qonunlar tizimi",
            "Faqat jinoyat jazolari",
            "Xalifalar chiqargan qonunlar"
        ],
        "answer": "Qur'on va Hadisga asoslangan ilohiy qonunlar tizimi"
    },
    {
        "id": 21,
        "question": "Islom aqidasi bo'yicha «iymonning olti rukni» qaysilar?",
        "options": [
            "Allohga, farishtalariga, kitoblariga, payg'ambarlariga, oxirat kuniga, qadar (taqdir)ga ishonish",
            "Allohga, Qur'onga, namozga, ro'zaga, zakotga, hajga ishonish",
            "Allohga, Payg'ambarga, Qur'onga, sunnatga, ijmo'ga, qiyosga ishonish",
            "Allohga, farishtalariga, jannatga, do'zaxga, hisob-kitobga, taqdirgа ishonish"
        ],
        "answer": "Allohga, farishtalariga, kitoblariga, payg'ambarlariga, oxirat kuniga, qadar (taqdir)ga ishonish"
    },
    {
        "id": 22,
        "question": "Mo'tazila — islomiy ratsionalistik maktab qaysi asrda paydo bo'lgan va asosiy g'oyasi nima?",
        "options": [
            "VII asrda; Qur'on yaratilmagan, abadiydir degan g'oya",
            "VIII — IX asrlarda; Qur'on yaratilgan va aql dindan ustun degan g'oya",
            "X asrda; faqat hadislarga tayanish kerak degan g'oya",
            "XII asrda; tasavvufni rad etish g'oyasi"
        ],
        "answer": "VIII — IX asrlarda; Qur'on yaratilgan va aql dindan ustun degan g'oya"
    },
    {
        "id": 23,
        "question": "Islom tarixida «Badr jangi» qachon va kimlar o'rtasida bo'lib o'tdi?",
        "options": [
            "623-yil, musulmonlar va Vizantiya o'rtasida",
            "624-yil, musulmonlar va Makka mushriklar o'rtasida",
            "625-yil, musulmonlar va Madinа yahudiylar o'rtasida",
            "627-yil, musulmonlar va forslar o'rtasida"
        ],
        "answer": "624-yil, musulmonlar va Makka mushriklar o'rtasida"
    },
    {
        "id": 24,
        "question": "Islom tarixida «Xandak jangi» (Ahzob jangi) qaysi yili bo'lib o'tdi va qanday mudofaa usuli qo'llanildi?",
        "options": [
            "624-yil, qal'a qurib himoya qilindi",
            "625-yil, tog' orqali chekinildi",
            "627-yil, Madina atrofiga xandaq (ariq) qazildi",
            "630-yil, sulh tuzildi"
        ],
        "answer": "627-yil, Madina atrofiga xandaq (ariq) qazildi"
    },
    {
        "id": 25,
        "question": "Qur'oni Karim birinchi marta to'liq kitob shaklida kim tomonidan jamlangan?",
        "options": [
            "Muhammad (s.a.v.) hayotligida o'zlari tomonidan",
            "Abu Bakr Siddiq (r.a.) xalifaligi davrida Zayd ibn Sobit boshchiligida",
            "Umar ibn Xattob (r.a.) tomonidan",
            "Ali ibn Abu Tolib (r.a.) tomonidan"
        ],
        "answer": "Abu Bakr Siddiq (r.a.) xalifaligi davrida Zayd ibn Sobit boshchiligida"
    },
    {
        "id": 26,
        "question": "Islomda «ijtihod» nima va kim ijtihod qila oladi?",
        "options": [
            "Oddiy namoz o'qish; har qanday musulmon",
            "Qur'on va Sunnat asosida yangi huquqiy hukm chiqarish; yuqori darajali olim-mujtahid",
            "Hadislarni yod olish; hafiz",
            "Qur'onni to'g'ri talaffuz qilish; qori"
        ],
        "answer": "Qur'on va Sunnat asosida yangi huquqiy hukm chiqarish; yuqori darajali olim-mujtahid"
    },
    {
        "id": 27,
        "question": "Islom tarixida «Oltin davr» (islom Uyg'onish davri) qaysi asrlarga to'g'ri keladi?",
        "options": [
            "VII — VIII asrlar",
            "VIII — XIII asrlar",
            "XIV — XVI asrlar",
            "XVII — XVIII asrlar"
        ],
        "answer": "VIII — XIII asrlar"
    },
    {
        "id": 28,
        "question": "Islomda «sunna» va «hadis» tushunchalari o'rtasidagi farq nima?",
        "options": [
            "Ular bir xil tushuncha, farqi yo'q",
            "Sunna — Payg'ambar (s.a.v.)ning amaliy hayoti va yo'li; hadis — uning so'z va xatti-harakatlarini yozma bayon etishi",
            "Sunna — Qur'on oyatlari; hadis — olimlar izohi",
            "Sunna — ixtiyoriy amallar; hadis — majburiy amallar"
        ],
        "answer": "Sunna — Payg'ambar (s.a.v.)ning amaliy hayoti va yo'li; hadis — uning so'z va xatti-harakatlarini yozma bayon etishi"
    },
    {
        "id": 29,
        "question": "Abbosiylar xalifaligi poytaxti — Bag'dod shahri qachon va kim tomonidan qurilgan?",
        "options": [
            "750-yil, Abbos ibn Abd al-Muttalib tomonidan",
            "762-yil, xalifa al-Mansur tomonidan",
            "786-yil, xalifa Horun ar-Rashid tomonidan",
            "800-yil, xalifa al-Ma'mun tomonidan"
        ],
        "answer": "762-yil, xalifa al-Mansur tomonidan"
    },
    {
        "id": 30,
        "question": "Islomda «halol» va «haram» tushunchalari nimani anglatadi?",
        "options": [
            "Halol — farz amallar; haram — sunnат amallar",
            "Halol — diniy jihatdan ruxsat etilgan narsa va amallar; haram — qat'iyan taqiqlangan narsa va amallar",
            "Halol — faqat ovqatga tegishli; haram — faqat ichimlikka tegishli",
            "Halol — dunyoviy ishlar; haram — diniy ishlar"
        ],
        "answer": "Halol — diniy jihatdan ruxsat etilgan narsa va amallar; haram — qat'iyan taqiqlangan narsa va amallar"
    },
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
    percentage = round((score / 30) * 100, 1)
    date_str = datetime.now().strftime("%Y-%m-%d %H:%M")
    answers_str = json.dumps(answers, ensure_ascii=False)
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
    return jsonify({'success': True, 'score': score, 'total': 30, 'details': result_details})

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
