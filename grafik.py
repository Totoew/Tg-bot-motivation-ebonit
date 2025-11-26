import pandas as pd
import re
import vk_api
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import numpy as np
import io
import requests
import os
EXCEL_FILE = '/Users/daniltotoev/Downloads/Новая таблица.xlsx'
VK_TOKEN = "vk1.a.s1sznd8vChzG2Cz4XiD5SQ__txXf9MNJOn6qJYbnqsoa5CVtyGkTQfdMVnxDDXJaK6Krs2BKl0Kvi1EANYHaCp8Q1YWgX-ZLZo_OEjA3dimeimvo2w2Q_7U1Pks1lxGXxWNIoPxxaU8LnJK7wCx_s7xjoFd4OGEtJaR4J4_2VSc9witYessjlxqr8lQEn6h5cBQRjgkxdwlPd-CxkUZv2g"  # ← ЗАМЕНИ!
#VK_TOKEN = os.getenv('VK_TOKEN')

def extract_lesson_number(header):
    match = re.match(r'^(\d+)', str(header))
    return int(match.group(1)) if match else None

def parse_lesson_range(user_input):
    user_input = user_input.strip()
    if '-' in user_input:
        start, end = map(int, user_input.split('-'))
        return list(range(start, end + 1))
    else:
        return [int(user_input)]

def create_detailed_graph(lesson_nums, student_scores, max_scores, lives, name):
    hard_lessons = []
    hard_percents = []
    test_lessons = []
    test_done = []

    for i, lesson in enumerate(lesson_nums):
        max_val = max_scores[i]
        stud_val = student_scores[i]

        if pd.isna(stud_val):
            stud_val = 0
        if pd.isna(max_val):
            max_val = 1

        if max_val > 1:  # сложное ДЗ
            percent = (stud_val / max_val * 100) if max_val > 0 else 0
            hard_lessons.append(lesson)
            hard_percents.append(percent)
        else:  # тестовое
            test_lessons.append(lesson)
            test_done.append(bool(stud_val >= 1))

    fig, ax1 = plt.subplots(figsize=(9, 5))
    hearts = "❤️" * lives
    fig.suptitle(f"Твои результаты, {name} | Жизней: {lives} {hearts}", fontsize=14, weight='bold')

    if hard_lessons:
        bars = ax1.bar(hard_lessons, hard_percents, color='#2ca02c', edgecolor='black', width=0.6)
        ax1.set_ylabel("Балл (% от максимума)", color='#2ca02c')
        ax1.set_ylim(0, 110)
        ax1.grid(axis='y', linestyle='--', alpha=0.6)
        for bar, pct in zip(bars, hard_percents):
            ax1.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 2,
                     f"{pct:.0f}%", ha='center', va='bottom', fontweight='bold')

    # Подпись тестовых ДЗ — текст вместо значков
    checkbox_parts = []
    for lesson in lesson_nums:
        if lesson in test_lessons:
            idx = test_lessons.index(lesson)
            status = "Сделано✓" if test_done[idx] else "Не сделано =("
            checkbox_parts.append(f"{lesson}: {status}")
    if checkbox_parts:
        plt.figtext(0.02, 0.02, "Тестовые ДЗ: " + " | ".join(checkbox_parts), fontsize=9, ha="left")

    ax1.set_xlabel("Номер урока")
    ax1.set_xticks(lesson_nums)
    ax1.set_xticklabels(lesson_nums)
    plt.tight_layout(rect=[0, 0.05, 1, 0.95])

    buf = io.BytesIO()
    plt.savefig(buf, format='png', dpi=150, bbox_inches='tight')
    buf.seek(0)
    plt.close()
    buf.seek(0)
    return buf

def main():
    # === Ввод диапазона ===
    lesson_input = input("Введите диапазон домашек (например, 12-17): ")
    try:
        target_lessons = parse_lesson_range(lesson_input)
    except:
        print("❌ Ошибка в формате. Используй '12-17' или '15'")
        return

    # === Чтение Excel ===
    try:
        df_full = pd.read_excel(EXCEL_FILE, header=None)
    except FileNotFoundError:
        print(f"❌ Файл не найден: {EXCEL_FILE}")
        return

    headers = df_full.iloc[0]           # строка 1 — заголовки
    max_scores_row = df_full.iloc[6]    # строка 7 — макс. баллы
    student_rows = df_full.iloc[21:22]   # первые 3 ученика (8,9,10)

    # === Находим столбцы, соответствующие нужным урокам ===
    hw_columns = []
    lesson_numbers = []
    for col_idx in headers[19:].index:  # начиная с T (индекс 19)
        num = extract_lesson_number(headers[col_idx])
        if num is not None and num in target_lessons:
            hw_columns.append(col_idx)
            lesson_numbers.append(num)

    # Сортируем по номеру урока
    combined = sorted(zip(lesson_numbers, hw_columns))
    if not combined:
        print(f"❌ Не найдено ДЗ для уроков: {target_lessons}")
        return
    lesson_numbers, hw_columns = zip(*combined)
    lesson_numbers = list(lesson_numbers)
    hw_columns = list(hw_columns)

    print(f"✅ Будут обработаны уроки: {lesson_numbers}")

    # === Отправка ===
    vk_session = vk_api.VkApi(token=VK_TOKEN)
    vk = vk_session.get_api()

    for _, row in student_rows.iterrows():
        full_name = row.iloc[1]
        vk_id_raw = row.iloc[2]
        lives_raw = row.iloc[4]

        if pd.isna(vk_id_raw) or not str(vk_id_raw).isdigit():
            print(f"⚠️ Пропущен: {full_name}")
            continue

        vk_id = int(vk_id_raw)
        lives = int(lives_raw) if pd.notna(lives_raw) else 0
        name = full_name.split()[0] if full_name and isinstance(full_name, str) else "Друг"

        # Собираем данные только по выбранным столбцам
        student_scores = [float(row[col]) if pd.notna(row[col]) else 0.0 for col in hw_columns]
        max_scores = [float(max_scores_row[col]) if pd.notna(max_scores_row[col]) else 1.0 for col in hw_columns]

        try:
            graph_buf = create_detailed_graph(lesson_numbers, student_scores, max_scores, lives, name)
        except Exception as e:
            print(f"Ошибка графика для {name}: {e}")
            continue

        try:
            upload_url = vk.photos.getMessagesUploadServer(peer_id=vk_id)['upload_url']
            response = requests.post(upload_url, files={'photo': ('results.png', graph_buf.read(), 'image/png')})
            result = response.json()
            saved_photo = vk.photos.saveMessagesPhoto(
                photo=result['photo'],
                server=result['server'],
                hash=result['hash']
            )[0]
            vk.messages.send(
                user_id=vk_id,
                message=f"📊 Твои результаты по домашкам {min(lesson_numbers)}–{max(lesson_numbers)}:",
                attachment=f"photo{saved_photo['owner_id']}_{saved_photo['id']}",
                random_id=0
            )
            print(f"✅ Отправлено: {name}")
        except Exception as e:
            print(f"❌ Ошибка отправки для {name}: {e}")

if __name__ == '__main__':
    main()
