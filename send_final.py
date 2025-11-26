import requests
import pandas as pd
import re
import os
import io
import matplotlib
import getpass
#import pymorphy2
import vk_api

from grafik import create_detailed_graph

matplotlib.use('Agg')
import matplotlib.pyplot as plt

from static_data import motivation_videos, future_wishes, quotes

EXCEL_FILE = None

# --- Вспомогательные функции ---
def detect_gender(full_name):
    """
    Определяет пол по окончанию имени/отчества
    """
    if pd.isna(full_name) or not isinstance(full_name, str):
        return "unknown"

    name_parts = full_name.strip().split()

    # Проверяем окончания
    for part in name_parts:
        if part.endswith(('ова', 'ева', 'ина', 'ская', 'цкая')):
            return "female"
        elif part.endswith(('ов', 'ев', 'ин', 'ский', 'цкий')):
            return "male"

    # Если не определили по фамилии, проверяем имя
    first_name = name_parts[0] if name_parts else ""
    female_endings = ('а', 'я', 'ья')
    male_endings = ('й', 'ь', 'н', 'р', 'т')

    if first_name.endswith(female_endings):
        return "female"
    elif first_name.endswith(male_endings):
        return "male"

    return "unknown"

def confirm_action(message="Продолжить?"):
    """
    Запрос подтверждения у пользователя
    """
    print(f"\n⚠️  {message}")
    print("1 - Да, продолжить")
    print("2 - Нет, отменить")

    while True:
        choice = input("Ваш выбор (1/2): ").strip()
        if choice == "1":
            return True
        elif choice == "2":
            return False
        else:
            print("❌ Неверный выбор. Введите 1 или 2")

def adapt_wish_by_gender(wish_text, gender):
    """
    Адаптирует текст пожелания под пол ученика
    """
    if gender == "male":
        return wish_text.replace('стал(а)', 'стал').replace('подготовился(лась)', 'подготовился').replace('уверен(а)', 'уверен')
    elif gender == "female":
        return wish_text.replace('стал(а)', 'стала').replace('подготовился(лась)', 'подготовилась').replace('уверен(а)', 'уверена')
    else:
        return wish_text  # Оставляем оба варианта если пол не определили

def get_vk_token_gui():
    """
    Для GUI версии - токен передаётся как параметр
    """
    return None

def extract_name(full_name):
   if pd.isna(full_name) or not isinstance(full_name, str):
       return "Друг"
   parts = full_name.strip().split()
   return parts[0] if parts else "Друг"

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


def load_template(category):
    from templates_embedded import TEMPLATES

    category_map = {1: 'strong', 2: 'medium', 3: 'weak'}
    template_key = category_map.get(category, 'medium')

    return TEMPLATES.get(template_key, "Шаблон не найден")

def format_best_hw(best_entries):
   if not best_entries or best_entries[0][0] == "—":
       return "—"
   parts = []
   for name, score, pct in best_entries:
       parts.append(f"«{name}» — {int(score)} баллов ({pct:.0f}%)")
   return "\n".join(parts)

def get_best_hw_info(headers, hw_columns, student_scores, max_scores, lesson_numbers):
   hard_hw = []
   percent_sum = 0
   count = 0

   for i, col in enumerate(hw_columns):
       max_val = max_scores[i]
       stud_val = student_scores[i]
       if max_val > 1 and stud_val > 0:
           percent = (stud_val / max_val) * 100
           header = str(headers[col])
           lesson_name = re.sub(r'^\d+\.\s*', '', header).strip()
           hard_hw.append((lesson_name, stud_val, percent))
           percent_sum += percent
           count += 1

   if count == 0:
       return 0.0, [("—", 0, 0)]

   avg_percent = round(percent_sum / count, 1)
   max_percent = max(hw[2] for hw in hard_hw)
   best_entries = [hw for hw in hard_hw if abs(hw[2] - max_percent) < 1e-5]
   return avg_percent, best_entries

def build_message_for_student(
        name, full_name, category, block_number, hw_count,
        test_done_count, test_total_count,
        avg_percent, best_hw_str, lives
):
    # ⭐⭐ ПРАВИЛЬНАЯ ИНДЕКСАЦИЯ ДЛЯ 9 ЭЛЕМЕНТОВ ⭐⭐
    quote_index = block_number - 1
    wish_index = block_number - 1
    video_index = block_number - 1

    # Защита от выхода за границы массивов
    quote_index = min(quote_index, len(quotes) - 1)
    wish_index = min(wish_index, len(future_wishes) - 1)
    video_index = min(video_index, len(motivation_videos) - 1)

    quote = quotes[quote_index].format(name=name)
    wish = future_wishes[wish_index]
    video_url = motivation_videos[video_index].strip()

    gender = detect_gender(full_name)
    wish = adapt_wish_by_gender(wish, gender)

    lives_message = "Ни одной жизни не потеряно! 🚘" if lives == 3 else f"Количество жизней: {lives}/3"
    template = load_template(category)
    message_text = template.format(
        BLOCK_NUMBER=block_number,
        HW_COUNT=hw_count,
        TEST_DONE_COUNT=test_done_count,
        TEST_TOTAL_COUNT=test_total_count,
        AVG_HARD_SCORE=f"{avg_percent}%",
        BEST_HW_BLOCK=best_hw_str,
        LIVES=lives,
        LIVES_MESSAGE=lives_message
    )
    return f"{quote}\n\n{message_text}\n\n{wish}", video_url

def get_curators_vk_ids(df_full):
    """Читает vk_id куратора из C5 (Excel строка 5 → индекс 4)"""
    curators = []
    if len(df_full) > 4:  # есть ли строка 5 (индекс 4)
        vk_id_raw = df_full.iloc[4, 2]  # C5 → строка 5 → индекс 4, столбец C → индекс 2
        if pd.notna(vk_id_raw) and str(vk_id_raw).isdigit():
            curators.append(int(vk_id_raw))
    return curators

def get_video_attachment(vk, video_url):
    """
    Преобразует URL видео VK в attachment вида 'video-123456_789'
    """
    try:
        # Извлекаем owner_id и video_id из URL
        match = re.search(r'video(-?\d+)_(\d+)', video_url)
        if match:
            owner_id = match.group(1)
            video_id = match.group(2)
            return f"video{owner_id}_{video_id}"

        # Альтернативный формат URL
        match = re.search(r'vk\.com\/video(\d+)_(\d+)', video_url)
        if match:
            owner_id = match.group(1)
            video_id = match.group(2)
            return f"video{owner_id}_{video_id}"

    except Exception as e:
        print(f"❌ Ошибка разбора URL видео: {e}")

    return None

def get_vk_token():
   """
   Запрашивает VK токен у пользователя через консоль
   """
   print("\n🔐 Введите VK API токен для отправки сообщений")
   print("⚠️  Токен не будет отображаться при вводе для безопасности")
   token = getpass.getpass("VK Token: ").strip()

   if not token:
       print("❌ Токен не может быть пустым")
       return None

   # Простая проверка формата токена
   if not token.startswith('vk1.a.') or len(token) < 50:
       print("❌ Неверный формат токена. Должен начинаться с 'vk1.a.'")
       return None
   return token

# --- Режим ПРЕВЬЮ ---
def preview_mode(vk_token=None, block_number=None, lesson_range=None, excel_file=None):
    """
    Режим превью с поддержкой GUI
    """
    # Используем переданный файл или глобальный EXCEL_FILE
    current_excel_file = excel_file if excel_file else EXCEL_FILE

    if vk_token is None:
        vk_token = get_vk_token()
    if block_number is None:
        block_number = int(input("Номер блока: "))
    if lesson_range is None:
        lesson_input = input("Диапазон домашек (например, 12-17): ")
    else:
        lesson_input = lesson_range

    # Используем current_excel_file вместо EXCEL_FILE
    df_full = pd.read_excel(current_excel_file, header=None)
    curators = get_curators_vk_ids(df_full)

    if not curators:
        print("❌ Не найдены vk_id кураторов в C5")
        return

    print(f"🎯 Превью будет отправлено кураторам: {curators}")

    target_lessons = parse_lesson_range(lesson_input)

    headers = df_full.iloc[0]
    max_scores_row = df_full.iloc[6]
    student_rows = list(df_full.iloc[7:].iterrows())

    hw_columns = []
    lesson_numbers = []
    for col_idx in headers[19:].index:
        num = extract_lesson_number(headers[col_idx])
        if num is not None and num in target_lessons:
            hw_columns.append(col_idx)
            lesson_numbers.append(num)

    if not hw_columns:
        print("❌ Не найдено ДЗ")
        return

    combined = sorted(zip(lesson_numbers, hw_columns))
    lesson_numbers, hw_columns = zip(*combined)
    lesson_numbers = list(lesson_numbers)
    hw_columns = list(hw_columns)

    # Собираем до 2 учеников на категорию
    representatives = {1: [], 2: [], 3: []}
    for _, row in student_rows:
        full_name = row.iloc[1]
        vk_id_raw = row.iloc[2]
        lives_raw = row.iloc[4]

        if pd.isna(vk_id_raw) or not str(vk_id_raw).isdigit():
            continue

        name = extract_name(full_name)
        lives = int(lives_raw) if pd.notna(lives_raw) else 0

        student_scores = []
        max_scores = []
        total_score = 0
        test_done_count = 0
        test_total_count = 0
        hard_scores = []

        for col in hw_columns:
            stud_val = row[col] if pd.notna(row[col]) else 0
            max_val = max_scores_row[col] if pd.notna(max_scores_row[col]) else 1
            stud_val = float(stud_val) if str(stud_val).replace('.', '').isdigit() else 0
            max_val = float(max_val) if str(max_val).replace('.', '').isdigit() else 1

            student_scores.append(stud_val)
            max_scores.append(max_val)
            total_score += stud_val

            if max_val <= 1:
                test_total_count += 1
                if stud_val >= 1:
                    test_done_count += 1
            else:
                if stud_val > 0:
                    hard_scores.append(stud_val)

        hard_submitted_all = len(hard_scores) == sum(1 for col in hw_columns if max_scores_row[col] > 1)
        max_possible_score = sum(max_scores_row[col] for col in hw_columns if not pd.isna(max_scores_row[col]))
        ratio = total_score / max_possible_score if max_possible_score > 0 else 0

        if hard_submitted_all and ratio >= 0.70:
            category = 1
        elif ratio >= 0.42:
            category = 2
        else:
            category = 3

        if len(representatives[category]) < 2:
            avg_percent, best_entries = get_best_hw_info(headers, hw_columns, student_scores, max_scores,
                                                         lesson_numbers)
            best_hw_str = format_best_hw(best_entries)
            msg_text, video_url = build_message_for_student(
                name, full_name, category, block_number, len(hw_columns),  # ⭐⭐ ДОБАВИЛИ full_name ⭐⭐
                test_done_count, test_total_count,
                avg_percent, best_hw_str, lives
            )
            representatives[category].append({
                'original_name': name,
                'message': (msg_text, video_url),
                'category': category
            })

        if all(len(representatives[cat]) >= 2 for cat in [1, 2, 3]):
            break

    # === Отправка куратору ===
    if not vk_token:
        print("❌ Не задан VK_TOKEN — не удастся отправить превью")
        return

    vk_session = vk_api.VkApi(token=vk_token)
    vk = vk_session.get_api()

    total_sent = 0
    for cat in [1, 2, 3]:
        for rep in representatives[cat]:
            # Найдём исходные данные ученика для графика
            student_found = None
            for _, row in student_rows:
                name = extract_name(row.iloc[1])
                if name == rep['original_name']:
                    student_found = row
                    break

            if student_found is None:
                print(f"⚠️ Не найдены данные для графика: {rep['original_name']}")
                continue

            # Собираем данные для графика
            student_scores = []
            max_scores = []
            lives = int(student_found.iloc[4]) if pd.notna(student_found.iloc[4]) else 0

            for col in hw_columns:
                stud_val = student_found[col] if pd.notna(student_found[col]) else 0
                max_val = max_scores_row[col] if pd.notna(max_scores_row[col]) else 1
                stud_val = float(stud_val) if str(stud_val).replace('.', '').isdigit() else 0
                max_val = float(max_val) if str(max_val).replace('.', '').isdigit() else 1
                student_scores.append(stud_val)
                max_scores.append(max_val)

            # Создаём график
            try:
                graph_buf = create_detailed_graph(lesson_numbers, student_scores, max_scores, lives,
                                                  rep['original_name'])
                graph_attach = upload_graph_to_vk(vk, curators[0], graph_buf)
            except Exception as e:
                print(f"Ошибка создания графика для {rep['original_name']}: {e}")
                graph_attach = None

            msg_text, video_url = rep['message']
            message_with_header = f"【ПРЕВЬЮ】Категория {cat} — ученик: {rep['original_name']}\n\n{msg_text}"

            for curator_id in curators:
                try:
                    # Первое сообщение с текстом и графиком
                    vk.messages.send(
                        user_id=curator_id,
                        message=message_with_header,
                        attachment=graph_attach,
                        random_id=0
                    )

                    # Второе сообщение с видео как attachment
                    video_attach = get_video_attachment(vk, video_url)
                    if video_attach:
                        vk.messages.send(
                            user_id=curator_id,
                            message="",
                            attachment=video_attach,
                            random_id=0
                        )
                    else:
                        # Fallback: если не удалось распарсить, отправляем ссылку
                        vk.messages.send(
                            user_id=curator_id,
                            message=f"{video_url}",
                            random_id=0
                        )

                    total_sent += 2
                    print(f"✅ Превью отправлено куратору: {curator_id}")
                except Exception as e:
                    print(f"❌ Ошибка отправки превью куратору {curator_id}: {e}")

    print(f"\n✅ Превью отправлено: {total_sent} сообщений ({len(curators)} кураторам)")


def send_mode(vk_token=None, block_number=None, lesson_range=None, skip_rows_input="", excel_file=None, chat_id=None):
    """
    Режим отправки с подтверждением через бота
    """
    # Используем переданный файл или глобальный EXCEL_FILE
    current_excel_file = excel_file if excel_file else EXCEL_FILE

    # Получение параметров
    if vk_token is None:
        vk_token = get_vk_token()
    if block_number is None:
        block_number = int(input("Номер блока: "))
    if lesson_range is None:
        lesson_input = input("Диапазон домашек (например, 12-17): ")
    else:
        lesson_input = lesson_range

    if skip_rows_input:
        skip_input = skip_rows_input
    else:
        skip_input = input("Номера студентов для пропуска (через запятую, например: 1,3): ").strip()

    skip_students = set()
    if skip_input:
        try:
            skip_students = {int(x.strip()) for x in skip_input.split(',')}
        except:
            print("⚠️ Ошибка в номерах студентов. Продолжаем без исключений.")

    # Используем current_excel_file вместо EXCEL_FILE
    df_full = pd.read_excel(current_excel_file, header=None)
    headers = df_full.iloc[0]
    max_scores_row = df_full.iloc[6]
    student_rows = list(df_full.iloc[7:].iterrows())

    # Выбор столбцов
    hw_columns = []
    lesson_numbers = []
    for col_idx in headers[19:].index:
        num = extract_lesson_number(headers[col_idx])
        if num is not None and num in parse_lesson_range(lesson_input):
            hw_columns.append(col_idx)
            lesson_numbers.append(num)

    if not hw_columns:
        print("❌ Не найдено ДЗ")
        return

    combined = sorted(zip(lesson_numbers, hw_columns))
    lesson_numbers, hw_columns = zip(*combined)
    lesson_numbers = list(lesson_numbers)
    hw_columns = list(hw_columns)

    # Подготовка списка студентов для отправки
    students_to_process = []
    for student_number, (original_idx, row) in enumerate(student_rows, 1):  # ← Начинаем с 1!
        excel_row_number = original_idx + 1

        # Пропускаем по номеру студента, а не по строке Excel
        if student_number in skip_students:
            print(f"🚫 Пропущен студент #{student_number} (строка Excel: {excel_row_number})")
            continue

        full_name = row.iloc[1]
        vk_id_raw = row.iloc[2]

        if pd.isna(vk_id_raw) or not str(vk_id_raw).isdigit():
            continue

        students_to_process.append((original_idx, row, excel_row_number, student_number))

    total_students = len(students_to_process)
    total_messages = total_students * 2

    print("\n" + "=" * 50)
    print("📊 СТАТИСТИКА ОТПРАВКИ")
    print("=" * 50)
    print(f"👥 Всего учеников: {total_students}")
    print(f"📨 Всего сообщений: {total_messages} (текст + видео)")
    print(f"🚫 Пропущено студентов: {len(skip_students)}")

    if students_to_process:
        print("\n📋 Первые 5 учеников для отправки:")
        for i, (_, row, excel_row, student_num) in enumerate(students_to_process[:5]):
            name = extract_name(row.iloc[1])
            vk_id = row.iloc[2]
            print(f"  {student_num}. {name} (VK ID: {vk_id}, строка Excel: {excel_row})")

        if total_students > 5:
            print(f"  ... и еще {total_students - 5} учеников")

    print("🔄 Начинаю отправку...")

    if not vk_token:
        print("❌ Не задан VK_TOKEN")
        return

    import vk_api
    vk_session = vk_api.VkApi(token=vk_token)
    vk = vk_session.get_api()

    sent_count = 0
    total_students = len(students_to_process)

    # Создаем сообщение с прогресс-баром
    progress_message = None
    if students_to_process and chat_id:
        try:
            progress_text = create_progress_bar(0, total_students, "Начинаю отправку...")
            progress_message = bot.send_message(
                chat_id,
                progress_text,
                parse_mode='Markdown'
            )
        except Exception as e:
            print(f"⚠️ Не удалось создать прогресс-бар: {e}")

    for i, (original_idx, row, excel_row_number, student_number) in enumerate(students_to_process):
        full_name = row.iloc[1]
        vk_id_raw = row.iloc[2]
        vk_id = int(vk_id_raw)
        name = extract_name(full_name)
        lives_raw = row.iloc[4]
        lives = int(lives_raw) if pd.notna(lives_raw) else 0

        student_scores = []
        max_scores = []
        total_score = 0
        test_done_count = 0
        test_total_count = 0
        hard_scores = []

        for col in hw_columns:
            stud_val = row[col] if pd.notna(row[col]) else 0
            max_val = max_scores_row[col] if pd.notna(max_scores_row[col]) else 1
            stud_val = float(stud_val) if str(stud_val).replace('.', '').isdigit() else 0
            max_val = float(max_val) if str(max_val).replace('.', '').isdigit() else 1

            student_scores.append(stud_val)
            max_scores.append(max_val)
            total_score += stud_val

            if max_val <= 1:
                test_total_count += 1
                if stud_val >= 1:
                    test_done_count += 1
            else:
                if stud_val > 0:
                    hard_scores.append(stud_val)

        hard_submitted_all = len(hard_scores) == sum(1 for col in hw_columns if max_scores_row[col] > 1)
        max_possible_score = sum(max_scores_row[col] for col in hw_columns if not pd.isna(max_scores_row[col]))
        ratio = total_score / max_possible_score if max_possible_score > 0 else 0

        if hard_submitted_all and ratio >= 0.70:
            category = 1
        elif ratio >= 0.42:
            category = 2
        else:
            category = 3

        avg_percent, best_entries = get_best_hw_info(headers, hw_columns, student_scores, max_scores, lesson_numbers)
        best_hw_str = format_best_hw(best_entries)
        message_text, video_url = build_message_for_student(
            name, full_name, category, block_number, len(hw_columns),
            test_done_count, test_total_count,
            avg_percent, best_hw_str, lives
        )

        try:
            # === Создание и отправка графика ===
            try:
                graph_buf = create_detailed_graph(lesson_numbers, student_scores, max_scores, lives, name)
                graph_attach = upload_graph_to_vk(vk, vk_id, graph_buf)
            except Exception as e:
                print(f"⚠️ Ошибка создания графика для {name}: {e}")
                graph_attach = None

            # Отправка текстового сообщения С ГРАФИКОМ
            vk.messages.send(
                user_id=vk_id,
                message=message_text,
                attachment=graph_attach,
                random_id=0
            )

            # Отправка видео
            video_attach = get_video_attachment(vk, video_url)
            if video_attach:
                vk.messages.send(
                    user_id=vk_id,
                    message="",
                    attachment=video_attach,
                    random_id=0
                )
            else:
                # Fallback: если не удалось распарсить, отправляем ссылку
                vk.messages.send(
                    user_id=vk_id,
                    message=f"{video_url}",
                    random_id=0
                )

            sent_count += 2
            print(f"✅ Отправлено: #{student_number} {name} (строка Excel: {excel_row_number})")

            # === ОБНОВЛЯЕМ ПРОГРЕСС-БАР ===
            if progress_message:
                try:
                    current_status = f"Отправлено: {name}"
                    progress_text = create_progress_bar(i + 1, total_students, current_status)

                    bot.edit_message_text(
                        chat_id=progress_message.chat.id,
                        message_id=progress_message.message_id,
                        text=progress_text,
                        parse_mode='Markdown'
                    )
                except Exception as e:
                    print(f"⚠️ Не удалось обновить прогресс-бар: {e}")

        except Exception as e:
            print(f"❌ Ошибка для #{student_number} {name}: {e}")

            # Обновляем прогресс-бар с ошибкой
            if progress_message:
                try:
                    current_status = f"Ошибка: {name}"
                    progress_text = create_progress_bar(i + 1, total_students, current_status)
                    bot.edit_message_text(
                        chat_id=progress_message.chat.id,
                        message_id=progress_message.message_id,
                        text=progress_text,
                        parse_mode='Markdown'
                    )
                except:
                    pass

    # Финальное обновление прогресс-бара
    if progress_message:
        try:
            success_count = sent_count // 2
            final_text = (
                f"✅ *Отправка завершена!*\n\n"
                f"📊 *Результат:*\n"
                f"• Студентов: {total_students}\n"
                f"• Сообщений: {sent_count}\n"
                f"• Успешно: {success_count}/{total_students}"
            )
            bot.edit_message_text(
                chat_id=progress_message.chat.id,
                message_id=progress_message.message_id,
                text=final_text,
                parse_mode='Markdown'
            )
        except Exception as e:
            print(f"⚠️ Не удалось обновить финальный прогресс-бар: {e}")

    print(f"\n🏁 Отправка завершена. Всего отправлено: {sent_count} сообщений")


def create_progress_bar(current, total, status="", length=10):
    """
    Создает текстовый прогресс-бар
    """
    percent = current / total if total > 0 else 0
    filled_length = int(length * percent)
    bar = '█' * filled_length + '▒' * (length - filled_length)

    return (
        f"📤 *Отправка сообщений*\n\n"
        f"`[{bar}]` {percent:.1%}\n"
        f"**{current}/{total}** студентов\n"
        f"_{status}_"
    )

def upload_graph_to_vk(vk, user_id, graph_buffer):
   """Загружает график в ВК и возвращает attachment вида 'photo123_456'"""
   upload_url = vk.photos.getMessagesUploadServer(peer_id=user_id)['upload_url']
   response = requests.post(upload_url, files={'photo': ('graph.png', graph_buffer.read(), 'image/png')})
   result = response.json()
   photo = vk.photos.saveMessagesPhoto(
       photo=result['photo'],
       server=result['server'],
       hash=result['hash']
   )[0]
   return f"photo{photo['owner_id']}_{photo['id']}"

if __name__ == '__main__':
   vk_token = get_vk_token()
   if not vk_token:
       print("❌ Не удалось получить токен. Программа завершена.")
       exit(1)

   mode = input("Режим: (1) Превью (к куратору) / (2) Отправка ученикам? ")
   if mode.strip() == "1":
       preview_mode(vk_token)
   elif mode.strip() == "2":
       send_mode(vk_token)
   else:
       print("Неверный выбор.")
