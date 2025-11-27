import time
import re
import telebot
from telebot.types import InlineKeyboardMarkup, InlineKeyboardButton, ReplyKeyboardRemove
import os
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import io
from send_final import (
    preview_mode,
    send_mode,
    extract_lesson_number,
    parse_lesson_range,
    get_best_hw_info,
    format_best_hw,
    create_detailed_graph,
    extract_name
)

def extract_lesson_number(header):
    """Извлекает номер урока из заголовка, включая пробники с дробными номерами"""
    if pd.isna(header):
        return None

    header_str = str(header).strip()

    # Для формата "24.1. Пробник №1" - возвращаем дробное число 24.1
    match_float = re.match(r'^(\d+)\.(\d+)', header_str)
    if match_float:
        whole_part = int(match_float.group(1))
        decimal_part = int(match_float.group(2))
        return whole_part + decimal_part * 0.1

    # Для обычных ДЗ: "24. Обычное ДЗ" - возвращаем целое число 24
    match_int = re.match(r'^(\d+)', header_str)
    return int(match_int.group(1)) if match_int else None

with open(r"C:\Users\Пользователь\Desktop\bot-token.txt", 'r', encoding='utf-8') as file:
    content = file.read()

with open(r"C:\Users\Пользователь\Desktop\bot-token.txt", 'r', encoding='utf-8') as file:
    content = file.read()

bot = telebot.TeleBot(content)

user_data = {}

@bot.message_handler(commands=['start'])
def start(message):
    """Главное меню"""
    user_id = message.from_user.id
    user_data[user_id] = {'step': 'main_menu'}

    keyboard = InlineKeyboardMarkup()
    keyboard.add(InlineKeyboardButton("👁️ ПРЕВЬЮ", callback_data="preview"))
    keyboard.add(InlineKeyboardButton("📤 ОТПРАВКА", callback_data="send"))
    keyboard.add(InlineKeyboardButton("📊 ПОЛУЧИТЬ СТАТИСТИКУ", callback_data="get_stats"))
    keyboard.add(InlineKeyboardButton("📈 ВЫВЕСТИ ГРАФИК ПРОБНИКОВ", callback_data="probniki_stats"))  # НОВАЯ КНОПКА
    keyboard.add(InlineKeyboardButton("⚙️ НАСТРОЙКИ", callback_data="settings"))
    keyboard.add(InlineKeyboardButton("🎥 ВИДЕО-ИНСТРУКЦИЯ",
                                      url="https://docs.google.com/document/d/1utGllba1nr1QqmnLpOK03hwYpY87NmVIyDgsfk3kJpA/edit?usp=sharing"))

    bot.send_message(
        message.chat.id,
        "🚀 *Physics Motivation Bot*\n\n"
        "Выберите действие:",
        reply_markup=keyboard,
        parse_mode='Markdown'
    )

@bot.callback_query_handler(func=lambda call: True)
def handle_callback(call):
    """Обработка кнопок главного меню"""
    user_id = call.from_user.id

    if call.data == "preview":
        bot.send_message(call.message.chat.id, "📁 Загрузите Excel файл:")
        user_data[user_id] = {'step': 'waiting_excel', 'mode': 'preview'}

    elif call.data == "send":
        bot.send_message(call.message.chat.id, "📁 Загрузите Excel файл:")
        user_data[user_id] = {'step': 'waiting_excel', 'mode': 'send'}

    elif call.data == "get_stats":
        bot.send_message(call.message.chat.id, "📁 Загрузите Excel файл для статистики:")
        user_data[user_id] = {'step': 'waiting_stats_excel'}

    elif call.data == "probniki_stats":  # НОВЫЙ ОБРАБОТЧИК
        bot.send_message(call.message.chat.id, "📁 Загрузите Excel файл для анализа пробников:")
        user_data[user_id] = {'step': 'waiting_probniki_excel'}

    elif call.data == "settings":
        show_instructions(call.message)

    elif call.data == "back_to_menu":
        try:
            bot.delete_message(call.message.chat.id, call.message.message_id)
        except:
            pass
        start(call.message)


def show_instructions(message):
    instructions = """
🔧 ИНСТРУКЦИЯ ПО ИСПОЛЬЗОВАНИЮ БОТА

🎯 РЕЖИМ «ПРЕВЬЮ»:
• Отправляет примеры сообщений кураторам для проверки
• Показывает по 2 студента из каждой категории
• Включает графики успеваемости и видео
• Не отправляет сообщения реальным студентам

📤 РЕЖИМ «ОТПРАВКА»:
• Отправляет сообщения реальным студентам
• Можно указать номера студентов для пропуска
• Каждый студент получает 2 сообщения: текст+график и видео

📊 СТАТИСТИКА:
• Генерирует графики успеваемости для всех студентов
• Показывает статистику по выполнению ДЗ
• Отправляет результаты в телеграм

📊 КАТЕГОРИИ СТУДЕНТОВ:
• Категория 1 - лучшие результаты (≥70% + все сложные ДЗ)
• Категория 2 - хорошие результаты (≥42%)
• Категория 3 - нужно улучшить результат (<42%)

🔄 ПРОЦЕСС РАБОТЫ:
1. Выберите режим (Превью/Отправка/Статистика)
2. Загрузите Excel файл
3. Введите необходимые параметры
4. Получите результат

⏱ *ВРЕМЯ ОБРАБОТКИ:*
• В среднем 2-3 минуты на 30 учеников

❓ ЧАСТЫЕ ПРОБЛЕМЫ:
• Файл не загружается - проверьте формат (.xlsx/.xls)
• Сообщения не отправляются - проверьте VK токен
• Студенты не находятся - проверьте VK ID в файле

📞 ПОДДЕРЖКА:
Если возникли проблемы - обратитесь к разработчику бота @totoevv.
"""
    keyboard = InlineKeyboardMarkup()
    keyboard.add(InlineKeyboardButton("⬅️ В главное меню", callback_data="back_to_menu"))

    bot.send_message(
        message.chat.id,
        instructions,
        reply_markup=keyboard,
        parse_mode='Markdown'
    )

@bot.message_handler(func=lambda message: user_data.get(message.from_user.id, {}).get('step') == 'waiting_probniki_excel',
                     content_types=['document'])
def handle_probniki_excel(message):
    """Обрабатывает загрузку Excel файла для пробников"""
    user_id = message.from_user.id

    if not message.document.file_name.endswith(('.xlsx', '.xls')):
        bot.send_message(message.chat.id, "❌ Пожалуйста, загрузите Excel файл (.xlsx или .xls)")
        return

    try:
        file_info = bot.get_file(message.document.file_id)
        downloaded_file = bot.download_file(file_info.file_path)

        file_path = f"temp_probniki_{user_id}_{message.document.file_name}"
        with open(file_path, 'wb') as new_file:
            new_file.write(downloaded_file)

        user_data[user_id]['excel_file'] = file_path
        user_data[user_id]['step'] = 'waiting_probniki_limit'

        bot.send_message(
            message.chat.id,
            "👥 *Сколько учеников вывести?*\n\n"
            "Введите число (например: 10) или 'все' для вывода всех учеников:",
            parse_mode='Markdown'
        )

    except Exception as e:
        bot.send_message(message.chat.id, f"❌ Ошибка загрузки файла: {e}")


@bot.message_handler(
    func=lambda message: user_data.get(message.from_user.id, {}).get('step') == 'waiting_probniki_limit')
def handle_probniki_limit(message):
    """Обрабатывает ввод лимита учеников для пробников"""
    user_id = message.from_user.id

    try:
        limit_input = message.text.strip().lower()

        if limit_input == 'все':
            limit = None
        else:
            try:
                limit = int(limit_input)
                if limit <= 0:
                    bot.send_message(message.chat.id, "❌ Число должно быть больше 0")
                    return
            except ValueError:
                bot.send_message(message.chat.id, "❌ Введите число или 'все'")
                return

        bot.send_message(message.chat.id, "⏳ Анализирую пробники...")
        generate_probniki_stats(
            message,
            user_data[user_id]['excel_file'],
            limit
        )

        # Очистка
        if os.path.exists(user_data[user_id]['excel_file']):
            os.remove(user_data[user_id]['excel_file'])
        user_data[user_id] = {'step': 'main_menu'}

    except Exception as e:
        bot.send_message(message.chat.id, f"❌ Ошибка: {e}")


def generate_probniki_stats(message, excel_file, limit=None):
    """Генерирует комбинированные графики по пробникам"""
    try:
        # Читаем Excel файл
        df_full = pd.read_excel(excel_file, header=None)
        headers = df_full.iloc[0]
        max_scores_row = df_full.iloc[6]
        student_rows = list(df_full.iloc[7:].iterrows())

        # Карта пробников
        probniki_info = {
            'AF': {'name': 'Пробник 1', 'search_terms': ['AF', 'ПРОБНИК 1', 'ПРОБНИК №1', '1.1']},
            'AS': {'name': 'Пробник 2', 'search_terms': ['AS', 'ПРОБНИК 2', 'ПРОБНИК №2', '2.1']},
            'BF': {'name': 'Пробник 3', 'search_terms': ['BF', 'ПРОБНИК 3', 'ПРОБНИК №3', '3.1']},
            'BS': {'name': 'Пробник 4', 'search_terms': ['BS', 'ПРОБНИК 4', 'ПРОБНИК №4', '4.1']},
            'CF': {'name': 'Пробник 5', 'search_terms': ['CF', 'ПРОБНИК 5', 'ПРОБНИК №5', '5.1']},
            'CS': {'name': 'Пробник 6', 'search_terms': ['CS', 'ПРОБНИК 6', 'ПРОБНИК №6', '6.1']},
            'DF': {'name': 'Пробник 7', 'search_terms': ['DF', 'ПРОБНИК 7', 'ПРОБНИК №7', '7.1']},
            'DS': {'name': 'Пробник 8', 'search_terms': ['DS', 'ПРОБНИК 8', 'ПРОБНИК №8', '8.1']}
        }

        # Находим столбцы пробников
        probniki_columns = {}

        for col_idx in headers[19:].index:
            header_text = str(headers[col_idx]).upper().strip()

            if not header_text or header_text == 'NAN':
                continue

            for probnik_key, probnik_data in probniki_info.items():
                for search_term in probnik_data['search_terms']:
                    if search_term.upper() in header_text:
                        if probnik_key not in probniki_columns:
                            probniki_columns[probnik_key] = col_idx
                        break

        print(f"🔍 Найдено пробников: {len(probniki_columns)}")

        if not probniki_columns:
            bot.send_message(message.chat.id, "❌ В файле не найдены пробники")
            return

        # Фильтруем студентов с VK ID
        students_to_process = []
        for original_idx, row in student_rows:
            full_name = row.iloc[1]
            vk_id_raw = row.iloc[2]

            if pd.notna(vk_id_raw) and str(vk_id_raw).isdigit():
                students_to_process.append((original_idx, row))

        # Применяем лимит
        if limit is not None and limit < len(students_to_process):
            students_to_process = students_to_process[:limit]

        total_to_process = len(students_to_process)
        processed_count = 0

        if total_to_process == 0:
            bot.send_message(message.chat.id, "❌ В файле нет студентов с корректными VK ID")
            return

        progress_msg = bot.send_message(
            message.chat.id,
            f"📈 Анализирую {len(probniki_columns)} пробников для {total_to_process} студентов..."
        )

        # Упорядочиваем пробники
        probnik_order = ['AF', 'AS', 'BF', 'BS', 'CF', 'CS', 'DF', 'DS']
        ordered_probniki_names = [probniki_info[key]['name'] for key in probnik_order if key in probniki_columns]

        for original_idx, row in students_to_process:
            full_name = row.iloc[1]
            name = extract_name(full_name)

            # Собираем данные для графика
            probniki_scores = []
            probniki_max_scores = []

            for probnik_key in probnik_order:
                if probnik_key in probniki_columns:
                    col_idx = probniki_columns[probnik_key]
                    stud_val = row[col_idx] if pd.notna(row[col_idx]) else 0
                    max_val = max_scores_row[col_idx] if pd.notna(max_scores_row[col_idx]) else 1

                    try:
                        stud_val = float(stud_val)
                    except:
                        stud_val = 0

                    try:
                        max_val = float(max_val)
                    except:
                        max_val = 1

                    probniki_scores.append(stud_val)
                    probniki_max_scores.append(max_val)

            # ИСПОЛЬЗУЕМ create_detailed_graph ДЛЯ СОЗДАНИЯ ГРАФИКА
            try:
                print(f"🔄 Создаю график пробников для {name}...")

                # Создаем номера для оси X (1, 2, 3, ...)
                lesson_numbers = list(range(1, len(ordered_probniki_names) + 1))

                # Используем проверенную функцию создания графика
                graph_buf = create_detailed_graph(
                    lesson_numbers,
                    probniki_scores,
                    probniki_max_scores,
                    3,  # lives - фиктивное значение
                    f"{name} - Пробники"
                )

                # Создаем текстовую статистику
                probniki_percentages = []
                for score, max_score in zip(probniki_scores, probniki_max_scores):
                    percentage = (score / max_score * 100) if max_score > 0 else 0
                    probniki_percentages.append(percentage)

                avg_percent = sum(probniki_percentages) / len(probniki_percentages) if probniki_percentages else 0

                # Детальная статистика по каждому пробнику
                details = "\n".join([
                    f"• {name}: {percent:.0f} баллов"
                    for name, score, max_score, percent in zip(
                        ordered_probniki_names, probniki_scores, probniki_max_scores, probniki_percentages
                    )
                ])

                caption = (
                    f"📊 *Пробники для {name}*\n\n"
                    f"{details}\n\n"
                    f"📈 Средний балл: {avg_percent:.0f}\n"
                    f"🏆 Лучший результат: {max(probniki_percentages):.0f}"
                )

                # Отправляем график
                bot.send_photo(
                    message.chat.id,
                    graph_buf,
                    caption=caption,
                    parse_mode='Markdown'
                )
                graph_buf.close()
                print(f"✅ График отправлен для {name}")

            except Exception as e:
                print(f"❌ Ошибка создания графика для {name}: {str(e)}")
                import traceback
                print(f"❌ Детали ошибки: {traceback.format_exc()}")

                # Текстовый вывод если график не создался
                results_text = "\n".join([
                    f"• {name}: {(score / max_score * 100) if max_score > 0 else 0:.0f} баллов"
                    for name, score, max_score in zip(ordered_probniki_names, probniki_scores, probniki_max_scores)
                ])

                bot.send_message(
                    message.chat.id,
                    f"📊 *Пробники для {name}*\n\n{results_text}",
                    parse_mode='Markdown'
                )

            processed_count += 1

            # Обновляем прогресс
            if processed_count % 2 == 0:
                try:
                    bot.edit_message_text(
                        f"📈 Обработано {processed_count}/{total_to_process} студентов...",
                        message.chat.id,
                        progress_msg.message_id
                    )
                except:
                    pass

            # Задержка между студентами
            time.sleep(2)

        # Завершение
        try:
            bot.delete_message(message.chat.id, progress_msg.message_id)
        except:
            pass

        bot.send_message(
            message.chat.id,
            f"✅ Анализ пробников завершен для {processed_count} студентов\n"
            f"📊 Обработано пробников: {len(probniki_columns)}"
        )

    except Exception as e:
        print(f"❌ Критическая ошибка: {e}")
        bot.send_message(message.chat.id, f"❌ Ошибка анализа пробников: {e}")

@bot.message_handler(func=lambda message: user_data.get(message.from_user.id, {}).get('step') == 'waiting_stats_excel',
                     content_types=['document'])
def handle_stats_excel(message):
    """Обрабатывает загрузку Excel файла для статистики"""
    user_id = message.from_user.id

    if not message.document.file_name.endswith(('.xlsx', '.xls')):
        bot.send_message(message.chat.id, "❌ Пожалуйста, загрузите Excel файл (.xlsx или .xls)")
        return

    try:
        file_info = bot.get_file(message.document.file_id)
        downloaded_file = bot.download_file(file_info.file_path)

        file_path = f"temp_stats_{user_id}_{message.document.file_name}"
        with open(file_path, 'wb') as new_file:
            new_file.write(downloaded_file)

        user_data[user_id]['excel_file'] = file_path
        user_data[user_id]['step'] = 'waiting_stats_lesson_range'

        bot.send_message(
            message.chat.id,
            "✅ Файл загружен! Теперь введите диапазон домашек (например: 12-21):"
        )

    except Exception as e:
        bot.send_message(message.chat.id, f"❌ Ошибка загрузки файла: {e}")

@bot.message_handler(
    func=lambda message: user_data.get(message.from_user.id, {}).get('step') == 'waiting_stats_lesson_range')
def handle_stats_lesson_range(message):
    """Обрабатывает ввод диапазона для статистики"""
    user_id = message.from_user.id

    try:
        lesson_range = message.text.strip()

        # Проверяем формат диапазона
        if '-' not in lesson_range:
            bot.send_message(message.chat.id, "❌ Неверный формат. Используйте формат: 12-21")
            return

        user_data[user_id]['lesson_range'] = lesson_range
        user_data[user_id]['step'] = 'waiting_stats_limit'  # МЕНЯЕМ ШАГ

        bot.send_message(
            message.chat.id,
            "👥 *Сколько учеников вывести?*\n\n"
            "Введите число (например: 10) или 'все' для вывода всех учеников:",
            parse_mode='Markdown'
        )

    except Exception as e:
        bot.send_message(message.chat.id, f"❌ Ошибка: {e}")

@bot.message_handler(func=lambda message: user_data.get(message.from_user.id, {}).get('step') == 'waiting_stats_limit')
def handle_stats_limit(message):
    """Обрабатывает ввод лимита учеников для статистики"""
    user_id = message.from_user.id

    try:
        limit_input = message.text.strip().lower()

        if limit_input == 'все':
            user_data[user_id]['limit'] = None  # Без лимита
        else:
            try:
                limit = int(limit_input)
                if limit <= 0:
                    bot.send_message(message.chat.id, "❌ Число должно быть больше 0")
                    return
                user_data[user_id]['limit'] = limit
            except ValueError:
                bot.send_message(message.chat.id, "❌ Введите число или 'все'")
                return

        # Сразу начинаем обработку
        bot.send_message(message.chat.id, "⏳ Начинаю анализ данных...")
        generate_and_send_stats(
            message,
            user_data[user_id]['excel_file'],
            user_data[user_id]['lesson_range'],
            user_data[user_id].get('limit')  # ПЕРЕДАЕМ ЛИМИТ
        )

        # Очистка
        if os.path.exists(user_data[user_id]['excel_file']):
            os.remove(user_data[user_id]['excel_file'])
        user_data[user_id] = {'step': 'main_menu'}

    except Exception as e:
        bot.send_message(message.chat.id, f"❌ Ошибка: {e}")


def extract_lesson_number(header):
    """Извлекает номер урока из заголовка, включая пробники с дробными номерами"""
    if pd.isna(header):
        return None

    header_str = str(header).strip()

    # Для формата "24.1. Пробник №1" - возвращаем дробное число 24.1
    match_float = re.match(r'^(\d+)\.(\d+)', header_str)
    if match_float:
        whole_part = int(match_float.group(1))
        decimal_part = int(match_float.group(2))
        return whole_part + decimal_part * 0.1

    # Для обычных ДЗ: "24. Обычное ДЗ" - возвращаем целое число 24
    match_int = re.match(r'^(\d+)', header_str)
    return int(match_int.group(1)) if match_int else None

def generate_and_send_stats(message, excel_file, lesson_range, limit=None):
    """Генерирует и отправляет статистику - с лимитом учеников"""
    try:
        # Читаем Excel файл
        df_full = pd.read_excel(excel_file, header=None)
        headers = df_full.iloc[0]
        max_scores_row = df_full.iloc[6]
        student_rows = list(df_full.iloc[7:].iterrows())

        # Выбор столбцов по диапазону
        hw_columns = []
        lesson_numbers = []
        for col_idx in headers[19:].index:
            num = extract_lesson_number(headers[col_idx])
            if num is not None and num in parse_lesson_range(lesson_range):
                hw_columns.append(col_idx)
                lesson_numbers.append(num)

        if not hw_columns:
            bot.send_message(message.chat.id, "❌ Не найдено ДЗ в указанном диапазоне")
            return

        combined = sorted(zip(lesson_numbers, hw_columns))
        lesson_numbers, hw_columns = zip(*combined)
        lesson_numbers = list(lesson_numbers)
        hw_columns = list(hw_columns)

        # Фильтруем студентов с VK ID
        students_to_process = []
        for original_idx, row in student_rows:
            full_name = row.iloc[1]
            vk_id_raw = row.iloc[2]

            if pd.notna(vk_id_raw) and str(vk_id_raw).isdigit():
                students_to_process.append((original_idx, row))

        # Применяем лимит если указан
        original_count = len(students_to_process)
        if limit is not None and limit < len(students_to_process):
            students_to_process = students_to_process[:limit]
            limit_text = f" (лимит: {limit})"
        else:
            limit_text = ""

        total_to_process = len(students_to_process)
        processed_count = 0

        if total_to_process == 0:
            bot.send_message(message.chat.id, "❌ В файле нет студентов с корректными VK ID")
            return

        # Отправляем начальное сообщение
        progress_msg = bot.send_message(
            message.chat.id,
            f"📊 Обрабатываю {total_to_process} студентов{limit_text}..."
        )

        # Обрабатываем студентов группами
        batch_size = 3
        for batch_start in range(0, len(students_to_process), batch_size):
            batch_end = min(batch_start + batch_size, len(students_to_process))
            batch = students_to_process[batch_start:batch_end]

            for original_idx, row in batch:
                full_name = row.iloc[1]
                vk_id_raw = row.iloc[2]
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

                # Определяем категорию студента
                hard_submitted_all = len(hard_scores) == sum(1 for col in hw_columns if max_scores_row[col] > 1)
                max_possible_score = sum(max_scores_row[col] for col in hw_columns if not pd.isna(max_scores_row[col]))
                ratio = total_score / max_possible_score if max_possible_score > 0 else 0

                if hard_submitted_all and ratio >= 0.70:
                    category = 1
                    category_emoji = "🔥"
                elif ratio >= 0.42:
                    category = 2
                    category_emoji = "📈"
                else:
                    category = 3
                    category_emoji = "📚"

                avg_percent, best_entries = get_best_hw_info(headers, hw_columns, student_scores, max_scores,
                                                             lesson_numbers)
                best_hw_str = format_best_hw(best_entries)

                # Генерируем сообщение со статистикой
                stats_message = generate_stats_message(
                    name, category_emoji, len(hw_columns), test_done_count, test_total_count,
                    avg_percent, best_hw_str, lives, lesson_range, category
                )

                try:
                    graph_buf = create_detailed_graph(lesson_numbers, student_scores, max_scores, lives, name)

                    # Отправляем сообщение с графиком
                    bot.send_photo(
                        message.chat.id,
                        graph_buf,
                        caption=stats_message,
                        parse_mode='Markdown'
                    )
                    graph_buf.close()

                except Exception as e:
                    print(f"Ошибка создания/отправки графика для {name}: {e}")
                    try:
                        bot.send_message(
                            message.chat.id,
                            f"📊 *Статистика для {name}* (график не удалось создать)\n\n{stats_message}",
                            parse_mode='Markdown'
                        )
                    except Exception as send_error:
                        print(f"Ошибка отправки сообщения для {name}: {send_error}")
                        continue

                processed_count += 1

            try:
                if limit_text:
                    progress_text = f"📊 Обработано {processed_count}/{total_to_process} студентов{limit_text}..."
                else:
                    progress_text = f"📊 Обработано {processed_count}/{total_to_process} студентов..."

                bot.edit_message_text(
                    progress_text,
                    message.chat.id,
                    progress_msg.message_id
                )
            except:
                pass

            # Задержка между группами
            if batch_end < len(students_to_process):
                time.sleep(5)
                print(f"⏳ Пауза между группами... Обработано: {processed_count}/{total_to_process}")

        try:
            bot.delete_message(message.chat.id, progress_msg.message_id)
        except:
            pass

        # Итоговое сообщение с учетом лимита
        if limit is not None and limit < original_count:
            bot.send_message(
                message.chat.id,
                f"✅ Статистика сгенерирована для {processed_count} студентов (выведено по лимиту {limit} из {original_count})"
            )
        else:
            bot.send_message(
                message.chat.id,
                f"✅ Статистика сгенерирована для {processed_count} студентов"
            )

    except Exception as e:
        bot.send_message(message.chat.id, f"❌ Ошибка генерации статистики: {e}")

def generate_stats_message(name, emoji, total_hw_count, test_done_count, test_total_count,
                           avg_percent, best_hw_str, lives, lesson_range, category):

    lives_status = " Ни одной жизни не потеряно! 🚘" if lives >= 3 else f" Потеряно жизней: {3 - lives}"

    category_text = {
        1: "🔥 *Категория 1 - Отличные результаты*",
        2: "📈 *Категория 2 - Хорошие результаты*",
        3: "📚 *Категория 3 - Требует улучшения*"
    }.get(category, "")

    message = (
        f" *Статистика для {name}*\n"
        f"{category_text}\n"
        f"*Диапазон:* занятия {lesson_range}\n\n"
        f"📊 *Общая статистика за {total_hw_count} занятий:*\n"
        f"- Выполнено тестовых ДЗ: {test_done_count}/{test_total_count}\n"
    )

    if avg_percent > 0:
        message += f"- Средний балл за сложные ДЗ: {avg_percent:.1f}%\n"

    if best_hw_str and best_hw_str != "—":
        message += f"- Лучшие результаты:\n{best_hw_str}\n\n"
    else:
        message += "\n"

    message += f"💫 {lives_status}"

    return message

@bot.message_handler(content_types=['document'])
def handle_document(message):
    """Обработка загруженных Excel файлов для обычных режимов"""
    user_id = message.from_user.id
    current_step = user_data.get(user_id, {}).get('step')

    # Если это режим статистики - пропускаем
    if current_step in ['waiting_stats_excel', 'waiting_stats_lesson_range']:
        return

    if current_step != 'waiting_excel':
        bot.send_message(message.chat.id, "❌ Сначала выберите действие в меню")
        return

    if not message.document.file_name.endswith(('.xlsx', '.xls')):
        bot.send_message(message.chat.id, "❌ Пожалуйста, загрузите Excel файл (.xlsx или .xls)")
        return

    try:
        file_info = bot.get_file(message.document.file_id)
        downloaded_file = bot.download_file(file_info.file_path)

        file_path = f"temp_{user_id}_{message.document.file_name}"
        with open(file_path, 'wb') as new_file:
            new_file.write(downloaded_file)

        user_data[user_id]['excel_file'] = file_path
        user_data[user_id]['step'] = 'waiting_vk_token'

        bot.send_message(
            message.chat.id,
            f"✅ Файл *{message.document.file_name}* загружен!\n\n"
            "Теперь введите ваш VK API токен:",
            parse_mode='Markdown'
        )

    except Exception as e:
        bot.send_message(message.chat.id, f"❌ Ошибка загрузки файла: {e}")


# ... остальные обработчики для preview/send режимов остаются без изменений ...
@bot.message_handler(func=lambda message: user_data.get(message.from_user.id, {}).get('step') == 'waiting_vk_token')
def handle_vk_token(message):
    """Обработка VK токена"""
    user_id = message.from_user.id
    user_data[user_id]['vk_token'] = message.text
    user_data[user_id]['step'] = 'waiting_block_number'

    bot.send_message(
        message.chat.id,
        "🔢 Введите номер блока (например: 9):",
        reply_markup=ReplyKeyboardRemove()
    )


@bot.message_handler(func=lambda message: user_data.get(message.from_user.id, {}).get('step') == 'waiting_block_number')
def handle_block_number(message):
    """Обработка номера блока"""
    user_id = message.from_user.id

    try:
        user_data[user_id]['block_number'] = int(message.text)
        user_data[user_id]['step'] = 'waiting_lesson_range'

        bot.send_message(
            message.chat.id,
            "📚 Введите диапазон домашек (например: 12-17):"
        )
    except ValueError:
        bot.send_message(message.chat.id, "❌ Введите число для номера блока")


@bot.message_handler(func=lambda message: user_data.get(message.from_user.id, {}).get('step') == 'waiting_lesson_range')
def handle_lesson_range(message):
    """Обработка диапазона уроков"""
    user_id = message.from_user.id
    user_data[user_id]['lesson_range'] = message.text

    if user_data[user_id]['mode'] == 'send':
        user_data[user_id]['step'] = 'waiting_skip_rows'
        bot.send_message(
            message.chat.id,
            "🚫 Введите номера студентов для пропуска (через запятую):\n"
            "*Пример:* Чтобы пропустить 1го и 3го студента в списке, введите `1,3`\n"
            "Или отправьте `нет` чтобы не пропускать никого:",
            parse_mode='Markdown'
        )
    else:
        show_confirmation(user_id, message)


@bot.message_handler(func=lambda message: user_data.get(message.from_user.id, {}).get('step') == 'waiting_skip_rows')
def handle_skip_rows(message):
    """Обработка пропуска студентов"""
    user_id = message.from_user.id

    if message.text.lower() in ['нет', 'no', '']:
        user_data[user_id]['skip_rows'] = ''
    else:
        user_data[user_id]['skip_rows'] = message.text
    show_confirmation(user_id, message)


def show_confirmation(user_id, message):
    """Показывает подтверждение отправки"""
    data = user_data[user_id]

    confirm_text = (
        "🚨 *ПОДТВЕРЖДЕНИЕ ОТПРАВКИ*\n\n"
        f"• Режим: {data['mode'].upper()}\n"
        f"• Блок: {data['block_number']}\n"
        f"• Диапазон: {data['lesson_range']}\n"
    )

    if data['mode'] == 'send' and data.get('skip_rows'):
        confirm_text += f"• Пропуск студентов: {data['skip_rows']}\n"

    confirm_text += "\n*Точно отправить сообщения? да/нет*"

    bot.send_message(
        message.chat.id,
        confirm_text,
        parse_mode='Markdown'
    )

    user_data[user_id]['step'] = 'waiting_confirmation'


@bot.message_handler(func=lambda message: user_data.get(message.from_user.id, {}).get('step') == 'waiting_confirmation')
def handle_confirmation(message):
    """Обработка подтверждения пользователя"""
    user_id = message.from_user.id
    user_response = message.text.lower().strip()

    if user_response in ['да', 'yes', 'y', 'д']:
        bot.send_message(message.chat.id, "🔄 Запускаю выполнение...")
        launch_program(user_id, message)

    elif user_response in ['нет', 'no', 'n', 'н']:
        bot.send_message(message.chat.id, "❌ Отправка отменена")
        user_data[user_id] = {'step': 'main_menu'}
        start(message)

    else:
        bot.send_message(message.chat.id, "❌ Пожалуйста, ответьте 'да' или 'нет'")


def launch_program(user_id, message):
    """Запуск выбранного режима"""
    data = user_data[user_id]

    try:
        if data['mode'] == 'preview':
            result = preview_mode(
                vk_token=data['vk_token'],
                block_number=data['block_number'],
                lesson_range=data['lesson_range'],
                excel_file=data['excel_file']
            )
            bot.send_message(message.chat.id, "✅ Превью отправлено куратору!")

        elif data['mode'] == 'send':
            result = send_mode(
                vk_token=data['vk_token'],
                block_number=data['block_number'],
                lesson_range=data['lesson_range'],
                skip_rows_input=data.get('skip_rows', ''),
                excel_file=data['excel_file']
            )
            bot.send_message(message.chat.id, "✅ Сообщения отправлены студентам!")

        # Очищаем временный файл
        if os.path.exists(data['excel_file']):
            os.remove(data['excel_file'])

        user_data[user_id] = {'step': 'main_menu'}
        start(message)

    except Exception as e:
        bot.send_message(message.chat.id, f"❌ Ошибка: {e}")
        user_data[user_id] = {'step': 'main_menu'}


if __name__ == "__main__":
    print("🤖 Бот запущен!")
    bot.infinity_polling()