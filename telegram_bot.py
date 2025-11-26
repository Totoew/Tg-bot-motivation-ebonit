import telebot
from telebot.types import InlineKeyboardMarkup, InlineKeyboardButton, ReplyKeyboardRemove
import os
import pandas as pd
from send_final import preview_mode, send_mode
with open(r"C:\Users\Пользователь\Desktop\bot-token.txt", 'r', encoding='utf-8') as file:
    content = file.read()

bot = telebot.TeleBot(content)

# Хранилище состояний пользователей
user_data = {}

def debug_user_data(user_id):
    """Функция для отладки - показывает текущее состояние пользователя"""
    if user_id in user_data:
        print(f"🔍 DEBUG user_{user_id}: {user_data[user_id]}")
    else:
        print(f"🔍 DEBUG user_{user_id}: NO DATA")

@bot.message_handler(commands=['start'])
def start(message):
    """Главное меню"""
    user_id = message.from_user.id
    user_data[user_id] = {'step': 'main_menu'}

    keyboard = InlineKeyboardMarkup()
    keyboard.add(InlineKeyboardButton("👁️ ПРЕВЬЮ", callback_data="preview"))
    keyboard.add(InlineKeyboardButton("📤 ОТПРАВКА", callback_data="send"))
    keyboard.add(InlineKeyboardButton("⚙️ НАСТРОЙКИ", callback_data="settings"))
    keyboard.add(InlineKeyboardButton("🎥 ВИДЕО-ИНСТРУКЦИЯ", url="https://docs.google.com/document/d/1utGllba1nr1QqmnLpOK03hwYpY87NmVIyDgsfk3kJpA/edit?usp=sharing"))

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

    elif call.data == "settings":
        show_instructions(call.message)

    elif call.data == "back_to_menu":
        # Удаляем сообщение с инструкцией и возвращаем в меню
        try:
            bot.delete_message(call.message.chat.id, call.message.message_id)
        except:
            pass
        start(call.message)

def show_instructions(message):
    """Показывает инструкцию по использованию бота"""

    instructions = """
🔧 ИНСТРУКЦИЯ ПО ИСПОЛЬЗОВАНИЮ БОТА
Видео по настройке: 

🎯 РЕЖИМ «ПРЕВЬЮ»:
• Отправляет примеры сообщений кураторам для проверки
• Показывает по 2 студента из каждой категории
• Включает графики успеваемости и видео
• Не отправляет сообщения реальным студентам

📤 РЕЖИМ «ОТПРАВКА»:
• Отправляет сообщения реальным студентам
• Можно указать номера студентов для пропуска
• Каждый студент получает 2 сообщения: текст+график и видео

📊 КАТЕГОРИИ СТУДЕНТОВ:
• Категория 1 - лучшие результаты (≥70% + все сложные ДЗ)
• Категория 2 - хорошие результаты (≥42%)
• Категория 3 - нужно улучшить результат (<42%)

🔄 ПРОЦЕСС РАБОТЫ:
1. Выберите режим (Превью/Отправка)
2. Загрузите Excel файл
3. Введите VK токен
4. Укажите номер блока (например: 9)
5. Укажите диапазон ДЗ (например: 12-17)
6. Для отправки - укажите студентов для пропуска
7. Подтвердите отправку

⏱ *ВРЕМЯ ОТПРАВКИ:*
• Зависит от количества студентов
• В среднем 5-10 секунд на студента
• Прогресс отображается в реальном времени

❓ ЧАСТЫЕ ПРОБЛЕМЫ:
• Файл не загружается - проверьте формат (.xlsx/.xls)
• Сообщения не отправляются - проверьте VK токен
• Студенты не находятся - проверьте VK ID в файле

📞 ПОДДЕРЖКА:
Если возникли проблемы - обратитесь к разработчику бота @totoevv.
"""
    # Создаем клавиатуру для возврата в меню
    keyboard = InlineKeyboardMarkup()
    keyboard.add(InlineKeyboardButton("⬅️ В главное меню", callback_data="back_to_menu"))

    bot.send_message(
        message.chat.id,
        instructions,
        reply_markup=keyboard,
        parse_mode='Markdown'
    )

@bot.message_handler(content_types=['document'])
def handle_document(message):
    """Обработка загруженных Excel файлов"""
    user_id = message.from_user.id

    if user_data.get(user_id, {}).get('step') != 'waiting_excel':
        bot.send_message(message.chat.id, "❌ Сначала выберите действие в меню")
        return

    # Проверяем что это Excel файл
    if not message.document.file_name.endswith(('.xlsx', '.xls')):
        bot.send_message(message.chat.id, "❌ Пожалуйста, загрузите Excel файл (.xlsx или .xls)")
        return

    try:
        # Скачиваем файл
        file_info = bot.get_file(message.document.file_id)
        downloaded_file = bot.download_file(file_info.file_path)

        # Сохраняем временно
        file_path = f"temp_{user_id}_{message.document.file_name}"
        with open(file_path, 'wb') as new_file:
            new_file.write(downloaded_file)

        # Сохраняем путь в данные пользователя
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

@bot.message_handler(
    func=lambda message: user_data.get(message.from_user.id, {}).get('step') == 'waiting_content_block')  # ← ИЗМЕНИЛИ ФИЛЬТР
def handle_block_for_content(message):
    """Обрабатывает ввод номера блока и показывает контент"""
    user_id = message.from_user.id

    try:
        block_number = int(message.text.strip())
        show_motivation_content(message, block_number)

    except ValueError:
        bot.send_message(
            message.chat.id,
            "❌ Пожалуйста, введите корректный номер блока (цифру):"
        )

def show_motivation_content(message, block_number):
    """Показывает мотивационный контент для указанного блока"""

    # Импортируем данные из static_data
    from static_data import quotes, motivation_videos, future_wishes

    # Проверяем, существует ли такой блок в данных
    if block_number not in quotes or block_number not in motivation_videos or block_number not in future_wishes:
        bot.send_message(
            message.chat.id,
            f"❌ Контент для блока {block_number} не найден.\n"
            f"Доступные блоки: {list(quotes.keys())}",
            parse_mode='Markdown'
        )
        user_data[message.from_user.id] = {'step': 'main_menu'}
        return

    # Получаем контент для блока
    quote = quotes[block_number]
    video_url = motivation_videos[block_number]
    wish = future_wishes[block_number]

    # Формируем сообщение
    content_message = (
        f"📚 *Мотивационные материалы для блока {block_number}*\n\n"
        f"💫 *Цитата:*\n{quote}\n\n"
        f"🎥 *Видео:* {video_url}\n\n"
        f"✨ *Пожелание:*\n{wish}"
    )

    # Создаем клавиатуру для возврата
    keyboard = InlineKeyboardMarkup()
    keyboard.add(InlineKeyboardButton("📝 Показать другой блок", callback_data="show_content"))
    keyboard.add(InlineKeyboardButton("⬅️ В главное меню", callback_data="back_to_menu"))

    bot.send_message(
        message.chat.id,
        content_message,
        reply_markup=keyboard,
        parse_mode='Markdown'
    )

    # Сбрасываем состояние пользователя
    user_data[message.from_user.id] = {'step': 'main_menu'}

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

    # Если режим отправки - запрашиваем пропуск строк
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
        # Для превью сразу показываем подтверждение
        show_confirmation(user_id, message)

@bot.message_handler(func=lambda message: user_data.get(message.from_user.id, {}).get('step') == 'waiting_skip_rows')
def handle_skip_rows(message):
    """Обработка пропуска строк"""
    user_id = message.from_user.id

    # Если пользователь ввел "нет" или пустое значение - не пропускаем никого
    if message.text.lower() in ['нет', 'нет', 'no', '']:
        user_data[user_id]['skip_rows'] = ''
    else:
        user_data[user_id]['skip_rows'] = message.text

    # Показываем подтверждение
    show_confirmation(user_id, message)

def show_confirmation(user_id, message):
    """Показать подтверждение отправки"""
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

        user_data[user_id] = {'step': 'main_menu'}
        start(message)

    except Exception as e:
        bot.send_message(message.chat.id, f"❌ Ошибка: {e}")
        user_data[user_id] = {'step': 'main_menu'}

if __name__ == "__main__":
    print("🤖 Бот запущен!")
    bot.infinity_polling()