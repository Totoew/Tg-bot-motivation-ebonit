#!/usr/bin/env python3
# -*- coding: utf-8 -*-
import re
import sys
import os
import platform

current_dir = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, current_dir)

def clear_screen():
    """Очистка экрана для разных ОС"""
    if platform.system() == "Windows":
        os.system('cls')
    else:
        os.system('clear')

def print_header():
    """Красивый заголовок"""
    clear_screen()
    print("🚀 PHYSICS MOTIVATION BOT")
    print("=" * 50)
    print("Автоматическая отправка мотивационных сообщений")
    print("=" * 50)
    print()

def get_input(prompt, required=True):
    """Безопасный ввод данных"""
    while True:
        try:
            value = input(prompt).strip()
            if required and not value:
                print("❌ Это поле обязательно для заполнения!")
                continue
            return value
        except KeyboardInterrupt:
            print("\n\n👋 Выход из программы.")
            sys.exit(0)
        except Exception as e:
            print(f"❌ Ошибка ввода: {e}")


def select_excel_file():
    """Выбор файла Excel"""
    import os
    import re

    print("\n📁 ВЫБОР ФАЙЛА EXCEL")
    print("-" * 30)

    default_path = "/Users/daniltotoev/Downloads/Новая таблица.xlsx"

    print(f"По умолчанию: {default_path}")
    print("Нажмите Enter чтобы использовать путь по умолчанию")
    print("Или введите полный путь к вашему файлу Excel")
    print("Пример: C:\\Users\\Имя\\Downloads\\таблица.xlsx")

    file_path = get_input("Путь к файлу: ", required=False)

    if not file_path:
        file_path = default_path
    else:
        # Убираем кавычки если пользователь их ввел
        file_path = file_path.strip('"\'')
        # Заменяем прямые слеши на обратные для Windows
        file_path = file_path.replace('/', '\\')
        # Если путь не полный (без диска), добавляем текущий диск
        if not re.match(r'^[a-zA-Z]:', file_path):
            current_drive = os.path.splitdrive(os.getcwd())[0]
            file_path = current_drive + '\\' + file_path.lstrip('\\')

    # Проверяем существование файла
    if not os.path.exists(file_path):
        print(f"❌ Файл не найден: {file_path}")
        print("Пожалуйста, укажите правильный путь к файлу Excel")
        return select_excel_file()

    print(f"✅ Файл найден: {os.path.basename(file_path)}")
    return file_path

def get_vk_token():
    """Получение VK токена"""
    print("\n🔑 VK API ТОКЕН")
    print("-" * 30)
    print("Токен необходим для отправки сообщений через VK API")
    print("Получить токен можно: https://vk.com/dev/access_token")
    print()

    token = get_input("Введите ваш VK токен: ")
    return token

def get_curators_vk_ids():
    """Получение VK ID кураторов для превью"""
    print("\n👥 VK ID КУРАТОРОВ ДЛЯ ПРЕВЬЮ")
    print("-" * 30)
    print("Введите VK ID кураторов, которым отправить превью")
    print("Можно указать несколько ID через запятую")
    print("Пример: 550891157, 123456789, 987654321")
    print()

    curators_input = get_input("VK ID кураторов: ")

    # Парсим введенные ID
    curators = []
    if curators_input:
        try:
            curators = [int(id.strip()) for id in curators_input.split(',') if id.strip().isdigit()]
        except ValueError:
            print("❌ Ошибка в формате VK ID. Используйте только числа через запятую.")
            return get_curators_vk_ids()

    if not curators:
        print("❌ Не указаны VK ID кураторов")
        return get_curators_vk_ids()

    print(f"✅ Будут отправлены кураторам: {curators}")
    return curators

def get_block_settings():
    """Настройки блока и домашек"""
    print("\n⚙️  НАСТРОЙКИ ОТПРАВКИ")
    print("-" * 30)

    block_number = int(get_input("Номер блока: "))
    lessons_range = get_input("Диапазон домашек (например, 12-17): ")

    return block_number, lessons_range

def get_skip_rows():
    """Настройка пропуска строк"""
    print("\n📋 ДОПОЛНИТЕЛЬНЫЕ НАСТРОЙКИ")
    print("-" * 30)
    print("Можно пропустить определенные строки из Excel")
    print("Например: 5,12,18 или оставить пустым")

    skip_rows = get_input("Пропустить строки (через запятую): ", required=False)
    return skip_rows

def main_menu():
    """Главное меню"""
    print_header()

    print("Выберите режим работы:")
    print("1. 👁️  ПРЕВЬЮ - отправка примеров кураторам")
    print("2. 📤 ОТПРАВКА - отправка сообщений ученикам")
    print("3. ℹ️  СПРАВКА")
    print("4. ❌ ВЫХОД")
    print()

    choice = get_input("Ваш выбор (1-4): ")
    return choice

def show_help():
    """Показать справку"""
    print_header()
    print("📖 СПРАВКА ПО ПРОГРАММЕ")
    print("=" * 50)
    print()
    print("🎯 НАЗНАЧЕНИЕ:")
    print("   Программа для автоматической отправки мотивационных")
    print("   сообщений ученикам на основе их успеваемости")
    print()
    print("📁 ТРЕБУЕМЫЕ ФАЙЛЫ:")
    print("   • Файл Excel с данными учеников")
    print("   • Папка templates/ с шаблонами сообщений")
    print("   • Файл static_data.py с видео и цитатами")
    print()
    print("🔑 VK ТОКЕН:")
    print("   • Получить: из сообщества ВК")
    print("   • Нужны права: messages, friends")
    print()
    print("⚙️  РЕЖИМЫ РАБОТЫ:")
    print("   • ПРЕВЬЮ - тестовая отправка кураторам")
    print("   • ОТПРАВКА - реальная отправка ученикам")
    print()
    input("Нажмите Enter чтобы вернуться в меню...")

def run_program():
    """Основная логика программы"""
    try:
        from send_final import preview_mode, send_mode

        # Получаем настройки
        excel_file = select_excel_file()
        vk_token = get_vk_token()
        block_number, lessons_range = get_block_settings()

        # Устанавливаем глобальную переменную
        import send_final
        send_final.EXCEL_FILE = excel_file

        print("\n⏳ Запуск программы...")
        print("Пожалуйста, подождите...")

        # Запускаем выбранный режим
        choice = get_input("\nВыберите режим (1-превью, 2-отправка): ")

        if choice == "1":
            preview_mode(vk_token=vk_token, block_number=block_number, lesson_range=lessons_range)
        elif choice == "2":
            skip_rows = get_skip_rows()
            send_mode(vk_token=vk_token, block_number=block_number, lesson_range=lessons_range,
                      skip_rows_input=skip_rows)
        else:
            print("❌ Неверный выбор режима")
            return

        print("\n✅ Программа успешно завершена!")
        input("\nНажмите Enter чтобы продолжить...")

    except ImportError as e:
        print(f"❌ Ошибка импорта модулей: {e}")
        print("Убедитесь что все файлы программы находятся в одной папке")
        input("\nНажмите Enter чтобы продолжить...")
    except Exception as e:
        print(f"❌ Произошла ошибка: {e}")
        input("\nНажмите Enter чтобы продолжить...")

def main():
    """Главная функция"""
    try:
        while True:
            choice = main_menu()

            if choice == "1" or choice == "2":
                run_program()
            elif choice == "3":
                show_help()
            elif choice == "4":
                print("\n👋 До свидания!")
                break
            else:
                print("❌ Неверный выбор. Попробуйте снова.")
                input("Нажмите Enter чтобы продолжить...")

    except KeyboardInterrupt:
        print("\n\n👋 Выход из программы.")
    except Exception as e:
        print(f"\n❌ Критическая ошибка: {e}")
        input("Нажмите Enter чтобы выйти...")

if __name__ == "__main__":
    main()
    