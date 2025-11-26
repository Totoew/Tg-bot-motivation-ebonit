#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import os
import sys

# Добавляем текущую директорию в путь
current_dir = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, current_dir)

def main():
    """Главная функция запуска"""
    try:
        from physics_bot_console import main as console_main
        console_main()
    except ImportError as e:
        print(f"❌ Ошибка импорта: {e}")
        print("Убедитесь, что все файлы находятся в одной папке")
        input("Нажмите Enter для выхода...")
    except Exception as e:
        print(f"❌ Критическая ошибка: {e}")
        input("Нажмите Enter для выхода...")

if __name__ == "__main__":
    main()
