# -*- coding: utf-8 -*-
"""
Установка Loodsman API как Windows Service
Запуск от администратора: python install_service.py
"""

import os
import sys
import subprocess

SERVICE_NAME = "LoodsmanAPI"
SERVICE_DISPLAY_NAME = "Loodsman API Service"
SERVICE_DESCRIPTION = "FastAPI сервис для выгрузки данных из ЛОЦМАН"

def install_service():
    """Установка сервиса"""
    # Получаем полный путь к текущему интерпретатору
    python_exe = sys.executable
    
    # Получаем путь к сервисному скрипту
    current_dir = os.path.dirname(os.path.abspath(__file__))
    service_script = os.path.join(current_dir, "windows_service_impl.py")
    
    # Команда для установки
    cmd = [
        python_exe, service_script, "install"
    ]
    
    result = subprocess.run(cmd, capture_output=True, text=True)
    if result.returncode == 0:
        print(f"✓ Сервис '{SERVICE_NAME}' успешно установлен")
    else:
        print(f"✗ Ошибка установки: {result.stderr}")

def start_service():
    """Запуск сервиса"""
    cmd = [
        sys.executable, 
        os.path.join(os.path.dirname(__file__), "windows_service_impl.py"), 
        "start"
    ]
    result = subprocess.run(cmd, capture_output=True, text=True)
    if result.returncode == 0:
        print(f"✓ Сервис '{SERVICE_NAME}' запущен")
    else:
        print(f"✗ Ошибка запуска: {result.stderr}")

if __name__ == "__main__":
    import argparse
    parser = argparse.ArgumentParser()
    parser.add_argument("action", choices=["install", "start", "stop", "remove"], 
                       help="Действие с сервисом")
    args = parser.parse_args()
    
    if args.action == "install":
        install_service()
    elif args.action == "start":
        start_service()
    elif args.action == "stop":
        subprocess.run([sys.executable, "windows_service_impl.py", "stop"])
    elif args.action == "remove":
        subprocess.run([sys.executable, "windows_service_impl.py", "remove"])
