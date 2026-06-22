# -*- coding: utf-8 -*-
"""
Удаление Loodsman API Windows Service
Запуск от администратора: python uninstall_service.py
"""

import os
import sys
import subprocess

if __name__ == "__main__":
    # Останавливаем сервис
    subprocess.run([sys.executable, "windows_service_impl.py", "stop"])
    # Удаляем сервис
    subprocess.run([sys.executable, "windows_service_impl.py", "remove"])
    print("✓ Сервис LoodsmanAPI остановлен и удалён")
