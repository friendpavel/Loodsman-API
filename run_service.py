# -*- coding: utf-8 -*-
"""
Запуск FastAPI сервиса
"""

import uvicorn

if __name__ == "__main__":
    print("=" * 60)
    print("Запуск Loodsman API Service")
    print("=" * 60)
    print("Адрес: http://localhost:8000")
    print("Документация: http://localhost:8000/docs")
    print("\nНажмите Ctrl+C для остановки")
    print("=" * 60)
    
    uvicorn.run(
        "app.main:app",
        host="0.0.0.0",
        port=8000,
        reload=True,
        log_level="info"
    )