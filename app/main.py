# -*- coding: utf-8 -*-
"""
FastAPI сервис для выгрузки данных из ЛОЦМАН
Swagger документация: /docs
"""

from fastapi import FastAPI, HTTPException
from fastapi.responses import Response
import io
from openpyxl import Workbook

from .models import BuildTreeRequest
from .loodsman_client import LoodsmanClient

# Создаем клиент (один на всё приложение)
client = LoodsmanClient()

# Создаем FastAPI приложение
app = FastAPI(
    title="Loodsman API",
    description="API для выгрузки дерева изделий и материалов из ЛОЦМАН",
    version="2.0.0",
    docs_url="/docs",
    redoc_url="/redoc"
)


@app.get("/")
async def root():
    """Корневой endpoint"""
    return {
        "service": "Loodsman API",
        "version": "2.0.0",
        "docs": "/docs",
        "health": "/health"
    }


@app.get("/health")
async def health_check():
    """Проверка работоспособности сервиса"""
    return {"status": "ok", "service": "loodsman-api"}


@app.post("/build_json")
async def build_json(request: BuildTreeRequest):
    """
    Построение дерева и возврат в JSON формате.
    """
    try:
        tree, report = client.get_tree(request.start_id, request.max_depth, request.db_name)
        
        if not tree:
            return {
                "success": False,
                "error": f"Объект с ID {request.start_id} не найден"
            }
        
        total_objects = len(report)
        materials_count = len([r for r in report if r['Обозначение материала']])
        
        return {
            "success": True,
            "start_id": request.start_id,
            "max_depth": request.max_depth,
            "total_objects": total_objects,
            "materials_count": materials_count,
            "tree": tree
        }
        
    except Exception as e:
        raise HTTPException(
            status_code=500,
            detail=f"Ошибка при построении дерева: {str(e)}"
        )


@app.post("/build_report")
async def build_report(request: BuildTreeRequest):
    """
    Построение отчета и возврат XLSX файла (4 столбца).
    """
    try:
        tree, report = client.get_tree(request.start_id, request.max_depth, request.db_name)
        
        if not tree:
            raise HTTPException(
                status_code=404,
                detail=f"Объект с ID {request.start_id} не найден"
            )
        
        # Создаем Excel файл в памяти
        wb = Workbook()
        ws = wb.active
        ws.title = "Отчет"
        
        # Заголовки
        headers = ["Наименование детали", "Количество на изделие", "Обозначение детали", "Обозначение материала"]
        ws.append(headers)
        
        # Данные
        for row in report:
            ws.append([
                row['Наименование детали'],
                row['Количество на изделие'],
                row['Обозначение детали'],
                row['Обозначение материала']
            ])
        
        # Автоподбор ширины колонок
        for column in ws.columns:
            max_length = 0
            column_letter = column[0].column_letter
            for cell in column:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except:
                    pass
            adjusted_width = min(max_length + 2, 50)
            ws.column_dimensions[column_letter].width = adjusted_width
        
        # Сохраняем в BytesIO
        output = io.BytesIO()
        wb.save(output)
        output.seek(0)
        
        return Response(
            content=output.getvalue(),
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={
                "Content-Disposition": f"attachment; filename=loodsman_report_{request.start_id}.xlsx"
            }
        )
        
    except Exception as e:
        raise HTTPException(
            status_code=500,
            detail=f"Ошибка при генерации отчета: {str(e)}"
        )


@app.post("/build_report_for_ko")
async def build_report_for_ko(request: BuildTreeRequest):
    """
    Построение отчета для КО с дополнительным столбцом ГОСТ материала.
    Исключает стандартные изделия (Винт, Шайба, Болт, Гайка и т.д.).
    
    Столбцы:
    1. Наименование детали
    2. Количество на изделие
    3. Обозначение детали
    4. Обозначение материала
    5. ГОСТ материала
    """
    try:
        tree, report = client.get_tree(request.start_id, request.max_depth, request.db_name)
        
        if not tree:
            raise HTTPException(
                status_code=404,
                detail=f"Объект с ID {request.start_id} не найден"
            )
        
        # Фильтруем отчет: исключаем стандартные изделия
        # Список типов стандартных изделий для исключения
        excluded_types = [
            "Винт", "Болт", "Шайба", "Гайка", "Шпилька", 
            "Шплинт", "Штифт", "Заклепка", "Стандартное изделие"
        ]
        
        filtered_report = []
        for row in report:
            # Проверяем, не является ли деталь стандартным изделием
            # Для этого нужно получить тип объекта из дерева
            # Используем product для определения (по наличию ГОСТ в названии или специальным маркерам)
            part_number = row['Обозначение детали']
            part_name = row['Наименование детали']
            
            is_standard = False
            for excluded in excluded_types:
                if excluded in part_name or excluded in part_number:
                    is_standard = True
                    break
            
            # Также исключаем если в обозначении есть явный ГОСТ (признак стандартного изделия)
            if not is_standard and ('ГОСТ' in part_number or 'ОСТ' in part_number or 'СТП' in part_number):
                # Проверяем, не является ли это материалом (у материалов тоже есть ГОСТ)
                if row['Обозначение материала']:  # Если есть материал, то это деталь, а не стандартное изделие
                    is_standard = False
                else:
                    is_standard = True
            
            if not is_standard:
                # Получаем ГОСТ материала из атрибутов
                # Нужно получить дополнительные данные из дерева
                filtered_report.append(row)
        
        # Создаем Excel файл с дополнительным столбцом ГОСТ
        wb = Workbook()
        ws = wb.active
        ws.title = "Отчет_КО"
        
        # Заголовки (5 столбцов)
        headers = ["Наименование детали", "Количество на изделие", "Обозначение детали", "Обозначение материала", "ГОСТ материала"]
        ws.append(headers)
        
        # Данные с дополнительным столбцом ГОСТ
        for row in filtered_report:
            # Получаем ГОСТ материала (из атрибутов, если есть)
            gost = row.get('ГОСТ материала', '')
            
            ws.append([
                row['Наименование детали'],
                row['Количество на изделие'],
                row['Обозначение детали'],
                row['Обозначение материала'],
                gost
            ])
        
        # Автоподбор ширины колонок
        for column in ws.columns:
            max_length = 0
            column_letter = column[0].column_letter
            for cell in column:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except:
                    pass
            adjusted_width = min(max_length + 2, 50)
            ws.column_dimensions[column_letter].width = adjusted_width
        
        # Сохраняем в BytesIO
        output = io.BytesIO()
        wb.save(output)
        output.seek(0)
        
        return Response(
            content=output.getvalue(),
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={
                "Content-Disposition": f"attachment; filename=loodsman_report_ko_{request.start_id}.xlsx"
            }
        )
        
    except Exception as e:
        raise HTTPException(
            status_code=500,
            detail=f"Ошибка при генерации отчета для КО: {str(e)}"
        )