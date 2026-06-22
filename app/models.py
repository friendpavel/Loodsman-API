# -*- coding: utf-8 -*-
from pydantic import BaseModel, Field
from typing import Optional, List, Dict, Any

class BuildTreeRequest(BaseModel):
    """Запрос на построение дерева"""
    start_id: int = Field(..., description="ID начального объекта (изделия или проекта)")
    max_depth: int = Field(15, description="Максимальная глубина рекурсии", ge=1, le=30)
    db_name: str = Field("SATOMS", description="Имя базы данных ЛОЦМАН")

class BuildTreeResponse(BaseModel):
    """Ответ с деревом"""
    success: bool
    start_id: int
    max_depth: int
    total_objects: Optional[int] = None
    materials_count: Optional[int] = None
    tree: Optional[Dict[str, Any]] = None
