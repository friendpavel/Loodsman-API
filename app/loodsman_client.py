# -*- coding: utf-8 -*-
import pythoncom
from win32com.client import Dispatch, gencache
from typing import Dict, Any, List, Tuple, Optional

class LoodsmanClient:
    """Клиент для работы с ЛОЦМАН через COM"""
    
    def __init__(self):
        self.LINK_TYPE_COMPOSITION = "Состоит из ..."
        self.LINK_TYPE_MATERIAL = "Изготавливается из ..."
        self._initialized = False
        
    def _ensure_connection(self):
        """Инициализация COM и подключение"""
        if not self._initialized:
            pythoncom.CoInitialize()
            self._initialized = True
            
    def _get_simple_api(self, db_name: str):
        """Получение API подключения к ЛОЦМАН"""
        data_provider_module = gencache.EnsureModule(
            "{76AA08AD-DEDA-40CB-962B-24440A3E96AD}", 0, 1, 0
        )
        loodsman_connection = data_provider_module.ILoodsmanConnection(
            Dispatch("{044259A0-FEAE-4B55-B6FE-E9B024FCCD4B}")._oleobj_.QueryInterface(
                data_provider_module.ILoodsmanConnection.CLSID, pythoncom.IID_IDispatch
            )
        )
        simple_api = loodsman_connection.API8.GetSimpleAPI()
        simple_api.RunMethod("ConnectToDB", (db_name,))
        return simple_api

    def dataset_to_list(self, data_set):
        """Конвертирует DataSet в список словарей"""
        if not data_set or data_set.RecordCount == 0:
            return []
        result = []
        data_set.First()
        while not data_set.Eof:
            row = {}
            for i in range(data_set.FieldCount):
                field_name = data_set.FieldName(i)
                try:
                    val = data_set.FieldValue(field_name)
                    if val is not None:
                        row[field_name] = str(val)
                except:
                    row[field_name] = None
            result.append(row)
            data_set.Next()
        return result

    def get_linked_objects(self, simple_api, object_id: int, link_type: str) -> List[Dict]:
        """Возвращает объекты по указанному типу связи"""
        try:
            ds = simple_api.GetDataSet("GetLinkedFast2", (object_id, link_type, False))
            if ds and ds.RecordCount > 0:
                return self.dataset_to_list(ds)
        except:
            pass
        return []

    def get_object_attributes(self, simple_api, object_id: int) -> Dict:
        """Получает все атрибуты объекта"""
        try:
            ds = simple_api.GetDataSet("GetInfoAboutVersion", (0, "", "", object_id, 2))
            if not ds or ds.RecordCount == 0:
                return {}
            attrs = {}
            ds.First()
            while not ds.Eof:
                name = ds.FieldValue("_NAME")
                value = ds.FieldValue("_VALUE")
                if name:
                    attrs[str(name)] = str(value) if value else ""
                ds.Next()
            return attrs
        except:
            return {}

    def get_object_info(self, simple_api, object_id: int) -> Dict:
        """Получает базовую информацию об объекте"""
        try:
            ds = simple_api.GetDataSet("GetInfoAboutVersion", (0, "", "", object_id, 15))
            if ds and ds.RecordCount > 0:
                ds.First()
                return {
                    "id": ds.FieldValue("_ID_VERSION"),
                    "type": ds.FieldValue("_TYPE"),
                    "product": ds.FieldValue("_PRODUCT"),
                    "state": ds.FieldValue("_STATE"),
                    "version": ds.FieldValue("_VERSION"),
                }
        except:
            pass
        return {}

    def build_full_tree(self, simple_api, object_id: int, object_name: str, object_type: str,
                        quantity: float = 1.0, depth: int = 0, max_depth: int = 15, 
                        visited: Optional[set] = None) -> Optional[Dict]:
        """Рекурсивно строит полное дерево объектов"""
        if visited is None:
            visited = set()
        if depth > max_depth or object_id in visited:
            return None
        visited.add(object_id)

        attrs = self.get_object_attributes(simple_api, object_id)
        
        node = {
            "id": object_id,
            "type": object_type,
            "name": attrs.get("Наименование", object_name),
            "product": object_name,
            "quantity": quantity,
            "attributes": attrs,
            "material": None,
            "material_attributes": None,
            "material_gost": None,
            "children": []
        }

        # Поиск материала по связи "Изготавливается из ..."
        materials = self.get_linked_objects(simple_api, object_id, self.LINK_TYPE_MATERIAL)
        
        if materials:
            mat = materials[0]
            mat_id = mat.get("_ID_VERSION")
            mat_product = mat.get("_PRODUCT", "")
            mat_attrs = self.get_object_attributes(simple_api, mat_id) if mat_id else {}
            
            material_name = mat_attrs.get("Марка материала", mat_product)
            material_gost = mat_attrs.get("НТД на материал", "")
            
            node["material"] = material_name
            node["material_attributes"] = mat_attrs
            node["material_gost"] = material_gost
            
            # Добавляем материал как отдельный дочерний объект
            material_node = {
                "id": mat_id,
                "type": mat.get("_TYPE", "Материал"),
                "name": mat_attrs.get("Наименование", mat_product),
                "product": mat_product,
                "quantity": 1.0,
                "attributes": mat_attrs,
                "is_material": True,
                "material": None,
                "material_attributes": None,
                "material_gost": None,
                "children": []
            }
            node["children"].append(material_node)

        # Получение состава (дочерние объекты по связи "Состоит из ...")
        children = self.get_linked_objects(simple_api, object_id, self.LINK_TYPE_COMPOSITION)
        
        if children:
            for child in children:
                child_id = child.get("_ID_VERSION")
                child_product = child.get("_PRODUCT", "")
                child_qty = float(child.get("_MIN_QUANTITY", 1.0))
                
                if not child_id or child_id in visited:
                    continue
                
                child_info = self.get_object_info(simple_api, child_id)
                child_type = child_info.get("type", "Unknown")
                
                child_node = self.build_full_tree(
                    simple_api, child_id, child_product, child_type,
                    child_qty, depth + 1, max_depth, visited
                )
                if child_node:
                    node["children"].append(child_node)

        return node

    def flatten_tree_to_report(self, node: Dict, result: Optional[List] = None) -> List[Dict]:
        """Преобразует дерево в плоский список для отчета (4 столбца)"""
        if result is None:
            result = []
        
        if not node.get("is_material"):
            row = {
                "Наименование детали": node.get("name", ""),
                "Количество на изделие": node.get("quantity", 1.0),
                "Обозначение детали": node.get("product", ""),
                "Обозначение материала": node.get("material", ""),
                "ГОСТ материала": node.get("material_gost", "")  # Добавляем ГОСТ
            }
            result.append(row)
        
        for child in node.get("children", []):
            self.flatten_tree_to_report(child, result)
        
        return result

    def get_tree(self, start_id: int, max_depth: int = 15, db_name: str = "SATOMS") -> Tuple[Optional[Dict], List[Dict]]:
        """Основной метод: получает дерево и отчет для указанного объекта"""
        self._ensure_connection()
        simple_api = self._get_simple_api(db_name)
        
        obj_info = self.get_object_info(simple_api, start_id)
        obj_name = obj_info.get("product", f"ID_{start_id}")
        obj_type = obj_info.get("type", "Unknown")
        
        tree = self.build_full_tree(simple_api, start_id, obj_name, obj_type, max_depth=max_depth)
        report = self.flatten_tree_to_report(tree) if tree else []
        
        return tree, report