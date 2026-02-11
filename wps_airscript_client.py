"""
WPS 智能表格 AirScript API2.0 Python客户端
 * WPS 智能表格 AirScript2.0 API通用工具函数库
 * @Repository：https://github.com/HnBigVolibear/wps_airscript2.0_online_tool
 * @Version：V20260207 豪华版
 * @Author：湖南大白熊RPA工作室
 * @Contact：https://github.com/HnBigVolibear/
 * @Liscense：MIT
 * 基于WPS AirScript2.0，官方文档：https://airsheet.wps.cn/docs/apiV2/overview.html
 * 现在已切换至 2.0 版本，不过要注意可能有部分函数不兼容AirScript1.0。。。
 * 部分方法（尤其是单元格插入图片），如果你用起来发现出现离奇报错，那么请切换至1.0版本！
"""

import requests
from typing import Dict, Any, Optional, List


class WPSAirScriptClient:
    """WPS 智能表格 AirScript API 客户端"""
    
    def __init__(self, file_id: str, token: str, script_id: str, base_url: str = "https://www.kdocs.cn"):
        """
        初始化 API 客户端

        Args:
            file_id: 文件 ID（从 URL 中获取）
            token: AirScript Token
            script_id: 脚本id
            base_url: API 基础 URL，默认为 https://www.kdocs.cn
        """
        self.file_id = file_id
        self.token = token
        self.script_id = script_id
        self.base_url = base_url.rstrip('/')

    def _get_headers(self) -> Dict[str, str]:
        """获取请求头"""
        return {
            'Content-Type': 'application/json',
            'AirScript-Token': self.token
        }

    def _request(self, context: Dict[str, Any]) -> Dict[str, Any]:
        """
        发送 HTTP 请求

        Args:
            context: 上下文参数

        Returns:
            API 响应的 JSON 数据
        """
        url = f"{self.base_url}/api/v3/ide/file/{self.file_id}/script/{self.script_id}/sync_task"
        
        try:
            response = requests.post(
                url=url,
                headers=self._get_headers(),
                json={"Context": context},
                timeout=123
            )
            response.raise_for_status()
            return response.json()
        except requests.exceptions.RequestException as e:
            print(f"请求失败: {e}")
            if hasattr(e, 'response') and e.response is not None:
                print(f"响应状态码: {e.response.status_code}")
                print(f"响应内容: {e.response.text}")
            raise

    def _extract_result(self, result: Any) -> Any:
        """
        提取结果：统一处理列表和字典格式
        
        Args:
            result: 原始结果
            
        Returns:
            提取后的结果
        """
        # 如果是列表，取第一个元素
        if isinstance(result, list) and len(result) > 0:
            return result[0]
        return result
    
    def _call_function(self, function_name: str, sheet_name: str = None, **params) -> Any: # type: ignore
        """
        调用脚本函数的通用方法

        Args:
            function_name: 函数名
            sheet_name: 工作表名称
            **params: 函数参数

        Returns:
            函数执行结果
        """
        context = {
            "argv": {
                "function": function_name,
                **params
            }
        }
        
        if sheet_name:
            # context["active_sheet"] = sheet_name
            context["sheet_name"] = sheet_name # type: ignore
            context["argv"]["thisSheetName"] = sheet_name # 特意设置一个专门指定自定义表名的字段！
        else:
            context["argv"]["thisSheetName"] = "" # 特意设置一个专门指定自定义表名的字段！
        
        # print(f"DEBUG: 发送的上下文 = {str(context)[:200]}")
        response = self._request(context)
        
        # 解析返回数据
        if response.get("data") and response["data"].get("result"):
            result_str = response["data"]["result"]
            if result_str != "[Undefined]":
                import json
                try:
                    result = json.loads(result_str)
                    return result[0] if result else None
                except:
                    return result_str
        
        return response

    # ==================== 单元格操作 ====================
    
    def get_cell_value(self, address: str, sheet_name: str = None) -> Any: # type: ignore
        """
        获取单元格值
        
        Args:
            address: 单元格地址，如 "A1"
            sheet_name: 工作表名称，可选。不指定则使用当前活动工作表
            
        Returns:
            单元格的值（可能是字符串、数字、布尔值等）
            
        Example:
            >>> client.get_cell_value("A1")
            'Hello'
            >>> client.get_cell_value("B2", "Sheet1")
            123
        """
        return self._call_function("getCellValue", sheet_name=sheet_name, address=address)
    
    def set_cell_value(self, address: str, value: Any, sheet_name: str = None) -> Dict: # type: ignore
        """
        设置单元格值
        
        Args:
            address: 单元格地址，如 "A1"
            value: 要设置的值（字符串、数字、布尔值等）
            sheet_name: 工作表名称，可选。不指定则使用当前活动工作表
            
        Returns:
            执行结果字典
            
        Example:
            >>> client.set_cell_value("A1", "Hello")
            >>> client.set_cell_value("B2", 123, "Sheet1")
        """
        return self._call_function("setCellValue", sheet_name=sheet_name, address=address, value=value)
    
    def get_range_values(self, address: str, sheet_name: str = None) -> List[Dict]: # type: ignore
        """
        获取区域值（返回二维数组）
        
        Args:
            address: 区域地址，如 "A1:C3"
            sheet_name: 工作表名称，可选。不指定则使用当前活动工作表
            
        Returns:
            二维数组，每个元素代表一行数据
            
        Example:
            >>> client.get_range_values("A1:B2")
            [['Name', 'Age'], ['Alice', 25]]
        """
        return self._call_function("getRangeValues", sheet_name=sheet_name, address=address)
    
    def set_range_values(self, address: str, values: List[List], sheet_name: str = None) -> List[Dict]: # type: ignore
        """
        设置区域值（批量写入）
        
        Args:
            address: 区域地址，如 "A1:C3"
            values: 二维数组数据，每个子数组代表一行
            sheet_name: 工作表名称，可选。不指定则使用当前活动工作表
            
        Returns:
            执行结果字典
            
        Example:
            >>> data = [['Name', 'Age'], ['Alice', 25], ['Bob', 30]]
            >>> client.set_range_values("A1:B3", data)
        """
        return self._call_function("setRangeValues", sheet_name, address=address, values=values)
    
    def batch_write(self, data: List[List], start_cell: str = "A1", sheet_name: str = None) -> List[Dict]: # type: ignore
        """
        批量写入数据到工作表
        
        Args:
            data: 二维数组数据
            start_cell: 起始单元格，默认 "A1"
            sheet_name: 工作表名称
            
        Returns:
            执行结果
        """
        # 计算范围
        if not data or len(data) == 0:
            return [{"success": False, "message": "数据为空"}]
        
        rows = len(data)
        cols = len(data[0]) if data[0] else 0
        
        if cols == 0:
            return [{"success": False, "message": "数据为空"}]
        
        # 计算结束单元格
        import re
        match = re.match(r'([A-Z]+)(\d+)', start_cell)
        if not match:
            return [{"success": False, "message": "起始单元格格式错误"}]
        
        start_col = match.group(1)
        start_row = int(match.group(2))
        
        # 计算结束列
        end_col_num = self._column_letter_to_number(start_col) + cols - 1
        end_col = self._column_number_to_letter(end_col_num)
        end_row = start_row + rows - 1
        
        address = f"{start_cell}:{end_col}{end_row}"
        return self.set_range_values(address, data, sheet_name)
    
    def _column_letter_to_number(self, column: str) -> int:
        """列字母转数字"""
        result = 0
        for char in column:
            result = result * 26 + (ord(char) - ord('A') + 1)
        return result
    
    def _column_number_to_letter(self, num: int) -> str:
        """列数字转字母"""
        letter = ""
        while num > 0:
            remainder = (num - 1) % 26
            letter = chr(ord('A') + remainder) + letter
            num = (num - 1) // 26
        return letter
    
    def clear_range(self, address: str, sheet_name: str = None) -> Dict: # type: ignore
        """
        清除区域内容和格式
        
        Args:
            address: 区域地址，如 "A1:C3"
            sheet_name: 工作表名称，可选
            
        Returns:
            执行结果字典
        """
        return self._call_function("clearRange", sheet_name=sheet_name, address=address)
    
    def clear_range_contents(self, address: str, sheet_name: str = None) -> Dict: # type: ignore
        """
        清除区域内容（保留格式）
        
        Args:
            address: 区域地址，如 "A1:C3"
            sheet_name: 工作表名称，可选
            
        Returns:
            执行结果字典
        """
        return self._call_function("clearRangeContents", sheet_name=sheet_name, address=address)
    
    def get_cell_formula(self, address: str, sheet_name: str = None) -> List[Dict]: # type: ignore
        """
        获取单元格公式
        
        Args:
            address: 单元格地址，如 "A1"
            sheet_name: 工作表名称，可选
            
        Returns:
            公式字符串，如 "=SUM(A1:A10)"
        """
        return self._call_function("getCellFormula", sheet_name=sheet_name, address=address)
    
    def set_cell_formula(self, address: str, formula: str, sheet_name: str = None) -> Dict: # type: ignore
        """
        设置单元格公式
        
        Args:
            address: 单元格地址，如 "A1"
            formula: 公式字符串，如 "=SUM(A1:A10)"
            sheet_name: 工作表名称，可选
            
        Returns:
            执行结果字典
        """
        return self._call_function("setCellFormula", sheet_name=sheet_name, address=address, formula=formula)

    # ==================== 格式化操作 ====================
    
    def set_font(self, address: str, font_options: Dict, sheet_name: str = None) -> Dict: # type: ignore
        """
        设置字体样式
        
        Args:
            address: 单元格或区域地址，如 "A1" 或 "A1:C3"
            font_options: 字体选项字典，可包含以下键：
                - name: 字体名称，如 "Arial", "微软雅黑"
                - size: 字体大小，如 12
                - bold: 是否粗体，True/False
                - italic: 是否斜体，True/False
                - color: 字体颜色（Excel 颜色值）
            sheet_name: 工作表名称，可选
            
        Returns:
            执行结果字典
            
        Example:
            >>> client.set_font("A1", {"name": "Arial", "size": 14, "bold": True})
        """
        return self._call_function("setCellFont", sheet_name=sheet_name, address=address, fontOptions=font_options)
    
    def set_background_color(self, address: str, color: int, sheet_name: str = None) -> Dict: # type: ignore
        """
        设置背景色
        
        Args:
            address: 单元格或区域地址，如 "A1" 或 "A1:C3"
            color: Excel 颜色值（可使用 rgb_to_excel_color 方法转换）
            sheet_name: 工作表名称，可选
            
        Returns:
            执行结果字典
            
        Example:
            >>> color = client.rgb_to_excel_color(255, 255, 0)  # 黄色
            >>> client.set_background_color("A1", color)
        """
        return self._call_function("setCellBackgroundColor", sheet_name=sheet_name, address=address, color=color)
    
    def set_alignment(self, address: str, align_options: Dict, sheet_name: str = None) -> Dict: # type: ignore
        """
        设置对齐方式
        
        Args:
            address: 单元格或区域地址，如 "A1" 或 "A1:C3"
            align_options:
                对齐选项字典，可包含以下键

                ▸ horizontal（设置区域的水平对齐方式）
                    - -4152 → 右对齐 (xlHAlignRight)
                    - -4131 → 左对齐 (xlHAlignLeft)
                    - -4130 → 两端对齐 (xlHAlignJustify)
                    - -4117 → 分散对齐 (xlHAlignDistributed)
                    - -4108 → 居中 (xlHAlignCenter)
                    - 1     → 自动 (xlHAlignGeneral)
                    - 5     → 填充 (xlHAlignFill)
                    - 7     → 跨列居中 (xlHAlignCenterAcrossSelection)

                ▸ vertical（设置区域的垂直对齐方式）
                    - -4160 → 顶部对齐 (xlVAlignTop)
                    - -4130 → 两端对齐 (xlVAlignJustify)
                    - -4117 → 分散对齐 (xlVAlignDistributed)
                    - -4108 → 垂直居中 (xlVAlignCenter)
                    - -4107 → 底部对齐 (xlVAlignBottom)
            sheet_name: 工作表名称，可选
            
        Returns:
            执行结果字典
            
        Example:
            >>> # 居中对齐
            >>> client.set_alignment("A1", {"horizontal": -4108, "vertical": -4108})
            >>> # 靠左靠上
            >>> client.set_alignment("A1:C3", {"horizontal": -4131, "vertical": -4160})
        """
        return self._call_function("setCellAlignment", sheet_name=sheet_name, address=address, alignOptions=align_options)
    
    def set_border(self, address: str, border_options: Dict, sheet_name: str = None) -> Dict: # type: ignore
        """
        设置边框
        
        Args:
            address: 单元格或区域地址，如 "A1" 或 "A1:C3"
            border_options: 边框选项字典，可包含以下键：
                - style: 边框样式，如 "thin", "medium", "thick"
                - color: 边框颜色（Excel 颜色值）
                - position: 边框位置，如 "all", "top", "bottom", "left", "right"
            sheet_name: 工作表名称，可选
            
        Returns:
            执行结果字典
        """
        return self._call_function("setCellBorder", sheet_name=sheet_name, address=address, borderOptions=border_options)
    
    def merge_cells(self, address: str, sheet_name: str = None) -> Dict: # type: ignore
        """
        合并单元格
        
        Args:
            address: 区域地址，如 "A1:C3"
            sheet_name: 工作表名称，可选
            
        Returns:
            执行结果字典
        """
        return self._call_function("mergeCells", sheet_name=sheet_name, address=address)
    
    def auto_fit_columns(self, address: str, sheet_name: str = None) -> Dict: # type: ignore
        """
        自动调整列宽
        
        Args:
            address: 列地址或区域，如 "A:A" 或 "A1:C10"
            sheet_name: 工作表名称，可选
            
        Returns:
            执行结果字典
        """
        return self._call_function("autoFitColumns", sheet_name=sheet_name, address=address)
    
    def set_number_format(self, address: str, format_str: str, sheet_name: str = None) -> Dict: # type: ignore
        """
        设置数字格式
        
        Args:
            address: 单元格或区域地址，如 "A1" 或 "A1:C3"
            format_str: 格式字符串，如：
                - "0.00": 保留两位小数
                - "#,##0": 千分位分隔符
                - "0%": 百分比
                - "yyyy-mm-dd": 日期格式
            sheet_name: 工作表名称，可选
            
        Returns:
            执行结果字典
            
        Example:
            >>> client.set_number_format("A1", "0.00")
        """
        return self._call_function("setCellNumberFormat", sheet_name=sheet_name, address=address, format=format_str)
    
    def unmerge_cells(self, address: str, sheet_name: str = None) -> Dict: # type: ignore
        """
        取消合并单元格
        
        Args:
            address: 区域地址，如 "A1:C3"
            sheet_name: 工作表名称，可选
            
        Returns:
            执行结果字典
        """
        return self._call_function("unmergeCells", sheet_name=sheet_name, address=address)

    # ==================== 行列操作 ====================
    
    def insert_rows(self, row_index: int, count: int = 1, sheet_name: str = None) -> Dict: # type: ignore
        """
        插入行
        
        Args:
            row_index: 行索引（从 1 开始）
            count: 插入行数，默认为 1
            sheet_name: 工作表名称，可选
            
        Returns:
            执行结果字典
            
        Example:
            >>> client.insert_rows(3, 2)  # 在第 3 行位置插入 2 行
        """
        return self._call_function("insertRows", sheet_name=sheet_name, rowIndex=row_index, count=count)
    
    def set_row_height(self, row_index: int, height: float, sheet_name: str = None) -> Dict: # type: ignore # type: ignore
        """
        设置行高
        
        Args:
            row_index: 行索引（从 1 开始）
            height: 行高（单位：磅）
            sheet_name: 工作表名称，可选
            
        Returns:
            执行结果字典
            
        Example:
            >>> client.set_row_height(1, 30)  # 设置第 1 行高度为 30 磅
        """
        return self._call_function("setRowHeight", sheet_name=sheet_name, rowIndex=row_index, height=height)
    
    def set_column_width(self, column_index: int, width: float, sheet_name: str = None) -> Dict: # type: ignore # type: ignore
        """
        设置列宽
        
        Args:
            column_index: 列索引（从 1 开始，A=1, B=2, ...）
            width: 列宽（单位：字符宽度）
            sheet_name: 工作表名称，可选
            
        Returns:
            执行结果字典
            
        Example:
            >>> client.set_column_width(1, 20)  # 设置第 A 列宽度为 20
        """
        return self._call_function("setColumnWidth", sheet_name=sheet_name, columnIndex=column_index, width=width)
    
    def delete_rows(self, row_index: int, count: int = 1, sheet_name: str = None) -> Dict: # type: ignore # type: ignore
        """
        删除行
        
        Args:
            row_index: 起始行索引（从 1 开始）
            count: 删除行数，默认为 1
            sheet_name: 工作表名称，可选
            
        Returns:
            执行结果字典
            
        Example:
            >>> client.delete_rows(3, 2)  # 从第 3 行开始删除 2 行
        """
        return self._call_function("deleteRows", sheet_name=sheet_name, rowIndex=row_index, count=count)
    
    def insert_columns(self, column_index: int, count: int = 1, sheet_name: str = None) -> Dict: # type: ignore # type: ignore
        """
        插入列
        
        Args:
            column_index: 列索引（从 1 开始，A=1, B=2, ...）
            count: 插入列数，默认为 1
            sheet_name: 工作表名称，可选
            
        Returns:
            执行结果字典
            
        Example:
            >>> client.insert_columns(2, 1)  # 在第 B 列位置插入 1 列
        """
        return self._call_function("insertColumns", sheet_name=sheet_name, columnIndex=column_index, count=count)
    
    def delete_columns(self, column_index: int, count: int = 1, sheet_name: str = None) -> Dict: # type: ignore # type: ignore
        """
        删除列
        
        Args:
            column_index: 起始列索引（从 1 开始，A=1, B=2, ...）
            count: 删除列数，默认为 1
            sheet_name: 工作表名称，可选
            
        Returns:
            执行结果字典
            
        Example:
            >>> client.delete_columns(2, 1)  # 删除第 B 列
        """
        return self._call_function("deleteColumns", sheet_name=sheet_name, columnIndex=column_index, count=count)

    # ==================== 查找和替换 ====================
    
    def find_cell(self, search_text: str, search_range: str, sheet_name: str = None) -> Dict: # type: ignore # type: ignore
        """
        查找单元格（返回所有匹配项）
        
        Args:
            search_text: 要查找的文本
            search_range: 搜索范围，如 "A1:Z100"
            sheet_name: 工作表名称，可选
            
        Returns:
            包含所有匹配单元格信息的字典
            
        Example:
            >>> result = client.find_cell("Apple", "A1:Z100")
        """
        return self._call_function("findCell", sheet_name=sheet_name, searchText=search_text, searchRange=search_range)
    
    def replace_in_range(self, search_text: str, replace_text: str, search_range: str, sheet_name: str = None) -> Dict: # type: ignore # type: ignore
        """
        替换内容并返回替换数量
        
        Args:
            search_text: 要查找的文本
            replace_text: 替换后的文本
            search_range: 搜索范围，如 "A1:Z100"
            sheet_name: 工作表名称，可选
            
        Returns:
            包含替换数量的字典
            
        Example:
            >>> result = client.replace_in_range("old", "new", "A1:Z100")
            >>> print(result['count'])  # 替换了多少个
        """
        return self._call_function("replaceInRangeWithCount", sheet_name=sheet_name, 
                                  searchText=search_text, replaceText=replace_text, searchRange=search_range)
    
    def find_all_cells(self, search_text: str, search_range: str, sheet_name: str = None) -> List: # type: ignore
        """
        查找所有匹配的单元格
        
        Args:
            search_text: 要查找的文本
            search_range: 搜索范围，如 "A1:Z100"
            sheet_name: 工作表名称，可选
            
        Returns:
            匹配单元格地址列表，如 ['A1', 'B5', 'C10']
            
        Example:
            >>> cells = client.find_all_cells("Apple", "A1:Z100")
            >>> print(cells)  # ['A1', 'C5']
        """
        result = self._call_function("findAllCells", sheet_name=sheet_name, searchText=search_text, searchRange=search_range)
        result = self._extract_result(result)
        
        # 返回实际的单元格数组
        if result and isinstance(result, dict) and 'cells' in result:
            return result['cells']
        
        return []

    # ==================== 筛选操作 ====================

    def set_filter(self, field: int = 1, operator: str = None, criteria1: str = None, criteria2: str = None, is_reSet: bool = True, sheet_name: str = None) -> Dict: # type: ignore
        """
        设置筛选条件
        # result = client.set_filter(3, "xlOr", "麦子", "过客", True, "工作表1")
        # result = client.set_filter(2, "xlBottom10Percent", "10", None, True, "工作表1") # xlBottom10Percent
        # result = client.set_filter(2, "xlFilterValues", "<20", None, True, "工作表1")
        """
        field = int(field)
        if criteria1 and criteria1=='*':
            raise Exception("筛选条件不能为 '*'，请使用具体的条件值!")
        if criteria2 and criteria2=='*':
            raise Exception("筛选条件不能为 '*'，请使用具体的条件值!")
        return self._call_function("setFilter", sheet_name=sheet_name, field=field, operator=operator, criteria1=criteria1,criteria2=criteria2, is_reSet=is_reSet)

    def clear_filter(self, sheet_name: str = None) -> Dict: # type: ignore
        """
        清除筛选

        Args:
            sheet_name: 工作表名称，可选

        Returns:
            执行结果字典

        Example:
            >>> client.clear_filter("Sheet1")
        """
        return self._call_function("clearFilter", sheet_name=sheet_name)

    def get_filtered_data(self, sheet_name: str = None) -> Dict: # type: ignore
        """
        获取筛选后的数据

        Args:
            sheet_name: 工作表名称，可选

        Returns:
            包含筛选后数据的字典，格式：
            {
                'success': True,
                'data': [[row1_col1, row1_col2, ...], [row2_col1, row2_col2, ...], ...],
                'rowCount': 行数,
                'columnCount': 列数
            }

        Example:
            >>> result = client.get_filtered_data("Sheet1")
            >>> print(result['data'])  # 筛选后的数据
            >>> print(result['rowCount'])  # 筛选后的行数
        """
        return self._call_function("getFilteredData", sheet_name=sheet_name)

    # ==================== 排序操作 ====================
    
    def sort_range(self, address: str, sort_options: Dict, sheet_name: str = None) -> Dict: # type: ignore # type: ignore
        """
        排序区域
        
        Args:
            address: 区域地址，如 "A1:C10"
            sort_options: 排序选项字典，可包含以下键：
                - key: 排序列索引（从 1 开始）
                - order: 排序顺序，"asc" 升序或 "desc" 降序
                - hasHeaders: 是否包含标题行，True/False
            sheet_name: 工作表名称，可选
            
        Returns:
            执行结果字典
            
        Example:
            >>> client.sort_range("A1:C10", {"key": 2, "order": "asc", "hasHeaders": True})
        """
        return self._call_function("sortRange", sheet_name=sheet_name, address=address, sortOptions=sort_options)

    # ==================== 复制粘贴 ====================
    
    def copy_paste_range(self, source_address: str, target_address: str, sheet_name: str = None) -> Dict: # type: ignore # type: ignore
        """
        复制粘贴区域
        
        Args:
            source_address: 源区域地址，如 "A1:C3"
            target_address: 目标区域起始地址，如 "E1"
            sheet_name: 工作表名称，可选
            
        Returns:
            执行结果字典
            
        Example:
            >>> client.copy_paste_range("A1:C3", "E1")
        """
        return self._call_function("copyPasteRange", sheet_name=sheet_name, 
                                  sourceAddress=source_address, targetAddress=target_address)
    
    def copy_range(self, source_address: str, sheet_name: str = None) -> Dict: # type: ignore # type: ignore
        """
        复制区域到剪贴板
        
        Args:
            source_address: 源区域地址，如 "A1:C3"
            sheet_name: 工作表名称，可选
            
        Returns:
            执行结果字典
        """
        return self._call_function("copyRange", sheet_name=sheet_name, sourceAddress=source_address)
    
    def paste_to_range(self, target_address: str, sheet_name: str = None) -> Dict: # type: ignore # type: ignore
        """
        从剪贴板粘贴到指定位置
        
        Args:
            target_address: 目标区域起始地址，如 "E1"
            sheet_name: 工作表名称，可选
            
        Returns:
            执行结果字典
        """
        return self._call_function("pasteToRange", sheet_name=sheet_name, targetAddress=target_address)

    # ==================== 工作簿/工作表信息 ====================
    
    def get_worksheet_count(self) -> int:
        """
        获取工作表数量
        
        Returns:
            工作表数量
            
        Example:
            >>> count = client.get_worksheet_count()
            >>> print(count)  # 3
        """
        result = self._call_function("getWorksheetCount", None) # type: ignore
        result = self._extract_result(result)
        
        if isinstance(result, dict):
            return result.get("count", 0)
        
        return 0
    
    def get_workbook_sheets(self) -> List[str]:
        """
        获取所有工作表名称列表
        
        Returns:
            工作表名称列表
            
        Example:
            >>> sheets = client.get_workbook_sheets()
            >>> print(sheets)  # ['Sheet1', 'Sheet2', 'Sheet3']
        """
        result = self._call_function("getWorkbookName", None) # type: ignore
        result = self._extract_result(result)
        
        if isinstance(result, dict):
            return result.get("sheets", [])
        
        return []
    
    def get_used_range_data(self, sheet_name: str = None) -> List[List]: # type: ignore
        """
        获取已使用区域的数据
        
        Args:
            sheet_name: 工作表名称，可选
            
        Returns:
            二维数组，包含所有已使用单元格的数据
            
        Example:
            >>> data = client.get_used_range_data("Sheet1")
            >>> print(data)  # [['Name', 'Age'], ['Alice', 25]]
        """
        result = self._call_function("getUsedRangeData", sheet_name=sheet_name)
        result = self._extract_result(result)
        
        # 返回实际的数据数组
        if result and isinstance(result, dict) and 'data' in result:
            return result['data']
        
        return []

    # ==================== 工作表操作 ====================
    
    def add_worksheet(self, sheet_name: str = None) -> Dict: # type: ignore # type: ignore
        """
        添加新工作表
        
        Args:
            sheet_name: 工作表名称，可选。不指定则自动命名
            
        Returns:
            执行结果字典
            
        Example:
            >>> client.add_worksheet("NewSheet")
        """
        return self._call_function("addWorksheet", sheet_name=sheet_name)
    
    def delete_worksheet(self, sheet_identifier: str) -> Dict:
        """
        删除工作表
        
        Args:
            sheet_identifier: 工作表名称或索引
            
        Returns:
            执行结果字典
            
        Example:
            >>> client.delete_worksheet("Sheet2")
        """
        return self._call_function("deleteWorksheet", sheetIdentifier=sheet_identifier)

    def rename_worksheet(self, old_sheet_name: str, new_sheet_name: str) -> Dict:
        """
        重命名工作表

        Args:
            old_sheet_name: 原工作表名称
            new_sheet_name: 新工作表名称

        Returns:
            执行结果字典

        Example:
            >>> client.rename_worksheet("Sheet1", "数据表")
        """
        return self._call_function("renameWorksheet", oldSheetName=old_sheet_name, newSheetName=new_sheet_name)
    
    def worksheet_exists(self, sheet_name: str) -> bool:
        """
        检查工作表是否存在
        
        Args:
            sheet_name: 工作表名称
            
        Returns:
            True 表示存在，False 表示不存在
            
        Example:
            >>> exists = client.worksheet_exists("Sheet1")
            >>> print(exists)  # True
        """
        result = self._call_function("worksheetExists", sheet_name=sheet_name)
        result = self._extract_result(result)
        
        if isinstance(result, dict):
            return result.get("exists", False)
        
        return False

    def insert_image(self, address: str, image_data: str, sheet_name: str = None) -> Dict: # type: ignore
        """
        在指定单元格插入图片

        Args:
            address: 单元格地址，如 "A1"
            image_data: 图片的base64编码字符串
            sheet_name: 工作表名称，可选。不指定则使用当前活动工作表

        Returns:
            执行结果字典

        Example:
            >>> # 将图片转换为base64
            >>> import base64
            >>> with open("image.png", "rb") as f:
            ...     image_base64 = base64.b64encode(f.read()).decode('utf-8')
            >>> client.insert_image("A1", image_base64, "Sheet1")
        """
        return self._call_function("insertImage", sheet_name=sheet_name, address=address, imageData=image_data)

    def insert_link(self, address: str, text: str, url: str, sheet_name: str = None) -> Dict: # type: ignore
        """
        在指定单元格插入超链接

        Args:
            address: 单元格地址，如 "A1"
            text: 单元格呈现文字
            url: 链接网址
            sheet_name: 工作表名称，可选。不指定则使用当前活动工作表

        Returns:
            执行结果字典

        Example:
            >>> client.insert_link("A1", "点击这里", "https://www.example.com", "Sheet1")
        """
        return self._call_function("insertLink", sheet_name=sheet_name, address=address, text=text, url=url)

    def sortUsedRange(self, sheet_name: str, sortList: List, sortOptions: Dict) -> bool:
        result = self._call_function("sortUsedRange", sheet_name=sheet_name, sortList=sortList, sortOptions=sortOptions)
        result = self._extract_result(result)
        if isinstance(result, dict):
            return result.get("success", False)
        return False

    # ==================== 工具函数 ====================
    
    @staticmethod
    def rgb_to_excel_color(r: int, g: int, b: int) -> int:
        """
        RGB 颜色转 Excel 颜色值
        
        Args:
            r: 红色分量 (0-255)
            g: 绿色分量 (0-255)
            b: 蓝色分量 (0-255)
            
        Returns:
            Excel 颜色值
            
        Example:
            >>> color = WPSAirScriptClient.rgb_to_excel_color(255, 255, 0)  # 黄色
            >>> client.set_background_color("A1", color)
        """
        return r + g * 256 + b * 256 * 256

    # ==================== 透视表操作 ====================

    def create_pivot(self, source_sheet_name: str, source_range: str, 
                     row_column_indices: List[int], column_column_indices: List[int], 
                     value_column_indices: List[int], function_type: str,
                     target_sheet_name: str, target_cell: str) -> Dict: # type: ignore
        """
        创建透视表

        Args:
            source_sheet_name: 源数据表名称
            source_range: 源数据区域，如 "A1:D100"
            row_column_indices: 作为行字段的列索引列表（从1开始），可为空。如：[1,2]
            column_column_indices: 作为列字段的列索引列表（从1开始），可为空。如：[2,3]
            value_column_indices: 作为值字段的列索引列表（从1开始）。不能为空。如：[3]
            function_type: 统计函数类型，可选值：
                - "sum": 求和（默认）
                - "count": 计数
                - "average": 平均值
                - "max": 最大值
                - "min": 最小值
                - "product": 乘积
                - "countNums": 计数（仅数字）
                - "stdDev": 标准偏差
                - "stdDevP": 总体标准偏差
                - "var": 方差
                - "varP": 总体方差
            target_sheet_name: 透视表放置的工作表名称
            target_cell: 透视表放置的起始单元格，如 "A1"

        Returns:
            执行结果字典，包含success、message、pivotSheetName、pivotTableName等字段

        Example:
            >>> result = client.create_pivot(
            ...     "Sheet1", "A1:D100", [1, 2], [3], [4], "sum", "PivotSheet", "A1"
            ... )
        """
        return self._call_function(
            "createPivot",
            sourceSheetName=source_sheet_name,
            sourceRange=source_range,
            rowColumnIndices=row_column_indices,
            columnColumnIndices=column_column_indices,
            valueColumnIndices=value_column_indices,
            functionType=function_type,
            targetSheetName=target_sheet_name,
            targetCell=target_cell
        )

    def update_all_pivot_tables(self, sheet_name: str) -> Dict: # type: ignore
        """
        更新指定工作表中的所有透视表

        Args:
            sheet_name: 工作表名称

        Returns:
            执行结果字典

        Example:
            >>> result = client.update_all_pivot_tables("PivotSheet")
        """
        return self._call_function("updateAllPivotTables", sheet_name=sheet_name)

    def delete_all_pivot_tables(self, sheet_name: str) -> Dict: # type: ignore
        """
        删除指定工作表中的所有透视表

        Args:
            sheet_name: 工作表名称

        Returns:
            执行结果字典

        Example:
            >>> result = client.delete_all_pivot_tables("PivotSheet")
        """
        return self._call_function("deleteAllPivotTables", sheet_name=sheet_name)

    