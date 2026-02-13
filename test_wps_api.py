"""
WPS 智能表格 API 测试用例
罗列了部分测试函数！
"""

from wps_airscript_client import WPSAirScriptClient

# ==================== 配置信息 ====================
# 务必先阅读Readme.md文件，学习如何获取自己的WPS在线智能表的以下这3个参数！
FILE_ID = "*********"  # 替换为你的智能表格ID
TOKEN = "*************************"  # 替换为你的 Token令牌
SCRIPT_ID = "V2-******************"  # 替换为你的脚本ID
SHEET_NAME = "工作表1"  # 默认工作表名称

# ==================== 初始化客户端 ====================

def get_client():
    """获取 API 客户端实例"""
    return WPSAirScriptClient(FILE_ID, TOKEN, SCRIPT_ID)


# ==================== 批量写入测试 ====================

def test_batch_write():
    """测试批量写入数据"""
    client = get_client()
    data = [
        ["姓名", "年龄", "部门", "工资"],
        ["张三", 25, "技术部", 9999],
        ["李四", 30, "市场部", 9000],
        ["王五", 28, "技术部", 8500],
    ]
    # batch_write 现在使用 write_data，可以指定起始单元格
    result = client.batch_write(data, start_cell="A1", sheet_name=SHEET_NAME)
    print("批量写入:", result)
    return result


# ==================== 单元格操作测试 ====================

def test_get_cell_value():
    """测试获取单元格值"""
    client = get_client()
    result = client.get_cell_value("A2", SHEET_NAME)
    print("获取单元格值:", result)
    return result
test_get_cell_value()

def test_set_cell_value():
    """测试设置单元格值"""
    client = get_client()
    result = client.set_cell_value("E1", "备注", SHEET_NAME)
    print("设置单元格值:", result)
    return result


def test_get_range_values():
    """测试获取区域值"""
    client = get_client()
    result = client.get_range_values("A1:D3", SHEET_NAME)
    print("获取区域值:", result)
    return result


def test_set_range_values():
    """测试设置区域值"""
    client = get_client()
    values = [["测试1", "测试2"], ["测试3", "测试4"]]
    result = client.set_range_values("F1:G2", values, SHEET_NAME)
    print("设置区域值:", result)
    return result


def test_clear_range():
    """测试清除区域内容"""
    client = get_client()
    result = client.clear_range("F1:G2", SHEET_NAME)
    print("清除区域:", result)
    return result


# ==================== 格式化操作测试 ====================

def test_set_font():
    """测试设置字体"""
    client = get_client()
    font_options = {
        "name": "微软雅黑",
        "size": 12,
        "bold": True,
        "color": client.rgb_to_excel_color(255, 0, 0)  # 红色
    }
    result = client.set_font("A1:D1", font_options, SHEET_NAME)
    print("设置字体:", result)
    return result


def test_set_background_color():
    """测试设置背景色"""
    client = get_client()
    color = client.rgb_to_excel_color(68, 114, 196)  # 蓝色
    result = client.set_background_color("A1:D1", color, SHEET_NAME)
    print("设置背景色:", result)
    return result


def test_set_alignment():
    """测试设置对齐方式"""
    client = get_client()
    align_options = {
        "horizontal": -4131,
        "vertical": -4160
    }
    result = client.set_alignment("A1:D1", align_options, SHEET_NAME)
    print("设置对齐:", result)
    return result


def test_set_border():
    """测试设置边框"""
    client = get_client()
    border_options = {
        "lineStyle": 1,  # 实线
        "weight": 2  # 细线
    }
    result = client.set_border("A1:D4", border_options, SHEET_NAME)
    print("设置边框:", result)
    return result


def test_merge_cells():
    """测试合并单元格"""
    client = get_client()
    result = client.merge_cells("A5:D5", SHEET_NAME)
    print("合并单元格:", result)
    return result


def test_auto_fit_columns():
    """测试自动调整列宽"""
    client = get_client()
    result = client.auto_fit_columns("A:D", SHEET_NAME)
    print("自动调整列宽:", result)
    return result


# ==================== 行列操作测试 ====================

def test_insert_rows():
    """测试插入行"""
    client = get_client()
    result = client.insert_rows(36, 10, SHEET_NAME)
    print("插入行:", result)
    return result


def test_set_row_height():
    """测试设置行高"""
    client = get_client()
    result = client.set_row_height(1, 30, SHEET_NAME)
    print("设置行高:", result)
    return result


def test_set_column_width():
    """测试设置列宽"""
    client = get_client()
    result = client.set_column_width(1, 20, SHEET_NAME)
    print("设置列宽:", result)
    return result


# ==================== 查找和替换测试 ====================

def test_find_cell():
    """测试查找单元格"""
    client = get_client()
    result = client.find_cell("27", "A1:D100", SHEET_NAME)
    print("查找单元格:", result)
    return result


def test_replace_in_range():
    """测试替换内容"""
    client = get_client()
    result = client.replace_in_range("27", "研发部", "A1:D100", SHEET_NAME)
    print("替换内容:", result)
    return result


# ==================== 排序操作测试 ====================

def test_sort_range():
    """测试排序"""
    client = get_client()
    sort_options = {
        "key": "B2",  # 按年龄列排序
        "order": 1,  # 升序
        "hasHeader": True
    }
    result = client.sort_range("A1:A60", sort_options, SHEET_NAME)
    print("排序:", result)
    return result


# ==================== 复制粘贴测试 ====================

def test_copy_paste_range():
    """测试复制粘贴"""
    client = get_client()
    result = client.copy_paste_range("A1:D1", "A20", SHEET_NAME)
    print("复制粘贴:", result)
    return result


# ==================== 工作簿/工作表信息测试 ====================

def test_get_worksheet_count():
    """测试获取工作表数量"""
    client = get_client()
    result = client.get_worksheet_count()
    print("工作表数量:", result)
    return result


def test_get_workbook_sheets():
    """测试获取所有工作表名称"""
    client = get_client()
    result = client.get_workbook_sheets()
    print("所有工作表:", result)
    if result:
        for i, sheet_name in enumerate(result, 1):
            print(f"  {i}. {sheet_name}")
    return result


# ==================== 公式操作测试 ====================

def test_get_cell_formula():
    """测试获取单元格公式"""
    client = get_client()
    result = client.get_cell_formula("D2", SHEET_NAME)
    print("获取公式:", result)
    return result


def test_set_cell_formula():
    """测试设置单元格公式"""
    client = get_client()
    result = client.set_cell_formula("D5", "=SUM(D2:D4)", SHEET_NAME)
    print("设置公式:", result)
    return result


# ==================== 更多格式化测试 ====================

def test_set_number_format():
    """测试设置数字格式"""
    client = get_client()
    result = client.set_number_format("D2:D4", "#,##0", SHEET_NAME)
    print("设置数字格式:", result)
    return result


def test_unmerge_cells():
    """测试取消合并单元格"""
    client = get_client()
    result = client.unmerge_cells("A5:D5", SHEET_NAME)
    print("取消合并:", result)
    return result


def test_clear_range_contents():
    """测试清除内容（保留格式）"""
    client = get_client()
    result = client.clear_range_contents("F1:G2", SHEET_NAME)
    print("清除内容:", result)
    return result


# ==================== 更多行列操作测试 ====================

def test_delete_rows():
    """测试删除行"""
    client = get_client()
    result = client.delete_rows(5, 1, SHEET_NAME)
    print("删除行:", result)
    return result


def test_insert_columns():
    """测试插入列"""
    client = get_client()
    result = client.insert_columns(2, 1, SHEET_NAME)
    print("插入列:", result)
    return result


def test_delete_columns():
    """测试删除列"""
    client = get_client()
    result = client.delete_columns(5, 1, SHEET_NAME)
    print("删除列:", result)
    return result


# ==================== 更多查找测试 ====================

def test_find_all_cells():
    """测试查找所有匹配的单元格"""
    client = get_client()
    result = client.find_all_cells("技术部", "A1:D100", SHEET_NAME)
    print(f"查找所有匹配项:{result}")
    return result


# ==================== 更多复制粘贴测试 ====================

def test_copy_range():
    """测试复制区域"""
    client = get_client()
    result = client.copy_range("A1:D1", SHEET_NAME)
    print("复制区域:", result)
    return result


def test_paste_to_range():
    """测试粘贴到指定位置"""
    client = get_client()
    result = client.paste_to_range("A30", SHEET_NAME)
    print("粘贴:", result)
    return result


# ==================== 数据读取测试 ====================

def test_get_used_range_data(isGetData):
    """测试获取已使用区域数据"""
    client = get_client()
    result = client.get_used_range_data(isGetData, SHEET_NAME)
    if result:
        if isGetData=='是':
            print(f"已使用区域数据: {len(result)} 行 x {len(result[0]) if result else 0} 列")
            print("前3行数据:")
            for i, row in enumerate(result[:3]):
                print(f"  行{i+1}: {row}")
        else:
            print(
                f'已使用区域-起始行：{result[0]}',
                f'已使用区域-起始列：{result[1]}',
                f'已使用区域-结束行：{result[2]}',
                f'已使用区域-结束列：{result[3]}',
            )
    else:
        print("未获取到数据")
    return result


# ==================== 工作表管理测试 ====================

def test_add_worksheet():
    """测试添加工作表"""
    client = get_client()
    result = client.add_worksheet("新工作表")
    print("添加工作表:", result)
    return result


def test_worksheet_exists():
    """测试检查工作表是否存在"""
    client = get_client()
    result1 = client.worksheet_exists(SHEET_NAME)
    result2 = client.worksheet_exists("不存在的工作表")
    print(f"{SHEET_NAME} 存在: {result1}")
    print(f"不存在的工作表 存在: {result2}")
    return result1


def test_delete_worksheet():
    """测试删除工作表"""
    client = get_client()
    result = client.delete_worksheet("新工作表")
    print("删除工作表:", result)
    return result


# ==================== 综合测试 ====================

def test_create_formatted_table():
    """综合测试：创建格式化表格"""
    client = get_client()
    
    print("\n=== 综合测试：创建格式化表格 ===\n")
    
    # 1. 写入数据
    data = [
        ["姓名", "年龄", "部门", "工资"],
        ["张三", 25, "技术部", 8000],
        ["李四", 30, "市场部", 9000],
        ["王五", 28, "技术部", 8500],
    ]
    print("1. 写入数据...")
    client.batch_write(data, start_cell="A1", sheet_name=SHEET_NAME)
    
    # 2. 设置标题行格式
    print("2. 设置标题行格式...")
    client.set_font("A1:D1", {
        "bold": True,
        "size": 12,
        "color": client.rgb_to_excel_color(255, 255, 255)
    }, SHEET_NAME)
    
    client.set_background_color("A1:D1", client.rgb_to_excel_color(68, 114, 196), SHEET_NAME)
    
    client.set_alignment("A1:D1", {
        "horizontal": -4108,
        "vertical": -4108
    }, SHEET_NAME)
    
    # 3. 设置边框
    print("3. 设置边框...")
    client.set_border("A1:D4", {
        "lineStyle": 1,
        "weight": 2
    }, SHEET_NAME)
    
    # 4. 自动调整列宽
    print("4. 自动调整列宽...")
    client.auto_fit_columns("A:D", SHEET_NAME)
    
    print("\n✅ 格式化表格创建完成！")


# ==================== 主函数 ====================

def main():
    """运行测试"""
    # print("\n" + "="*60)
    # print("WPS AirScript API 测试")
    # print("="*60 + "\n")

    client = get_client()
    
    # 选择要运行的测试（取消注释即可运行）
    
    # 批量写入测试
    # test_batch_write() # test success
    #
    # # 单元格操作测试
    test_get_cell_value() # test success
    # test_set_cell_value() # test success
    test_get_range_values() # test success
    # test_set_range_values() # test success
    # test_clear_range() # test success
    #
    # # 格式化操作测试
    # test_set_font() # test success
    # test_set_background_color() # test success
    # test_set_alignment() # test success
    # test_set_border() # test success
    # test_merge_cells() # test success
    # test_auto_fit_columns() # test success
    #
    # # 行列操作测试
    # test_insert_rows() # test success
    # test_set_row_height() # test success
    # test_set_column_width() # test success
    #
    # # 查找和替换测试
    # test_find_cell() # test success
    # test_replace_in_range() # test success
    #
    # # 排序测试
    # test_sort_range() # test success
    #
    # # 复制粘贴测试
    # test_copy_paste_range() # test success
    #
    # # 工作簿/工作表信息测试
    # test_get_worksheet_count() # test success
    # test_get_workbook_sheets() # test success
    #
    # # 公式操作测试
    # test_get_cell_formula() # test success
    # test_set_cell_formula() # test success
    #
    # # 更多格式化测试
    # test_set_number_format() # test success
    # test_unmerge_cells() # test success
    # test_clear_range_contents() # test success
    #
    # # 更多行列操作测试
    # test_delete_rows() # test success
    # test_insert_columns() # test success
    # test_delete_columns() # test success
    #
    # # 更多查找测试
    # test_find_all_cells() # test success
    #
    # # 更多复制粘贴测试
    # test_copy_range() # test success
    # test_paste_to_range() # test success
    #
    # 单元格嵌入图片（注：这个命令必须把JS粘贴到AirScript1.0环境里！）
    # image_data = "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAYAAAAf8/9hAAABMklEQVR4AcxS21HDMBC8kxuBQgJOB0zCP04lCZWYAiBDB3hCI6YQ+di9SEKx+WCGn3hmR7qHdvckB/nndwUE1j7cTKvtxxzxbrPP09lq08X7bT+HtY9tkBhaUVlARQ8kd6j2atLNIfiCfh5fsI6AaIi34fSmarZjPFmzJ7jXoGuvoYcxMOrwOgRsBAeeuU5TeOJaSM1GqiLnzVil9Jj5mTNBcpFtsxFu1lyJLMBx2IPcmETECZBYuPAc7gFrac7qJsaxUZKKYOaibubt1+rN6ej2yVAcMChWY9NrUj/npcuXWauzdkmQXAifFVUSBpMvbDGidCIy1uqIf0ZgQFQKPnu6LH/mqsZWx4UDZqgA5R3e3f8F5vgiOHxgjXGNBQGLVOVPwj2hw/vCOvPErwQs/BXfAAAA//86U9wbAAAABklEQVQDAAVetyEAc7DHAAAAAElFTkSuQmCC"
    # res = client.insert_image("A1", image_data, SHEET_NAME)
    # print(res)
    # # 数据读取测试
    test_get_used_range_data("是") # test success
    test_get_used_range_data("否") # test success
    # time.sleep(1)
    #
    # # 工作表管理测试
    # test_add_worksheet() # test success
    # test_worksheet_exists() # test success
    # test_delete_worksheet() # test success
    #
    # # 综合测试
    # test_create_formatted_table() # test success


if __name__ == "__main__":
    main()
