/**
 * WPS 智能表格 AirScript2.0 API通用工具函数库
 * @Repository1：https://github.com/HnBigVolibear/dify_plugin_wps_airscript2.0_online_tool
 * @Repository2：https://github.com/HnBigVolibear/wps_airscript2.0_online_tool
 * @Version：V20260213豪华版
 * @Author：湖南大白熊RPA工作室
 * @Contact：https://github.com/HnBigVolibear/
 * @License：MIT
 * 基于WPS AirScript2.0，官方文档：https://airsheet.wps.cn/docs/apiV2/overview.html
 * 现在已切换至 2.0 版本，不过要注意可能有部分函数不兼容AirScript1.0。。。
 * 部分方法（尤其是单元格插入图片），如果你用起来发现出现离奇报错，那么请切换至1.0版本！
 */


// ==================== HTTP API 调用入口 ====================
/**
 * HTTP API 调用的主入口函数
 * 当通过 Python HTTP 请求调用时，会自动执行此函数
 *
 * 重要：WPS AirScript 需要脚本最后一个表达式作为返回值
 */
// 定义全局结果变量
var globalResult = [];

// 检查是否是 HTTP API 调用（存在 Context 对象）
if (typeof Context !== "undefined" && Context.argv) {
  try {
    console.log("接收到 HTTP API 调用");
    console.log("Context:", JSON.stringify(Context));

    var argv = Context.argv;
    // 注：WPS这个框架，存在一个坑，它擅自把用户从接口传入的sheet_name字段，自动改成了active_sheet这个名称！！！
    // Context示例: {"active_sheet":"工作表1","range":"$E$39","argv":{"woa_app":"db_assistant"},"link_from":""}
    var sheetName = argv.thisSheetName || Application.ActiveSheet.Name;
    
    // 如果有 items 数据，使用 setRangeValues 批量写入
    if (argv.items && Array.isArray(argv.items)) {
      try {
        const data = argv.items;
        const rows = data.length;
        const cols = data[0] ? data[0].length : 0;

        if (rows > 0 && cols > 0) {
          // 计算范围 (从 A1 开始)
          const endCol = columnNumberToLetter(cols);
          const address = `A1:${endCol}${rows}`;
          setRangeValues(address, data, sheetName);

          globalResult.push({
            success: true,
            message: "数据写入成功",
            rowsWritten: rows,
            range: address,
          });
        } else {
          globalResult.push({
            success: false,
            message: "数据为空",
          });
        }
        // console.log("返回结果:", JSON.stringify(globalResult));
      } catch (error) {
        globalResult.push({
          success: false,
          error: error.message,
        });
      }
    }
    // 如果有 function 参数，执行指定函数
    else if (argv.function) {
      globalResult = executeFunction(argv.function, argv, sheetName);
      // console.log("返回结果:", JSON.stringify(globalResult));
    }
    // 未指定操作
    else {
      globalResult.push({
        success: false,
        message: "未指定操作",
      });

      if (Object.keys(argv).length === 1 && Object.keys(argv).includes("woa_app")) {
        // 此时这里是在WPS在线脚本编辑器里进行本地调试时的情况！
        run_test_online();
      }
    }
  } catch (error) {
    console.error("HTTP API 调用出错:", error.message);
    globalResult = [];
    globalResult.push({
      success: false,
      error: error.message,
    });
  }
}

globalResult;

// ==================== HTTP API 辅助函数 ====================

/**
 * 执行指定函数（HTTP API 专用）
 * @param {string} functionName - 函数名
 * @param {Object} params - 参数对象
 * @param {string} sheetName - 工作表名称
 * @returns {Array} 执行结果数组
 */
function executeFunction(functionName, params, sheetName) {
  const result = [];
  console.log("执行函数:", functionName);
  console.log("目标工作表:", sheetName || "当前工作表");

  try {
    switch (functionName) {
      case "getCellValue":
        result.push({
          success: true,
          value: getCellValue(params.address, sheetName),
        });
        break;

      case "setCellValue":
        setCellValue(params.address, params.value, sheetName);
        result.push({ success: true, message: "设置成功" });
        break;

      case "getRangeValues":
        result.push({
          success: true,
          values: getRangeValues(params.address, sheetName),
        });
        break;

      case "setRangeValues":
        setRangeValues(params.address, params.values, sheetName);
        result.push({ success: true, message: "设置成功" });
        break;

      case "setCellFont":
        setCellFont(params.address, params.fontOptions, sheetName);
        result.push({ success: true, message: "字体设置成功" });
        break;

      case "setCellBackgroundColor":
        setCellBackgroundColor(params.address, params.color, sheetName);
        result.push({ success: true, message: "背景色设置成功" });
        break;

      case "setCellAlignment":
        setCellAlignment(params.address, params.alignOptions, sheetName);
        result.push({ success: true, message: "对齐方式设置成功" });
        break;

      case "setCellBorder":
        setCellBorder(params.address, params.borderOptions, sheetName);
        result.push({ success: true, message: "边框设置成功" });
        break;

      case "mergeCells":
        mergeCells(params.address, sheetName);
        result.push({ success: true, message: "合并成功" });
        break;

      case "autoFitColumns":
        autoFitColumns(params.address, sheetName);
        result.push({ success: true, message: "列宽调整成功" });
        break;

      case "insertRows":
        insertRows(params.rowIndex, params.count, sheetName);
        result.push({ success: true, message: "插入行成功" });
        break;

      case "setRowHeight":
        setRowHeight(params.rowIndex, params.height, sheetName);
        result.push({ success: true, message: "行高设置成功" });
        break;

      case "setColumnWidth":
        setColumnWidth(params.columnIndex, params.width, sheetName);
        result.push({ success: true, message: "列宽设置成功" });
        break;

      case "findCell":
        const cells = findCell(
          params.searchText,
          params.searchRange,
          sheetName
        );
        result.push({
          success: true,
          found: cells.length > 0,
          cells: cells,
        });
        break;

      case "replaceInRangeWithCount":
        const count = replaceInRangeWithCount(
          params.searchText,
          params.replaceText,
          params.searchRange,
          sheetName
        );
        result.push({ success: true, count: count });
        break;

      case "sortRange":
        sortRange(params.address, params.sortOptions, sheetName);
        result.push({ success: true, message: "排序成功" });
        break;
      
      case "sortUsedRange":
        sortUsedRange(sheetName, params.sortList, params.sortOptions);
        result.push({ success: true, message: "自定义排序成功" });
        break;

      case "copyPasteRange":
        copyPasteRange(
          params.sourceAddress,
          params.targetAddress,
          sheetName,
          sheetName
        );
        result.push({ success: true, message: "复制粘贴成功" });
        break;

      case "clearRange":
        clearRange(params.address, sheetName);
        result.push({ success: true, message: "清除成功" });
        break;

      case "clearRangeContents":
        clearRangeContents(params.address, sheetName);
        result.push({ success: true, message: "清除内容成功" });
        break;

      case "getCellFormula":
        result.push({
          success: true,
          formula: getCellFormula(params.address, sheetName),
        });
        break;

      case "setCellFormula":
        setCellFormula(params.address, params.formula, sheetName);
        result.push({ success: true, message: "设置公式成功" });
        break;

      case "setCellNumberFormat":
        setCellNumberFormat(params.address, params.format, sheetName);
        result.push({ success: true, message: "设置数字格式成功" });
        break;

      case "unmergeCells":
        unmergeCells(params.address, sheetName);
        result.push({ success: true, message: "取消合并成功" });
        break;

      case "deleteRows":
        deleteRows(params.rowIndex, params.count, sheetName);
        result.push({ success: true, message: "删除行成功" });
        break;

      case "insertColumns":
        insertColumns(params.columnIndex, params.count, sheetName);
        result.push({ success: true, message: "插入列成功" });
        break;

      case "deleteColumns":
        deleteColumns(params.columnIndex, params.count, sheetName);
        result.push({ success: true, message: "删除列成功" });
        break;

      case "findAllCells":
        const allCells = findAllCells(
          params.searchText,
          params.searchRange,
          sheetName
        );
        // 转换为标准格式
        const cellsInfo = allCells.map((cell) => ({
          address: cell.Address,
          value: cell.Value2,
          row: cell.Row,
          column: cell.Column,
        }));
        result.push({
          success: true,
          cells: cellsInfo,
          count: cellsInfo.length,
        });
        break;

      case "copyRange":
        copyRange(params.sourceAddress, sheetName);
        result.push({ success: true, message: "复制成功" });
        break;

      case "pasteToRange":
        pasteToRange(params.targetAddress, sheetName);
        result.push({ success: true, message: "粘贴成功" });
        break;

      case "getUsedRangeData":
        result.push({
          success: true,
          data: getUsedRangeData(params.isGetData, sheetName),
        });
        break;

      case "addWorksheet": 
        const newSheetName = addWorksheet(sheetName);
        result.push({
          success: true,
          message: "添加工作表成功！",
          sheetName: newSheetName,
        });
        break;

      case "deleteWorksheet":
        deleteWorksheet(params.sheetIdentifier);
        result.push({ success: true, message: "删除工作表成功" });
        break;

      case "worksheetExists":
        result.push({
          success: true,
          exists: worksheetExists(sheetName),
        });
        break;
      
      case "renameWorksheet":
        renameWorksheet(params.oldSheetName, params.newSheetName);
        result.push({ success: true, message: "重命名工作表成功" });
        break;
      
      case "createPivot":
        createPivot(params.sourceSheetName, params.sourceRange, params.rowColumnIndices, params.columnColumnIndices, params.valueColumnIndices, params.functionType, params.targetSheetName, params.targetCell)  
        result.push({ success: true, message: "创建数据透视表成功" });
        break;
      
      case "updateAllPivotTables":
        updateAllPivotTables(sheetName)
        result.push({ success: true, message: "更新数据透视表成功" });
        break;
      
      case "deleteAllPivotTables":
        deleteAllPivotTables(sheetName)
        result.push({ success: true, message: "删除数据透视表成功" });
        break;

      case "getWorksheetCount":
        result.push({ success: true, count: getWorksheetCount() });
        break;

      case "getWorkbookName":
        result.push({ success: true, sheets: getWorkbookName() });
        break;

      case "setFilter":
        setFilter(params.field, params.operator, params.criteria1,params.criteria2, params.is_reSet, sheetName);
        result.push({ success: true, message: "设置筛选成功" });
        break;

      case "clearFilter":
        clearFilter(sheetName);
        result.push({ success: true, message: "清除筛选成功" });
        break;

      case "getFilteredData":
        const filteredRes = getFilteredData(sheetName);
        result.push(filteredRes);
        break;

      case "insertImage":
        const insertImageRes = insertImage(params.address, params.imageData, sheetName);
        // const insertImageRes = insertImageByKSDrive(params.address, params.imageData, sheetName);
        result.push({ success: true, message: insertImageRes });
        break;
      
      case "insertLink":
        insertLink(params.address, params.text, params.url, sheetName);
        result.push({ success: true, message: "单元格插入链接成功" });
        break;

      default:
        result.push({
          success: false,
          message: "未知函数: " + functionName,
        });
    }
  } catch (error) {
    result.push({
      success: false,
      error: error.message,
    });
  }
  return result;
}

// ==================== 工作簿 (Workbook) 相关操作 ====================

/**
 * 获取当前活动的工作簿对象
 * @returns {Object} 工作簿对象
 */
function getActiveWorkbook() {
  return Application.ActiveWorkbook;
}

/**
 * 获取工作簿名称
 * @param {Object} workbook - 工作簿对象，不传则使用当前活动工作簿
 * @returns {string} 工作簿名称
 */
function getWorkbookName(workbook) {
  try {
    const wb = workbook || Application.ActiveWorkbook;

    // WPS AirScript 可能不支持获取工作簿名称
    // 返回所有工作表名称作为替代
    if (wb && wb.Sheets) {
      const sheets = wb.Sheets;
      const sheetNames = [];

      for (let i = 1; i <= sheets.Count; i++) {
        sheetNames.push(sheets.Item(i).Name);
      }

      return sheetNames;
    }

    return [];
  } catch (error) {
    console.error("getWorkbookName 错误:", error.message);
    return [];
  }
}

/**
 * 保存工作簿
 * @param {Object} workbook - 工作簿对象，不传则使用当前活动工作簿
 */
function saveWorkbook(workbook) {
  const wb = workbook || getActiveWorkbook();
  wb.Save();
}

/**
 * 关闭工作簿
 * @param {Object} workbook - 工作簿对象，不传则使用当前活动工作簿
 * @param {boolean} saveChanges - 是否保存更改，默认 false
 */
function closeWorkbook(workbook, saveChanges = false) {
  const wb = workbook || getActiveWorkbook();
  wb.Close(saveChanges);
}

// ==================== 工作表 (Worksheet) 相关操作 ====================

/**
 * 获取当前活动的工作表对象
 * @param {Object} workbook - 工作簿对象，不传则使用当前活动工作簿
 * @returns {Object} 工作表对象
 */
function getActiveWorksheet(workbook) {
  const wb = workbook || getActiveWorkbook();
  return wb.ActiveSheet;
}

/**
 * 根据名称获取工作表（支持模糊匹配）
 * @param {string} sheetName - 工作表名称，不传则使用当前活动工作表
 * @returns {Object} 工作表对象
 */
function getWorksheetByName(sheetName) {
  // 如果没有传入工作表名称，返回当前活动工作表
  if (!sheetName) {
    return Application.ActiveSheet;
  }

  const workbook = Application.ActiveWorkbook;
  const sheetCount = workbook.Sheets.Count;

  // 精确匹配
  for (let i = 1; i <= sheetCount; i++) {
    const sheet = workbook.Sheets(i);
    if (sheet.Name === sheetName) {
      return sheet;
    }
  }

  // 模糊匹配（包含）
  for (let i = 1; i <= sheetCount; i++) {
    const sheet = workbook.Sheets(i);
    if (sheet.Name.includes(sheetName)) {
      // console.log("找到匹配的工作表:", sheet.Name);
      return sheet;
    }
  }

  // 未找到，返回 null
  console.error("未找到工作表:", sheetName);
  return null;
}

/**
 * 根据索引获取工作表
 * @param {number} index - 工作表索引（从1开始）
 * @param {Object} workbook - 工作簿对象，不传则使用当前活动工作簿
 * @returns {Object} 工作表对象
 */
function getWorksheetByIndex(index, workbook) {
  const wb = workbook || getActiveWorkbook();
  return wb.Worksheets.Item(index);
}

/**
 * 检查工作表是否存在（支持模糊匹配）
 * @param {string} sheetName - 工作表名称
 * @param {Object} workbook - 工作簿对象，不传则使用当前活动工作簿
 * @returns {boolean} 是否存在
 */
function worksheetExists(sheetName, workbook) {
  if (!sheetName) {
    return false;
  }
  const wb = workbook || getActiveWorkbook();
  const sheetCount = wb.Sheets.Count;

  // 精确匹配
  for (let i = 1; i <= sheetCount; i++) {
    const sheet = wb.Sheets(i);
    if (sheet.Name === sheetName) {
      return true;
    }
  }

  // 模糊匹配（包含）
  // for (let i = 1; i <= sheetCount; i++) {
  //   const sheet = wb.Sheets(i);
  //   if (sheet.Name.includes(sheetName)) {
  //     return true;
  //   }
  // }

  return false;
}

function renameWorksheet(oldSheetName, newSheetName) {
  const ws1 = getWorksheetByName(oldSheetName);
  if (!ws1) {
    throw new Error("未找到原工作表: " + oldSheetName);
  }
  const ws2 = getWorksheetByName(newSheetName);
  if (ws2) {
    throw new Error("对不起，你要命的名，当前已存在同名工作表： " + newSheetName);
  }
  // 现在没冲突了，可以重命名工作表了。。
  ws1.Name = newSheetName
}

/**
 * 添加新工作表
 * @param {string} sheetName - 工作表名称（可选）
 * @param {Object} workbook - 工作簿对象，不传则使用当前活动工作簿
 * @returns {Object} 新创建的工作表对象
 */
function addWorksheet(sheetName) {
  const wb = getActiveWorkbook();
  const is_Exists = worksheetExists(sheetName, wb)
  if (is_Exists) {
    throw new Error("新建工作表失败！名称已存在: " + sheetName);
  } else {
    const newSheet = wb.Worksheets.Add();
    if (sheetName) { 
      newSheet.Name = sheetName;
    }
    return newSheet.Name
  }
}

/**
 * 删除工作表
 * @param {string|number} sheetIdentifier - 工作表名称或索引
 * @param {Object} workbook - 工作簿对象，不传则使用当前活动工作簿
 */
function deleteWorksheet(sheetIdentifier, workbook) {
  const wb = workbook || getActiveWorkbook();
  const sheet =
    typeof sheetIdentifier === "string"
      ? getWorksheetByName(sheetIdentifier, wb)
      : getWorksheetByIndex(sheetIdentifier, wb);
  sheet.Delete();
}

/**
 * 获取工作表数量
 * @param {Object} workbook - 工作簿对象，不传则使用当前活动工作簿
 * @returns {number} 工作表数量
 */
function getWorksheetCount(workbook) {
  const wb = workbook || getActiveWorkbook();
  return wb.Worksheets.Count;
}

// ==================== 单元格 (Range) 相关操作 ====================

/**
 * 获取单元格区域对象
 * @param {string} address - 单元格地址，如 "A1" 或 "A1:B10"
 * @param {string|Object} worksheetOrName - 工作表对象或工作表名称，不传则使用当前活动工作表
 * @returns {Object} 单元格区域对象
 */
function getRange(address, worksheetOrName) {
  let ws;

  if (!worksheetOrName) {
    // 没有传入参数，使用当前活动工作表
    ws = Application.ActiveSheet;
  } else if (typeof worksheetOrName === "string") {
    // 传入的是工作表名称
    ws = getWorksheetByName(worksheetOrName);
    if (!ws) {
      throw new Error("未找到工作表: " + worksheetOrName);
    }
  } else {
    // 传入的是工作表对象
    ws = worksheetOrName;
  }

  return ws.Range(address);
}

/**
 * 获取单元格的值
 * @param {string} address - 单元格地址，如 "A1"
 * @param {string} sheetName - 工作表名称，不传则使用当前活动工作表
 * @returns {*} 单元格的值
 */
function getCellValue(address, sheetName) {
  const range = getRange(address, sheetName);
  return range.Value2;
}

/**
 * 设置单元格的值
 * @param {string} address - 单元格地址，如 "A1"
 * @param {*} value - 要设置的值
 * @param {string} sheetName - 工作表名称，不传则使用当前活动工作表
 */
function setCellValue(address, value, sheetName) {
  const range = getRange(address, sheetName);
  range.Value2 = value;
}

/**
 * 获取单元格区域的值（二维数组）
 * @param {string} address - 单元格区域地址，如 "A1:B10"
 * @param {string} sheetName - 工作表名称，不传则使用当前活动工作表
 * @returns {Array} 二维数组
 */
function getRangeValues(address, sheetName) {
  const range = getRange(address, sheetName);
  return range.Value2;
}

/**
 * 设置单元格区域的值（二维数组）
 * @param {string} address - 单元格区域地址，如 "A1:B10"
 * @param {Array} values - 二维数组
 * @param {string} sheetName - 工作表名称，不传则使用当前活动工作表
 */
function setRangeValues(address, values, sheetName) {
  const range = getRange(address, sheetName);
  range.Value2 = values;
}

/**
 * 清除单元格内容
 * @param {string} address - 单元格地址，如 "A1" 或 "A1:B10"
 * @param {string} sheetName - 工作表名称，不传则使用当前活动工作表
 */
function clearRange(address, sheetName) {
  const range = getRange(address, sheetName);
  range.Clear();
}

/**
 * 清除单元格内容（保留格式）
 * @param {string} address - 单元格地址，如 "A1" 或 "A1:B10"
 * @param {string} sheetName - 工作表名称，不传则使用当前活动工作表
 */
function clearRangeContents(address, sheetName) {
  const range = getRange(address, sheetName);
  range.ClearContents();
}

/**
 * 获取单元格公式
 * @param {string} address - 单元格地址，如 "A1"
 * @param {string} sheetName - 工作表名称，不传则使用当前活动工作表
 * @returns {string} 单元格公式
 */
function getCellFormula(address, sheetName) {
  const range = getRange(address, sheetName);
  return range.Formula;
}

/**
 * 设置单元格公式
 * @param {string} address - 单元格地址，如 "A1"
 * @param {string} formula - 公式字符串，如 "=SUM(A1:A10)"
 * @param {string} sheetName - 工作表名称，不传则使用当前活动工作表
 */
function setCellFormula(address, formula, sheetName) {
  const range = getRange(address, sheetName);
  range.Formula = formula;
}

// ==================== 单元格格式化操作 ====================

/**
 * 设置单元格字体样式
 * @param {string} address - 单元格地址
 * @param {Object} fontOptions - 字体选项 { name, size, bold, italic, color }
 * @param {string} sheetName - 工作表名称，不传则使用当前活动工作表
 */
function setCellFont(address, fontOptions, sheetName) {
  const range = getRange(address, sheetName);
  const font = range.Font;

  if (fontOptions.name) font.Name = fontOptions.name;
  if (fontOptions.size) font.Size = fontOptions.size;
  if (fontOptions.bold !== undefined) font.Bold = fontOptions.bold;
  if (fontOptions.italic !== undefined) font.Italic = fontOptions.italic;
  if (fontOptions.color) font.Color = hexColorToRGB( fontOptions.color );
}

/**
 * 设置单元格背景色
 * @param {string} address - 单元格地址
 * @param {number} color - 颜色值（RGB）
 * @param {string} sheetName - 工作表名称，不传则使用当前活动工作表
 */
function setCellBackgroundColor(address, color, sheetName) {
  const range = getRange(address, sheetName);
  range.Interior.Color = hexColorToRGB(color);
}

/**
 * 设置单元格对齐方式
 * @param {string} address - 单元格地址
 * @param {Object} alignOptions - 对齐选项 { horizontal, vertical }
 * @param {string} sheetName - 工作表名称，不传则使用当前活动工作表
 */
function setCellAlignment(address, alignOptions, sheetName) {
  const range = getRange(address, sheetName);

  if (alignOptions.horizontal) {
    range.HorizontalAlignment = alignOptions.horizontal;
  }
  if (alignOptions.vertical) {
    range.VerticalAlignment = alignOptions.vertical;
  }
}

/**
 * 设置单元格边框
 * @param {string} address - 单元格地址
 * @param {Object} borderOptions - 边框选项 { lineStyle, weight, color }
 * @param {string} sheetName - 工作表名称，不传则使用当前活动工作表
 */
function setCellBorder(address, borderOptions, sheetName) {
  const range = getRange(address, sheetName);
  const borders = range.Borders;

  if (borderOptions.lineStyle) borders.LineStyle = borderOptions.lineStyle;
  if (borderOptions.weight) borders.Weight = borderOptions.weight;
  if (borderOptions.color) borders.Color = hexColorToRGB(borderOptions.color)
}

/**
 * 设置单元格数字格式
 * @param {string} address - 单元格地址
 * @param {string} format - 数字格式，如 "0.00", "#,##0", "yyyy-mm-dd"
 * @param {string} sheetName - 工作表名称，不传则使用当前活动工作表
 */
function setCellNumberFormat(address, format, sheetName) {
  const range = getRange(address, sheetName);
  range.NumberFormat = format;
}

// ==================== 行列操作 ====================

/**
 * 插入行
 * @param {number} rowIndex - 行索引（从1开始）
 * @param {number} count - 插入行数，默认1
 * @param {string} sheetName - 工作表名称，不传则使用当前活动工作表
 */
function insertRows(rowIndex, count = 1, sheetName) {
  const ws = getWorksheetByName(sheetName);
  const range = ws.Rows(rowIndex);
  for (let i = 0; i < count; i++) {
    range.Insert();
  }
}

/**
 * 删除行
 * @param {number} rowIndex - 行索引（从1开始）
 * @param {number} count - 删除行数，默认1
 * @param {string} sheetName - 工作表名称，不传则使用当前活动工作表
 */
function deleteRows(rowIndex, count = 1, sheetName) {
  const ws = getWorksheetByName(sheetName);
  for (let i = 0; i < count; i++) {
    const range = ws.Rows(rowIndex);
    range.Delete();
  }
}

/**
 * 插入列
 * @param {number} columnIndex - 列索引（从1开始）
 * @param {number} count - 插入列数，默认1
 * @param {string} sheetName - 工作表名称，不传则使用当前活动工作表
 */
function insertColumns(columnIndex, count = 1, sheetName) {
  const ws = getWorksheetByName(sheetName);
  const range = ws.Columns(columnIndex);
  for (let i = 0; i < count; i++) {
    range.Insert();
  }
}

/**
 * 删除列
 * @param {number} columnIndex - 列索引（从1开始）
 * @param {number} count - 删除列数，默认1
 * @param {string} sheetName - 工作表名称，不传则使用当前活动工作表
 */
function deleteColumns(columnIndex, count = 1, sheetName) {
  const ws = getWorksheetByName(sheetName);
  for (let i = 0; i < count; i++) {
    const range = ws.Columns(columnIndex);
    range.Delete();
  }
}

/**
 * 设置行高
 * @param {number} rowIndex - 行索引（从1开始）
 * @param {number} height - 行高
 * @param {string} sheetName - 工作表名称，不传则使用当前活动工作表
 */
function setRowHeight(rowIndex, height, sheetName) {
  const ws = getWorksheetByName(sheetName);
  ws.Rows(rowIndex).RowHeight = height;
}

/**
 * 设置列宽
 * @param {number} columnIndex - 列索引（从1开始）
 * @param {number} width - 列宽
 * @param {string} sheetName - 工作表名称，不传则使用当前活动工作表
 */
function setColumnWidth(columnIndex, width, sheetName) {
  const ws = getWorksheetByName(sheetName);
  ws.Columns(columnIndex).ColumnWidth = width;
}

/**
 * 自动调整列宽
 * @param {string} address - 单元格区域地址，如 "A:A" 或 "A1:C10"
 * @param {string} sheetName - 工作表名称，不传则使用当前活动工作表
 */
function autoFitColumns(address, sheetName) {
  const range = getRange(address, sheetName);
  range.Columns.AutoFit();
}

// ==================== 查找和筛选操作 ====================

/**
 * 查找单元格（返回所有匹配项）
 * @param {string} searchText - 要查找的文本
 * @param {string} searchRange - 查找范围，如 "A1:Z100"
 * @param {string} sheetName - 工作表名称，不传则使用当前活动工作表
 * @returns {Array} 找到的所有单元格信息数组 [{address, value, row, column}]，未找到返回空数组
 */
function findCell(searchText, searchRange, sheetName) {
  const range = getRange(searchRange, sheetName);
  const result = [];

  // 查找第一个匹配项
  const firstCell = range.Find(searchText);

  if (!firstCell) {
    return result;
  }

  // 记录第一个单元格的行列，用于判断是否循环回到起点
  const firstRow = firstCell.Row;
  const firstCol = firstCell.Column;
  let currentCell = firstCell;
  let count = 0;
  const maxIterations = 10000; // 防止无限循环

  // 循环查找所有匹配项
  do {
    result.push({
      address: currentCell.Address,
      value: currentCell.Value2,
      row: currentCell.Row,
      column: currentCell.Column,
    });

    // 查找下一个匹配项
    currentCell = range.FindNext(currentCell);
    count++;

    // 安全检查：防止无限循环
    if (count > maxIterations) {
      console.error("查找循环次数超过限制，可能存在问题");
      break;
    }

    // 如果找不到或者回到第一个单元格（通过行列判断），则退出循环
  } while (
    currentCell &&
    !(currentCell.Row === firstRow && currentCell.Column === firstCol)
  );

  return result;
}

/**
 * 查找所有匹配的单元格
 * @param {string} searchText - 要查找的文本
 * @param {string} searchRange - 查找范围，如 "A1:Z100"
 * @param {string} sheetName - 工作表名称，不传则使用当前活动工作表
 * @returns {Array} 找到的单元格对象数组
 */
function findAllCells(searchText, searchRange, sheetName) {
  const range = getRange(searchRange, sheetName);
  const results = [];
  const firstCell = range.Find(searchText);

  if (!firstCell) return results;

  // 记录第一个单元格的行列
  const firstRow = firstCell.Row;
  const firstCol = firstCell.Column;
  let currentCell = firstCell;
  let count = 0;
  const maxIterations = 10000; // 防止无限循环

  do {
    results.push(currentCell);
    currentCell = range.FindNext(currentCell);
    count++;

    // 安全检查：防止无限循环
    if (count > maxIterations) {
      console.error("查找循环次数超过限制，可能存在问题");
      break;
    }
  } while (
    currentCell &&
    !(currentCell.Row === firstRow && currentCell.Column === firstCol)
  );

  return results;
}

/**
 * 替换单元格内容
 * @param {string} searchText - 要查找的文本
 * @param {string} replaceText - 替换的文本
 * @param {string} searchRange - 查找范围，如 "A1:Z100"
 * @param {string} sheetName - 工作表名称，不传则使用当前活动工作表
 * @returns {boolean} 是否成功替换（true=成功，false=未找到）
 */
function replaceInRange(searchText, replaceText, searchRange, sheetName) {
  const range = getRange(searchRange, sheetName);
  return range.Replace(searchText, replaceText);
}

/**
 * 替换单元格内容并返回替换数量
 * @param {string} searchText - 要查找的文本
 * @param {string} replaceText - 替换的文本
 * @param {string} searchRange - 查找范围，如 "A1:Z100"
 * @param {string} sheetName - 工作表名称，不传则使用当前活动工作表
 * @returns {number} 替换的数量
 */
function replaceInRangeWithCount(
  searchText,
  replaceText,
  searchRange,
  sheetName
) {
  // 先查找所有匹配项（用于计数）
  const cells = findAllCells(searchText, searchRange, sheetName);
  const count = cells.length;

  // 如果找到匹配项，执行替换
  if (count > 0) {
    const range = getRange(searchRange, sheetName);
    range.Replace(searchText, replaceText);
  }

  return count;
}

// ==================== 筛选操作 ====================

/**
 * 设置筛选条件
 * @param {string} field - 筛选字段（列字母），如 "A"
 * @param {string} criteria1 - 筛选条件，如 ">100", "Apple", ">=2023-01-01"
 * @param {string} criteria2 - 筛选条件，如 ">100", "Apple", ">=2023-01-01"
 * @param {string} operator - 操作符，如 ">", "<", "=", ">=", "<=", "<>", "contains", "beginsWith", "endsWith"
 * @param {string} is_reSet - 是否先清除筛选，true 或 false
 * @param {string} sheetName - 工作表名称，不传则使用当前活动工作表
 */
function setFilter(field, operator, criteria1,criteria2, is_reSet, sheetName) {
  // AutoFilter参考文档：https://airsheet.wps.cn/docs/apiV2/excel/workbook/Range/方法/AutoFilter 方法.html
  const is_Exists = worksheetExists(sheetName)
  if (!is_Exists) {
    throw new Error("设置筛选失败！提供的表名不存在！");
  }
  const ws = getWorksheetByName(sheetName);
  if (is_reSet) {
    // 如果存在筛选，则先清除筛选
    if (ws && ws.AutoFilter) {
      // 清除筛选
      try {
        ws.AutoFilterMode = false; // 第一种方法：Vibe Coding 给出的方法。反而奏效了。。。
      } catch {
        ws.AutoFilter.ShowAllData(); // 第二种方法：WPS官方AI给出的方法，但是这一行好像报错。。。
      }
    }
  }

  // 获取已使用区域
  ws.AutoFilterMode = true; // 再重新开启。
  let filterRange;
  try {
    filterRange = ws.AutoFilter.Range; // 当前已有筛选区域情况下，直接获取即可。
  } catch {
    filterRange = ws.UsedRange; // 此时则是未开启筛选状态的！于是直接获取已使用区域！
  }
  
  if ( operator=="xlFilterCellColor" || operator=="xlFilterFontColor" ) {
    criteria1 = hexColorToRGB( criteria1 )
  }

  // 获取operator规则对象
  const ExcelConstants = {
    xlAnd: xlAnd, // 条件 1 和条件 2 的逻辑与。只有operator=xlAnd和xlOr时，criteria2条件2才会实际起作用！其他operator时候只有criteria1有用！
    xlOr: xlOr,   // 条件 1 和条件 2 的逻辑或。只有operator=xlAnd和xlOr时，criteria2条件2才会实际起作用！其他operator时候只有criteria1有用！
    
    xlBottom10Items: xlBottom10Items, // 显示最低值项（条件 1 中指定的项数）
    xlBottom10Percent: xlBottom10Percent, // 显示最低值项（条件 1 中指定的百分数）
    xlFilterCellColor: xlFilterCellColor, // 单元格颜色
    // xlFilterDynamic: xlFilterDynamic, // 动态筛选
    xlFilterFontColor: xlFilterFontColor, // 字体颜色
    // xlFilterIcon: xlFilterIcon,   // 筛选图标
    xlFilterValues: undefined,  // 筛选值，比如“<30”，此时反而不需要这个规则对象，因此直接设置为None！
    xlTop10Items: xlTop10Items,    // 显示最高值项（条件 1 中指定的项数）
    xlTop10Percent: xlTop10Percent   // 显示最高值项（条件 1 中指定的百分数）
  };
  operator = ExcelConstants[operator]

  // --- 辅助函数：格式化 Criteria (核心修复点) ---
  // WPS AirScript 要求通配符条件最好带上 "=" 前缀
  const formatCriteria = (s) => {
    if (s == null || typeof s !== 'string') return s;
    // const s = s.trim(); // 容忍空格，因为有时候就是有空格参与筛选的！
    // 如果包含通配符 * 或 ? ，
    if (s.includes('*') || s.includes('?')) {
      // 且还没有以运算符开头 (=, >, <, <>)
      // 则强制加上 "=" 前缀，确保 WPS 识别为模式匹配
      if (!s.startsWith('=') && !s.startsWith('>') && !s.startsWith('<') && !s.startsWith('<>')) {
        return '=' + s;
      }
    }
    return s;
  };
  
  if (criteria2 == null || criteria2 == undefined || criteria2.trim()=='' || typeof criteria2 === 'object') {
    // ↑不知道为什么，当python前端传入criteria2=None时，JS这里接收到的criteria2  ↑ 却变成了 {} 服了。。。因此上面这里判断它是否为字典对象object！
    filterRange.AutoFilter(field, criteria1, operator)
    console.log('---打印--- 条件分支1')
  } else if (operator == xlAnd || operator == xlOr ) {
    criteria1 = formatCriteria(criteria1); 
    criteria2 = formatCriteria(criteria2); 
    filterRange.AutoFilter(field, criteria1, operator, criteria2)
    console.log('---打印--- 条件分支2')
  } else {
    filterRange.AutoFilter(field, criteria1, operator)
    console.log('---打印--- 条件分支3')
  }
  // // 以下是 1.0 版本的写法。但是在2.0版本里，AutoFilter对象的所有属性都是只读的，无法修改！因此下面这个旧方法已经废弃！
  // const autoFilter = ws.AutoFilter;
  // const filterItem = autoFilter.Filters.Item(field);
  // filterItem.Operator = operator; // 这里已经是数值（xlAnd/xlOr）
  // filterItem.Criteria1 = criteria1;
  // filterItem.Criteria2 = criteria2;
  // autoFilter.ApplyFilter();
}

/**
 * 清除筛选
 * @param {string} sheetName - 工作表名称，不传则使用当前活动工作表
 */
function clearFilter(sheetName) {
  // AutoFilter参考文档：https://airsheet.wps.cn/docs/apiV2/excel/workbook/Range/方法/AutoFilter 方法.html
  const is_Exists = worksheetExists(sheetName)
  if (!is_Exists) {
    throw new Error("设置筛选失败！提供的表名不存在！");
  }
  const ws = getWorksheetByName(sheetName);
  // 如果存在筛选，清除筛选
  if (ws && ws.AutoFilter) {
    // 清除筛选
    try {
      ws.AutoFilterMode = false; // 第一种方法：Vibe Coding 给出的方法。AI反而奏效了。。。
    } catch {
      ws.AutoFilter.ShowAllData(); // 第二种方法：WPS官方文档里给出的方法，但是这一行好像报错。。。WPS官方拉胯了。。。不及时更新文档（底部）：https://airsheet.wps.cn/docs/api/excel/workbook/AutoFilter.html
    }
  }
}


/**
 * 获取工作表中筛选后显示的数据
 * @param {string} sheetName - 工作表名称，不传则使用当前活动工作表
 * @returns {Object} 操作结果，包含筛选后的数据
 */
function getFilteredData(sheetName) {
// function getFilteredData(sheetName, range) { // 作废参数 * @param {string} range - 数据区域，如 "A1:D100"
  // AutoFilter参考文档：https://airsheet.wps.cn/docs/apiV2/excel/workbook/Range/方法/AutoFilter 方法.html
  // 【格外注意】截至20260201，WPS官方好像暂不支持该方法！
  // 因为查看官方文档发现，Range.Hidden这个属性，是只写的，无法读！即你可以设置某一行是否隐藏，但是无法直接读取这个状态！
  // 但是，本例通过直接用if进行比对，巧妙绕过了这个限制，顺利实现了这个功能！！！
  const is_Exists = worksheetExists(sheetName)
  if (!is_Exists) {
    throw new Error("设置筛选失败！提供的表名不存在！");
  }
  try {
    const sheet = getWorksheetByName(sheetName);
    // const sheet = Application.ActiveWorkbook.Sheets.Item(sheetName);
    // 检查是否有筛选
    const filter = sheet.AutoFilter;
    if (!filter) {
      return { success: false, message: "工作表未启用筛选功能" };
    }
    // 获取筛选后的可见行
    const visibleRows = [];
    let filterRange;
    try {
      filterRange = sheet.AutoFilter.Range; // 当前已有筛选区域情况下，直接获取即可。
    } catch {
      filterRange = sheet.UsedRange; // 此时则是未开启筛选状态的！于是直接获取已使用区域！
    }
    const rowCount = filterRange.Rows.Count;
    const colCount = filterRange.Columns.Count;
    // 遍历每一行
    for (let i = 1; i <= rowCount; i++) {
      const row = filterRange.Rows.Item(i);
      // 检查行是否可见（未被筛选隐藏）
      if (row.Hidden === false) {  // 至关重要的一行判断！
        const rowData = [];
        for (let j = 1; j <= colCount; j++) {
          rowData.push(row.Cells.Item(1, j).Value2);
        }
        visibleRows.push(rowData);
      }
    }
    return {
      success: true,
      message: `成功获取 ${visibleRows.length} 行的筛选后数据`,
      data: visibleRows,
      rowCount: visibleRows.length,
      colCount: colCount
    };
  } catch (error) {
    return {
      success: false,
      message: `获取筛选数据失败: ${error.message}`
    };
  }
}

// ==================== 排序操作 ====================

/**
 * 对区域进行排序
 * @param {string} address - 要排序的区域地址
 * @param {Object} sortOptions - 排序选项 { key, order, hasHeader }
 *   - key: 排序关键列地址，如 "A1"
 *   - order: 排序顺序，1=升序，2=降序
 *   - hasHeader: 是否包含标题行，默认 false
 * @param {string} sheetName - 工作表名称，不传则使用当前活动工作表
 */
function sortRange(address, sortOptions, sheetName) {
  const range = getRange(address, sheetName);
  const key = getRange(sortOptions.key, sheetName);
  const order = sortOptions.order || 1;
  const header = sortOptions.hasHeader ? 1 : 2;

  range.Sort(key, order, null, null, null, null, null, header);
}


// 上面那个自带的排序方法，好像用不了，只能自己实现了。。。
// 直接对当前已使用区域进行排序！
/**
 * 对区域进行自定义排序 - 大白熊自研
 * @param {string} sheetName - 要排序的表
 * @param {List} sortList - 定义参与排序的列，形如：[ ["C", "desc"], ["D", "asc"], ... ]
 * @param {Object} sortOptions - 排序选项 { key, order, hasHeader }
 *   - modeHeader: 是否有表头参与，xlGuess为自动，xlYes为包含表头，xlNo为不包含表头。默认设置为xlGuess
 *   - modeMatchCase: 是否大小写敏感，“是”为区分大小写，“否”为不区分大小写，默认设置“否”
 */
function sortUsedRange(sheetName, sortList , sortOptions) {
  // 获取当前表格区域
  let range;
  if (sheetName) {
    const ws = getWorksheetByName(sheetName);
    range = ws.UsedRange
  } else {
    range = ActiveSheet.UsedRange
  }

  // 获取到排序对象
  const sort = ActiveSheet.Sort
  // 获取排序范围
  const sortFields = sort.SortFields
  // 清除之前的范围
  sortFields.Clear()

  // sortList示例：[ ["C", "desc"], ["D", "asc"], ... ]
  for (let i = 0; i < sortList.length; i++) {
    let col_str = sortList[i][0] +":"+ sortList[i][0]
    let sort_asc_or_desc = sortList[i][1]=="asc" ? xlAscending: xlDescending; 
    // 写死xlSortOnValues, 按值排序。暂不支持API预留的颜色排序。。。 
    sortFields.Add(Range(col_str).Item(1, 1), xlSortOnValues, sort_asc_or_desc);
  }

  // 设置是否包含表头参数，xlGuess为自动，xlYes为包含表头，xlNo为不包含表头。默认设置为xlGuess
  let modeHeader = sortOptions.modeHeader;
  modeHeader = modeHeader=='xlGuess' ? xlGuess : (modeHeader=='xlYes' ? xlYes : xlNo)
  sort.Header = modeHeader

  // 设置是否大小写敏感，true为区分大小写，false为不区分大小写，默认设置false
  let modeMatchCase = sortOptions.modeMatchCase=='是' ? true : false
  sort.MatchCase = modeMatchCase

  // 【这里写死】设置中文排序方法，xlPinYin为拼音排序，xlStroke为比划数排序。默认设置为xlPinYin
  sort.SortMethod = xlPinYin
  // 【这里写死】设置排序的方法，xlSortColumns为按列排序，xlSortRows为按行排序，默认设置为xlSortColumns
  sort.Orientation = xlSortColumns

  // 排序前必须设置SetRange
  sort.SetRange(range)
  // 开始排序
  sort.Apply()
  console.log('自定义排序完成！')
}



// ==================== 复制粘贴操作 ====================

/**
 * 复制单元格区域
 * @param {string} sourceAddress - 源区域地址
 * @param {string} sheetName - 工作表名称，不传则使用当前活动工作表
 */
function copyRange(sourceAddress, sheetName) {
  const range = getRange(sourceAddress, sheetName);
  range.Copy();
}

/**
 * 粘贴到指定位置
 * @param {string} targetAddress - 目标区域地址
 * @param {string} sheetName - 工作表名称，不传则使用当前活动工作表
 */
function pasteToRange(targetAddress, sheetName) {
  const range = getRange(targetAddress, sheetName);
  range.Select();
  const ws = getWorksheetByName(sheetName);
  ws.Paste();
}

/**
 * 复制并粘贴单元格区域
 * @param {string} sourceAddress - 源区域地址
 * @param {string} targetAddress - 目标区域地址
 * @param {Object} sourceWorksheet - 源工作表对象
 * @param {Object} targetWorksheet - 目标工作表对象
 */
function copyPasteRange(
  sourceAddress,
  targetAddress,
  sourceWorksheet,
  targetWorksheet
) {
  const sourceRange = getRange(sourceAddress, sourceWorksheet);
  const targetRange = getRange(targetAddress, targetWorksheet);
  sourceRange.Copy(targetRange);
}

// ==================== 合并单元格操作 ====================

/**
 * 合并单元格
 * @param {string} address - 要合并的区域地址，如 "A1:B2"
 * @param {string} sheetName - 工作表名称，不传则使用当前活动工作表
 */
function mergeCells(address, sheetName) {
  const range = getRange(address, sheetName);
  range.Merge();
}

/**
 * 取消合并单元格
 * @param {string} address - 要取消合并的区域地址
 * @param {string} sheetName - 工作表名称，不传则使用当前活动工作表
 */
function unmergeCells(address, sheetName) {
  const range = getRange(address, sheetName);
  range.UnMerge();
}

function insertImage(address, imageData, sheetName) {
  // 夭寿了！2.0环境好像暂无法实现这个方法！
  // 如果你真滴有插入图片到单元格的需求，那么请将本JS代码，直接复制到1.0版本的脚本里面去！
  // 但是当前这个脚本里，其他很多方法都是无法适配1.0版本的！
  // 因此你可能需要同时维护两边，使用2个script_id了！！！
  // 另外，不要尝试改用KSDrive云文档API来强行写！使用KSDrive.openFile读取表格后，结果发现还是没用！
  try {
    // 以下这个插入图片的方法，仅在1.0版本有效！
    const range = getRange(address, sheetName);
    range.InsertImage(imageData); // 这个方法，目前好像只有1.0版本支持。。。
    return "插入图片成功！ 温馨提醒，若你要使用本插件的其他海量操作方法，那么再次初始化时，你切记要使用AirScript2.0版本的脚本ID哦！"
  } catch (error) {
    console.log('插入图片失败：', error.message);
    if (error.message == 'range.InsertImage is not a function') {
      return "【严重错误】插入图片失败：当前你的AirScript脚本，是放入你WPS在线智能表格的2.0版本脚本里面，但是2.0与1.0版本并不完全兼容，尤其是这个插入图片的range.InsertImage方法！！！ 如果你真的需要使用这个功能，请再额外创建一个1.0版本的脚本，全选复制粘贴放入本插件这里的全部代码即可！！！"
    } else {
      return "【未知错误】插入图片失败：" + JSON.stringify(error)
    }
  }
}


function insertLink(address, text, url, sheetName) {
  const range = getRange(address, sheetName);
  // Range("C1").Hyperlinks.Add(Range("C1"), "https://www.wps.com/")
  range.Hyperlinks.Add(range, url);
  // 2.0环境里，必须最后设置单元格显示文字！否则会被link覆盖！
  range.Value2 = text;
}

// ==================== 批量数据操作 ====================

/**
 * 获取已使用区域的数据
 * @param {string} isGetData - 是否返回数据。否则只返回当前已使用区域的位置（起始单元格~结束单元格）
 * @param {string} sheetName - 工作表名称，不传则使用当前活动工作表
 * @returns {Array} 二维数组数据
 */
function getUsedRangeData(isGetData, sheetName) {
  const ws = getWorksheetByName(sheetName);
  const usedRange = ws.UsedRange;
  if (isGetData=='是') {
    return usedRange.Value2;
  } else {
    return [
      usedRange.Row, // 起始行
      usedRange.Column, // 起始列
      usedRange.Row+usedRange.Rows.Count-1, // 最后一行
      usedRange.Column+usedRange.Columns.Count-1 // 最后一列
    ]
  }
}


// ==================== 透视表操作 =================
/**
 * 创建透视表函数（可设置统计方式版）
 * @param {string} sourceSheetName - 源数据表名称
 * @param {string} sourceRange - 源数据区域，如 "A1:D100"
 * @param {Array<number>} rowColumnIndices - 作为行字段的列索引列表（从1开始），可为空。如：[1,2]
 * @param {Array<number>} columnColumnIndices - 作为列字段的列索引列表（从1开始），可为空。如：[2,3]
 * @param {Array<number>} valueColumnIndices - 作为值字段的列索引列表（从1开始）。如：[3]
 * @param {string} functionType - 统计函数类型，可选值：
 *   - "sum": 求和（默认）
 *   - "count": 计数
 *   - "average": 平均值
 *   - "max": 最大值
 *   - "min": 最小值
 *   - "product": 乘积
 *   - "countNums": 计数（仅数字）
 *   - "stdDev": 标准偏差
 *   - "stdDevP": 总体标准偏差
 *   - "var": 方差
 *   - "varP": 总体方差
 * @param {string} targetSheetName - 透视表放置的工作表名称
 * @param {string} targetCell - 透视表放置的起始单元格，如 "A1"
 * @returns {Object} 操作结果
 */
function createPivot(sourceSheetName, sourceRange, rowColumnIndices, columnColumnIndices, valueColumnIndices, functionType, targetSheetName, targetCell) {
  try {
    // 验证行列不能同时为空
    if ((!rowColumnIndices || rowColumnIndices.length === 0) && 
        (!columnColumnIndices || columnColumnIndices.length === 0)) {
      return { success: false, message: "行字段和列字段不能同时为空" };
    }

    // 验证值字段不能为空
    if (!valueColumnIndices || valueColumnIndices.length === 0) {
      return { success: false, message: "值字段不能为空" };
    }

    // 获取源数据工作表
    const sourceSheet = Application.ActiveWorkbook.Sheets.Item(sourceSheetName);
    if (!sourceSheet) {
      return { success: false, message: `未找到源数据表: ${sourceSheetName}` };
    }

    // 获取源数据区域
    const sourceRangeObj = sourceSheet.Range(sourceRange);
    if (!sourceRangeObj) {
      return { success: false, message: `无效的源数据区域: ${sourceRange}` };
    }

    // 验证列索引
    const maxColumn = sourceRangeObj.Columns.Count;
    
    // 验证行字段索引
    if (rowColumnIndices) {
      for (let i = 0; i < rowColumnIndices.length; i++) {
        if (rowColumnIndices[i] < 1 || rowColumnIndices[i] > maxColumn) {
          return { success: false, message: `行字段列索引超出范围: ${rowColumnIndices[i]}` };
        }
      }
    }
    
    // 验证列字段索引
    if (columnColumnIndices) {
      for (let i = 0; i < columnColumnIndices.length; i++) {
        if (columnColumnIndices[i] < 1 || columnColumnIndices[i] > maxColumn) {
          return { success: false, message: `列字段列索引超出范围: ${columnColumnIndices[i]}` };
        }
      }
    }
    
    // 验证值字段索引
    for (let i = 0; i < valueColumnIndices.length; i++) {
      if (valueColumnIndices[i] < 1 || valueColumnIndices[i] > maxColumn) {
        return { success: false, message: `值字段列索引超出范围: ${valueColumnIndices[i]}` };
      }
    }

    // 处理目标工作表
    let targetSheet = Application.ActiveWorkbook.Sheets.Item(targetSheetName);
    if (!targetSheet) {
      // 工作表不存在，创建新工作表
      targetSheet = Application.ActiveWorkbook.Sheets.Add();
      targetSheet.Name = targetSheetName;
    } else {
      // 工作表存在，删除所有透视表
      deleteAllPivotTables(targetSheetName)
    }

    // 创建透视表缓存
    const pivotCache = Application.ActiveWorkbook.PivotCaches().Create(
      1, // xlDatabase
      sourceRangeObj
    );

    // 创建透视表
    const pivotTable = pivotCache.CreatePivotTable(
      targetSheet.Range(targetCell),
      "透视表_" + new Date().getTime()
    );

    // 添加行字段（如果指定）
    if (rowColumnIndices && rowColumnIndices.length > 0) {
      for (let i = 0; i < rowColumnIndices.length; i++) {
        const rowField = pivotTable.PivotFields(rowColumnIndices[i]);
        rowField.Orientation = 1; // xlRowField
      }
    }

    // 添加列字段（如果指定）
    if (columnColumnIndices && columnColumnIndices.length > 0) {
      for (let i = 0; i < columnColumnIndices.length; i++) {
        const columnField = pivotTable.PivotFields(columnColumnIndices[i]);
        columnField.Orientation = 2; // xlColumnField
      }
    }

    // 统计函数类型映射
    const functionMap = {
      "sum": -4157,        // xlSum
      "count": -4112,      // xlCount
      "average": -4106,    // xlAverage
      "max": -4136,        // xlMax
      "min": -4139,        // xlMin
      "product": -4149,    // xlProduct
      "countNums": -4113,  // xlCountNums
      "stdDev": -4155,     // xlStDev
      "stdDevP": -4156,    // xlStDevP
      "var": -4164,        // xlVar
      "varP": -4165        // xlVarP
    };

    // 获取统计函数类型，默认为求和
    const funcType = functionType || "sum";
    const xlFunction = functionMap[funcType] || functionMap["sum"];

    // 添加值字段
    for (let i = 0; i < valueColumnIndices.length; i++) {
      const valueField = pivotTable.PivotFields(valueColumnIndices[i]);
      valueField.Orientation = 4; // xlDataField
      valueField.Function = xlFunction;
    }

    return {
      success: true,
      message: "透视表创建成功",
      pivotSheetName: targetSheetName,
      pivotTableName: pivotTable.Name
    };

  } catch (error) {
    return {
      success: false,
      message: `创建透视表失败: ${error.message}`
    };
  }
}

/**
 * 更新指定工作表里的所有透视表
 * @param {string} sheetName - 工作表名称
 * @returns {Object} 操作结果
 */
function updateAllPivotTables(sheetName) {
  try {
    // 验证参数
    if (!sheetName) {
      return { success: false, message: "工作表名称不能为空" };
    }

    // 获取工作表
    const sheet = Application.ActiveWorkbook.Sheets.Item(sheetName);
    if (!sheet) {
      return { success: false, message: `未找到工作表: ${sheetName}` };
    }

    // 获取所有透视表
    const pivotTables = sheet.PivotTables();
    if (!pivotTables || pivotTables.Count === 0) {
      return { success: false, message: "工作表中没有透视表" };
    }

    // 更新所有透视表
    for (let i = 1; i <= pivotTables.Count; i++) {
      const pivotTable = pivotTables.Item(i);
      pivotTable.RefreshTable();
    }

    return {
      success: true,
      message: `成功更新 ${pivotTables.Count} 个透视表`,
      count: pivotTables.Count
    };

  } catch (error) {
    return {
      success: false,
      message: `更新透视表失败: ${error.message}`
    };
  }
}


/**
 * 通过清空数据区域删除透视表
 * @param {string} sheetName - 工作表名称
 * @returns {Object} 操作结果
 */
function deleteAllPivotTables(sheetName) {
  try {
    // 验证参数
    if (!sheetName) {
      return { success: false, message: "工作表名称不能为空" };
    }

    // 获取工作表
    const sheet = Application.ActiveWorkbook.Sheets.Item(sheetName);
    if (!sheet) {
      return { success: false, message: `未找到工作表: ${sheetName}` };
    }

    // 获取所有透视表
    const pivotTables = sheet.PivotTables();
    if (!pivotTables || pivotTables.Count === 0) {
      return { success: false, message: "工作表中没有透视表" };
    }

    // 记录删除的透视表数量
    let deletedCount = 0;
    let failedCount = 0;
    
    // 这个坑爹的框架，好像pivotTables.Count的数量，与表里实际的透视表数量是不一致的！
    // 因此这里重试多次，确保删光所有透视表。。。
    for (let j = 1; j <= 10; j++) {
      // 遍历所有透视表
      for (let i = 1; i <= pivotTables.Count; i++) {
        try {
          // 获取透视表对象
          const pivotTable = pivotTables.Item(i);
          
          // 获取透视表的数据区域
          const dataRange = pivotTable.TableRange2;
          
          if (dataRange) {
            // 清空数据区域
            dataRange.Clear();
            deletedCount++;
          } else {
            failedCount++;
          }
        } catch (e) {
          console.error(`处理第 ${i} 个透视表失败:`, e);
          failedCount++;
        }
      }
    }
    // 构建返回消息
    let message = "";
    if (deletedCount > 0) {
      message += `成功删除 ${deletedCount} 个透视表`;
    }
    if (failedCount > 0) {
      message += (message ? "，" : "") + `删除失败 ${failedCount} 个`;
    }

    return {
      success: deletedCount > 0,
      message: message || "没有透视表被删除",
      deletedCount: deletedCount,
      failedCount: failedCount
    };

  } catch (error) {
    return {
      success: false,
      message: `删除透视表失败: ${error.message}`
    };
  }
}

// ==================== 工具函数 ====================

/**
 * 列字母转数字索引
 * @param {string} column - 列字母，如 "A", "AB"
 * @returns {number} 列索引（从1开始）
 */
function columnLetterToNumber(column) {
  let result = 0;
  for (let i = 0; i < column.length; i++) {
    result = result * 26 + (column.charCodeAt(i) - 64);
  }
  return result;
}

/**
 * 列数字索引转字母
 * @param {number} columnNumber - 列索引（从1开始）
 * @returns {string} 列字母
 */
function columnNumberToLetter(columnNumber) {
  let letter = "";
  while (columnNumber > 0) {
    const remainder = (columnNumber - 1) % 26;
    letter = String.fromCharCode(65 + remainder) + letter;
    columnNumber = Math.floor((columnNumber - 1) / 26);
  }
  return letter;
}

/**
 * RGB 颜色转换为 Excel 颜色值
 * @param {number} r - 红色值 (0-255)
 * @param {number} g - 绿色值 (0-255)
 * @param {number} b - 蓝色值 (0-255)
 * @returns {number} Excel 颜色值 
 */
function rgbToExcelColor(r, g, b) {
  return r + g * 256 + b * 256 * 256;
}
function hexColorToRGB(hex) {
  // 将十六进制色值字符串，转换为RGB对象！
  // 例如：#FF0000 -> RGB(255, 0, 0)
  hex = hex.replace('#', '');
  // 解析 R、G、B 分量
  const r = parseInt(hex.substring(0, 2), 16);
  const g = parseInt(hex.substring(2, 4), 16);
  const b = parseInt(hex.substring(4, 6), 16);
  
  return RGB(r, g, b);
}


// 本地测试调试专用
function run_test_online() {
  let sheetName = '工作表4' 
  // let sheetName = '测试激活'

  // image_data = "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAYAAAAf8/9hAAABMklEQVR4AcxS21HDMBC8kxuBQgJOB0zCP04lCZWYAiBDB3hCI6YQ+di9SEKx+WCGn3hmR7qHdvckB/nndwUE1j7cTKvtxxzxbrPP09lq08X7bT+HtY9tkBhaUVlARQ8kd6j2atLNIfiCfh5fsI6AaIi34fSmarZjPFmzJ7jXoGuvoYcxMOrwOgRsBAeeuU5TeOJaSM1GqiLnzVil9Jj5mTNBcpFtsxFu1lyJLMBx2IPcmETECZBYuPAc7gFrac7qJsaxUZKKYOaibubt1+rN6ej2yVAcMChWY9NrUj/npcuXWauzdkmQXAifFVUSBpMvbDGidCIy1uqIf0ZgQFQKPnu6LH/mqsZWx4UDZqgA5R3e3f8F5vgiOHxgjXGNBQGLVOVPwj2hw/vCOvPErwQs/BXfAAAA//86U9wbAAAABklEQVQDAAVetyEAc7DHAAAAAElFTkSuQmCC"
  // insertImage("A1", image_data, sheetName)

  // setFilter(field, operator, criteria1,criteria2, is_reSet, sheetName)
  // setFilter(2, 'xlOr', '*宜*',"*客*", true, sheetName)
  // setFilter(2, 'xlAnd', '*胡*',"*豪", true, sheetName)
  // setFilter(4, 'xlAnd', '>20',"<30", true, sheetName)
  // setFilter(2, 'xlAnd', '*客*',"??豪*", true, sheetName)
  // setFilter(2, "xlBottom10Percent", "99", None, true, sheetName) // 示例
  // setFilter(4, "xlFilterValues", "<20", None, true, sheetName) // 示例

  // let t = getFilteredData('工作表4') // 筛选功能，仅在 2.0 版本才能使用！！！
  // console.log(t)

  // createPivot(
  //   "工作表4", "A:D", 
  //   [2,3], 
  //   [], 
  //   [4],
  //   'sum',
  //   "透视表测试_测试02", "B1"
  // )

  // updateAllPivotTables("透视表测试_测试02")
  // deleteAllPivotTables("透视表测试_测试02")

  // addWorksheet("工作表4")
  
  console.log( getUsedRangeData( "否", sheetName ) )
  // console.log( getUsedRangeData( "是", sheetName ) )
  
}



// ==================== 返回结果 ====================
return globalResult;
