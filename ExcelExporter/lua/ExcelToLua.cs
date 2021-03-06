﻿using ExcelExporter.common;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;

namespace ExcelExporter.lua
{
    public class ExcelToLua
    {
        string DirExcel { get; set; } = "";
        string ExcelFileName { get; set; } = "";
        string ExcelSheetName { get; set; } = "";

        // 解析目录下所有的excel文件
        public void PackageDirectory(string relativePath = "")
        {
            ClearGlobalData();

            string rootPath = Directory.GetCurrentDirectory();
            string dir = relativePath != string.Empty ? rootPath + relativePath : rootPath + "\\..\\table_xlsx\\";
            DirExcel = dir;
            var xlsFiles = Directory.GetFiles(dir, "*.xlsx");

            foreach (var file in xlsFiles)
            {
                if (!file.Contains("~$"))
                {
                    AnalysisExcelFile(file);
                }
            }
        }

        Dictionary<int, string> ColumDescMap = new Dictionary<int, string>();    // 第一行 列描述
        Dictionary<int, string> ColumTypeMap = new Dictionary<int, string>();    // 第二行 列类型 int string json等
        List<string> ColumSrvField = new List<string>();// 第三行 服务器导出字段
        Dictionary<int, string> ColumCltFieldMap = new Dictionary<int, string>(); // 第四行 客户端导出字段 k: colum id    v: cell string
        List<string> ColumFint = new List<string>();    // 第五行描述，做表无需处理
        string table_data = string.Empty;   // 记录数据表内容，一行对应一条字符串

        List<string> TableFieldDefines = new List<string>();
        Dictionary<string, string> TableFieldMap = new Dictionary<string, string>();    // k: sheetName v: fields define lua string

        void ClearSheetData()
        {
            ColumDescMap.Clear();
            ColumTypeMap.Clear();
            ColumSrvField.Clear();
            ColumCltFieldMap.Clear();
            ColumFint.Clear();
            table_data = string.Empty;
        }

        void ClearGlobalData()
        {
            TableFieldDefines.Clear();
            TableFieldMap.Clear();
        }

        // 解析单个文件
        public bool AnalysisExcelFile(string fileName)
        {
            bool ret = false;
            ISheet sheet = null;

            //try
            {
                //fs = new FileStream(fileName, FileMode.Open, FileAccess.Read);
                //workbook = WorkbookFactory.Create(fs);

                XSSFWorkbook workbook;
                using (FileStream file = new FileStream(fileName, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                {
                    workbook = new XSSFWorkbook(file);
                    string exls_file_name = Path.GetFileName(fileName);
                    string exls_file_name_wihtout_suffix = exls_file_name.Split('.')[0];
                    int sheetCount = workbook.NumberOfSheets;
                    for (int i = 0; i < sheetCount; i++)
                    {
                        sheet = workbook.GetSheetAt(i);

                        ClearSheetData();
                        var content = ProcessSheet(exls_file_name_wihtout_suffix, sheet);
                        if (!content.Equals(string.Empty))
                        {
                            string table_name = exls_file_name_wihtout_suffix + "_" + sheet.SheetName;
                            WriteLuaFile(content, table_name + ".lua");
                        }
                    }
                }

                ret = true;
            }
            //catch(Exception e)
            //{
            //    Console.WriteLine(e.ToString());
            //}
            //finally
            //{
            //}

            return ret;
        }

        string FormatErrorMsg(string msg, int rowIndex, int colIndex)
        {
            return string.Format("Excel: {0} Sheet: {1} Row: {2} Colum: {3}, Error: {4}",
                ExcelFileName, ExcelSheetName, rowIndex.ToString(), colIndex.ToString(), msg);
        }

        private string ProcessSheet(string exls_name, ISheet sheet)
        {
            ExcelFileName = exls_name;
            ExcelSheetName = sheet.SheetName;

            // 第一行 记录列数 记录列描述
            IRow row1 = sheet.GetRow(0);
            // 第一个单元格不是array字段，该表不需要导出
            if (row1 == null)
                return string.Empty;
            var cell011 = row1.GetCell(0);
            if (cell011 == null || cell011.StringCellValue != "array")
            {
                return string.Empty;
            }
            int colums = 0;
            foreach (var cell in row1.Cells)
            {
                if (cell == null || cell.CellType == CellType.Blank)
                {
                    continue;
                }
                ColumDescMap.Add(cell.Address.Column, cell.ToString());
                colums++;
            }

            int TotalRowCount = sheet.LastRowNum;
            if (TotalRowCount < 5)
            {
                Console.Error.WriteLine("Excel文件格式不正确行数过少");
                return string.Empty;
            }

            IRow row2 = sheet.GetRow(1);
            int type_colums_count = 0;
            foreach (var cell in row2.Cells)
            {
                if (cell == null || cell.CellType == CellType.Blank)
                {
                    continue;
                }
                type_colums_count++;
//                ColumType.Add(cell.ToString());
                ColumTypeMap.Add(cell.Address.Column, cell.ToString());
            }

            IRow row3 = sheet.GetRow(2);
            int srv_colums_count = 0;
            foreach (var cell in row3.Cells)
            {
                if (cell == null || cell.CellType == CellType.Blank)
                {
                    continue;
                }
                srv_colums_count++;
            }

            IRow row4 = sheet.GetRow(3);    // client 导出标记
            int clt_colums_count = 0;
            foreach (var cell in row4.Cells)
            {
                if (cell == null || cell.CellType == CellType.Blank)
                {
                    continue;
                }
                clt_colums_count++;
//                ColumCltField.Add(cell.ToString());
                ColumCltFieldMap.Add(cell.Address.Column, cell.ToString());
            }

            ProcessSheetFields(exls_name, sheet.SheetName, row4, row1, row2);

            SheetIds.Clear();

            for (int row = Define.StartRowIndex; row <= TotalRowCount; row++)
            {
                IRow rowData = sheet.GetRow(row);
                if (rowData != null) //null is when the row only contains empty cells 
                {
                    //MessageBox.Show(string.Format("Row {0} = {1}", row, sheet.GetRow(row).GetCell(0).StringCellValue));
                    ICell cell0 = rowData.GetCell(0);
                    if (cell0 != null && cell0.ToString().StartsWith("##"))
                    {
                        Console.WriteLine(string.Format("Ignore Sheet {0} Row {1}.", sheet.SheetName, row));
                    }
                    else
                    {
                        if (exls_name.Equals("characters_red") && sheet.SheetName.Equals("skill"))
                        {
                            CheckCharacterRowJsonConfig(row1, rowData, 6, 9); // effect_id, hit_data
                            CheckCharacterRowJsonConfig(row1, rowData, 6, 10); // effect_id, hit_time
                            CheckCharacterRowJsonConfig(row1, rowData, 6, 5); // effect_id, target_relation
                        }
                        table_data += ProcessRow(row, rowData);
                    }
                }
            }

            string table_name = exls_name + "_" + sheet.SheetName;
            string content = string.Format(Define.TABLE_DATA_TEMP, table_data, table_name);
            //Console.WriteLine(content);

            return content;

            //WriteLuaFile(content, table_name + ".lua");
        }

        List<string> SheetIds = new List<string>();
        // rowIndex 1-5 文件头
        private string ProcessRow(int rowIndex, IRow row)
        {
            string row_data = string.Empty;
            int lastCellNum = row.LastCellNum;

            row_data += "{";
            // 第一列默认忽略
            foreach (KeyValuePair<int, string> kv in ColumCltFieldMap)
            {
                int field_colum_id = kv.Key;
                if (field_colum_id < Define.StartColumIndex)
                {
                    continue;
                }
                if (!ColumDescMap.ContainsKey(field_colum_id))
                {
                    // 后续列都认为无效
                    break;
                }
                // 根据类型区分处理
                string type = string.Empty;
                if (!ColumTypeMap.TryGetValue(field_colum_id, out type))
                {
                    Console.Error.WriteLine(string.Format("ColumTypeMap TryGetValue exception row: {0} colum: {1}", rowIndex, field_colum_id));
                    continue;
                }
                ICell cell = row.GetCell(field_colum_id);

                if (field_colum_id == 1) // Key列做一些规则和格式检查
                {
                    if (cell == null || cell.ToString() == string.Empty)
                    {
                        string message = FormatErrorMsg("ID不能填空", rowIndex + 1, field_colum_id);
                        throw new ArgumentException(message, "CellType");
                    }
                    string cell_id = cell.ToString();
                    if (SheetIds.Contains(cell_id))
                    {
                        string message = FormatErrorMsg("ID重复", rowIndex + 1, field_colum_id);
                        throw new ArgumentException(message, "CellType");
                    }
                    SheetIds.Add(cell_id);
                }

                if (type == "int" || type == "number" || type == "num" || type == "float")
                {
                    row_data += (cell != null) ? cell.NumericCellValue.ToString() : Define.DefaultLuaNum;// Utils.GetCellValue(row.GetCell(i)); //
                    row_data += ",";
                }
                else if (type == "string" || type == "str")
                {
                    try
                    {
                        if (cell != null)
                        {
                            cell.SetCellType(CellType.String);
                            string content = cell.StringCellValue;

                            content = content.Replace("\n", "");
                            content = content.Replace("\r", "");
                            StringBuilder sb = new StringBuilder();
                            foreach (char ch in content)
                            {
                                if (ch == '\"' || ch == '\\')
                                {
                                    sb.Append('\\');
                                }
                                sb.Append(ch);
                            }
                            content = sb.ToString();

                            row_data += "\"" + content + "\"";
                        }
                        else
                        {
                            row_data += Define.DefaultLuaStr;
                        }
                        row_data += ",";
                    }
                    catch (Exception e)
                    {
                        Console.Error.WriteLine("StringCellValue Exception: " + e.ToString() + " --- colum:" + field_colum_id.ToString() + " --- row: " + rowIndex.ToString());
                    }
                }
                else if (type == "json")
                {
                    if (cell != null)
                    {
                        try
                        {
                            //DataFormatter formatter = new DataFormatter(); //creating formatter using the default locale
                            //string json = formatter.FormatCellValue(cell); //Returns the formatted value of a cell as a String regardless of the cell type.

                            string json = cell.StringCellValue;
                            json = json.Replace("\n", "");
                            json = json.Replace("\r", "");
                            var jsonObj = JsonConvert.DeserializeObject(json);
                            //Console.WriteLine(JsonConvert.DeserializeObject(json) is Array);
                            if (jsonObj != null)
                            {
                                row_data += JsonObjectToLuaStr(jsonObj);
                            }
                            else
                            {
                                row_data += Define.DefaultLuaTable + ",";
                            }
                        }
                        catch (Exception)
                        {
                            string message = FormatErrorMsg("解析Json异常", rowIndex + 1, field_colum_id);
                            throw new ArgumentException(message, cell.CellType.ToString());
                        }
                    }
                    else
                    {
                        row_data += Define.DefaultLuaTable + ",";
                    }
                }
                else
                {
                    throw new ArgumentException("Cell Type Unknow", "CellType");
                }
            }

            row_data += "},\n";

            return row_data;
        }

        string JsonObjectToLuaStr(object obj)
        {
            if (obj is JArray)
            {
                string s = "{";
                foreach (var item in (JArray)obj)
                {
                    if (item.Type == JTokenType.Object
                        || item.Type == JTokenType.Array)
                    {
                        s += JsonObjectToLuaStr(item);
                    }
                    else if (item.Type == JTokenType.String)
                    {
                        s += "\"" + item.ToString() + "\",";
                    }
                    else if (item.Type == JTokenType.Boolean)
                    {
                        s += item.ToString().ToLower() + ",";
                    }
                    else if (item.Type == JTokenType.Integer || item.Type == JTokenType.Float)
                    {
                        s += item.ToString() + ",";
                    }
                }
                s += "},";
                return s;
            }
            else
            {
                string s = "{";
                foreach (var kv in (JObject)obj)
                {
                    if (kv.Value.Type == JTokenType.Object
                        || kv.Value.Type == JTokenType.Array)
                    {
                        s += JsonObjectToLuaStr(kv.Value);
                    }
                    else if (kv.Value.Type == JTokenType.String)
                    {
                        s += "\"" + kv.Value.ToString() + "\",";
                    }
                    else if (kv.Value.Type == JTokenType.Boolean)
                    {
                        s += kv.Value.ToString().ToLower() + ",";
                    }
                    else // TODO 处理更完整的对象类型
                    {
                        s += kv.Value.ToString() + ",";
                    }
                }
                s += "},";
                return s;
            }
        }

        // 解析sheet时同时解析并存储sheet的fields数据
        // @clientFields: ##client 行数据
        // @desc: 表首行数据
        void ProcessSheetFields(string excelName, string sheetName, IRow clientFields, IRow desc, IRow dataType)
        {
            string fieldsStr = @"";
            int lua_start_index = 1;
            foreach (KeyValuePair<int, string> kv in ColumCltFieldMap)
            {
                int colum_index = kv.Key;
                if (colum_index < Define.StartColumIndex)
                {
                    continue;
                }
                if (!ColumDescMap.ContainsKey(colum_index))
                {
                    // 后续列都认为无效
                    break;
                }
                ICell cell = clientFields.GetCell(colum_index);
                ICell cellDataType = dataType.GetCell(colum_index);
                if (cellDataType.StringCellValue == "json")
                {
                    if (cell != null && cell.CellType == CellType.String && cell.StringCellValue != string.Empty)
                    {
                        // 格式：field = i, -- desc\r\n
                        fieldsStr += "        " + cell.StringCellValue + " = " + lua_start_index + ", -- " + desc.GetCell(colum_index).ToString() + "\r\n";
                    }
                }
                else
                {
                    if (cell != null && cell.CellType == CellType.String && cell.StringCellValue != string.Empty)
                    {
                        // 格式：field = i, -- desc\r\n
                        fieldsStr += "        " + cell.StringCellValue + " = " + lua_start_index + ", -- " + desc.GetCell(colum_index).ToString() + "\r\n";
                    }
                }

                lua_start_index++;
            }

            string table_name = excelName + "_" + sheetName;
            string data = string.Format(Define.SINGLE_TABLE_FIELD_TEMPLATE,
                        table_name, fieldsStr, table_name + ".lua");
            if (TableFieldMap.ContainsKey(table_name))
            {
                throw new Exception(string.Format("sheet name {0} has existed.", table_name));
            }
            else
            {
                TableFieldMap.Add(table_name, data);
            }
        }

        private void WriteLuaFile(string content, string fileName)
        {
            string luaPath = DirExcel + "/lua/";
            if (!Directory.Exists(luaPath))
            {
                Directory.CreateDirectory(luaPath);
            }
            using (StreamWriter file = new StreamWriter(luaPath + fileName))
            {
                file.Write(content);
            }
        }

        public void genTableFieldLua()
        {
            string fieldsStr = string.Empty;
            foreach (var k in TableFieldMap.Keys)
            {
                string str = TableFieldMap[k];
                fieldsStr += str;
            }
            // Define.TABLE_FIELD_TEMPLATE 应该包含了多个 SINGLE_TABLE_FIELD_TEMPLATE
            string data = string.Format(Define.TABLE_FIELD_TEMPLATE, fieldsStr);
            WriteLuaFile(data, "TableFieldDef.lua");
        }

        // 检测角色配置表的一行中具体某几列的json数组长度一致性
        void CheckCharacterRowJsonConfig(IRow head, IRow row, int accord, int check)
        {
            ICell cellHead0 = head.GetCell(accord);
            ICell cellHead1 = head.GetCell(check);
            ICell cellId = row.GetCell(1);

            int accordArrLen = 0;
            int checkArrLen = 0;
            ICell cell0 = row.GetCell(accord);
            try
            {
                string json = cell0.StringCellValue;
                json = json.Replace("\n", "");
                json = json.Replace("\r", "");
                var jsonObj = JsonConvert.DeserializeObject(json);
                accordArrLen = (jsonObj as JArray).Count;
            }
            catch (Exception)
            {
                throw new Exception(string.Format("Json 格式错误 技能ID: {0} 出错列：{1}", cellId.NumericCellValue, cellHead0.StringCellValue));
            }

            ICell cell1 = row.GetCell(check);
            try
            {
                string json = cell1.StringCellValue;
                json = json.Replace("\n", "");
                json = json.Replace("\r", "");
                var jsonObj = JsonConvert.DeserializeObject(json);
                checkArrLen = (jsonObj as JArray).Count;
            }
            catch (Exception)
            {
                throw new Exception(string.Format("Json 格式错误 技能ID: {0} 出错列：{1}", cellId.NumericCellValue, cellHead1.StringCellValue));
            }

            if (accordArrLen > 0 && accordArrLen != checkArrLen)
            {
                throw new Exception(string.Format("Json 格式错误 数组长度不一致 技能ID: {0} {1} --- {2}", cellId.NumericCellValue, cellHead0.StringCellValue, cellHead1.StringCellValue));
            }
        }
    }
}
