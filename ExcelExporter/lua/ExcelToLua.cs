using ExcelExporter.common;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;

namespace ExcelExporter.lua
{
    public class ExcelToLua
    {
        // 解析目录下所有的excel文件
        public void PackageDirectory()
        {
            ClearGlobalData();

            var xlsFiles = Directory.GetFiles(Define.ExcelPath, "*.xlsx");

            foreach (var file in xlsFiles)
            {
                if (!file.Contains("~$"))
                {
                    AnalysisExcelFile(file);
                }
            }
        }

        List<string> ColumDesc = new List<string>();    // 第一行 列描述
        List<string> ColumType = new List<string>();    // 第二行 列类型 int string json等
        List<string> ColumSrvField = new List<string>();// 第三行 服务器导出字段
        List<string> ColumCltField = new List<string>();// 第四行 客户端导出字段
        List<string> ColumFint = new List<string>();    // 第五行描述，做表无需处理
        string table_data = string.Empty;   // 记录数据表内容，一行对应一条字符串

        List<string> TableFieldDefines = new List<string>();
        Dictionary<string, string> TableFieldMap = new Dictionary<string, string>();    // k: sheetName v: fields define lua string

        void ClearSheetData()
        {
            ColumDesc.Clear();
            ColumType.Clear();
            ColumSrvField.Clear();
            ColumCltField.Clear();
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
                    int sheetCount = workbook.NumberOfSheets;
                    for (int i = 0; i < sheetCount; i++)
                    {
                        sheet = workbook.GetSheetAt(i);

                        ClearSheetData();
                        ProcessSheet(sheet);
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

        private void ProcessSheet(ISheet sheet)
        {
            // 第一行 记录列数 记录列描述
            IRow row1 = sheet.GetRow(0);
            // 第一个单元格不是array字段，该表不需要导出
            if (row1 == null || row1.GetCell(0) == null || row1.GetCell(0).StringCellValue != "array")
            {
                return;
            }
            int colums = 0;
            foreach (var cell in row1.Cells)
            {
                if (cell == null || cell.CellType == CellType.Blank)
                {
                    break;
                }
                colums++;
            }

            int TotalRowCount = sheet.LastRowNum;
            if (TotalRowCount < 5)
            {
                Console.Error.WriteLine("Excel文件格式不正确行数过少");
                return;
            }

            IRow row2 = sheet.GetRow(1);
            int type_colums_count = 0;
            foreach (var cell in row2.Cells)
            {
                if (cell == null || cell.CellType == CellType.Blank)
                {
                    break;
                }
                type_colums_count++;
                ColumType.Add(cell.ToString());
            }

            IRow row3 = sheet.GetRow(2);
            int srv_colums_count = 0;
            foreach (var cell in row3.Cells)
            {
                if (cell == null || cell.CellType == CellType.Blank)
                {
                    break;
                }
                srv_colums_count++;
            }

            IRow row4 = sheet.GetRow(3);    // client 导出标记
            int clt_colums_count = 0;
            foreach (var cell in row4.Cells)
            {
                if (cell == null || cell.CellType == CellType.Blank)
                {
                    break;
                }
                clt_colums_count++;
                ColumCltField.Add(cell.ToString());
            }

            ProcessSheetFields(sheet.SheetName, row4, row1, row2);

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
                        table_data += ProcessRow(row, rowData);
                    }
                }
            }

            string content = string.Format(Define.TABLE_DATA_TEMP, table_data, sheet.SheetName);
            //Console.WriteLine(content);

            WriteLuaFile(content, sheet.SheetName + ".lua");
        }

        // rowIndex 1-5 文件头
        private string ProcessRow(int rowIndex, IRow row)
        {
            string row_data = string.Empty;
            int lastCellNum = row.LastCellNum;

            row_data += "{";
            // 第一列默认忽略
            for (int i = Define.StartColumIndex; i < ColumCltField.Count; i++)
            {
                // 根据类型区分处理
                string type = ColumType[i];
                ICell cell = row.GetCell(i);

                if (type == "int")
                {
                    row_data += (cell != null) ? cell.NumericCellValue.ToString() : Define.DefaultNum;// Utils.GetCellValue(row.GetCell(i)); //
                }
                else if (type == "string")
                {
                    row_data += (cell != null) ? "\"" + cell.StringCellValue + "\"" : Define.DefaultStr;
                }
                else if (type == "json")
                {
                    if (cell != null)
                    {
                        string json = cell.StringCellValue;
                        var jsonObj = JsonConvert.DeserializeObject(json);
                        //Console.WriteLine(JsonConvert.DeserializeObject(json) is Array);
                        if (jsonObj != null)
                        {
                            row_data += JsonObjectToLuaStr(jsonObj);
                        }
                        else
                        {
                            row_data += Define.DefaultTable;
                        }
                    }
                    else
                    {

                    }
                }
                else
                {
                    throw new ArgumentException("Cell Type Unknow", "CellType");
                }
                row_data += ",";
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
                    else if (item.Type == JTokenType.Integer || item.Type == JTokenType.Float)
                    {
                        s += item.ToString() + ",";
                    }
                }
                s += "}";
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
                    else // TODO 处理更完整的对象类型
                    {
                        s += kv.Value.ToString() + ",";
                    }
                }
                s += "}";
                return s;
            }
        }

        // 解析sheet时同时解析并存储sheet的fields数据
        // @clientFields: ##client 行数据
        // @desc: 表首行数据
        void ProcessSheetFields(string sheetName, IRow clientFields, IRow desc, IRow dataType)
        {
            string fieldsStr = @"";
            for (int i = Define.StartColumIndex; i < ColumCltField.Count; i++)
            {
                ICell cell = clientFields.GetCell(i);
                ICell cellDataType = dataType.GetCell(i);
                if (cellDataType.StringCellValue == "json")
                {
                    if (cell != null && cell.CellType == CellType.String && cell.StringCellValue != string.Empty)
                    {
                        // 格式：field = i, -- desc\r\n
                        fieldsStr += "        " + cell.StringCellValue + " = " + i + ", -- " + desc.GetCell(i).ToString() + "\r\n";
                    }
                }
                else
                {
                    if (cell != null && cell.CellType == CellType.String && cell.StringCellValue != string.Empty)
                    {
                        // 格式：field = i, -- desc\r\n
                        fieldsStr += "        " + cell.StringCellValue + " = " + i + ", -- " + desc.GetCell(i).ToString() + "\r\n";
                    }
                }
            }

            string data = string.Format(Define.SINGLE_TABLE_FIELD_TEMPLATE,
                        sheetName, fieldsStr, sheetName + ".lua");
            TableFieldMap.Add(sheetName, data);
        }

        private void WriteLuaFile(string content, string fileName)
        {
            if (!Directory.Exists(Define.OutputLuaPath))
            {
                Directory.CreateDirectory(Define.OutputLuaPath);
            }
            using (StreamWriter file = new StreamWriter(Define.OutputLuaPath + fileName))
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
    }
}
