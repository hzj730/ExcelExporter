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

        List<string> ColumDesc = new List<string>();    // 第一行 列描述
        List<string> ColumType = new List<string>();    // 第二行 列类型 int string json等
        List<string> ColumSrvField = new List<string>();// 第三行 服务器导出字段
        List<string> ColumCltField = new List<string>();// 第四行 客户端导出字段
        List<string> ColumFint = new List<string>();    // 第五行描述，做表无需处理
        string table_data = string.Empty;   // 记录数据表内容，一行对应一条字符串

        private void ProcessSheet(ISheet sheet)
        {
            // rest 
            ColumDesc.Clear();
            ColumType.Clear();
            ColumSrvField.Clear();
            ColumCltField.Clear();
            ColumFint.Clear();

            int TotalRowCount = sheet.LastRowNum;
            if (TotalRowCount < 5)
            {
                Console.Error.WriteLine("Excel文件格式不正确行数过少");
                return;
            }
            // 第一行 记录列数 记录列描述
            IRow row1 = sheet.GetRow(0);
            int colums = 0;
            foreach (var cell in row1.Cells)
            {
                if (cell == null || cell.CellType == CellType.Blank)
                {
                    break;
                }
                colums++;
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

            for (int row = 5; row <= TotalRowCount; row++)
            {
                IRow rowData = sheet.GetRow(row);
                if (rowData != null) //null is when the row only contains empty cells 
                {
                    //MessageBox.Show(string.Format("Row {0} = {1}", row, sheet.GetRow(row).GetCell(0).StringCellValue));
                    table_data += ProcessRow(row, rowData);
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
            for (int i = 0; i < ColumCltField.Count; i++)
            {
                string field = ColumCltField[i];
                if (!field.StartsWith("##"))
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
                            if (jsonObj is JArray)
                            {
                                var arr = JArray.Parse(json);
                                string s = "{";
                                foreach (var item in arr)
                                {
                                    if (item.Type == JTokenType.Object)
                                    {
                                        var obj = (JObject)item;
                                        s += JsonObjectToLuaStr(obj);
                                    }
                                }
                                s += "}";
                                row_data += s;
                            }
                            else
                            {
                                var obj = JObject.Parse(json);
                                row_data += JsonObjectToLuaStr(obj);
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
            }

            row_data += "},\n";

            return row_data;
        }

        string JsonObjectToLuaStr(JObject obj)
        {
            string s = "{";
            foreach (var kv in obj)
            {
                if (kv.Value.Type == JTokenType.Object)
                {
                    var sub_obj = (JObject)kv.Value;
                    s += JsonObjectToLuaStr(sub_obj);
                }
                else // TODO 处理更完整的对象类型
                {
                    s += kv.Value.ToString() + ",";
                }
            }
            s += "}";
            return s;
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
            WriteLuaFile(Define.TABLE_FIELD_TEMP, "TableFieldDef.lua");
        }
    }
}
