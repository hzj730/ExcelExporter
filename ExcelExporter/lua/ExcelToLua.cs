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
        public static bool AnalysisExcelFile(string fileName)
        {
            bool ret = false;
            ISheet sheet = null;

            try
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

                        DealSheet(sheet);
                    }
                }

                ret = true;
            }
            catch(Exception e)
            {
                Console.WriteLine(e.ToString());
            }
            finally
            {
            }

            return ret;
        }

        private static void DealSheet(ISheet sheet)
        {
            for (int row = 0; row <= sheet.LastRowNum; row++)
            {
                if (sheet.GetRow(row) != null) //null is when the row only contains empty cells 
                {
                    //MessageBox.Show(string.Format("Row {0} = {1}", row, sheet.GetRow(row).GetCell(0).StringCellValue));
                }
            }
        }
    }
}
