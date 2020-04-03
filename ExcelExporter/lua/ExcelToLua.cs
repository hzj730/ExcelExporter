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
            IWorkbook workbook = null;
            ISheet sheet = null;
            FileStream fs = null;

            try
            {
                fs = new FileStream(fileName, FileMode.Open, FileAccess.Read);
                workbook = WorkbookFactory.Create(fs);

                if (fileName.IndexOf(".xlsx") > 0) // 2007版本
                    workbook = new XSSFWorkbook();
                else if (fileName.IndexOf(".xls") > 0) // 2003版本
                    workbook = new HSSFWorkbook();

                ret = true;
            }
            catch
            {

            }
            finally
            {
                if (fs != null)
                {
                    fs.Close();
                    fs = null;
                }
            }

            return ret;
        }
    }
}
