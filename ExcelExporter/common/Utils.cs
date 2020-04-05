using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelExporter.common
{
    public class Utils
    {
        public static string GetCellValue(ICell cell)
        {
            if (cell.CellType == CellType.Formula)
            {
                return cell.StringCellValue;
            }
            else
            {
                return cell.ToString();
            }
        }
    }
}
