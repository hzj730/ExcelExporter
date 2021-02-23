using ExcelExporter.json;
using ExcelExporter.lua;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelExporter
{
    class Program
    {
        // 缺省输出格式
        static string DefaultExportType = "json";
        static void Main(string[] args)
        {
            int argsLength = args.Length;
            
            Console.WriteLine("CurrentDirectory: " + Environment.CurrentDirectory);
            Console.WriteLine(argsLength);
            for (int i = 0; i < argsLength; i++)
            {
                Console.Write("第" + (i + 1) + "个参数是：");
                Console.WriteLine(args[i]);
            }

            string dir = (args.Length > 1 && args[1] != null) ? args[1] : string.Empty;
            string exportType = args.Length > 0 ? args[0] : DefaultExportType; // defalut export lua
            string exportPlat = (args.Length >= 3 && args[2] != null) ? args[2] : "client";

            if (exportType == "lua")
            {
                try
                {
                    ExcelToLua excelToLua = new ExcelToLua();
                    excelToLua.PackageDirectory(dir);

                    excelToLua.genTableFieldLua();
                }
                catch (Exception e)
                {
                    Console.Error.WriteLine(e.ToString());
                    Console.ReadKey();
                }
            }
            else if (exportType == "json")
            {
                var excelToJson = new ExcelToJson();
                if (dir.Contains(".xlsx"))
                {
                    excelToJson.PackageFile(exportPlat, dir);
                }
                else
                {
                    excelToJson.PackageDirectory(exportPlat, dir);
                }
            }
            else if (exportType == "tars")
            {

            }
            else
            {
                Console.Error.WriteLine("参数错误");
                Console.ReadKey();
            }

            //Console.ReadKey();
        }
    }
}
