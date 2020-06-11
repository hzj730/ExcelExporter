using ExcelExporter.lua;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelExporter
{
    class Program
    {
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

            if (args.Length == 0 || args[0] == "lua")
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
                }
            }
            else if (args[0] == "tars")
            {

            }
            else
            {
                Console.WriteLine("参数错误");
            }

            //Console.ReadKey();
        }
    }
}
