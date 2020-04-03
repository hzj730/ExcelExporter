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

            if (args[0] == "lua")
            {

            }
            else if (args[0] == "tars")
            {

            }
            else
            {
                Console.WriteLine("参数错误");
            }
        }
    }
}
