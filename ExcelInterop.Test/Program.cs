using System;
using System.Collections.Generic;

namespace ExcelInterop.Test
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Hello World!");

            var data = new List<IList<string>>()
            {
                new List<string>() { "hello", "high"},
                new List<string>() { "hi", "bye"},
            };

            ExcelUtility.SaveAsExcel(data, @"F:\validation\demo\codes.xlsx");
        }
    }
}
