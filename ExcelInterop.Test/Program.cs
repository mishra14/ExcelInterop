using System;
using System.Collections.Generic;

namespace ExcelInterop.Sample
{
    class Program
    {
        static void Main(string[] args)
        {
			TestDataLoad();
		}

		private static void TestDataSave()
		{
			var data = new List<IList<string>>()
			{
				new List<string>() { "hello", "high"},
				new List<string>() { "hi", "bye"},
			};

			ExcelUtility.SaveAsExcel(data, @"F:\validation\demo\codes.xlsx");
		}

		private static void TestDataLoad()
		{
			var data = ExcelUtility.ReadFromExcel(@"F:\validation\demo\codes.xlsx", "Pack");
		}
	}
}
