using System;
using Newtonsoft.Json;
using System.IO;

using exceltools.helpers;

namespace exceltools
{
    class Program
    {      
		static void Main(string[] args)
		{
			if (args.Length < 3)
			{
				Console.WriteLine("Arguments not found!");
				Console.WriteLine("1 - required - action (excel2csv / csv2excel)");
                Console.WriteLine("2 - required - FileIn");
				Console.WriteLine("3 - required - FileOut");
                Console.WriteLine("4 - optional - JSON config");
				Environment.Exit(1);
			}

            string inFile = args[1];
			string outFile = args[2];

            ExcelTools obj = new ExcelTools();

            switch (args[0])
			{
				case "excel2csv":
                    converterSettings[] settings = null;

                    if (args.Length > 3)
                    {
                        settings = JsonConvert.DeserializeObject<converterSettings[]>(args[2]);
                    }

                    obj.csv2excel(inFile, outFile, settings);
					break;
				case "csv2excel":
					throw new Exception("Action not implemented");
                    break;
			}

			Console.WriteLine(File.Exists(outFile) ? "ok" : "error");
        }
        
    }
}
