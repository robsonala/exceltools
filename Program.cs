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
            //runTest();

			if (args.Length < 3)
			{
				Console.WriteLine("Arguments not found!");
				Console.WriteLine("0 - required - action (excel2csv / csv2excel)");
                Console.WriteLine("1 - required - FileIn");
				Console.WriteLine("2 - required - FileOut");
                Console.WriteLine("3 - optional - JSON config");
				Environment.Exit(1);
			}

            string inFile = args[1];
			string outFile = args[2];
			         
            ExcelTools obj = new ExcelTools();

            switch (args[0])
			{
				case "excel2csv":
                    converterToExcelSettings[] settingsExcel = null;

                    if (args.Length > 3)
                    {
                        settingsExcel = JsonConvert.DeserializeObject<converterToExcelSettings[]>(args[3]);
                    }

                    obj.csv2excel(inFile, outFile, settingsExcel);
					break;
                case "csv2excel":
                    converterToCsvSettings settingsCsv = null;

                    if (args.Length > 3)
                    {
                        settingsCsv = JsonConvert.DeserializeObject<converterToCsvSettings>(args[3]);
                    }

                    obj.excel2csv(inFile, outFile, settingsCsv);
                    break;
			}

			Console.WriteLine(File.Exists(outFile) ? "ok" : "error");
        }

        static void runTest()
        {
            ExcelTools objx = new ExcelTools();

            string _inFile = "/Users/robsonaugustolazzarinalviani/projects/files/out/112_salesout-week_kobo_20180507093048_569695af00eb813a6e.xlsx";
            string _outFile = "/Users/robsonaugustolazzarinalviani/projects/files/out/112_salesout-week_kobo_20180507093048_569695af00eb813a6e.xlsx.csv";

            converterToCsvSettings testSettings = new converterToCsvSettings();
            testSettings.Sheets = new string[] { "Sheet1" };
            testSettings.SkipHidden = true;

            objx.excel2csv(_inFile, _outFile, testSettings);

            //string _inFile = "/Users/robsonaugustolazzarinalviani/projects/files/in/112_salesout-week_kobo_20180507093048_569695af00eb813a6e.csv";
            //string _outFile = "/Users/robsonaugustolazzarinalviani/projects/files/out/112_salesout-week_kobo_20180507093048_569695af00eb813a6e.xlsx";

            //objx.csv2excel(_inFile, _outFile);

            Console.WriteLine(File.Exists(_outFile) ? "ok" : "error");

            Environment.Exit(1);
        }
        
    }
}
