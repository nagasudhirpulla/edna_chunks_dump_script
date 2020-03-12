using EdnaExcelConfigUtils;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EdnaChunksDumpConsoleApp
{
    class Program
    {
        static void Main(string[] args)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            ExcelUtils excelUtils = new ExcelUtils();
            //EdnaFetchConfig config = excelUtils.ReadFromExcel(@"C:\Nagasudhir\Desktop\EdnaFetchConfig.xlsx");
            string configFilePath = "EdnaFetchConfig.xlsx";
            for (int argIter = 0; argIter < args.Length; argIter++)
            {
                // Console.WriteLine(args[argIter]);
                if (args[argIter] == "--config")
                {
                    configFilePath = args[argIter + 1];
                }
            }
            EdnaFetchConfig config = excelUtils.ReadFromExcel(configFilePath);
            DataFetcher fetcher = new DataFetcher(config);
            fetcher.FetchAndDumpData();
        }
    }
}
