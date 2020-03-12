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
            EdnaFetchConfig config = excelUtils.ReadFromExcel("EdnaFetchConfig.xlsx");
            DataFetcher fetcher = new DataFetcher(config);
            fetcher.FetchAndDumpData();
        }
    }
}
