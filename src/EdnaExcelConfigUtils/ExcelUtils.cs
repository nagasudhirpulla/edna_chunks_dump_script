using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using OfficeOpenXml;

namespace EdnaExcelConfigUtils
{
    public class ExcelUtils
    {
        // get directory of a running exe file
        public const string PointsSheetName = "pnts";
        public const string ConfigSheetName = "config";

        public EdnaFetchConfig ReadFromExcel(string configExcelPath)
        {
            FileInfo fileInfo = new FileInfo(configExcelPath);
            if (!fileInfo.Exists)
            {
                Console.WriteLine($"Edna fetch config file {fileInfo.FullName} does not exist");
                return null;
            }
            ExcelPackage package = new ExcelPackage(fileInfo);

            EdnaFetchConfig fetchConfig = new EdnaFetchConfig
            {
                FetchWindow = new TimeSpan()
            };

            // get points from pnts worksheet
            bool isPntsSheetExists = package.Workbook.Worksheets.Any(sheet => sheet.Name == ExcelUtils.PointsSheetName);
            if (!isPntsSheetExists)
            {
                Console.WriteLine($"{ExcelUtils.PointsSheetName} Sheet is absent in {configExcelPath}, hence points can't be read");
                return null;
            }
            ExcelWorksheet worksheet = package.Workbook.Worksheets[ExcelUtils.PointsSheetName];
            List<EdnaPnt> points = new List<EdnaPnt>();
            // loop all rows
            for (int i = worksheet.Dimension.Start.Row; i <= worksheet.Dimension.End.Row; i++)
            {
                //loop first 2 column in the row
                object cellVal = worksheet.Cells[i, 1].Value;
                if (cellVal == null)
                {
                    continue;
                }
                string pntId = cellVal.ToString();
                cellVal = worksheet.Cells[i, 2].Value;
                if (cellVal == null)
                {
                    continue;
                }
                string pntName = cellVal.ToString();

                points.Add(new EdnaPnt { Id = pntId, Name = pntName });
            }

            fetchConfig.Pnts = points;

            // get config from "config" worksheet
            bool isConfigSheetExists = package.Workbook.Worksheets.Any(sheet => sheet.Name == ExcelUtils.ConfigSheetName);
            if (!isPntsSheetExists)
            {
                Console.WriteLine($"{ExcelUtils.ConfigSheetName} Sheet is absent in {configExcelPath}, hence config can't be read");
                return null;
            }
            worksheet = package.Workbook.Worksheets[ExcelUtils.ConfigSheetName];
            // loop all rows
            for (int i = worksheet.Dimension.Start.Row; i <= worksheet.Dimension.End.Row; i++)
            {
                //loop first 2 column in the row
                ExcelRange cell = worksheet.Cells[i, 1];
                if (cell.Value == null)
                    continue;
                string key = cell.GetValue<string>().ToLower();

                cell = worksheet.Cells[i, 2];
                if (cell.Value == null)
                    continue;

                if (key == "dumpFolder".ToLower())
                    fetchConfig.DumpFolder = cell.GetValue<string>();
                else if (key == "dummyFetch".ToLower())
                    fetchConfig.IsDummyFetch = cell.GetValue<string>().ToLower() == "true" ? true : false;
                else if (key == "rowsPerChunk".ToLower())
                    fetchConfig.RowsPerChunk = worksheet.Cells[i, 2].GetValue<int>();
                else if (key == "absoluteStartTime".ToLower())
                    fetchConfig.StartTime.AbsoluteTime = worksheet.Cells[i, 2].GetValue<DateTime>();
                else if (key == "varStartYears".ToLower())
                    fetchConfig.StartTime.YearsOffset = worksheet.Cells[i, 2].GetValue<int>();
                else if (key == "varStartMonths".ToLower())
                    fetchConfig.StartTime.MonthsOffset = worksheet.Cells[i, 2].GetValue<int>();
                else if (key == "varStartDays".ToLower())
                    fetchConfig.StartTime.DaysOffset = worksheet.Cells[i, 2].GetValue<int>();
                else if (key == "varStartHours".ToLower())
                    fetchConfig.StartTime.HoursOffset = worksheet.Cells[i, 2].GetValue<int>();
                else if (key == "varStartMinutes".ToLower())
                    fetchConfig.StartTime.MinutesOffset = worksheet.Cells[i, 2].GetValue<int>();
                else if (key == "varStartSeconds".ToLower())
                    fetchConfig.StartTime.SecondsOffset = worksheet.Cells[i, 2].GetValue<int>();
                else if (key == "absoluteEndTime".ToLower())
                    fetchConfig.EndTime.AbsoluteTime = worksheet.Cells[i, 2].GetValue<DateTime>();
                else if (key == "varEndYears".ToLower())
                    fetchConfig.EndTime.YearsOffset = worksheet.Cells[i, 2].GetValue<int>();
                else if (key == "varEndMonths".ToLower())
                    fetchConfig.EndTime.MonthsOffset = worksheet.Cells[i, 2].GetValue<int>();
                else if (key == "varEndDays".ToLower())
                    fetchConfig.EndTime.DaysOffset = worksheet.Cells[i, 2].GetValue<int>();
                else if (key == "varEndHours".ToLower())
                    fetchConfig.EndTime.HoursOffset = worksheet.Cells[i, 2].GetValue<int>();
                else if (key == "varEndMinutes".ToLower())
                    fetchConfig.EndTime.MinutesOffset = worksheet.Cells[i, 2].GetValue<int>();
                else if (key == "varEndSeconds".ToLower())
                    fetchConfig.EndTime.SecondsOffset = worksheet.Cells[i, 2].GetValue<int>();
                else if (key == "fetchWindowDays".ToLower())
                    fetchConfig.FetchWindow += TimeSpan.FromDays(worksheet.Cells[i, 2].GetValue<int>());
                else if (key == "fetchWindowHours".ToLower())
                    fetchConfig.FetchWindow += TimeSpan.FromHours(worksheet.Cells[i, 2].GetValue<int>());
                else if (key == "fetchWindowMinutes".ToLower())
                    fetchConfig.FetchWindow += TimeSpan.FromMinutes(worksheet.Cells[i, 2].GetValue<int>());
                else if (key == "fetchWindowSeconds".ToLower())
                    fetchConfig.FetchWindow += TimeSpan.FromSeconds(worksheet.Cells[i, 2].GetValue<int>());
                else if (key == "FetchStrategy".ToLower())
                {
                    string strategy = worksheet.Cells[i, 2].GetValue<string>();
                    fetchConfig.FetchStrategy = new List<string>() { "snap", "average", "max", "min" }.Exists(e => e == strategy.ToLower()) ? strategy : "snap";
                }
                else if (key == "FetchPeriodicitySecs".ToLower())
                    fetchConfig.FetchPeriodicitySecs = worksheet.Cells[i, 2].GetValue<int>();
                else if (key == "FilenamePrefix".ToLower())
                    fetchConfig.FilenamePrefix = worksheet.Cells[i, 2].GetValue<string>();
            }
            // ensure fetch window is not zero
            if (fetchConfig.FetchWindow.TotalSeconds == 0)
            {
                fetchConfig.FetchWindow = new TimeSpan(1, 1, 1, 1);
            }
            return fetchConfig;
        }
    }
}
