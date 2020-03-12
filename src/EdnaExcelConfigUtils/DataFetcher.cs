using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using InStep.eDNA.EzDNAApiNet;

namespace EdnaExcelConfigUtils
{
    public class DataFetcher
    {
        public EdnaFetchConfig FetchConfig { get; set; }
        private StreamWriter DumpFile;
        private int numDumpedRows;
        private int fileCount = 0;

        public DataFetcher(EdnaFetchConfig config) { FetchConfig = config; numDumpedRows = FetchConfig.RowsPerChunk; }

        public void FetchAndDumpData()
        {
            DirectoryInfo folderInfo = new DirectoryInfo(FetchConfig.DumpFolder);
            // check if dump folder exists
            if (!folderInfo.Exists)
            {
                Console.WriteLine($"{folderInfo.FullName} Directory doesnot exist to dump files");
                return;
            }
            if (FetchConfig.Pnts.Count == 0)
            {
                Console.WriteLine("No Edna points specified for dumping files");
                return;
            }
            DateTime startTime = FetchConfig.StartTime.GetTime();
            DateTime endTime = FetchConfig.EndTime.GetTime();
            if (startTime > endTime)
            {
                Console.WriteLine("Start time > End time, hence not fetching data");
                return;
            }
            // get fetch windows
            List<(DateTime, DateTime)> fetchWindows = GetFetchWindows();

            foreach (EdnaPnt pnt in FetchConfig.Pnts)
            {
                foreach (var window in fetchWindows)
                {
                    // get data
                    List<(DateTime, double)> data = FetchData(pnt.Id, window.Item1, window.Item2);
                    // dump to file
                    PushDataToDumpFile(data, pnt);
                }
            }
            if (DumpFile != null)
            {
                // closing and disposing the file
                DumpFile.Flush();
                DumpFile.Dispose();
            }
        }

        private List<(DateTime, DateTime)> GetFetchWindows()
        {
            List<(DateTime, DateTime)> fetchWindows = new List<(DateTime, DateTime)>();
            DateTime startTime = FetchConfig.StartTime.GetTime();
            DateTime endTime = FetchConfig.EndTime.GetTime();
            if (startTime > endTime)
            {
                return fetchWindows;
            }
            if (startTime == endTime)
            {
                fetchWindows.Add((startTime, startTime));
                return fetchWindows;
            }
            DateTime windowStartTime = startTime;
            DateTime windowEndTime;
            while (windowStartTime < endTime)
            {
                windowEndTime = windowStartTime + FetchConfig.FetchWindow;
                if (windowEndTime > endTime)
                {
                    windowEndTime = endTime;
                }
                fetchWindows.Add((windowStartTime, windowEndTime));
                windowStartTime = windowEndTime + TimeSpan.FromSeconds(FetchConfig.FetchPeriodicitySecs);
            }
            return fetchWindows;
        }

        private void GetNewFileForDump()
        {
            if (DumpFile != null)
            {
                // closing and disposing the file
                DumpFile.Flush();
                DumpFile.Dispose();
            }
            fileCount += 1;
            DumpFile = new StreamWriter($@"{(new DirectoryInfo(FetchConfig.DumpFolder)).FullName}\{FetchConfig.FilenamePrefix}_{DateTime.Now.ToString("dd_MM_yyyy_HH_mm_ss")}_{fileCount}.csv");
            numDumpedRows = 0;
        }

        private void PushDataToDumpFile(List<(DateTime, double)> data, EdnaPnt pnt)
        {
            int currStartRowIndex = 0;
            // decide number of rows to be dumped
            while (currStartRowIndex < data.Count)
            {
                // check if new file creation is required
                if (numDumpedRows >= FetchConfig.RowsPerChunk)
                {
                    // get new file for dumping
                    GetNewFileForDump();
                }
                int currEndRowIndex = currStartRowIndex + Math.Min(data.Count - currStartRowIndex - 1, FetchConfig.RowsPerChunk - numDumpedRows - 1);
                string dumpString = string.Join("\n", data.Skip(currStartRowIndex).Take(currEndRowIndex - currStartRowIndex + 1).Select(i => $"{pnt.Id},{i.Item1.ToString("dd_MM_yyyy_HH_mm_ss_fff")},{i.Item2}"));
                DumpFile.WriteLine(dumpString);
                //for (int rowInd = currStartRowIndex; rowInd <= currEndRowIndex; rowInd++)
                //{
                //    DumpFile.WriteLine($"{pnt.Id},{data[rowInd].Item1.ToString("dd_MM_yyyy_HH_mm_ss_fff")},{data[rowInd].Item2}");
                //}
                numDumpedRows += currEndRowIndex - currStartRowIndex + 1;
                currStartRowIndex = currEndRowIndex + 1;
            }
        }
        public List<(DateTime, double)> FetchData(string measId, DateTime startTime, DateTime endTime)
        {
            if (!FetchConfig.IsDummyFetch)
            {
                return FetchEdnaData(measId, startTime, endTime);
            }
            else
            {
                return FetchDummyData(measId, startTime, endTime);
            }
        }
        public List<(DateTime, double)> FetchEdnaData(string measId, DateTime startTime, DateTime endTime)
        {
            try
            {
                int nret = 0;
                uint s = 0;
                double dval = 0;
                DateTime timestamp = DateTime.Now;
                string status = "";
                TimeSpan period = TimeSpan.FromSeconds(FetchConfig.FetchPeriodicitySecs);
                // History request initiation
                if (FetchConfig.FetchStrategy == "raw")
                { nret = History.DnaGetHistRaw(measId, startTime, endTime, out s); }
                else if (FetchConfig.FetchStrategy == "snap")
                { nret = History.DnaGetHistSnap(measId, startTime, endTime, period, out s); }
                else if (FetchConfig.FetchStrategy == "average")
                { nret = History.DnaGetHistAvg(measId, startTime, endTime, period, out s); }
                else if (FetchConfig.FetchStrategy == "min")
                { nret = History.DnaGetHistMin(measId, startTime, endTime, period, out s); }
                else if (FetchConfig.FetchStrategy == "max")
                { nret = History.DnaGetHistMax(measId, startTime, endTime, period, out s); }

                // Get history values
                List<(DateTime, double)> historyResults = new List<(DateTime, double)>();
                while (nret == 0)
                {
                    nret = History.DnaGetNextHist(s, out dval, out timestamp, out status);
                    if (status != null)
                    {
                        historyResults.Add((timestamp, dval));
                    }
                }
                return historyResults;
            }
            catch (Exception e)
            {
                // Todo send this to console printing of the dashboard
                Console.WriteLine($"Error while fetching history data of point {measId}");
                Console.WriteLine($"The exception is {e}");
            }
            return new List<(DateTime, double)>();
        }

        public List<(DateTime, double)> FetchDummyData(string measId, DateTime startTime, DateTime endTime)
        {
            TimeSpan period = TimeSpan.FromSeconds(FetchConfig.FetchPeriodicitySecs);
            List<(DateTime, double)> historyResults = new List<(DateTime, double)>();
            Random random = new Random();
            for (DateTime targetTime = startTime; targetTime <= endTime; targetTime += period)
            {
                historyResults.Add((targetTime, random.Next(0, 100)));
            }

            return historyResults;
        }
    }
}
