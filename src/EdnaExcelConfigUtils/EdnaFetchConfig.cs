using System;
using System.Collections.Generic;

namespace EdnaExcelConfigUtils
{
    public class EdnaFetchConfig
    {
        // folder location at which chunks need to be dumped
        public string DumpFolder { get; set; }
        public List<EdnaPnt> Pnts { get; set; } = new List<EdnaPnt>();
        public VariableTime StartTime { get; set; } = new VariableTime();
        public VariableTime EndTime { get; set; } = new VariableTime();
        public string FetchStrategy { get; set; } = "snap";
        // if dummy data to be fetched for testing purposes
        public bool IsDummyFetch { get; set; }
        public int RowsPerChunk { get; set; }
        public int FetchPeriodicitySecs { get; set; } = 60;
        public TimeSpan FetchWindow { get; set; } = new TimeSpan(1, 0, 0, 0, 0);
        public string FilenamePrefix { get; set; } = "ednaDump";
    }
}
