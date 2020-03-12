using System;

namespace EdnaExcelConfigUtils
{
    public class VariableTime
    {
        public int? YearsOffset { get; set; }
        public int? MonthsOffset { get; set; }
        public int? DaysOffset { get; set; }
        public int? HoursOffset { get; set; }
        public int? MinutesOffset { get; set; }
        public int? SecondsOffset { get; set; }
        public DateTime AbsoluteTime { get; set; } = DateTime.Now;

        public VariableTime()
        {

        }

        public VariableTime(int? yearsOffset, int? monthsOffset, int? daysOffset, int? hoursOffset, int? minutesOffset, int? secondsOffset, DateTime absoluteTime)
        {
            YearsOffset = yearsOffset;
            MonthsOffset = monthsOffset;
            DaysOffset = daysOffset;
            HoursOffset = hoursOffset;
            MinutesOffset = minutesOffset;
            SecondsOffset = secondsOffset;
            AbsoluteTime = absoluteTime;
        }

        internal VariableTime Clone()
        {
            VariableTime variableTime = new VariableTime
            {
                YearsOffset = YearsOffset,
                MonthsOffset = MonthsOffset,
                DaysOffset = DaysOffset,
                HoursOffset = HoursOffset,
                MinutesOffset = MinutesOffset,
                SecondsOffset = SecondsOffset,
                AbsoluteTime = AbsoluteTime
            };
            return variableTime;
        }

        public DateTime GetTime()
        {
            DateTime absTime = AbsoluteTime;
            DateTime nowTime = DateTime.Now;

            // Make millisecond component as zero for the absolute time and now time
            absTime = absTime.AddMilliseconds(-1 * absTime.Millisecond);
            nowTime = nowTime.AddMilliseconds(-1 * nowTime.Millisecond);
            DateTime resultTime = nowTime;

            // Add offsets to current time as per the settings
            if (YearsOffset.HasValue)
            {
                resultTime = resultTime.AddYears(YearsOffset.Value);
            }
            if (MonthsOffset.HasValue)
            {
                resultTime = resultTime.AddMonths(MonthsOffset.Value);
            }
            if (DaysOffset.HasValue)
            {
                resultTime = resultTime.AddDays(DaysOffset.Value);
            }
            if (HoursOffset.HasValue)
            {
                resultTime = resultTime.AddHours(HoursOffset.Value);
            }
            if (MinutesOffset.HasValue)
            {
                resultTime = resultTime.AddMinutes(MinutesOffset.Value);
            }
            if (SecondsOffset.HasValue)
            {
                resultTime = resultTime.AddSeconds(SecondsOffset.Value);
            }

            // Set absolute time settings to the result time
            if (!YearsOffset.HasValue)
            {
                resultTime = resultTime.AddYears(absTime.Year - resultTime.Year);
            }
            if (!MonthsOffset.HasValue)
            {
                resultTime = resultTime.AddMonths(absTime.Month - resultTime.Month);
            }
            if (!DaysOffset.HasValue)
            {
                resultTime = resultTime.AddDays(absTime.Day - resultTime.Day);
            }
            if (!HoursOffset.HasValue)
            {
                resultTime = resultTime.AddHours(absTime.Hour - resultTime.Hour);
            }
            if (!MinutesOffset.HasValue)
            {
                resultTime = resultTime.AddMinutes(absTime.Minute - resultTime.Minute);
            }
            if (!SecondsOffset.HasValue)
            {
                resultTime = resultTime.AddSeconds(absTime.Second - resultTime.Second);
            }
            return resultTime;
        }
    }
}
