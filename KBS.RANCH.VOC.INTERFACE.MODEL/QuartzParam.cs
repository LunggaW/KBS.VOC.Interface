using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace KBS.RANCH.VOCOLLECT.INTERFACE.MODEL
{
    public class QuartzParam
    {
        public int ID { get; set; }
        public string ShortDesc { get; set; }
        public string LongDesc { get; set; }
        public int IntervalSecond { get; set; }
        public int IntervalMinute { get; set; }
        public int IntervalHour { get; set; }
        public int IntervalDay { get; set; }
        public int StartHour { get; set; }
        public int StartMinute { get; set; }
        public bool IsDaily { get; set; }
        public bool IsMonthly { get; set; }
        public bool IsWeekly { get; set; }
        public string DayOfWeek { get; set; }
        public int DayOfMonth { get; set; }

    }
}
