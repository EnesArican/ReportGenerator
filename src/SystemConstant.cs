using System;
using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;

namespace ReportGenerator
{
    public static class SystemConstant
    {
        public const string CsvFileSearchPattern = "daily_report_*.csv";

        public const string CsvClassName = "Fatih,,";
        public const string CsvHeaderRow = "Last Name,First Name,Attendance,Phone,Notes,";
        public const string CsvDateRow = "Date:";


        public const int DateRow = 2;
        public const int HeaderRow = 3;

        public const int MobileCol = 4;
        public const int AtStrtCol = 5;

    }
}
