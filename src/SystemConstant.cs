using System;
using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;

namespace ReportGenerator
{
    public static class SystemConstant
    {
        public const string CsvFileSearchPattern = "daily_report_*.csv";

        public const string ClassName = "Fatih,,";
        public const string HeaderRow = "Last Name,First Name,Attendance,Phone,Notes,";
        public const string dateRow = "Date:";
    }
}
