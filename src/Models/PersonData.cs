using ReportGenerator.Enums;
using System;
using System.Collections.Generic;
using System.Text;

namespace ReportGenerator.Models
{
    public class PersonData
    {
        public string FirstName { get; set; }
        public string LastName { get; set; }
        public string PhoneNumber { get; set; }

        public IDictionary<DateTime, Attendance> AttendanceRecord { get; set; }
    }
}
