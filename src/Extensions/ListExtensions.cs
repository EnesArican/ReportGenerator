using Microsoft.Office.Interop.Excel;
using ReportGenerator.Enums;
using ReportGenerator.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ReportGenerator.Extensions
{
    public static class ListExtensions
    {
        public static void AddRecord(this IList<PersonData> personDataList, DateTime date, string line)
        {
            string[] details = line.Split(',');
            var firstName = details[1];
            var lastName = details[0];

            var person = personDataList.FirstOrDefault(p => p.FirstName == firstName && p.LastName == lastName);

            if (person == null) 
            {
                person = new PersonData
                {
                    FirstName = firstName,
                    LastName = lastName,
                    PhoneNumber = details[3],
                    AttendanceRecord = new Dictionary<DateTime, Attendance>()
                };

                personDataList.Add(person);
            };


            person.AttendanceRecord.Add(date, ConvertAttendance(details[2]));
        }

        private static Attendance ConvertAttendance(string value) =>
        value switch
        {
            "P"  => Attendance.VAR,
            "A"  => Attendance.YOK,
            "TU" => Attendance.IZINLI,
            "M"  => Attendance.HASTA,
            _    => throw new ArgumentException("value is not a valid attendance type")
        };
    }

   
}
