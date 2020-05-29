using System;
using System.Linq;
using ReportGenerator.Models;
using SC = ReportGenerator.SystemConstant;
using ReportGenerator.Interfaces;
using System.Collections.Generic;
using Excel = Microsoft.Office.Interop.Excel;

using System.Globalization;

namespace ReportGenerator.Services
{
    public class DataWriterService : IDataWriterService
    {
        private readonly IFormatterService _formatterService;

        public DataWriterService(IFormatterService formatterService) => _formatterService = formatterService;
        public void WriteData(IList<PersonData> personDataList, Excel.Workbook workbook)
        {
            //var test = personDataList.SelectMany(p => p.AttendanceRecord.Select(a => new { name = p.FirstName, surname = p.LastName, date = a.Key, attedance = a.Value}));

            var dateGroupsPerMonth = personDataList.SelectMany(
                p => p.AttendanceRecord.Keys).Distinct().GroupBy(
                    g => new DateGroup { Year = g.Year, Month = g.Month });

            GenerateWorkbook(workbook, dateGroupsPerMonth, personDataList);

            var sheet = (Excel.Worksheet)workbook.Sheets[1];
            sheet.Activate();

        }


        private void GenerateWorkbook(Excel.Workbook workbook, IEnumerable<IGrouping<DateGroup, DateTime>> dateGroupsPerMonth, IList<PersonData> personDataList) 
        {
            var sheet = (Excel.Worksheet)workbook.ActiveSheet;

            CultureInfo cultureInfoTR = new CultureInfo("tr-TR");
            CultureInfo cultureInfoGB = new CultureInfo("en-GB");

            int y = 0, sheetCount = 1;
            foreach (var monthDateGroup in dateGroupsPerMonth.OrderByDescending(d => new DateTime(d.Key.Year, d.Key.Month, 1)))
            {
                sheet = (Excel.Worksheet)workbook.Sheets[sheetCount];
                sheet.Activate();

                y = SC.HeaderRow + 1;
                foreach (var personData in personDataList.OrderBy(p => p.FirstName))
                {
                    sheet.Cells[y, 1] = personData.FirstName;
                    sheet.Cells[y, 2] = personData.LastName;
                    sheet.Cells[y, 3] = personData.PhoneNumber;

                    CultureInfo.CurrentCulture = cultureInfoTR;

                    sheet.Cells[1, SC.AtStrtCol] = monthDateGroup.Select(d => d.ToString("MMMM")).First();

                    AddAttendanceAndDates(sheet, monthDateGroup, personData, y);
                    y++;
                    CultureInfo.CurrentCulture = cultureInfoGB;
                }
                sheetCount++;
                workbook.Sheets.Add(After: sheet);

                _formatterService.FormatWorksheet(sheet, monthDateGroup.Count());

            }

            // last sheet is empty so remove it
            sheet = (Excel.Worksheet)workbook.Sheets[sheetCount];
            sheet.Delete();


        }


        private void AddAttendanceAndDates(Excel.Worksheet sheet, IGrouping<DateGroup, DateTime> monthDateGroup, PersonData personData, int y)
        {
            var x = SC.AtStrtCol;
            foreach (var date in monthDateGroup.OrderBy(d => d.Day))
            {
                sheet.Cells[SC.HeaderRow, x] = date.ToString("dddd");
                sheet.Cells[SC.DateRow, x] = date;

                if (personData.AttendanceRecord.TryGetValue(date, out var value))
                    sheet.Cells[y, x] = value.ToString();
                x++;
            }
        }
    }
}
