using Excel = Microsoft.Office.Interop.Excel;
using ReportGenerator.Interfaces;
using ReportGenerator.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ReportGenerator.Enums;
using System.Globalization;

namespace ReportGenerator.Services
{
    public class DataWriterService : IDataWriterService
    {
        public void WriteData(IList<PersonData> personDataList, Excel.Workbook workbook)
        {


            //var test = personDataList.SelectMany(p => p.AttendanceRecord.Select(a => new { name = p.FirstName, surname = p.LastName, date = a.Key, attedance = a.Value}));

            var dateGroupsPerMonth = personDataList.SelectMany(p => p.AttendanceRecord.Keys).Distinct().GroupBy(g => new DateGroup{ Year = g.Year, Month = g.Month });


            GenerateWorkbook(workbook, dateGroupsPerMonth, personDataList);

            var sheet = (Excel.Worksheet)workbook.ActiveSheet;
              
            CultureInfo cultureInfoTR = new CultureInfo("tr-TR");
            CultureInfo cultureInfoGB = new CultureInfo("en-GB");

            int y = 0, x = 0, sheetCount = 1;
            int headerRow = 3, dateRow = 2, atStrtCol = 4, atEndCol;

            foreach (var monthDateGroup in dateGroupsPerMonth.OrderByDescending(d => new DateTime(d.Key.Year, d.Key.Month, 1))) 
            {
                sheet = (Excel.Worksheet)workbook.Sheets[sheetCount];
                sheet.Activate();

                sheet.Cells[headerRow, 1] = "Adi";
                sheet.Cells[headerRow, 2] = "Soyadi";
                sheet.Cells[headerRow, 3] = "Telefon";

                y = headerRow + 1;
                foreach (var personData in personDataList.OrderBy(p => p.FirstName))
                {
                    sheet.Cells[y, 1] = personData.FirstName;
                    sheet.Cells[y, 2] = personData.LastName;
                    sheet.Cells[y, 3] = personData.PhoneNumber;
                    

                    CultureInfo.CurrentCulture = cultureInfoTR;
                    
                    //atEndCol = atStrtCol + dateGroup.Count();
                    //var strtCell = sheet.Cells[1, atStrtCol];
                    //var endCell = sheet.Cells[1, atEndCol];
                    //var monthHeader = sheet.Range[strtCell,endCell];
                    //monthHeader.Merge();

                    sheet.Cells[1,atStrtCol] = monthDateGroup.Select(d => d.ToString("MMMM")).First(); 


                    x = atStrtCol;
                    foreach (var date in monthDateGroup.OrderBy(d => d.Day))
                    {
                        sheet.Cells[headerRow, x] = date.ToString("dddd");
                        sheet.Cells[dateRow, x] = date;

                        if (personData.AttendanceRecord.TryGetValue(date, out var value)) 
                            sheet.Cells[y, x] = value.ToString();
                        x++;
                    }
                    y++;
                    CultureInfo.CurrentCulture = cultureInfoGB;
                }
                sheetCount++;
                workbook.Sheets.Add(After: sheet);

                Excel.Range mobileColumn = (Excel.Range)sheet.Columns[3];
                mobileColumn.NumberFormat = "#### ### #### ###";

                Excel.Range dateRowRange = (Excel.Range)sheet.Rows[dateRow];
                dateRowRange.NumberFormat = "dd-MM-yyyy";


            }
            sheet = (Excel.Worksheet)workbook.Sheets[1];
            sheet.Activate();

        }


        private void GenerateWorkbook(Excel.Workbook workbook, IEnumerable<IGrouping<DateGroup, DateTime>> dateGroupsPerMonth, IList<PersonData> personDataList) 
        {
            var sheet = (Excel.Worksheet)workbook.ActiveSheet;

            CultureInfo cultureInfoTR = new CultureInfo("tr-TR");
            CultureInfo cultureInfoGB = new CultureInfo("en-GB");

            int y = 0, x = 0, sheetCount = 1;
            int headerRow = 3, dateRow = 2, atStrtCol = 4, atEndCol;

            foreach (var monthDateGroup in dateGroupsPerMonth.OrderByDescending(d => new DateTime(d.Key.Year, d.Key.Month, 1)))
            {
                sheet = (Excel.Worksheet)workbook.Sheets[sheetCount];
                sheet.Activate();

                sheet.Cells[headerRow, 1] = "Adi";
                sheet.Cells[headerRow, 2] = "Soyadi";
                sheet.Cells[headerRow, 3] = "Telefon";

                y = headerRow + 1;
                foreach (var personData in personDataList.OrderBy(p => p.FirstName))
                {
                    sheet.Cells[y, 1] = personData.FirstName;
                    sheet.Cells[y, 2] = personData.LastName;
                    sheet.Cells[y, 3] = personData.PhoneNumber;


                    CultureInfo.CurrentCulture = cultureInfoTR;

                    //atEndCol = atStrtCol + dateGroup.Count();
                    //var strtCell = sheet.Cells[1, atStrtCol];
                    //var endCell = sheet.Cells[1, atEndCol];
                    //var monthHeader = sheet.Range[strtCell,endCell];
                    //monthHeader.Merge();

                    sheet.Cells[1, atStrtCol] = monthDateGroup.Select(d => d.ToString("MMMM")).First();


                    x = atStrtCol;
                    foreach (var date in monthDateGroup.OrderBy(d => d.Day))
                    {
                        sheet.Cells[headerRow, x] = date.ToString("dddd");
                        sheet.Cells[dateRow, x] = date;

                        if (personData.AttendanceRecord.TryGetValue(date, out var value))
                            sheet.Cells[y, x] = value.ToString();
                        x++;
                    }
                    y++;
                    CultureInfo.CurrentCulture = cultureInfoGB;
                }
                sheetCount++;
                workbook.Sheets.Add(After: sheet);

                Excel.Range mobileColumn = (Excel.Range)sheet.Columns[3];
                mobileColumn.NumberFormat = "#### ### #### ###";

                Excel.Range dateRowRange = (Excel.Range)sheet.Rows[dateRow];
                dateRowRange.NumberFormat = "dd-MM-yyyy";
            }

        }
    }
}
