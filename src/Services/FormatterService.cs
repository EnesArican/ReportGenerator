using ReportGenerator.Interfaces;
using SC = ReportGenerator.SystemConstant;
using Excel = Microsoft.Office.Interop.Excel;
using System.Drawing;
using ReportGenerator.Enums;

namespace ReportGenerator.Services
{
    public class FormatterService : IFormatterService
    {
        public void FormatWorksheet(Excel.Worksheet sheet, int AtColsCount)
        {
            sheet.Cells[SC.HeaderRow, 1] = "Adi";
            sheet.Cells[SC.HeaderRow, 2] = "Soyadi";
            sheet.Cells[SC.HeaderRow, 3] = "Telefon";

            Excel.Range mobileColumn = (Excel.Range)sheet.Columns[3];
            mobileColumn.NumberFormat = "#### ### #### ###";

            Excel.Range dateRowRange = (Excel.Range)sheet.Rows[SC.DateRow];
            dateRowRange.NumberFormat = "dd-MM-yyyy";


            var atEndCol = SC.AtStrtCol + AtColsCount - 1;

            AddFormatConditions(sheet, atEndCol);

            MergeMonthHeader(sheet, atEndCol);


           
        }


        private void AddFormatConditions(Excel.Worksheet sheet, int atEndCol) 
        {
            var atStrtRow = SC.HeaderRow + 1;
            var atEndRow = sheet.UsedRange.Rows.Count;
            var strtCell = sheet.Cells[atStrtRow, SC.AtStrtCol];
            var endCell = sheet.Cells[atEndRow, atEndCol];
            var AttendanceRange = sheet.Range[strtCell, endCell];

            var formatCond = AddFormatCondition(AttendanceRange, Attendance.VAR.ToString());
            formatCond.Interior.Color = Color.FromArgb(198, 239, 206);
            formatCond.Font.Color = Color.FromArgb(0, 97, 0);

            formatCond = AddFormatCondition(AttendanceRange, Attendance.YOK.ToString());
            formatCond.Interior.Color = Color.FromArgb(255, 199, 206);
            formatCond.Font.Color = Color.FromArgb(156, 0, 6);

            formatCond = AddFormatCondition(AttendanceRange, Attendance.İZİNLİ.ToString());
            formatCond.Interior.Color = Color.FromArgb(255, 235, 156);
            formatCond.Font.Color = Color.FromArgb(156, 101, 0);

            formatCond = AddFormatCondition(AttendanceRange, Attendance.HASTA.ToString());
            formatCond.Interior.Color = Color.FromArgb(230, 184, 183);
            formatCond.Font.Color = Color.FromArgb(0, 32, 96);

            
        }

        private Excel.FormatCondition AddFormatCondition(Excel.Range range, string text) => 
            (Excel.FormatCondition)range.FormatConditions.Add(
                Excel.XlFormatConditionType.xlCellValue, Excel.XlFormatConditionOperator.xlEqual, text
            );

        private void MergeMonthHeader(Excel.Worksheet sheet, int atEndCol) 
        {
            var strtCell = sheet.Cells[1, SC.AtStrtCol];
            var endCell = sheet.Cells[1, atEndCol];
            var monthHeader = sheet.Range[strtCell, endCell];
            monthHeader.Merge();
        }

    }
}
