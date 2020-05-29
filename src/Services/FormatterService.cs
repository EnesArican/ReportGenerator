using ReportGenerator.Interfaces;
using SC = ReportGenerator.SystemConstant;
using System.Drawing;
using ReportGenerator.Enums;
using Microsoft.Office.Interop.Excel;

namespace ReportGenerator.Services
{
    public class FormatterService : IFormatterService
    {
        public void FormatWorksheet(Worksheet sheet, int AtColsCount)
        {

            AddColumHeaders(sheet);

            

            Range mobileColumn = (Range)sheet.Columns[SC.MobileCol];
            mobileColumn.NumberFormat = "#### ### #### ###";

            Range dateRowRange = (Range)sheet.Rows[SC.DateRow];
            dateRowRange.NumberFormat = "dd-MM-yyyy";


            var atEndCol = SC.AtStrtCol + AtColsCount - 1;

            AddFormatConditions(sheet, atEndCol);

            MergeMonthHeader(sheet, atEndCol);

            sheet.UsedRange.Cells.HorizontalAlignment = XlHAlign.xlHAlignCenter;
            sheet.UsedRange.Cells.VerticalAlignment = XlVAlign.xlVAlignCenter;



            var titleRow = (Range)sheet.UsedRange.Rows[SC.DateRow];
            titleRow.Interior.Color = Color.FromArgb(255, 242, 204);
            titleRow.Font.Color = Color.FromArgb(0, 32, 96);

            titleRow = (Range)sheet.UsedRange.Rows[SC.HeaderRow];
            titleRow.Interior.Color = Color.FromArgb(255, 242, 204);
            titleRow.Font.Color = Color.FromArgb(0, 32, 96);

        }


        private void AddColumHeaders(Worksheet sheet) 
        {
            sheet.Cells[1, 1] = "Yoklama Ayi";
            sheet.Cells[2, 1] = "Yoklama Tarihleri";

            sheet.Cells[SC.HeaderRow, 1] = "SN.";
            sheet.Cells[SC.HeaderRow, 2] = "Adi";
            sheet.Cells[SC.HeaderRow, 3] = "Soyadi";
            sheet.Cells[SC.HeaderRow, 4] = "Telefon";
        }


        private void AddFormatConditions(Worksheet sheet, int atEndCol) 
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

        private FormatCondition AddFormatCondition(Range range, string text) => 
            (FormatCondition)range.FormatConditions.Add(
                XlFormatConditionType.xlCellValue, XlFormatConditionOperator.xlEqual, text
            );

        private void MergeMonthHeader(Worksheet sheet, int atEndCol) 
        {



            var strtCell = sheet.Cells[1, SC.AtStrtCol];
            var endCell = sheet.Cells[1, atEndCol];
            var monthHeader = sheet.Range[strtCell, endCell];
            monthHeader.Merge();
        }

    }
}
