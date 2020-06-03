using System.Drawing;
using ReportGenerator.Enums;
using ReportGenerator.Interfaces;
using Microsoft.Office.Interop.Excel;
using C = ReportGenerator.Constants.ColourScheme;
using SC = ReportGenerator.Constants.SystemConstant;

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

            var attendanceRange = GetAttendanceRange(sheet, atEndCol);

            //FormatAttendanceRange();

            AddBorders(attendanceRange.Borders);
            attendanceRange.Borders[XlBordersIndex.xlInsideVertical].Weight = XlBorderWeight.xlMedium;
            AddFormatConditions(attendanceRange);

            MergeMonthHeader(sheet, atEndCol);

            var allHeaders = GetAllHeadersRange(sheet, atEndCol);
            AddBorders(allHeaders.Borders);

            var headerColsRange = GetHeaderColsRange(sheet);
            AddBorders(headerColsRange.Borders);

            sheet.UsedRange.Cells.HorizontalAlignment = XlHAlign.xlHAlignCenter;
            sheet.UsedRange.Cells.VerticalAlignment = XlVAlign.xlVAlignCenter;


            var numbersRange = GetNumbersRange(sheet);
            FormatRange(numbersRange, C.Cream, C.Red, 10);

            var titleRow = (Range)sheet.UsedRange.Rows[1];
            FormatRange(titleRow, C.Cream, C.Red, 13, true);

            var dateRow = (Range)sheet.UsedRange.Rows[SC.DateRow];
            FormatRange(dateRow, C.Cream, C.Navy, 9, true);
          
            var headerRow = (Range)sheet.UsedRange.Rows[SC.HeaderRow];
            FormatRange(headerRow, C.Cream, C.Navy, 9, true);        
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



        private void AddBorders(Borders borders) 
        {
            borders.Color = C.Brown;
            borders.LineStyle = XlLineStyle.xlContinuous;
            borders.Weight = XlBorderWeight.xlThin;

            //borders[XlBordersIndex.xlEdgeLeft].Weight = XlBorderWeight.xlThick;
            //borders[XlBordersIndex.xlEdgeRight].Weight = XlBorderWeight.xlThick;
        }


        private void FormatRange(Range range, Color interiorColour, Color fontColour, int fontSize, bool isBold = false) 
        {
            range.Interior.Color = interiorColour;
            range.Font.Color = fontColour;
            range.Font.Bold = isBold;
            range.Font.Size = fontSize;
        }

        private Range GetAttendanceRange(Worksheet sheet, int atEndCol) 
        {
            var atStrtRow = SC.HeaderRow + 1;
            var atEndRow = sheet.UsedRange.Rows.Count;
            var strtCell = sheet.Cells[atStrtRow, SC.AtStrtCol];
            var endCell = sheet.Cells[atEndRow, atEndCol];
            return sheet.Range[strtCell, endCell];
        }


        private Range GetNumbersRange(Worksheet sheet)
        {
            var atEndRow = sheet.UsedRange.Rows.Count;
            var strtCell = sheet.Cells[SC.HeaderRow+1, 1];
            var endCell = sheet.Cells[atEndRow, 1];
            return sheet.Range[strtCell, endCell];
        }

        private Range GetAllHeadersRange(Worksheet sheet, int atEndCol)
        {
            var strtCell = sheet.Cells[1, 1];
            var endCell = sheet.Cells[SC.HeaderRow, atEndCol];
            return sheet.Range[strtCell, endCell];
        }

        private Range GetHeaderColsRange(Worksheet sheet)
        {
            var atEndRow = sheet.UsedRange.Rows.Count;
            var strtCell = sheet.Cells[SC.HeaderRow + 1, 1];
            var endCell = sheet.Cells[atEndRow, SC.MobileCol];
            return sheet.Range[strtCell, endCell];
        }







        private void AddFormatConditions(Range attendanceRange) 
        {
            var formatCond = AddFormatCondition(attendanceRange, Attendance.VAR.ToString());
            formatCond.Interior.Color = Color.FromArgb(198, 239, 206);
            formatCond.Font.Color = Color.FromArgb(0, 97, 0);

            formatCond = AddFormatCondition(attendanceRange, Attendance.YOK.ToString());
            formatCond.Interior.Color = Color.FromArgb(255, 199, 206);
            formatCond.Font.Color = Color.FromArgb(156, 0, 6);

            formatCond = AddFormatCondition(attendanceRange, Attendance.İZİNLİ.ToString());
            formatCond.Interior.Color = Color.FromArgb(255, 235, 156);
            formatCond.Font.Color = Color.FromArgb(156, 101, 0);

            formatCond = AddFormatCondition(attendanceRange, Attendance.HASTA.ToString());
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

            strtCell = sheet.Cells[1, 1];
            endCell = sheet.Cells[1, SC.AtStrtCol - 1];
            var monthTitle = sheet.Range[strtCell, endCell];
            monthTitle.Merge();

            strtCell = sheet.Cells[2, 1];
            endCell = sheet.Cells[2, SC.AtStrtCol - 1];
            var datesTitle = sheet.Range[strtCell, endCell];
            datesTitle.Merge();

        }

    }
}
