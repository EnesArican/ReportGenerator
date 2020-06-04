using System.Drawing;
using ReportGenerator.Interfaces;
using Microsoft.Office.Interop.Excel;
using C = ReportGenerator.Constants.ColourScheme;
using SC = ReportGenerator.Constants.SystemConstant;
using ReportGenerator.Enums;

namespace ReportGenerator.Services
{
    public class RangeEditorService : IRangeEditorService
    {
        public void AddBorders(Range range)
        {
            var borders = range.Borders;
            borders.Color = C.Brown;
            borders.LineStyle = XlLineStyle.xlContinuous;
            borders.Weight = XlBorderWeight.xlThin;

            //borders[XlBordersIndex.xlEdgeLeft].Weight = XlBorderWeight.xlThick;
            //borders[XlBordersIndex.xlEdgeRight].Weight = XlBorderWeight.xlThick;
        }

        public void FormatRange(Range range, Color interiorColour, Color fontColour, int fontSize, bool isBold = false)
        {
            range.Interior.Color = interiorColour;
            range.Font.Color = fontColour;
            range.Font.Bold = isBold;
            range.Font.Size = fontSize;
        }

        public Range GetRange(Worksheet sheet, int strtRow, int strtCol, int endRow, int endCol)
        {
            var strtCell = sheet.Cells[strtRow, strtCol];
            var endCell = sheet.Cells[endRow, endCol];
            return sheet.Range[strtCell, endCell];
        }

        public void AddAlternateRowColours(Range range)
        {
            var formatCond = (FormatCondition)range.FormatConditions.Add(XlFormatConditionType.xlExpression, Formula1: "=MOD(ROW(),2)=0");
            formatCond.Interior.Color = C.LightGrey;
        }

        public void AddFormatConditions(Range attendanceRange)
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
    }
}
