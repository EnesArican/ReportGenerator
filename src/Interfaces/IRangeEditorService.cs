using System.Drawing;
using Microsoft.Office.Interop.Excel;

namespace ReportGenerator.Interfaces
{
    public interface IRangeEditorService
    {
        Range GetRange(Worksheet sheet, int strtRow, int strtCol, int endRow, int endCol);

        void AddFormatConditions(Range attendanceRange);

        void AddAlternateRowColours(Range range);

        void FormatRange(Range range, Color interiorColour, Color fontColour, int fontSize, bool isBold = false);

        void AddBorders(Range range);
    }
}
