using ReportGenerator.Interfaces;
using SC = ReportGenerator.SystemConstant;
using Excel = Microsoft.Office.Interop.Excel;

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
            var strtCell = sheet.Cells[1, SC.AtStrtCol];
            var endCell = sheet.Cells[1, atEndCol];
            var monthHeader = sheet.Range[strtCell, endCell];
            monthHeader.Merge();
        }
    }
}
