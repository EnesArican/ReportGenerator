using Excel = Microsoft.Office.Interop.Excel;

namespace ReportGenerator.Interfaces
{
    public interface IFormatterService
    {
        void FormatWorksheet(Excel.Worksheet sheet, int AtColsCount);
    }
}
