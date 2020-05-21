using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;

namespace ReportGenerator.Interfaces
{
    public interface IFileHandlerService
    {
        string FindCsvFile();

        Excel.Workbook GetWorkbook();
        void SaveAndCloseWorkbook(Excel.Workbook workbook);
    }
}
