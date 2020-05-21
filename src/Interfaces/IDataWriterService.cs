using ReportGenerator.Models;
using System;
using System.Collections.Generic;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;

namespace ReportGenerator.Interfaces
{
    public interface IDataWriterService
    {
        void WriteData(IList<PersonData> personDataList, Excel.Workbook workbook);
    }
}
