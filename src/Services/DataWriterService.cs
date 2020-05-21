using Excel = Microsoft.Office.Interop.Excel;
using ReportGenerator.Interfaces;
using ReportGenerator.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ReportGenerator.Services
{
    public class DataWriterService : IDataWriterService
    {
        public void WriteData(IList<PersonData> personDataList, Excel.Workbook workbook)
        {

            var sheet = (Excel.Worksheet)workbook.ActiveSheet;

            var i = 1;
            foreach (var personData in personDataList.OrderBy(p => p.FirstName)) 
            {
                sheet.Cells[i, 1] = personData.FirstName;
                sheet.Cells[i, 2] = personData.LastName;
                sheet.Cells[i, 3] = personData.PhoneNumber;
                i++;
            }
        }
    }
}
