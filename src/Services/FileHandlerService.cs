using Excel = Microsoft.Office.Interop.Excel;
using ReportGenerator.Interfaces;
using Syroot.Windows.IO;
using System;
using System.IO;
using System.Linq;
using Microsoft.Office.Interop.Excel;

namespace ReportGenerator.Services
{
    public class FileHandlerService : IFileHandlerService
    {
        private Excel.Application ExcelApp;
        
        public string FindCsvFile()
        {
            string downloadsPath = KnownFolders.Downloads.DefaultPath;
            Console.WriteLine("Downloads folder path: " + downloadsPath);

            var directory = new DirectoryInfo(downloadsPath);
            var myFile = directory.GetFiles(SystemConstant.CsvFileSearchPattern).OrderByDescending(f => f.LastWriteTime).First();
            Console.WriteLine("file you are looking for: " + myFile);

            return myFile.FullName;
        }

        public Excel.Workbook GetWorkbook() 
        {
            ExcelApp = new Excel.Application();
            return ExcelApp.Workbooks.Add(Type.Missing);
        }

        public void SaveAndCloseWorkbook(Workbook workbook)
        {
            workbook.SaveAs(@"C:\Temp\test.xlsx"); ;
            workbook.Close();
            ExcelApp.Quit();
        }
    }
}
