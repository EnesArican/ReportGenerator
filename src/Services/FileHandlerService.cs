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
        
        public string FindCsvFile()
        {
            string downloadsPath = KnownFolders.Downloads.DefaultPath;
            Console.WriteLine("Downloads folder path: " + downloadsPath);

            var directory = new DirectoryInfo(downloadsPath);
            var myFile = directory.GetFiles(SystemConstant.CsvFileSearchPattern).OrderByDescending(f => f.LastWriteTime).First();
            Console.WriteLine("file you are looking for: " + myFile);

            return myFile.FullName;
        }


        public void OpenXLFile() 
        {
            Application xlApp = new Application();

            var xlWorkBook = xlApp.Workbooks.Add(Type.Missing);

            xlWorkBook.SaveAs(@"C:\Temp\test.xlsx"); ;
            xlWorkBook.Close();
            xlApp.Quit();


        }
    }
}
