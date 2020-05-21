using ReportGenerator.Interfaces;
using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.Office.Core;
using Excel = Microsoft.Office.Interop.Excel;

namespace ReportGenerator
{
    public class ConsoleApp
    {
        private readonly IDataReaderService _fileReader;
        private readonly IFileHandlerService _fileHandler;

        public ConsoleApp(IDataReaderService fileReader,
                          IFileHandlerService fileHandler) 
        {
            _fileReader = fileReader;
            _fileHandler = fileHandler;
        }

        public void Run() 
        {

            var filePath = _fileHandler.FindCsvFile();
            var personDataList =  _fileReader.GetPersonData(filePath);

            var personTest = personDataList[1];
            var test = personTest.AttendanceRecord.Select(kvp => kvp.Key + ": " + kvp.Value.ToString());
            Console.WriteLine(string.Join(Environment.NewLine, test));

            Excel.Application excel = new Excel.Application();

            var xlWorkBook = excel.Workbooks.Add(Type.Missing);

            xlWorkBook.SaveAs(@"C:\Temp\test.xlsx"); ;
            xlWorkBook.Close();
            excel.Quit();

        }
    }
}
