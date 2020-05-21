using ReportGenerator.Interfaces;
using System;
using Excel = Microsoft.Office.Interop.Excel;

namespace ReportGenerator
{
    public class ConsoleApp
    {
        private readonly IDataReaderService _fileReader;
        private readonly IFileHandlerService _fileHandler;
        private readonly IDataWriterService _fileWriter;


        public ConsoleApp(IDataReaderService fileReader,
                          IFileHandlerService fileHandler,
                          IDataWriterService fileWriter) 
        {
            _fileReader = fileReader;
            _fileHandler = fileHandler;
            _fileWriter = fileWriter;
        }

        public void Run() 
        {

            var filePath = _fileHandler.FindCsvFile();
            var personDataList =  _fileReader.GetPersonData(filePath);


            var workbook = _fileHandler.GetWorkbook();

            _fileWriter.WriteData(personDataList, workbook);

            
            _fileHandler.SaveAndCloseWorkbook(workbook);

          

        }
    }
}
