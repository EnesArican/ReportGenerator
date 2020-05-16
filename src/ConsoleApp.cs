using ReportGenerator.Interfaces;
using System;
using System.Collections.Generic;
using System.Text;

namespace ReportGenerator
{
    public class ConsoleApp
    {
        private readonly IFileReaderService _fileReader;

        public ConsoleApp(IFileReaderService fileReader) 
        {
            _fileReader = fileReader;
        }

        public void Run() 
        {
            var test = 1;

            _fileReader.GetPersonData();
            Console.WriteLine("Hello World");
        }
    }
}
