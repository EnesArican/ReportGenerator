using Microsoft.Office.Interop.Excel;
using ReportGenerator.Interfaces;
using System;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices.ComTypes;

namespace ReportGenerator.Services
{
    public class FileReaderService : IFileReaderService
    {
        private readonly IFileFinderService _fileFinder;
        public FileReaderService(IFileFinderService fileFinder) => _fileFinder = fileFinder; 
        public void GetPersonData()
        {
            var filePath = _fileFinder.FindFile();

            var lines = File.ReadAllLines(filePath);

            DateTime date;
            foreach (var line in lines) 
            {
                if (line.Contains("Date:")) 
                {
                    var end = line.IndexOf(",") - 6;
                    var dateString = line.Substring(6, end);
                    date = DateTime.Parse(dateString);
                    Console.WriteLine(dateString);
                    Console.WriteLine(date);
                }




            }
        }



        

    }
}
