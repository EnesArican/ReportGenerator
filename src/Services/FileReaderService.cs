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
        private readonly string ClassName = "Fatih,,";
        public FileReaderService(IFileFinderService fileFinder) => _fileFinder = fileFinder; 
        public void GetPersonData()
        {
            var filePath = _fileFinder.FindFile();

            var lines = File.ReadAllLines(filePath).Where(l => !string.IsNullOrEmpty(l));

            DateTime date;
            var dateFound = false;
            var classFound = false;
            foreach (var line in lines) 
            {
                Console.WriteLine(line);
                //if line does not contain 5 commas then set datefound to false and class found to false 

                if (line.Contains("Date:"))
                {
                    date = GetDate(line);
                    dateFound = true;
                }

                if (dateFound && line.Contains(ClassName)) classFound = true;


                if (dateFound && classFound) 
                { 
                    
                }




            }
        }


        private DateTime GetDate(string line) 
        {
            var end = line.IndexOf(",") - 6;
            var dateString = line.Substring(6, end);
            return DateTime.Parse(dateString);
        }


    }
}
