using Microsoft.Office.Interop.Excel;
using ReportGenerator.Interfaces;
using System;
using System.Data;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices.ComTypes;

namespace ReportGenerator.Services
{
    public class FileReaderService : IFileReaderService
    {
        private readonly IFileFinderService _fileFinder;
        private readonly string ClassName = "Fatih,,";
        private readonly string HeaderRow = "Last Name,First Name,Attendance,Phone,Notes,";
        public FileReaderService(IFileFinderService fileFinder) => _fileFinder = fileFinder; 
        public void GetPersonData()
        {
            var filePath = _fileFinder.FindFile();

            var lines = File.ReadAllLines(filePath).Where(l => !string.IsNullOrEmpty(l) && !l.Contains(HeaderRow)).ToList();

            DateTime date;
            const string dateRow = "Date:";
            var dateFound = false;
            var classFound = false;

            var dateIndexes = lines.Select((value, index) => new { value, index })
                                   .Where(l => l.value.Contains(dateRow))
                                   .Select(l => l.index).ToList();

            foreach (var index in dateIndexes)
            {
                if (dateIndexes.LastOrDefault().Equals(index)) continue;


                var nextIndex = dateIndexes.FindIndex(d => d == index) + 1;
                var nextDateIndex = dateIndexes[nextIndex];
                Console.WriteLine($"{index}  {nextDateIndex}");


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
