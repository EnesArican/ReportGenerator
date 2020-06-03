using System;
using System.IO;
using System.Linq;
using System.Data;
using ReportGenerator.Models;
using ReportGenerator.Extensions;
using ReportGenerator.Interfaces;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using SC = ReportGenerator.Constants.SystemConstant;

namespace ReportGenerator.Services
{
    public class DataReaderService : IDataReaderService
    {
        public readonly Regex personDetailRegex = new Regex(@"\w.*,\w.*,\w.*,.*,,$");

        public IList<PersonData> GetPersonData(string filePath)
        {
            DateTime date;
            IList<PersonData> personDataList = new List<PersonData>();
            
            var lines = File.ReadAllLines(filePath).Where(l => !string.IsNullOrEmpty(l) && !l.Contains(SC.CsvHeaderRow)).ToList();

            var dateIndexes = lines.Select((value, index) => new { value, index })
                                   .Where(l => l.value.Contains(SC.CsvDateRow))
                                   .Select(l => l.index).ToList();

            foreach (var index in dateIndexes)
            {
                if (dateIndexes.LastOrDefault().Equals(index)) continue;

                var dateLine = lines[index];
                date = GetDate(dateLine);

                GetPersonDetails(personDataList, lines, date, index, dateIndexes);
            }
            return personDataList;
        }


        private void GetPersonDetails(IList<PersonData> PersonDataList, IList<string> lines, DateTime date, int i, List<int> dateIndexes) 
        {
            var nextIndex = dateIndexes.FindIndex(d => d == i) + 1;
            var nextDateIndex = dateIndexes[nextIndex];

            var classFound = false;
            while (i < nextDateIndex)
            {
                var line = lines[i];

                if (!personDetailRegex.IsMatch(lines[i])) classFound = false;

                if (classFound) PersonDataList.AddRecord(date, lines[i]);

                if (line.Contains(SC.CsvClassName)) classFound = true;

                i++;
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
