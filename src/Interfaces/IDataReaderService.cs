using ReportGenerator.Models;
using System;
using System.Collections.Generic;
using System.Text;

namespace ReportGenerator.Interfaces
{
    public interface IDataReaderService
    {
        IList<PersonData> GetPersonData(string filePath);
    }
}
