using ReportGenerator.Interfaces;
using Syroot.Windows.IO;
using System;
using System.IO;
using System.Linq;

namespace ReportGenerator.Services
{
    public class FileFinderService : IFileFinderService
    {
        private const string FileSearchPattern = "daily_report_*.csv";
        public string FindFile()
        {
            string downloadsPath = KnownFolders.Downloads.DefaultPath;
            Console.WriteLine("Downloads folder path: " + downloadsPath);
            Console.ReadLine();

            Console.WriteLine("looking for file");
            var directory = new DirectoryInfo(downloadsPath);
            var myFile = directory.GetFiles(FileSearchPattern).OrderByDescending(f => f.LastWriteTime).First();
            Console.WriteLine("file you are looking for: " + myFile);
            Console.ReadLine();

            return downloadsPath;
        }
    }
}
