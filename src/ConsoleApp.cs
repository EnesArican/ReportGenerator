using ReportGenerator.Interfaces;
using System;
using System.Collections.Generic;
using System.Text;

namespace ReportGenerator
{
    public class ConsoleApp
    {
        private readonly IFileFinderService _fileFinder;

        public ConsoleApp(IFileFinderService fileFinder) 
        {
            _fileFinder = fileFinder;
        }

        public void Run() 
        {
            _fileFinder.FindFile();
            Console.WriteLine("Hello World");
        }
    }
}
