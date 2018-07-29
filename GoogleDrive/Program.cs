using System;
using System.Configuration;

namespace GoogleDrive
{
    static class Program
    {
        static void Main(string[] args)
        {
            var drive = new Drive();
            LogOutput("Started...");
            drive.ProcessPresentation(ConfigurationManager.AppSettings["SpecificPresentationId"]);
            LogOutput("Finished...(hit any key)");
            Console.ReadLine();
        }

        static void LogOutput(string line)
        {
            Console.WriteLine(string.Format("{0}: {1}", DateTime.Now, line));
        }
    }
}