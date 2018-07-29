using System;
using System.Configuration;

namespace GoogleDrive
{
    static class Program
    {
        static void Main(string[] args)
        {
            #region Validate args

            if (args.Length > 0)
            {
                var invalidArgs = false;
                if (args.Length > 1)
                {
                    invalidArgs = true;
                }
                else if (args[0] != "/RefreshList" && (!args[0].StartsWith("/Id=") || args[0].Length<5))
                {
                    invalidArgs = true;
                }
                if (invalidArgs)
                {
                    Console.WriteLine("GoogleDrive [/Id=<PresentationId>] [/RefreshList]");
                    Console.WriteLine("Only one of the parameters can be specified at a time:");
                    Console.WriteLine("/Id              Process only this presentation");
                    Console.WriteLine("/RefreshList     Processes all presentations - forces refresh of the local cache of presentations list");
                    return;
                }
            }
            
            #endregion

            LogOutput("Started...");

            var drive = new Drive();

            #region Parse args

            string presentationId = null;
            
            var refreshCache = false;
            if (args.Length > 0) {
                if (args[0].StartsWith("/Id="))
                {
                    presentationId = args[0].Split('=')[1];
                }
                else
                {
                    refreshCache = true;
                }
            }

            #endregion

            if (presentationId != null)
            {
                #region Process specfic presentation

                drive.ProcessPresentation(presentationId);

                #endregion
            }
            else
            {
                #region Process all presentations

                if (refreshCache || drive.Presentations.Count == 0)
                {
                    LogOutput("Building presentations list...");
                    drive.BuildPresentationsList(ConfigurationManager.AppSettings["rootFolderId"]);
                    drive.SavePresentationsList();
                    LogOutput("Finished building presentations list...");
                }
                LogOutput(string.Format("Processing {0} presentations...", drive.Presentations.Count));
                for (var i=0; i<drive.Presentations.Count; i++)
                {
                    drive.ProcessPresentation(drive.Presentations[i]);
                    Console.Write(string.Format("\rProcessed {0} of {1} presentations...", i+1, drive.Presentations.Count));
                }
                #endregion
            }
            LogOutputWithNewLine("Finished...");
        }

        static void LogOutput(string line)
        {
            Console.WriteLine(string.Format("{0}: {1}", DateTime.Now, line));
        }
        static void LogOutputWithNewLine(string line)
        {
            Console.WriteLine(string.Format("\n{0}: {1}", DateTime.Now, line));
        }

    }
}