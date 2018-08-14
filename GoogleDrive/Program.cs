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
                if (args.Length > 1)
                {
                    PrintUsageAndExit(1);
                }
                else if (args[0] != "/RefreshList" && (!args[0].StartsWith("/Id=") || args[0].Length<5) && (!args[0].StartsWith("/StartFrom=") || args[0].Length < 12))
                {
                    PrintUsageAndExit(2);
                }
            }
            
            #endregion

            LogOutput("Started...");

            var drive = new Drive();

            #region Parse args

            string presentationId = null;
            int startFromIndex = 0;
            
            var refreshCache = false;
            if (args.Length > 0) {
                if (args[0] == "/?")
                {
                    PrintUsageAndExit(0);
                }
                if (args[0].StartsWith("/Id="))
                {
                    presentationId = args[0].Split('=')[1];
                }
                else if (args[0] == "/RefreshList")
                {
                    refreshCache = true;
                }
                else
                {
                    if (!Int32.TryParse(args[0].Split('=')[1], out startFromIndex))
                    {
                        PrintUsageAndExit(3);
                    }
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

                var rootFolderId = ConfigurationManager.AppSettings["rootFolderId"];
                if (refreshCache || drive.Presentations.Count == 0)
                {
                    LogOutput("Building presentations list...");
                    drive.ClearPresentationsList();
                    drive.BuildPresentationsList(rootFolderId, true);
                    drive.SavePresentationsList();
                    LogOutput("Finished building presentations list...");
                }
                LogOutput(string.Format("Processing {0} presentations, root folder: {1}", drive.Presentations.Count, drive.GetFolderName(rootFolderId)));

                if (startFromIndex > drive.Presentations.Count-1)
                {
                    PrintUsageAndExit(4);
                }
                if (startFromIndex > 0)
                {
                    LogOutput(string.Format("Skipping {0} presentations...", startFromIndex));
                }
                for (var i=startFromIndex; i<drive.Presentations.Count; i++)
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

        static void PrintUsageAndExit(int exitCode)
        {
            Console.WriteLine("GoogleDrive [/?] [/Id=<PresentationId>] [/RefreshList] [/StartFrom=<Index>]");
            Console.WriteLine("Only one of the parameters can be specified at a time:");
            Console.WriteLine("/Id          Process only this presentation");
            Console.WriteLine("/RefreshList Forces refresh of the local cache");
            Console.WriteLine("/StartFrom   Skips succeeded presentations");
            Console.WriteLine("/?           Prints this help");
            Environment.Exit(exitCode);
        }
    }
}