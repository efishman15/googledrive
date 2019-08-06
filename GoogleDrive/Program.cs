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
            string specificFolderId = null;
            
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
                    specificFolderId = args[0].Split('=')[1];
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

                if (refreshCache || drive.Cache.Folders.Count == 0)
                {
                    LogOutput("Building presentations list...");
                    drive.ClearCache();
                    drive.BuildPresentationsList(rootFolderId, true, null);
                    drive.BuildFoldersPath(drive.Cache.Folders, string.Empty);
                    drive.SaveCache();

                    LogOutput("Finished building presentations list...");
                }

                CacheFolder rootFolder;
                if (specificFolderId != null)
                {
                    rootFolder = drive.Cache.GetFolder(specificFolderId, null);
                    if (rootFolder == null)
                    {
                        //Specified folder id not found in cache
                        PrintUsageAndExit(3);
                    }
                    LogOutput(string.Format("Processing {0} presentations, root folder: {1}", rootFolder.TotalPresentations, rootFolder.FolderName));
                    drive.ProcessFolderPresentations(rootFolder);
                }
                else
                {
                    foreach (var folderKey in drive.Cache.Folders.Keys)
                    {
                        LogOutput(string.Format("Processing {0} presentations, root folder: {1}", drive.Cache.Folders[folderKey].TotalPresentations, drive.Cache.Folders[folderKey].FolderName));
                        drive.ProcessFolderPresentations(drive.Cache.Folders[folderKey]);
                    }
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