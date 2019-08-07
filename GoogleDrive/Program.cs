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
                else if (args[0] != "/RefreshList" && (!args[0].StartsWith("/RootFolderId=") || args[0].Length<15) && (!args[0].StartsWith("/PresentationId=") || args[0].Length < 17))
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
                if (args[0].StartsWith("/RootFolderId="))
                {
                    specificFolderId = args[0].Split('=')[1];
                }
                if (args[0].StartsWith("/PresentationId="))
                {
                    presentationId = args[0].Split('=')[1];
                }
                else if (args[0] == "/RefreshList")
                {
                    refreshCache = true;
                }
            }

            #endregion

            try
            {
                if (presentationId != null)
                {
                    #region Process specfic presentation

                    var cachePresentation = drive.Cache.GetPresentation(presentationId, drive.Cache.Folders);
                    if (cachePresentation != null)
                    {
                        drive.ProcessPresentation(cachePresentation);
                    }
                    else
                    {
                        LogOutputWithNewLine(string.Format("Presentation {0} not found in cache", presentationId));
                    }

                    #endregion
                }
                else
                {
                    #region Process folder presentations

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
                        rootFolder = drive.Cache.GetFolder(specificFolderId, drive.Cache.Folders);
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
                        LogOutput(string.Format("Processing {0} presentations", drive.Cache.TotalPresentations));
                        foreach (var folderKey in drive.Cache.Folders.Keys)
                        {
                            LogOutput(string.Format("Processing {0} presentations in folder: {1}", drive.Cache.Folders[folderKey].TotalPresentations, drive.Cache.Folders[folderKey].FolderName));
                            drive.ProcessFolderPresentations(drive.Cache.Folders[folderKey]);
                        }
                    }


                    #endregion
                }
            }
            catch(Exception)
            {

            }
            finally
            {
                LogSlideErrors(drive);
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
            Console.WriteLine("/RootFolderId    Process only this Root Folder and its subfolders");
            Console.WriteLine("/RefreshList     Forces refresh of the local cache");
            Console.WriteLine("/PresentationId  Skips succeeded presentations");
            Console.WriteLine("/?               Prints this help");
            Environment.Exit(exitCode);
        }

        static void LogSlideErrors(Drive drive)
        {
            if (drive.SlideErrors.Count > 0)
            {
                LogOutput(string.Format("{0} errors found", drive.SlideErrors.Count));
                foreach (var slideError in drive.SlideErrors)
                {
                    LogOutput(string.Format("Presentation: {0} {1}, Slide: {2}, Error: {3}", slideError.PresentationId, slideError.PresentationName, slideError.SlideId, slideError.Error));
                }
            }
        }
    }
}