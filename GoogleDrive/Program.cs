using System;
using System.Configuration;

namespace GoogleDrive
{
    internal static class Program
    {
        private enum ProcessingType
        {
            PRESENTATION,
            FOLDER,
            STUDENTS
        };
        private static Drive drive;
        private static int slidesProcessed;
        private static int slidesSkipped;
        private static ProcessingType processingType;

        private static void Main(string[] args)
        {
            #region Validate args

            if (args.Length > 0)
            {
                if (args.Length > 1)
                {
                    PrintUsageAndExit(1);
                }
                else if (
                            args[0] != "/RefreshList" && 
                            (!args[0].StartsWith("/RootFolderId=") || args[0].Length<15) && 
                            (!args[0].StartsWith("/PresentationId=") || args[0].Length < 17) &&
                            args[0] != "/Students"
                        )
                {
                    PrintUsageAndExit(2);
                }
            }
            
            #endregion

            LogOutputWithNewLine("Started...");

            drive = new Drive();
            drive.PresentationProcessed += Drive_PresentationProcessed;
            drive.PresentationError += Drive_PresentationError;
            drive.PresentationSkipped += Drive_PresentationSkipped;

            #region Parse args

            string presentationId = null;
            string specificFolderId = null;
            
            if (args.Length > 0) {
                if (args[0] == "/?")
                {
                    PrintUsageAndExit(0);
                }
                if (args[0].StartsWith("/RootFolderId="))
                {
                    specificFolderId = args[0].Split('=')[1];
                    processingType = ProcessingType.FOLDER;
                    
                }
                if (args[0].StartsWith("/PresentationId="))
                {
                    presentationId = args[0].Split('=')[1];
                    processingType = ProcessingType.PRESENTATION;
                }
                else if (args[0] == "/Students")
                {
                    processingType = ProcessingType.STUDENTS;
                }
            }

            #endregion

            switch (processingType) 
            {
                case ProcessingType.PRESENTATION:

                    #region Process specfic presentation

                    LogOutputWithNewLine(string.Format("Processing specific teacher presentation: {0}", presentationId));

                    var cachePresentation = drive.TeacherCache.GetPresentation(presentationId, drive.TeacherCache.Folders);
                    if (cachePresentation != null)
                    {
                        drive.ProcessTeacherPresentation(cachePresentation);
                    }
                    else
                    {
                        LogOutputWithNewLine(string.Format("Presentation {0} not found in cache", presentationId));
                    }
                    break;

                #endregion

                case ProcessingType.FOLDER:

                    #region Process folder presentations

                    var rootFolderId = ConfigurationManager.AppSettings["rootFolderId"];

                    CacheFolder rootFolder;
                    if (specificFolderId != null)
                    {
                        rootFolder = drive.TeacherCache.GetFolder(specificFolderId, drive.TeacherCache.Folders);
                        if (rootFolder == null)
                        {
                            //Specified folder id not found in cache
                            PrintUsageAndExit(3);
                        }
                        LogOutputWithNewLine(string.Format("Processing {0} teacher presentations in folder: {1}", rootFolder.TotalPresentations, rootFolder.FolderName));

                        drive.ProcessTeacherPresentations(rootFolder);
                    }
                    else
                    {
                        LogOutputWithNewLine(string.Format("Processing {0} teacher presentations", drive.TeacherCache.TotalPresentations));
                        foreach (var folderKey in drive.TeacherCache.Folders.Keys)
                        {
                            LogOutputWithNewLine(string.Format("Processing {0} teacher presentations in folder: {1}", drive.TeacherCache.Folders[folderKey].TotalPresentations, drive.TeacherCache.Folders[folderKey].FolderName));
                            drive.ProcessTeacherPresentations(drive.TeacherCache.Folders[folderKey]);
                        }
                    }

                    break;

                #endregion

                case ProcessingType.STUDENTS:

                    LogOutputWithNewLine("Processing student presentations...");
                    drive.ProcessStudentsPresentations();
                    break;
            }

            LogOutputWithNewLine("Finished...");
        }

        private static void Drive_PresentationSkipped(object sender, EventArgs e)
        {
            slidesSkipped++;
            OutputProgress();
        }

        private static void Drive_PresentationError(object sender, EventArgs e)
        {
            var slideErrorEventArgs = (SlideErrorEventArgs)e;
            LogOutputWithNewLine(string.Format("Presentation: {0} {1}, Slide: {2}, Error: {3}", slideErrorEventArgs.SlideError.PresentationId, slideErrorEventArgs.SlideError.PresentationName, slideErrorEventArgs.SlideError.SlideId, slideErrorEventArgs.SlideError.Error));
        }

        private static void Drive_PresentationProcessed(object sender, EventArgs e)
        {
            slidesProcessed++;
            OutputProgress();
        }

        private static void OutputProgress()
        {
            Console.Write(string.Format("\rStatus: processed: {0}, skipped: {1}, total: {2}...", slidesProcessed,slidesSkipped, slidesProcessed+slidesSkipped));
        }

        private static void LogOutputWithNewLine(string line)
        {
            Console.WriteLine(string.Format("\n{0}: {1}", DateTime.Now, line));
        }

        private static void PrintUsageAndExit(int exitCode)
        {
            Console.WriteLine("GoogleDrive [/?] [/RootFolderId=<FolderId>] [/Id=<PresentationId>] [/Students]");
            Console.WriteLine("Only one of the parameters can be specified at a time:");
            Console.WriteLine("/RootFolderId    Process only teacher presentations from this Root Folder and its subfolders");
            Console.WriteLine("/PresentationId  Process only this teacher presentation");
            Console.WriteLine("/Students        Process students presentations");
            Console.WriteLine("/?               Prints this help");
            Environment.Exit(exitCode);
        }
    }
}