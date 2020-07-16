using System;
using System.Configuration;

namespace GoogleDrive
{
    internal static class Program
    {
        private enum ProcessingType
        {
            TEACHER_PRESENTATION,
            TEACHER_FOLDER,
            STUDENTS
        };
        private static Drive drive;
        private static int totalSlidesProcessed;
        private static int totalSlidesSkipped;
        private static string currentProcessingFolder;
        private static int lastFolderSlidesProcessed;
        private static int lastFolderSlidesSkipped;
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
                            args[0] != "/Students" &&
                            (!args[0].StartsWith("/StartFromSheet=") || args[0].Length < 17)

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
            drive.FolderProcessingStarted += Drive_FolderProcessingStarted;
            drive.PresentationSkipped += Drive_PresentationSkipped;

            #region Parse args

            string presentationId = null;
            string specificFolderId = null;
            string startFromSheet = null;
            
            if (args.Length > 0) {
                if (args[0] == "/?")
                {
                    PrintUsageAndExit(0);
                }
                else if (args[0].StartsWith("/RootFolderId="))
                {
                    specificFolderId = args[0].Split('=')[1];
                    processingType = ProcessingType.TEACHER_FOLDER;                    
                }
                else if (args[0].StartsWith("/PresentationId="))
                {
                    presentationId = args[0].Split('=')[1];
                    processingType = ProcessingType.TEACHER_PRESENTATION;
                }
                else if (args[0] == "/Students")
                {
                    processingType = ProcessingType.STUDENTS;
                }
                else if (args[0].StartsWith("/StartFromSheet="))
                {
                    startFromSheet = args[0].Split('=')[1];
                    processingType = ProcessingType.STUDENTS;
                }
            }
            else
            {
                processingType = ProcessingType.TEACHER_FOLDER;
            }

            #endregion

            switch (processingType) 
            {
                case ProcessingType.TEACHER_PRESENTATION:

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

                case ProcessingType.TEACHER_FOLDER:

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

                        drive.ProcessTeacherPresentations(rootFolder);
                    }
                    else
                    {
                        LogOutputWithNewLine(string.Format("Start processing {0} teacher presentations...", drive.TeacherCache.TotalPresentations));
                        foreach (var folderKey in drive.TeacherCache.Folders.Keys)
                        {
                            drive.ProcessTeacherPresentations(drive.TeacherCache.Folders[folderKey]);
                        }
                    }

                    break;

                #endregion

                case ProcessingType.STUDENTS:
                    if (startFromSheet != null)
                    {
                        if (drive.StudentsCache.GetSubFolderByName(startFromSheet) != null)
                        {
                            LogOutputWithNewLine(string.Format("Processing student presentations, starting from sheet {0}...", startFromSheet));
                            drive.ProcessStudentsPresentations(startFromSheet);
                        }
                        else
                        {
                            LogOutputWithNewLine(string.Format("Sheet {0} does not exist", startFromSheet));
                        }
                    }
                    else
                    {
                        LogOutputWithNewLine("Processing student presentations...");
                        drive.ProcessStudentsPresentations();
                    }
                    break;
            }

            LogOutputWithNewLine("Finished...");
        }

        private static void Drive_PresentationSkipped(object sender, EventArgs e)
        {
            lastFolderSlidesSkipped++;
            totalSlidesSkipped++;
            OutputProgress();
        }

        private static void Drive_PresentationError(object sender, EventArgs e)
        {
            var slideErrorEventArgs = (SlideErrorEventArgs)e;
            LogOutputWithNewLine(string.Format("Presentation: {0} {1}, Slide: {2}, Error: {3}", slideErrorEventArgs.SlideError.PresentationId, slideErrorEventArgs.SlideError.PresentationName, slideErrorEventArgs.SlideError.SlideId, slideErrorEventArgs.SlideError.Error));
        }

        private static void Drive_FolderProcessingStarted(object sender, EventArgs e)
        {
            lastFolderSlidesProcessed = 0;
            lastFolderSlidesSkipped = 0;

            var processFolderEventArgs = (ProcessFolderEventArgs)e;
            currentProcessingFolder = processFolderEventArgs.FolderName;

            LogOutputWithNewLine(string.Format("\nStarted folder: {0}, {1} presentations", processFolderEventArgs.FolderName, processFolderEventArgs.TotalPresentations));
        }

        private static void Drive_PresentationProcessed(object sender, EventArgs e)
        {
            lastFolderSlidesProcessed++;
            totalSlidesProcessed++;
            OutputProgress();
        }

        private static void OutputProgress()
        {
            Console.Write(string.Format("\r{0}: {1} ({2}: done, {3}: skipped), Total: {4} ({5} done, {6} skipped)...", currentProcessingFolder, lastFolderSlidesProcessed + lastFolderSlidesSkipped, lastFolderSlidesProcessed, lastFolderSlidesSkipped, totalSlidesProcessed + totalSlidesSkipped, totalSlidesProcessed, totalSlidesSkipped));
        }

        private static void LogOutputWithNewLine(string line)
        {
            Console.WriteLine(string.Format("\n{0}: {1}", DateTime.Now, line));
        }

        private static void PrintUsageAndExit(int exitCode)
        {
            Console.WriteLine("GoogleDrive [/?] [/RootFolderId=<FolderId>] [/Id=<PresentationId>] [/Students] [/StartFromSheet=<SheetName>]");
            Console.WriteLine("Only one of the parameters can be specified at a time:");
            Console.WriteLine("/RootFolderId                Process only teacher presentations from this Root Folder and its subfolders");
            Console.WriteLine("/PresentationId              Process only this teacher presentation");
            Console.WriteLine("/Students                    Process students presentations");
            Console.WriteLine("/StartFromSheet=<SheetName>  Process students presentations, start from sheet by its name");
            Console.WriteLine("/?                           Prints this help");
            Console.WriteLine("");
            Console.WriteLine("If no parameter is specified, default will process teacher presentations from a root folder in the configuration file");
            Environment.Exit(exitCode);
        }
    }
}