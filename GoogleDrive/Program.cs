using System;
using Fclp;

namespace GoogleDrive
{
    internal static class Program
    {
        private static Drive drive;
        private static int totalSlidesProcessed;
        private static int totalSlidesSkipped;
        private static string currentProcessingFolder;
        private static int lastFolderSlidesProcessed;
        private static int lastFolderSlidesSkipped;

        private static void Main(string[] args)
        {
            #region Variables

            var commandLineParser = new FluentCommandLineParser();
            string mode = null;
            string teacherRootFolder = null;
            string teacherPresentationId = null;
            string studentsSpecificSheet = null;
            bool skipTimeStampCheck = false;
            bool clearAndRebuildCache = false;

            #endregion

            #region Parse command line

            commandLineParser.SetupHelp("h", "help", "?")
                  .Callback(callback => PrintUsageAndExit(0));

            commandLineParser.Setup<string>('m', "mode")
                .Callback(value => mode = value)
                .WithDescription("Select either 'Teacher' or 'Students'. If not supplied - default is set to 'Teacher'");

            commandLineParser.Setup<string>('r', "teacherrootfolder")
                .Callback(value => teacherRootFolder = value)
                .WithDescription("In teacher mode: override configuration's top root folder");

           commandLineParser.Setup<string>('p', "teacherpresentationid")
                .Callback(value => teacherPresentationId = value)
                .WithDescription("In teacher mode: work on this specific presentation only");

           commandLineParser.Setup<string>('s', "studentsspecificsheet")
                .Callback(value => studentsSpecificSheet = value)
                .WithDescription("In students mode: process only this students sheet");

           commandLineParser.Setup<bool>('t', "teacherskiptimestampcheck")
                .Callback(value => skipTimeStampCheck = value)
                .SetDefault(false)
                .WithDescription("allows skipping time stamp check - default is false - time stamp will be checked");

            commandLineParser.Setup<bool>('c', "clearandrebuildcache")
                 .Callback(value => clearAndRebuildCache = value)
                 .SetDefault(false)
                 .WithDescription("clears the json local cache files");

            var commandLineArgs = commandLineParser.Parse(args);

            if (commandLineArgs.HasErrors)
            {
                Console.WriteLine(commandLineArgs.ErrorText);
                PrintUsageAndExit(0);
                return;
            }

            #endregion

            LogOutputWithNewLine("Started...");

            drive = new Drive(clearAndRebuildCache);
            drive.PresentationProcessed += Drive_PresentationProcessed;
            drive.PresentationError += Drive_PresentationError;
            drive.FolderProcessingStarted += Drive_FolderProcessingStarted;
            drive.PresentationSkipped += Drive_PresentationSkipped;

            switch (mode)
            {
                case "Teacher":
                    {
                        #region Validate Teacher arguments

                        if (studentsSpecificSheet != null)
                        {
                            Console.WriteLine("'studentsstartsheet' argument is valid only in 'Students' mode");
                            PrintUsageAndExit(1);
                        }
                        if (teacherPresentationId != null && teacherRootFolder != null)
                        {
                            Console.WriteLine("Only one of: 'teacherrootfolder', 'teacherpresentationid' can be specified in 'Teacher' mode");
                            PrintUsageAndExit(2);
                        }

                        #endregion

                        #region Teacher cases

                        if (teacherPresentationId != null)
                        {
                            //Processing specific presentation
                            var cachePresentation = drive.TeacherCache.GetPresentation(teacherPresentationId, drive.TeacherCache.Folders);
                            if (cachePresentation != null)
                            {
                                LogOutputWithNewLine(string.Format("Processing specific teacher presentation: {0}", teacherPresentationId));
                                drive.ProcessTeacherPresentation(cachePresentation, skipTimeStampCheck);
                            }
                            else
                            {
                                Console.WriteLine(string.Format("Presentation {0} not found in cache", teacherPresentationId));
                                PrintUsageAndExit(3);
                            }
                        }
                        else
                        {
                            CacheFolder rootFolder;
                            if (teacherRootFolder != null)
                            {
                                rootFolder = drive.TeacherCache.GetSubFolderByName(teacherRootFolder);
                                //Process only a specified root folder
                                if (rootFolder != null)
                                {
                                    drive.ProcessTeacherPresentations(rootFolder, skipTimeStampCheck);
                                }
                                else
                                {
                                    Console.WriteLine(string.Format("Teacher root folder {0} not found in cache", teacherRootFolder));
                                    PrintUsageAndExit(4);
                                }
                            }
                            else
                            {
                                //Process all folders
                                LogOutputWithNewLine(string.Format("Start processing {0} teacher presentations...", drive.TeacherCache.TotalPresentations));
                                foreach (var folderKey in drive.TeacherCache.Folders.Keys)
                                {
                                    drive.ProcessTeacherPresentations(drive.TeacherCache.Folders[folderKey], skipTimeStampCheck);
                                }
                            }
                        }

                        #endregion

                        break;
                    }
                case "Students":
                    {
                        #region Validate Students arguments

                        if (teacherRootFolder != null)
                        {
                            Console.WriteLine("'teacherrootfolder' argument is valid only in 'Teacher' mode");
                            PrintUsageAndExit(5);
                        }
                        else if (teacherPresentationId != null)
                        {
                            Console.WriteLine("teacherrootfolder is valid only in 'Teacher' mode");
                            PrintUsageAndExit(6);
                        }

                        #endregion

                        #region Students cases

                        if (studentsSpecificSheet != null)
                        {
                            if (drive.StudentsCache.GetSubFolderByName(studentsSpecificSheet) != null)
                            {
                                LogOutputWithNewLine(string.Format("Processing student presentations, only for sheet {0}...", studentsSpecificSheet));
                                drive.ProcessStudentsPresentations(studentsSpecificSheet, skipTimeStampCheck);
                            }
                            else
                            {
                                Console.WriteLine(string.Format("Sheet {0} does not exist", studentsSpecificSheet));
                                PrintUsageAndExit(7);
                            }
                        }
                        else
                        {
                            LogOutputWithNewLine("Processing students presentations...");
                            drive.ProcessStudentsPresentations();
                        }

                        #endregion

                        break;
                    }

                default:
                    //Empty mode is allowed only when clearing and rebuilding cache
                    if (!clearAndRebuildCache)
                    {
                        Console.WriteLine(string.Format("Mode {0} is invalid. Supported modes are only: 'Teacher' or 'Students'", mode));
                        PrintUsageAndExit(8);
                    }

                    break;
            }

            LogOutputWithNewLine("Finished...");
        }

        private static void Drive_PresentationSkipped(object sender, EventArgs e)
        {
            lastFolderSlidesSkipped++;
            totalSlidesSkipped++;
            if (e != null)
            {
                var slideSkippedEventArgs = (SlideSkippedEventArgs)e;
                if (slideSkippedEventArgs.SlideSkipped.SheetRowNumber > 0)
                {
                    //Students mode - assuming empty row
                    LogOutputWithNewLine(string.Format("Row number {0} skipped in sheet, Reason: {1}", slideSkippedEventArgs.SlideSkipped.SheetRowNumber, slideSkippedEventArgs.SlideSkipped.SkipReason));
                }
                else
                {
                    //Teachers mode - assuming empty row
                    LogOutputWithNewLine(string.Format("Presentation: {0} {1}, Slide: {2}, Error: {3}", slideSkippedEventArgs.SlideSkipped.PresentationId, slideSkippedEventArgs.SlideSkipped.PresentationName, slideSkippedEventArgs.SlideSkipped.SlideId, slideSkippedEventArgs.SlideSkipped.SkipReason));
                }
            }
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
            Console.WriteLine("GoogleDrive [/h /help / ?] [/m:Teacher|Students] [/r:teacherrootfolder] [/p:presentationid] [/t]");
            Console.WriteLine("/h /help /?                  Prints this screen");
            Console.WriteLine("/m:Teacher|Students          Mode: 'Teacher' or 'Students'");
            Console.WriteLine("/p:presentationId            Process only the teacher presentation");
            Console.WriteLine("/r:rootfolder                Process only a specific teacher root folder (e.g. 182, 381, ...)");
            Console.WriteLine("/s:sheetname                 Process students presentations, only this specific sheet");
            Console.WriteLine("/t                           Skip time stamp check");
            Console.WriteLine("/c                           Clear and rebuild local cache files");
            Environment.Exit(exitCode);
        }
    }
}