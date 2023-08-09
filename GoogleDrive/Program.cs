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
            string studentsStartSheet = null;
            bool skipTimeStampCheck = false;

            #endregion

            #region Parse command line

            commandLineParser.SetupHelp("h", "help", "?")
                  .Callback(callback => PrintUsageAndExit(0));

            commandLineParser.Setup<string>('m', "mode")
                .Required()
                .Callback(value => mode = value)
                .WithDescription("Select either 'Teacher' or 'Students'. If not supplied - default is set to 'Teacher'");

            commandLineParser.Setup<string>('r', "teacherrootfolder")
                .Callback(value => teacherRootFolder = value)
                .WithDescription("In teacher mode: override configuration's top root folder");

           commandLineParser.Setup<string>('p', "teacherpresentationid")
                .Callback(value => teacherPresentationId = value)
                .WithDescription("In teacher mode: work on this specific presentation only");

           commandLineParser.Setup<string>('s', "studentsstartsheet")
                .Callback(value => studentsStartSheet = value)
                .WithDescription("In students mode: skip and start working from this sheet");

           commandLineParser.Setup<bool>('t', "teacherskiptimestampcheck")
                .Callback(value => skipTimeStampCheck = value)
                .SetDefault(false)
                .WithDescription("allows skipping time stamp check - default is false - time stamp will be checked");

           var commandLineArgs = commandLineParser.Parse(args);

            if (commandLineArgs.HasErrors)
            {
                Console.WriteLine(commandLineArgs.ErrorText);
                PrintUsageAndExit(0);
                return;
            }

            #endregion

            LogOutputWithNewLine("Started...");

            drive = new Drive();
            drive.PresentationProcessed += Drive_PresentationProcessed;
            drive.PresentationError += Drive_PresentationError;
            drive.FolderProcessingStarted += Drive_FolderProcessingStarted;
            drive.PresentationSkipped += Drive_PresentationSkipped;

            switch (mode)
            {
                case "Teacher":
                    {
                        #region Validate Teacher arguments

                        if (studentsStartSheet != null)
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

                        if (studentsStartSheet != null)
                        {
                            if (drive.StudentsCache.GetSubFolderByName(studentsStartSheet) != null)
                            {
                                LogOutputWithNewLine(string.Format("Processing student presentations, starting from sheet {0}...", studentsStartSheet));
                                drive.ProcessStudentsPresentations(studentsStartSheet, skipTimeStampCheck);
                            }
                            else
                            {
                                Console.WriteLine(string.Format("Sheet {0} does not exist", studentsStartSheet));
                                PrintUsageAndExit(7);
                            }
                        }
                        else
                        {
                            LogOutputWithNewLine("Processing student presentations...");
                            drive.ProcessStudentsPresentations();
                        }

                        #endregion

                        break;
                    }

                default:

                    Console.WriteLine(string.Format("Mode {0} is invalid. Supported modes are only: 'Teacher' or 'Students'", mode));
                    PrintUsageAndExit(8);

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
            Console.WriteLine("GoogleDrive [/h /help / ?] [/m:Teacher|Students] [/r:teacherrootfolder] [/p:presentationid] [/t]");
            Console.WriteLine("/h /help /?                  Prints this screen");
            Console.WriteLine("/m:Teacher|Students          Mode: 'Teacher' or 'Students'");
            Console.WriteLine("/p:presentationId            Process only the teacher presentation");
            Console.WriteLine("/r:rootfolder                Process only a specific teacher root folder (e.g. 801, 802, ...)");
            Console.WriteLine("/s:sheetname                 Process students presentations, start from sheet by its name");
            Console.WriteLine("/t                           Skip time stamp check");
            Environment.Exit(exitCode);
        }
    }
}