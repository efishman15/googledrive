using Google.Apis.Auth.OAuth2;
using Google.Apis.Drive.v3;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Slides.v1;
using Google.Apis.Slides.v1.Data;
using Google.Apis.Script.v1;
using Google.Apis.Script.v1.Data;
using Google.Apis.Util.Store;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Newtonsoft.Json;
using Newtonsoft.Json.Converters;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Threading;
using System.Text.RegularExpressions;
using Google.Apis.Sheets.v4.Data;
using Link = Google.Apis.Slides.v1.Data.Link;
using File = System.IO.File;
using System.Globalization;

namespace GoogleDrive
{
    #region Enums

    public enum AlignImage
    {
        TOP,
        BOTTOM
    }

    #endregion

    #region Class CachePresentation

    public class CachePresentation
    {
        #region Properties

        public string PresentationId { get; private set; }
        public string PresentationName { get; private set; }
        public string FooterText { get; set; }

        #endregion

        #region C'Tor/Dtor
        public CachePresentation(string presentationId, string presentationName)
        {
            PresentationId = presentationId;
            PresentationName = presentationName;
        }
        #endregion
    }
    #endregion

    #region Class CacheFolder
    public class CacheFolder
    {
        #region Properties

        public string FolderId { get; private set; }
        public string FolderName { get; private set; }
        public Dictionary<string, CacheFolder> Folders { get; set; }
        public string ParentFolderId { get; private set; }
        public List<CachePresentation> Presentations { get; private set; }
        public int TotalPresentations { get; set; }
        public int Level { get; set; }
        public string Path { get; set; }

        #endregion

        #region C'Tor/Dtor
        public CacheFolder(string folderId, string folderName, string parentFolderId)
        {
            FolderId = folderId;
            FolderName = folderName;
            Folders = new Dictionary<string, CacheFolder>();
            ParentFolderId = parentFolderId;
            Presentations = new List<CachePresentation>();
            Path = string.Empty;
        }
        #endregion

        #region Methods

        /// <summary>
        /// Adds a folder to the collection
        /// </summary>
        /// <param name="folderId"></param>
        /// <param name="folderName"></param>
        public void AddFolder(string folderId, string folderName, CacheFolder parent)
        {
            Folders.Add(folderId, new CacheFolder(folderId, folderName, this.FolderId));
        }

        /// <summary>
        /// Adds a presentation to a folder
        /// </summary>
        /// <param name="presentationId"></param>
        /// <param name="presentationName"></param>
        public CachePresentation AddPresentation(string presentationId, string presentationName)
        {
            var newCachePresentation = new CachePresentation(presentationId, presentationName);
            Presentations.Add(newCachePresentation);
            TotalPresentations++;
            return newCachePresentation;
        }

        /// <summary>
        /// Check if subfolder exists by its name
        /// </summary>
        /// <param name="name"></param>
        /// <returns></returns>
        public CacheFolder GetSubFolderByName(string name)
        {
            foreach(var key in Folders.Keys)
            {
                if (Folders[key].FolderName == name)
                {
                    return Folders[key];
                }
            }

            return null;
        }
        #endregion
    }
    #endregion

    #region Class Cache
    public class Cache
    {
        #region Private Members

        private string foldersFilter;
        private DriveService driveService;
        private int pathStartLevel;
        private string pathSeparator;
        private string folderNameSeparator;
        private string folderMimeType;
        private string presentationMimeType;


        #endregion

        #region Properties

        public string Name { get; private set; }
        public Dictionary<string, CacheFolder> Folders { get; set; }
        public int TotalPresentations { get; set; }
        public string DateCreated { get; set; }

        #endregion

        #region C'Tor/Dtor

        public Cache(DriveService driveService, string name)
        {
            Name = name;
            this.driveService = driveService;
            TotalPresentations = 0;
            Folders = new Dictionary<string, CacheFolder>();
        }

        public static Cache Load(DriveService driveService, string name)
        {
            // deserialize JSON directly from a file
            string cacheFileName = ConfigurationManager.AppSettings[name + "Cache"];
            if (File.Exists(cacheFileName))
            {
                using (StreamReader file = System.IO.File.OpenText(cacheFileName))
                {
                    var jsonSerializer = new JsonSerializer();
                    Cache cache = (Cache)jsonSerializer.Deserialize(file, typeof(Cache));
                    cache.Init();
                    return cache;
                }
            }
            else
            {
                var cache = new Cache(driveService, name);
                cache.Init();

                cache.Build(ConfigurationManager.AppSettings[name + "RootFolderId"], true, null);

                cache.BuildFoldersPath(cache.Folders, string.Empty);
                cache.Save();
                return cache;
            }
        }

        #endregion

        #region Methods

        /// <summary>
        /// Deletes a local cache file
        /// </summary>
        /// <param name="name"></param>
        public static void DeleteLocalFile(string name)
        {
            string cacheFileName = ConfigurationManager.AppSettings[name + "Cache"];
            if (File.Exists(cacheFileName))
            {
                File.Delete(cacheFileName);
            }

        }

        /// <summary>
        /// Get a folder in the tree by its folder id (recurssive)
        /// </summary>
        /// <param name="folderId"></param>
        /// <param name="folders"></param>
        /// <returns></returns>
        public CacheFolder GetFolder(string folderId, Dictionary<string, CacheFolder> folders = null)
        {
            if (folders == null)
            {
                folders = Folders;
            }

            if (folders.ContainsKey(folderId))
            {
                return folders[folderId];
            }
            else
            {
                foreach (var folderKey in folders.Keys)
                {
                    var folder = GetFolder(folderId, folders[folderKey].Folders);
                    if (folder != null)
                    {
                        return folder;
                    }
                }
            }

            return null;
        }

        /// <summary>
        /// Get a presentation in the tree by its presentation id (recurssive)
        /// </summary>
        /// <param name="presentationId"></param>
        /// <param name="folders"></param>
        /// <returns></returns>
        public CachePresentation GetPresentation(string presentationId, Dictionary<string, CacheFolder> folders)
        {
            foreach (var folderKey in folders.Keys)
            {
                foreach (var presentation in folders[folderKey].Presentations)
                {
                    if (presentation.PresentationId == presentationId)
                    {
                        return presentation;
                    }
                }

                var presentationInSubFolders = GetPresentation(presentationId, folders[folderKey].Folders);
                if (presentationInSubFolders != null)
                {
                    return presentationInSubFolders;
                }
            }

            return null;
        }

        /// <summary>
        /// Returns folder mime type in google drive as in app.config
        /// </summary>
        /// <returns></returns>
        public string GetFolderMimeType()
        {
            return folderMimeType;
        }

        /// <summary>
        /// Returns presentation mime type in google drive as in app.config
        /// </summary>
        /// <returns></returns>
        public string GetPresentationMimeType()
        {
            return presentationMimeType;
        }

        /// <summary>
        /// Check if subfolder exists by its name
        /// </summary>
        /// <param name="name"></param>
        /// <returns></returns>
        public CacheFolder GetSubFolderByName(string name)
        {
            foreach (var key in Folders.Keys)
            {
                if (Folders[key].FolderName == name)
                {
                    return Folders[key];
                }
            }

            return null;
        }

        /// <summary>
        /// Save cache to a local file
        /// </summary>
        public void Save()
        {
            var outputFileName = ConfigurationManager.AppSettings[Name + "Cache"];
            if (System.IO.File.Exists(outputFileName))
            {
                System.IO.File.Delete(outputFileName);
            }

            var jsonSerializer = new JsonSerializer();

            DateCreated = DateTime.Now.ToString();
            jsonSerializer.Converters.Add(new JavaScriptDateTimeConverter());
            jsonSerializer.NullValueHandling = NullValueHandling.Ignore;
            using (StreamWriter sw = new StreamWriter(outputFileName))
            using (JsonWriter writer = new JsonTextWriter(sw))
            {
                jsonSerializer.Serialize(writer, this);
            }

        }

        #endregion

        #region Private Methods

        /// <summary>
        /// Initialize variables from config
        /// </summary>
        /// <param name="driveService"></param>
        /// <param name="name"></param>
        /// <param name="pathStartLevel"></param>
        private void Init()
        {
            pathStartLevel = Convert.ToInt32(ConfigurationManager.AppSettings[Name + "PathStartLevel"]);

            pathSeparator = ConfigurationManager.AppSettings["PathSeparator"];
            folderNameSeparator = ConfigurationManager.AppSettings["FolderNameSeparator"];

            folderMimeType = ConfigurationManager.AppSettings["FolderMimeType"];
            presentationMimeType = ConfigurationManager.AppSettings["PresentationMimeType"];
        }

        /// <summary>
        /// Build recursivelly a cache of all folders and presentations under a top root folder in Google Drive
        /// </summary>
        /// <param name="rootFolderId"></param>
        /// <param name="isTop"></param>
        /// <param name=""></param>
        private void Build(string rootFolderId, bool isTop, CacheFolder parentFolder)
        {
            string filter = "'" + rootFolderId + "' in parents AND (mimeType = '" + folderMimeType + "') AND trashed=false";
            string pageToken = null;

            if (isTop)
            {
                foldersFilter = string.Empty;
            }
            do
            {
                var folderRequest = driveService.Files.List();
                folderRequest.Q = filter;
                folderRequest.OrderBy = "name";
                folderRequest.Spaces = "drive";
                folderRequest.Fields = "nextPageToken, files(id,name)";
                folderRequest.PageToken = pageToken;
                var folderResult = folderRequest.Execute();
                foreach (var folder in folderResult.Files)
                {
                    if (foldersFilter == string.Empty)
                    {
                        foldersFilter += "(";
                    }
                    else
                    {
                        foldersFilter += " or ";
                    }
                    foldersFilter += "'" + folder.Id + "' in parents";

                    var newFolder = new CacheFolder(folder.Id, folder.Name, parentFolder?.FolderId);
                    if (parentFolder == null)
                    {
                        newFolder.Level = 1;
                        Folders.Add(newFolder.FolderId, newFolder);
                    }
                    else
                    {
                        parentFolder.Folders.Add(newFolder.FolderId, newFolder);
                        newFolder.Level = parentFolder.Level + 1;
                    }

                    Build(folder.Id, false, newFolder);
                }
                pageToken = folderResult.NextPageToken;
            } while (pageToken != null);

            if (isTop)
            {
                if (foldersFilter == string.Empty)
                {
                    //No folders in root folder - just files
                    foldersFilter = "('" + rootFolderId + "' in parents) ";
                }
                else
                {
                    foldersFilter += ")";
                }
                filter = foldersFilter + " AND (mimeType = '" + presentationMimeType + "') AND trashed=false";
                do
                {
                    var filesRequest = driveService.Files.List();
                    filesRequest.Q = filter;
                    filesRequest.Spaces = "drive";
                    filesRequest.Fields = "nextPageToken, files(id, name, parents)";
                    filesRequest.PageToken = pageToken;
                    var fileResult = filesRequest.Execute();

                    foreach (var file in fileResult.Files)
                    {
                        //If file is filed in more than 1 folder, add it only under the first folder
                        //For the sake of this program - each presentation should be processed only once
                        AddPresentationToFolder(file.Parents[0], file.Id, file.Name);
                    }
                    pageToken = fileResult.NextPageToken;
                } while (pageToken != null);
            }
        }

        /// <summary>
        /// Process the cache:
        /// Add path to folders starting at "PathStartLevel" as given when cache was created/loaded
        /// </summary>
        private void BuildFoldersPath(Dictionary<string, CacheFolder> root, string parentPath)
        {
            foreach (var folderKey in root.Keys)
            {
                if (root[folderKey].Level >= pathStartLevel)
                {
                    if (parentPath == string.Empty)
                    {
                        root[folderKey].Path = NormalizeFolderName(root[folderKey].FolderName);
                    }
                    else
                    {
                        root[folderKey].Path = parentPath + pathSeparator + NormalizeFolderName(root[folderKey].FolderName);
                    }
                    foreach (var presentation in root[folderKey].Presentations)
                    {
                        presentation.FooterText = root[folderKey].Path;
                    }
                }
                else
                {
                    root[folderKey].Path = string.Empty;
                }
                BuildFoldersPath(root[folderKey].Folders, root[folderKey].Path);
            }
        }

        /// <summary>
        /// Adds a presentation to the folder in the tree
        /// </summary>
        /// <param name="folderId"></param>
        /// <param name="presentationId"></param>
        /// <param name="presentationName"></param>f
        public void AddPresentationToFolder(string folderId, string presentationId, string presentationName)
        {
            var parentFolder = GetFolder(folderId, Folders);
            if (parentFolder != null)
            {
                parentFolder.AddPresentation(presentationId, presentationName);
                TotalPresentations++;

                //Bubble counter up
                var currentFolder = GetFolder(parentFolder.ParentFolderId, Folders);
                while (currentFolder != null)
                {
                    currentFolder.TotalPresentations++;
                    if (currentFolder.ParentFolderId != null)
                    {
                        currentFolder = GetFolder(currentFolder.ParentFolderId, Folders);
                    }
                    else
                    {
                        currentFolder = null;
                    }
                }

            }
        }

        /// <summary>
        /// Returns a normalized folder name - everything AFTER the folderNameSeparator in config (usually ".")
        /// </summary>
        /// <param name="folderName"></param>
        /// <returns></returns>
        private string NormalizeFolderName(string folderName)
        {
            if (!folderName.Contains(folderNameSeparator))
            {
                return folderName.Trim();
            }

            var splitArray = folderName.Split(folderNameSeparator.ToCharArray());
            return splitArray[splitArray.Length - 1].Trim();
        }

        #endregion
    }
    #endregion

    #region Class Slide Error

    public class SlideError
    {
        #region Properties

        public string PresentationId { get; private set; }
        public string PresentationName { get; private set; }
        public int SlideId { get; private set; }
        public string Error { get; private set; }

        #endregion

        #region C'Tor/Dtor
        public SlideError(string presentationId, string presentationName, int slideId, string error)
        {
            PresentationId = presentationId;
            PresentationName = presentationName;
            SlideId = slideId;
            Error = error;
        }
        #endregion
    }

    #endregion

    #region Class Slide Skipped

    public class SlideSkipped
    {
        #region Properties

        public string PresentationId { get; private set; }
        public string PresentationName { get; private set; }
        public int SlideId { get; private set; }
        public int SheetRowNumber { get; private set; }

        public string SkipReason { get; private set; }

        #endregion

        #region C'Tor/Dtor
        public SlideSkipped(string presentationId, string presentationName, int slideId, int sheetRowNumber, string skipReason)
        {
            PresentationId = presentationId;
            PresentationName = presentationName;
            SlideId = slideId;
            SheetRowNumber = sheetRowNumber;
            SkipReason = skipReason;
        }
        #endregion
    }

    #endregion

    #region Class ProcessFolderEventArgs

    public class ProcessFolderEventArgs : EventArgs
    {
        #region Properties

        public string FolderName { get; set; }
        public int TotalPresentations { get; set; }

        #endregion

        #region C'Tor/D'Tor

        public ProcessFolderEventArgs(string folderName, int totalPresentations)
        {
            FolderName = folderName;
            TotalPresentations = totalPresentations;
        }

        #endregion
    }

    #endregion

    #region Class SlideErrorEventArgs

    public class SlideErrorEventArgs : EventArgs
    {
        #region Properties

        public SlideError SlideError { get; set; }

        #endregion

        #region C'Tor/D'Tor

        public SlideErrorEventArgs(SlideError slideError)
        {
            SlideError = slideError;
        }

        #endregion
    }

    public class SlideSkippedEventArgs : EventArgs
    {
        #region Properties

        public SlideSkipped SlideSkipped { get; set; }

        #endregion

        #region C'Tor/D'Tor

        public SlideSkippedEventArgs(SlideSkipped slideSkipped)
        {
            SlideSkipped = slideSkipped;
        }

        #endregion
    }


    #endregion

    #region Class Drive

    public class Drive
    {
        #region Constants 

        const string APP_PROPERTY_NORMALIZE_TIME = "NormalizeTime";

        #endregion

        #region Class Members

        private DriveService driveService;
        private SlidesService slidesService;
        private SheetsService sheetService;
        private ScriptService scriptService;

        private static string[] Scopes = { DriveService.Scope.DriveMetadata, DriveService.Scope.Drive, DriveService.Scope.DriveFile, SlidesService.Scope.Presentations, SheetsService.Scope.Spreadsheets};
        private static string ApplicationName = "Google Drive";

        private readonly Google.Apis.Slides.v1.Data.Link lastSlidelink;
        private readonly Google.Apis.Slides.v1.Data.Link nextSlidelink;
        private readonly Google.Apis.Slides.v1.Data.Link prevSlidelink;
        private readonly Google.Apis.Slides.v1.Data.Link firstSlidelink;

        private readonly string firstSlideText;
        private readonly string prevSlideText;
        private readonly string nextSlideText;
        private readonly string lastSlideText;

        private readonly string speakerNotesTextStyleFields;

        private readonly string slideHeaderTextBoxTextStyleFields;
        private readonly string slideFooterTextBoxTextStyleFields;
        private readonly string slideIdTextBoxTextStyleFields;

        private readonly Size slidePageIdSize;
        private readonly AffineTransform slidePageIdTransform;

        private readonly Size slideHeaderSize;
        private readonly AffineTransform slideHeaderTransform;

        private readonly Size slideFooterSize;
        private readonly AffineTransform slideFooterTransform;

        private readonly AlignImage alignImage;

        private readonly string lookForTextInHeader;

        private readonly string msPowerPointMimeType;
        private readonly string msPowerPointTempLocalFileName;
        private readonly int msPowerPointPageNumberMinTop;
        private readonly int msPowerPointAutoShapeMaxHeight;

        private readonly Application pptApplication;
        private readonly Presentations pptPresentations;
        private readonly int whiteColor;

        private readonly string masterPlanSpreadsheetId;
        private readonly string masterPlanSpreadsheetReadRangePattern;
        private readonly string masterPlanSpreadsheetUpdateRangePattern;
        private readonly string regexSpreadsheetHyperlinkExtractIdPattern;
        private readonly string regexSpreadsheetHyperlinkExtractNamePattern;
        private readonly string spreadsheetHyperlinkFormat;
        private readonly string regexNonTeachingSlidePattern;
        private readonly string copySlidesAppScriptId;
        private readonly string copySlidesAppScriptFunction;
        private readonly CultureInfo dateTimeCulture;

        private Regex regexSpreadsheetHyperlinkExtractId;
        private Regex regexSpreadsheetHyperlinkExtractName;
        private Regex regexNonTeachingSlide;

        #endregion

        #region C'Tor/D'Tor
        public Drive(bool clearAndRebuildCache)
        {
            #region Authentication to Google
            UserCredential credential;

            using (var stream =
                new FileStream("credentials.json", FileMode.Open, FileAccess.Read))
            {
                string credPath = "~/.credentials/token.json";
                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.FromStream(stream).Secrets,
                    Scopes,
                    ConfigurationManager.AppSettings["user"],
                    CancellationToken.None,
                    new FileDataStore(credPath, true)).Result;
            }

            // Create Drive, Slides, Sheets, Scripts API services.
            driveService = new DriveService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });

            slidesService = new SlidesService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });

            sheetService = new SheetsService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });

            scriptService = new ScriptService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });

            #endregion

            #region Cache

            if (clearAndRebuildCache)
            {
                Cache.DeleteLocalFile("Teacher");
                Cache.DeleteLocalFile("Students");
            }

            TeacherCache = Cache.Load(driveService, "Teacher");
            StudentsCache = Cache.Load(driveService, "Students");

            #endregion

            #region Load Configiguration Variables

            lastSlidelink = new Link() { RelativeLink = "LAST_SLIDE" };
            nextSlidelink = new Link() { RelativeLink = "NEXT_SLIDE" };
            prevSlidelink = new Link() { RelativeLink = "PREVIOUS_SLIDE" };
            firstSlidelink = new Link() { RelativeLink = "FIRST_SLIDE" };

            firstSlideText = ConfigurationManager.AppSettings["FirstSlideText"] + "\t";
            prevSlideText = ConfigurationManager.AppSettings["PrevSlideText"] + "\t";
            nextSlideText = ConfigurationManager.AppSettings["NextSlideText"] + "\t";
            lastSlideText = ConfigurationManager.AppSettings["LastSlideText"] + "\t";

            speakerNotesTextStyleFields = ConfigurationManager.AppSettings["SpeakerNotestTextStyleFields"];

            slideHeaderTextBoxTextStyleFields = ConfigurationManager.AppSettings["SlideHeaderTextBoxTextStyleFields"];
            slideFooterTextBoxTextStyleFields = ConfigurationManager.AppSettings["SlideFooterTextBoxTextStyleFields"];
            slideIdTextBoxTextStyleFields = ConfigurationManager.AppSettings["SlideIdTextBoxTextStyleFields"];

            slidePageIdSize = JsonConvert.DeserializeObject<Size>(ConfigurationManager.AppSettings["SlidePageIdSize"]);
            slidePageIdTransform = JsonConvert.DeserializeObject<AffineTransform>(ConfigurationManager.AppSettings["SlidePageIdTransform"]);

            slideHeaderSize = JsonConvert.DeserializeObject<Size>(ConfigurationManager.AppSettings["SlideHeaderSize"]);
            slideHeaderTransform = JsonConvert.DeserializeObject<AffineTransform>(ConfigurationManager.AppSettings["SlideHeaderTransform"]);

            slideFooterSize = JsonConvert.DeserializeObject<Size>(ConfigurationManager.AppSettings["SlideFooterSize"]);
            slideFooterTransform = JsonConvert.DeserializeObject<AffineTransform>(ConfigurationManager.AppSettings["SlideFooterTransform"]);

            alignImage = (AlignImage)Enum.Parse(typeof(AlignImage), ConfigurationManager.AppSettings["ImageAlign"]);

            lookForTextInHeader = ConfigurationManager.AppSettings["LookForTextInHeader"];

            msPowerPointMimeType = ConfigurationManager.AppSettings["MSPowerPointMimeType"];
            msPowerPointTempLocalFileName = ConfigurationManager.AppSettings["MSPowerPointTempLocalFileName"];
            whiteColor = int.Parse(ConfigurationManager.AppSettings["WhiteColor"]);
            msPowerPointPageNumberMinTop = int.Parse(ConfigurationManager.AppSettings["MSPowerPointPageNumberMinTop"]);
            msPowerPointAutoShapeMaxHeight = int.Parse(ConfigurationManager.AppSettings["MSPowerPointAutoShapeMaxHeight"]);

            masterPlanSpreadsheetId = ConfigurationManager.AppSettings["MasterPlanSpreadsheetId"];
            masterPlanSpreadsheetReadRangePattern = ConfigurationManager.AppSettings["MasterPlanSpreadsheetReadRangePattern"];
            masterPlanSpreadsheetUpdateRangePattern = ConfigurationManager.AppSettings["MasterPlanSpreadsheetUpdateRangePattern"];
            regexSpreadsheetHyperlinkExtractIdPattern = ConfigurationManager.AppSettings["RegexSpreadsheetHyperlinkExtractIdPattern"];
            regexSpreadsheetHyperlinkExtractNamePattern = ConfigurationManager.AppSettings["RegexSpreadsheetHyperlinkExtractNamePattern"];
            spreadsheetHyperlinkFormat = ConfigurationManager.AppSettings["SpreadsheetHyperlinkFormat"];
            regexNonTeachingSlidePattern = ConfigurationManager.AppSettings["RegexNonTeachingSlidePattern"];
            copySlidesAppScriptId = ConfigurationManager.AppSettings["CopySlidesAppScriptId"];
            copySlidesAppScriptFunction = ConfigurationManager.AppSettings["CopySlidesAppScriptFunction"];
            dateTimeCulture = new CultureInfo(ConfigurationManager.AppSettings["DateTimeCulture"]);

            regexSpreadsheetHyperlinkExtractId = new Regex(regexSpreadsheetHyperlinkExtractIdPattern);
            regexSpreadsheetHyperlinkExtractName = new Regex(regexSpreadsheetHyperlinkExtractNamePattern);
            regexNonTeachingSlide = new Regex(regexNonTeachingSlidePattern);

            pptApplication = new Application();
            pptPresentations = pptApplication.Presentations;



            #endregion
        }

        private void TeacherCache_BeforeBuildingCache(object sender, EventArgs e)
        {
            throw new NotImplementedException();
        }
        #endregion

        #region Properties

        public Cache TeacherCache { get; private set; }
        public Cache StudentsCache { get; private set; }

        #endregion

        #region Methods

        /// <summary>
        /// Process all the presentations in a root folder and its sub folders
        /// </summary>
        /// <param name="rootFolder"></param>
        public void ProcessTeacherPresentations(CacheFolder rootFolder, bool skipTimeStampCheck = false)
        {
            if (rootFolder.Level == 1)
            {
                FolderProcessingStarted.Invoke(this, new ProcessFolderEventArgs(rootFolder.FolderName, rootFolder.TotalPresentations));
            }

            foreach (var cachePresentation in rootFolder.Presentations)
            {
                ProcessTeacherPresentation(cachePresentation, skipTimeStampCheck);

            }
            //Process presentations in all subfolders
            foreach (var cachedFolderKey in rootFolder.Folders.Keys)
            {
                ProcessTeacherPresentations(rootFolder.Folders[cachedFolderKey], skipTimeStampCheck);
            }
        }

        /// <summary>
        /// Adjusts the presentation:
        /// 1) Adds an empty slide in the end, if it does not exist ("Empty board")
        /// 2) For each slide (except the last one "empty board"):
        ///     a) Delete existing speaker notes
        ///     b) Add Links to: "Prev Slide", "Next Slide" (to skip animated hints/solutions, "Last Slide" (empty board)
        ///     c) Adjust slide number text box
        ///     d) if slide contains only a single image in the body - align it to top/bottom 
        /// 3) For the last slide: add "TOC": a link to each slide (except this last slide)
        /// </summary>
        /// <param name="presentationId"></param>
        public void ProcessTeacherPresentation(CachePresentation cachePresentation, bool skipTimeStampCheck = false)
        {
            try
            {
                #region Local variables

                string objectId;
                int currentStartIndex;
                var slideHeaderCreated = false;
                var slideFooterCreated = false;
                var slidePageIdCreated = false;
                var desiredFooter = cachePresentation.FooterText + "\n";
                string desiredHeader = null;

                #endregion

                #region Load Presentation and check time stamp

                var presentationFile = LoadPresentationForTimeStampCheck(cachePresentation.PresentationId);

                if (presentationFile.AppProperties == null)
                {
                    presentationFile.AppProperties = new Dictionary<string, string>();
                }
                if (presentationFile.AppProperties.ContainsKey(APP_PROPERTY_NORMALIZE_TIME))
                {
                    if (!skipTimeStampCheck && presentationFile.ModifiedTimeDateTimeOffset <= Convert.ToDateTime(presentationFile.AppProperties[APP_PROPERTY_NORMALIZE_TIME],dateTimeCulture))
                    {
                        //File was not modified since it was last processed - skip
                        PresentationSkipped.Invoke(this, null);
                        return;
                    }
                }

                #endregion

                #region Load Presentation

                var presentationRequest = slidesService.Presentations.Get(cachePresentation.PresentationId);
                var presentation = presentationRequest.Execute();

                var myBatchRequest = new MyBatchRequest(slidesService, cachePresentation.PresentationId);

                #endregion

                #region Create Empty Slide (if neccessary)

                if (presentation.Slides[presentation.Slides.Count - 1].PageElements.Count > 2)
                {
                    //Create empty slide as the last slide
                    var createNewSlideBatchRequest = new MyBatchRequest(slidesService, cachePresentation.PresentationId);
                    createNewSlideBatchRequest.AddCreateSlideRequest(presentation.Slides.Count);
                    createNewSlideBatchRequest.Execute();

                    //Read presentation with the newly created slide
                    presentation = presentationRequest.Execute();
                }
                else
                {
                    //Deals with the case that the empty slide contains an unneccessary header/footer text
                    for (var i = 0; i < presentation.Slides[presentation.Slides.Count - 1].PageElements.Count; i++)
                    {
                        myBatchRequest.AddDeleteTextRequest(presentation.Slides[presentation.Slides.Count - 1].PageElements[i].ObjectId, presentation.Slides[presentation.Slides.Count - 1].PageElements[i].Shape);
                    }
                }

                #endregion

                #region Slides loop - processing all but last slide

                for (var i = 0; i < presentation.Slides.Count - 1; i++)
                {
                    #region Align Image (if single) - and not in place

                    if (presentation.Slides[i].PageElements.Count == 4)
                    {
                        //A template slide contains 3 page elements: header text box, footer text box, slide id text box
                        int mainImageIndex = -1;
                        for (var k = 0; k < presentation.Slides[i].PageElements.Count; k++)
                        {
                            if (presentation.Slides[i].PageElements[k].Image != null)
                            {
                                mainImageIndex = k;
                                break;
                            }
                        }
                        if (mainImageIndex >= 0)
                        {
                            myBatchRequest.AddUpdatePageElementTransformRequest(presentation.Slides[i].PageElements[mainImageIndex], alignImage);
                        }
                    }

                    #endregion

                    #region Process Text Boxes: Header/Footer/Slide Id

                    var parsedSlideTextElements = ParseSlideTextElements(presentation.Slides[i]);

                    desiredHeader = null;
                    if (parsedSlideTextElements.Header != null)
                    {
                        desiredHeader = parsedSlideTextElements.Header.Text;
                        if (!IsTextElementValid(parsedSlideTextElements.Header, parsedSlideTextElements.Header.Text, slideHeaderTransform, null))
                        {
                            //Delete the invalid object and create a new one
                            myBatchRequest.AddDeleteObjectRequest(parsedSlideTextElements.Header.ObjectId);
                            myBatchRequest.AddCreateShapeRequest(presentation.Slides[i].ObjectId, slideHeaderSize, slideHeaderTransform);
                            slideHeaderCreated = true;
                        }
                        else if (string.IsNullOrEmpty(parsedSlideTextElements.Header.Text))
                        {
                            //Header contains empty text
                            PresentationError.Invoke(this, new SlideErrorEventArgs(new SlideError(presentation.PresentationId, presentation.Title, i + 1, "No Header")));
                        }
                        else if (parsedSlideTextElements.Header.Text.Contains(lookForTextInHeader))
                        {
                            //Header contains forbidden text
                            PresentationError.Invoke(this, new SlideErrorEventArgs(new SlideError(presentation.PresentationId, presentation.Title, i + 1, "Header contains " + lookForTextInHeader)));
                        }
                    }
                    else
                    {
                        PresentationError.Invoke(this, new SlideErrorEventArgs(new SlideError(presentation.PresentationId, presentation.Title, i + 1, "No Header")));
                    }

                    //Footer - exclude last 2 slides (homework + toc)
                    if (i < presentation.Slides.Count - 2)
                    {
                        if (parsedSlideTextElements.Footer != null)
                        {
                            if (!IsTextElementValid(parsedSlideTextElements.Footer, desiredFooter, slideFooterTransform, null))
                            {
                                //Delete the invalid object and create a new one
                                myBatchRequest.AddDeleteObjectRequest(parsedSlideTextElements.Footer.ObjectId);
                                myBatchRequest.AddCreateShapeRequest(presentation.Slides[i].ObjectId, slideFooterSize, slideFooterTransform);
                                slideFooterCreated = true;
                            }
                        }
                        else
                        {
                            //Create a new text box to hold the slide footer
                            myBatchRequest.AddCreateShapeRequest(presentation.Slides[i].ObjectId, slideFooterSize, slideFooterTransform);
                            slideFooterCreated = true;
                        }
                    }

                    string desiredPageId = (i + 1).ToString();

                    //Page Id
                    if (parsedSlideTextElements.SlidePageId != null)
                    {
                        if (!IsTextElementValid(parsedSlideTextElements.SlidePageId, desiredPageId, slidePageIdTransform, nextSlidelink))
                        {
                            //Delete the invalid object and create a new one
                            myBatchRequest.AddDeleteObjectRequest(parsedSlideTextElements.SlidePageId.ObjectId);
                            myBatchRequest.AddCreateShapeRequest(presentation.Slides[i].ObjectId, slidePageIdSize, slidePageIdTransform);
                            slidePageIdCreated = true;
                        }
                    }
                    else
                    {
                        //Page Id text box does not exist - create a new text box to hold the slide number
                        myBatchRequest.AddCreateShapeRequest(presentation.Slides[i].ObjectId, slidePageIdSize, slidePageIdTransform);
                        slidePageIdCreated = true;
                    }

                    var batchResponse = myBatchRequest.Execute();
                    myBatchRequest.ClearRequests();

                    var addTextBoxesBatchRequest = new MyBatchRequest(slidesService, cachePresentation.PresentationId);
                    var textBoxesAdded = false;
                    var repliesCount = 0;
                    if (batchResponse != null)
                    {
                        repliesCount = batchResponse.Replies.Count;
                    }
                    for (var k = 0; k < repliesCount; k++)
                    {
                        if (batchResponse.Replies[k].CreateShape != null)
                        {
                            if (slideHeaderCreated)
                            {
                                addTextBoxesBatchRequest.AddInsertTextRequest(batchResponse.Replies[k].CreateShape.ObjectId, desiredHeader, 0);
                                addTextBoxesBatchRequest.AddUpdateTextStyleRequest(batchResponse.Replies[k].CreateShape.ObjectId, "SlideHeaderTextBoxTextStyle", slideHeaderTextBoxTextStyleFields, 0, desiredHeader.Length, null);
                                addTextBoxesBatchRequest.AddUpdateParagraphStyleRequest(batchResponse.Replies[k].CreateShape.ObjectId, false);
                                slideHeaderCreated = false;
                                textBoxesAdded = true;
                                continue; //To the next CreateShape reply
                            }
                            if (slideFooterCreated)
                            {
                                addTextBoxesBatchRequest.AddInsertTextRequest(batchResponse.Replies[k].CreateShape.ObjectId, desiredFooter, 0);
                                addTextBoxesBatchRequest.AddUpdateTextStyleRequest(batchResponse.Replies[k].CreateShape.ObjectId, "SlideFooterTextBoxTextStyle", slideFooterTextBoxTextStyleFields, 0, desiredFooter.Length, null);
                                addTextBoxesBatchRequest.AddUpdateParagraphStyleRequest(batchResponse.Replies[k].CreateShape.ObjectId, false);
                                slideFooterCreated = false;
                                textBoxesAdded = true;
                                continue; //To the next CreateShape reply
                            }
                            if (slidePageIdCreated)
                            {
                                addTextBoxesBatchRequest.AddInsertTextRequest(batchResponse.Replies[k].CreateShape.ObjectId, desiredPageId, 0);
                                addTextBoxesBatchRequest.AddUpdateTextStyleRequest(batchResponse.Replies[k].CreateShape.ObjectId, "SlideIdTextBoxTextStyle", slideIdTextBoxTextStyleFields, 0, desiredPageId.Length, nextSlidelink, false);
                                addTextBoxesBatchRequest.AddUpdateParagraphStyleRequest(batchResponse.Replies[k].CreateShape.ObjectId, false);
                                slidePageIdCreated = false;
                                textBoxesAdded = true;
                                continue; //To the next CreateShape reply
                            }
                        }
                    }
                    if (textBoxesAdded)
                    {
                        //Execute the requests to edit the text boxes created
                        addTextBoxesBatchRequest.Execute();
                    }

                    #endregion
                }

                #endregion

                #region Process Last Slide (TOC)

                var createTOCBatchRequest = new MyBatchRequest(slidesService, cachePresentation.PresentationId);
                var lastSlideNotesPage = presentation.Slides[presentation.Slides.Count - 1].SlideProperties.NotesPage;

                //Check if the last slide contains the TOC
                //Number f text elements should be 2n+1 where n=number of slides (excluding the last) and extra element for the paragraph style
                if (lastSlideNotesPage.PageElements.Count != 2 ||
                    lastSlideNotesPage.PageElements[1].Shape == null ||
                    lastSlideNotesPage.PageElements[1].Shape.Text == null ||
                    lastSlideNotesPage.PageElements[1].Shape.Text.TextElements == null ||
                    lastSlideNotesPage.PageElements[1].Shape.Text.TextElements.Count != 2 * (presentation.Slides.Count - 1) + 1)
                {
                    objectId = presentation.Slides[presentation.Slides.Count - 1].SlideProperties.NotesPage.PageElements[1].ObjectId;

                    createTOCBatchRequest.AddDeleteTextRequest(objectId, presentation.Slides[presentation.Slides.Count - 1].SlideProperties.NotesPage.PageElements[1].Shape);

                    currentStartIndex = 0;
                    string currentPageIdString;
                    for (var i = 1; i <= presentation.Slides.Count - 1; i++)
                    {
                        var link = new Google.Apis.Slides.v1.Data.Link()
                        {
                            SlideIndex = i - 1
                        };
                        currentPageIdString = (i).ToString("00") + "\t";
                        createTOCBatchRequest.AddInsertTextRequest(objectId, currentPageIdString, currentStartIndex);
                        //Link - will not contain the tab ("\t")
                        createTOCBatchRequest.AddUpdateTextStyleRequest(objectId, "SpeakerNotesTextStyle", speakerNotesTextStyleFields, currentStartIndex, currentStartIndex + currentPageIdString.Length - 1, link);
                        currentStartIndex += currentPageIdString.Length;
                    }
                    createTOCBatchRequest.AddUpdateParagraphStyleRequest(objectId, true);
                    createTOCBatchRequest.Execute();
                }
                #endregion

                #region Validate presentation animations using download and local inspect via PowerPoint
                ValidatePresentationAnimations(cachePresentation);
                #endregion

                #region Mark as processed in app properties

                MarkPresentationAsProcessed(presentationFile);

                #endregion

                #region Raise events

                PresentationProcessed.Invoke(this, null);

                #endregion
            }
            catch (Exception e)
            {
                PresentationError.Invoke(this, new SlideErrorEventArgs(new SlideError(cachePresentation.PresentationId, cachePresentation.PresentationName, 0, e.StackTrace)));
                throw (e);
            }
        }

        /// <summary>
        /// Process students presentations
        /// </summary>
        public void ProcessStudentsPresentations(string specificSheet = null, bool skipTimeStampCheck = false)
        {
            //Open masterplan spreadsheet
            var masterPlanSpreadsheetRequest = sheetService.Spreadsheets.Get(masterPlanSpreadsheetId);
            var masterPlanSpreadsheet = masterPlanSpreadsheetRequest.Execute();

            //Loop through all non-hidden sheets in the masterplan spreadsheet
            foreach (var sheet in masterPlanSpreadsheet.Sheets)
            {
                if (
                    (sheet.Properties.Hidden.HasValue && sheet.Properties.Hidden.Value == true) ||
                    (specificSheet != null && sheet.Properties.Title != specificSheet)
                    )
                {
                    continue; //To the next sheet
                }

                //Load sheet data
                var masterPlanSheetReadValuesRequest = sheetService.Spreadsheets.Values.Get(masterPlanSpreadsheetId, string.Format(masterPlanSpreadsheetReadRangePattern, sheet.Properties.Title));
                masterPlanSheetReadValuesRequest.ValueRenderOption = SpreadsheetsResource.ValuesResource.GetRequest.ValueRenderOptionEnum.FORMULA;
                var masterPlanSheetData = masterPlanSheetReadValuesRequest.Execute();

                //Raise event so parent can print sheet summary
                FolderProcessingStarted.Invoke(this, new ProcessFolderEventArgs(sheet.Properties.Title, masterPlanSheetData.Values.Count));

                //Skip the header line in the sheet
                int rowNumber = 2;

                //Variables to allow zeroing counter of files in a folder
                var prevMainFolderId = string.Empty;
                var prevSubFolderId = string.Empty;
                var currentFileNumberInFolder = 1;

                //Loop through the rows - each row represents a presentation to process
                foreach (var row in masterPlanSheetData.Values)
                {
                    #region Parse columns

                    //Check if this row is only a placeholder for a master presentation to come
                    if (row.Count <= 1)
                    {
                        rowNumber++;
                        continue;
                    }
                    var mainFolderId = regexSpreadsheetHyperlinkExtractId.Match(row[0].ToString()).Groups[1].Value;
                    var mainFolderName = ReplaceDoubleQuotesWithSingle(regexSpreadsheetHyperlinkExtractName.Match(row[0].ToString()).Groups[1].Value);
                    var subFolderId = regexSpreadsheetHyperlinkExtractId.Match(row[1].ToString()).Groups[1].Value;
                    var subFolderName = ReplaceDoubleQuotesWithSingle(regexSpreadsheetHyperlinkExtractName.Match(row[1].ToString()).Groups[1].Value);
                    var sourcePresentationId = regexSpreadsheetHyperlinkExtractId.Match(row[3].ToString()).Groups[1].Value;
                    var sourcePresentationName = regexSpreadsheetHyperlinkExtractName.Match(row[3].ToString()).Groups[1].Value;
                    Google.Apis.Drive.v3.Data.File targetPresentationDriveFile = null;
                    string targetPresentationId = null;

                    if (row[4] != null)
                    {
                        var match = regexSpreadsheetHyperlinkExtractId.Match(row[4].ToString());
                        if (match.Groups != null && match.Groups.Count > 1)
                        {
                            targetPresentationId = ReplaceDoubleQuotesWithSingle(match.Groups[1].Value);
                        }
                    }

                    //Check if main folder or sub folder has changed - start file counter
                    if (mainFolderId != prevMainFolderId || subFolderId != prevSubFolderId)
                    {
                        currentFileNumberInFolder = 1;
                    }
                    prevMainFolderId = mainFolderId;
                    prevSubFolderId = subFolderId;

                    var targetPresentationName = string.Format("{0:00}. {1}", currentFileNumberInFolder, sourcePresentationName);
                    //Replace double quotes to single quotes
                    targetPresentationName = targetPresentationName.Replace("\"\"", "\"");

                    #endregion

                    #region Check time stamps

                    //If the target presentation exists - check if source.modifiedDate > target.appProperties.NormalizeTime
                    //and if not - skip updating

                    if (targetPresentationId != null)
                    {
                        currentFileNumberInFolder++;
                        var sourcePresentationDriveFile = LoadPresentationForTimeStampCheck(sourcePresentationId);
                        targetPresentationDriveFile = LoadPresentationForTimeStampCheck(targetPresentationId);

                        if (targetPresentationDriveFile.Name != targetPresentationName)
                        {
                            RenamePresentation(targetPresentationDriveFile, targetPresentationName);
                        }

                        if (targetPresentationDriveFile.AppProperties == null)
                        {
                            targetPresentationDriveFile.AppProperties = new Dictionary<string, string>();
                        }
                        if (targetPresentationDriveFile.AppProperties.ContainsKey(APP_PROPERTY_NORMALIZE_TIME))
                        {
                            if (!skipTimeStampCheck && sourcePresentationDriveFile.ModifiedTimeDateTimeOffset <= Convert.ToDateTime(targetPresentationDriveFile.AppProperties[APP_PROPERTY_NORMALIZE_TIME],dateTimeCulture))
                            {
                                //Source file was not modified since the target was last processed - skip
                                PresentationSkipped.Invoke(this, null);
                                rowNumber++;
                                continue;
                            }
                        }
                    }

#endregion

#region Check if and which slides are to be copied

                    //Get the teachers presentation
                    if (sourcePresentationId == null || sourcePresentationId == string.Empty)
                    {
                        //Source file was not modified since the target was last processed - skip
                        PresentationSkipped.Invoke(this, new SlideSkippedEventArgs(new SlideSkipped(null,null,0,rowNumber,"Empty presentationId")));
                        rowNumber++;
                        continue;
                    }

                    var sourcePresentationRequest = slidesService.Presentations.Get(sourcePresentationId);
                    var sourcePresentation = sourcePresentationRequest.Execute();

                    //Check which slides to copy
                    var slidesToCopy = new List<int>();
                    var pageElementIndexList = new List<int>();
                    for (var i = 0; i < sourcePresentation.Slides.Count; i++)
                    {
                        var parseSlideTextElements = ParseSlideTextElements(sourcePresentation.Slides[i]);
                        if (parseSlideTextElements.Header != null && parseSlideTextElements.Header.Text != null && !regexNonTeachingSlide.IsMatch(parseSlideTextElements.Header.Text))
                        {
                            //Teaching slide
                            slidesToCopy.Add(i);
                            pageElementIndexList.Add(parseSlideTextElements.SlidePageId.PageElementIndex);
                        }
                    }

                    if (slidesToCopy.Count == 0)
                    {
                        //Skip this row - no slides to copy
                        rowNumber++;
                        PresentationSkipped.Invoke(this, null);
                        continue; //To the next row
                    }

#endregion

#region Create Folders if neccessary

                    var studentsMainFolder = CheckToCreateDriveFolder(StudentsCache.GetSubFolderByName(sheet.Properties.Title), mainFolderName);

                    //Sometimes presentations are filed directly under a main folder without a subfolder
                    CacheFolder studentsSubFolder;
                    if (subFolderId != mainFolderId)
                    {
                        studentsSubFolder = CheckToCreateDriveFolder(studentsMainFolder, subFolderName);
                    }
                    else
                    {
                        studentsSubFolder = studentsMainFolder;
                    }

#endregion

#region Create target presentation if neccessary

                    if (targetPresentationId == null)
                    {
                        //Create an empty presentation (created with 1 blank new slide)

                        //Presentation names in the spreadsheet do not contain 
                        //preceding numbers as the actual drive files are stored
                        //In google Drive, so use the original "Title" of the presentation object
                        //For the name of the file (to preserve the number in the name)
                        var targetPresentation = CreateEmptyPresentation(studentsSubFolder, targetPresentationName);

                        targetPresentationId = targetPresentation.PresentationId;

                        //Update the spreadsheet with the new presentation id
                        UpdateSheetHyperlinkCell(masterPlanSpreadsheetId, sheet.Properties.Title, rowNumber, targetPresentationId);

                        currentFileNumberInFolder++;
                    }

#endregion

#region Invoke App Script to copy the slide

                    //Unfortunatelly Presentation.appendSlide method (with linking option) exists only in App Scripts
                    var scriptExecutionBody = new ExecutionRequest
                    {
                        Parameters = new List<Object>() {
                            sourcePresentationId,
                            targetPresentationId,
                            targetPresentationName,
                            slidesToCopy.ToArray(),
                            pageElementIndexList.ToArray()
                        },
                        Function = copySlidesAppScriptFunction
                    };

                    var scriptRunRequest = scriptService.Scripts.Run(scriptExecutionBody, copySlidesAppScriptId);
                    var scriptResult = scriptRunRequest.Execute();

                    if (scriptResult.Response != null && scriptResult.Response.ContainsKey("result") && scriptResult.Response["result"].ToString() == "0")
                    {
                        PresentationProcessed.Invoke(this, null);
                    }
                    else
                    {
                        PresentationError.Invoke(this, new SlideErrorEventArgs(new SlideError(targetPresentationId, string.Empty, 0, scriptResult.Response.ToString())));
                    }

#endregion

#region Mark presentation as processed

                    if (targetPresentationDriveFile == null)
                    {
                        targetPresentationDriveFile = LoadPresentationForTimeStampCheck(targetPresentationId);
                    }
                    MarkPresentationAsProcessed(targetPresentationDriveFile);

#endregion

                    rowNumber++;
                }
            }

            //During processing, Students cache might have been updated - save it locally
            StudentsCache.Save();

        }

#endregion

#region Events

        public event EventHandler PresentationProcessed;
        public event EventHandler PresentationSkipped;
        public event EventHandler PresentationError;
        public event EventHandler FolderProcessingStarted;

#endregion

#region Private Methods

        /// <summary>
        /// Exports and Download the presentation as a Microsoft Powerpoint file (.pptx) and validates each slide:
        /// if a slide contains at least one image object (without a border, e.g. a question) - 
        /// it must contain animations if there are text boxes or lines/arrows
        /// </summary>
        /// <param name="cachePresentation"></param>
        private void ValidatePresentationAnimations(CachePresentation cachePresentation)
        {
            bool slideHasStudentQuestion;
            bool slideHasSolution;

#region Download presentation as powerpoint

            var exportRequest = driveService.Files.Export(cachePresentation.PresentationId, msPowerPointMimeType);
            var driveStream = exportRequest.ExecuteAsStream();
            var exportPath = Path.Combine(Path.GetTempPath(), msPowerPointTempLocalFileName);
            var fileStream = new FileStream(exportPath, FileMode.CreateNew);
            driveStream.CopyTo(fileStream);
            fileStream.Close();

#endregion

            var localPresentation = pptPresentations.Open(exportPath, MsoTriState.msoTrue, MsoTriState.msoFalse, MsoTriState.msoFalse);
            for (var i = 1; i <= localPresentation.Slides.Count; i++)
            {
                var slide = localPresentation.Slides[i];
                slideHasStudentQuestion = false;
                slideHasSolution = false;
                for (var j = 1; j <= slide.Shapes.Count; j++)
                {
                    var shape = slide.Shapes[j];
                    if (shape.Type == MsoShapeType.msoPicture && shape.Line.ForeColor.RGB == whiteColor && slide.TimeLine.MainSequence.Count == 0)
                    {
                        slideHasStudentQuestion = true;
                    }
                    else if (shape.Type == MsoShapeType.msoLine ||
                             (shape.Type == MsoShapeType.msoAutoShape && shape.Height < msPowerPointAutoShapeMaxHeight) ||
                             shape.Type == MsoShapeType.msoFreeform ||
                             shape.Type == MsoShapeType.msoTable ||
                             (shape.Type == MsoShapeType.msoTextBox && shape.Top < msPowerPointPageNumberMinTop))
                    {
                        slideHasSolution = true;
                    }
                }
                if (slideHasStudentQuestion && slideHasSolution)
                {
                    PresentationError.Invoke(this, new SlideErrorEventArgs(new SlideError(cachePresentation.PresentationId, cachePresentation.PresentationName, i, "Should have animations")));
                }
            }
            localPresentation.Close();

            System.IO.File.Delete(exportPath);
        }

        /// <summary>
        /// Retrives text from a shape page element
        /// </summary>
        /// <param name="shape"></param>
        /// <returns></returns>
        private string GetTextFromShape(Google.Apis.Slides.v1.Data.Shape shape)
        {
            if (shape == null ||
                shape.Text == null ||
                shape.Text.TextElements == null)
            {
                return null;
            }
            string text = "";
            for(var i=0; i < shape.Text.TextElements.Count; i++)
            {
                if (shape.Text.TextElements[i].TextRun != null &&
                    shape.Text.TextElements[i].TextRun.Content != null)
                {
                    text += shape.Text.TextElements[i].TextRun.Content;
                }
            }
            return text;
        }

        /// <summary>
        /// Retrives text from a shape page element
        /// </summary>
        /// <param name="shape"></param>
        /// <returns></returns>
        private Link GetLinkFromShape(Google.Apis.Slides.v1.Data.Shape shape)
        {
            if (shape != null &&
                shape.Text != null &&
                shape.Text.TextElements != null &&
                shape.Text.TextElements.Count > 1 &&
                shape.Text.TextElements[1].TextRun != null &&
                shape.Text.TextElements[1].TextRun.Style != null &&
                shape.Text.TextElements[1].TextRun.Style.Link != null
                )
            {
                return shape.Text.TextElements[1].TextRun.Style.Link;
            }

            return null;
        }

        /// <summary>
        /// Parses the slide and looks for the Header/Footer/PageId text boxes
        /// </summary>
        /// <param name="slide"></param>
        /// <returns></returns>
        private SlideParsedTextElements ParseSlideTextElements(Page slide)
        {
            var slideParsedTextElements = new SlideParsedTextElements();

            for (var j = 0; j < slide.PageElements.Count; j++)
            {
                if (slide.PageElements[j].Shape != null &&
                    slide.PageElements[j].Shape.ShapeType == "TEXT_BOX")
                {
                    //Found the header object
                    if (slide.PageElements[j].Transform.ScaleX == slideHeaderTransform.ScaleX &&
                        slide.PageElements[j].Transform.ScaleY == slideHeaderTransform.ScaleY)
                    {
                        slideParsedTextElements.Header = new SlideParsedTextElement(slide.PageElements[j].ObjectId, j, GetTextFromShape(slide.PageElements[j].Shape), slide.PageElements[j].Transform, slide.PageElements[j].Shape);
                    }

                    //Found the footer object
                    else if (slide.PageElements[j].Transform.ScaleX == slideFooterTransform.ScaleX &&
                             slide.PageElements[j].Transform.ScaleY == slideFooterTransform.ScaleY)
                    {
                        slideParsedTextElements.Footer = new SlideParsedTextElement(slide.PageElements[j].ObjectId, j, GetTextFromShape(slide.PageElements[j].Shape), slide.PageElements[j].Transform, slide.PageElements[j].Shape);
                    }

                    //Found the page id object
                    else if (slide.PageElements[j].Transform.ScaleX == slidePageIdTransform.ScaleX &&
                             slide.PageElements[j].Transform.ScaleY == slidePageIdTransform.ScaleY)

                    {
                        slideParsedTextElements.SlidePageId = new SlideParsedTextElement(slide.PageElements[j].ObjectId, j, GetTextFromShape(slide.PageElements[j].Shape), slide.PageElements[j].Transform, slide.PageElements[j].Shape);
                    }
                }
            }

            return slideParsedTextElements;
        }

        /// <summary>
        /// Validates a text box to check:
        /// 1. Is located in the designated place
        /// 2. Contains the correct text
        /// 3. Contains a desired link
        /// </summary>
        /// <param name="slideParsedTextElement"></param>
        /// <returns></returns>
        private bool IsTextElementValid(SlideParsedTextElement slideParsedTextElement, string desiredText, AffineTransform desiredTransform, Link desiredLink)
        {
            if (slideParsedTextElement.Transform.TranslateX != desiredTransform.TranslateX ||
                slideParsedTextElement.Transform.TranslateY != desiredTransform.TranslateY ||
                slideParsedTextElement.Text != desiredText ||
                desiredLink != GetLinkFromShape(slideParsedTextElement.Shape))
            {
                return false;
            }

            return true;
        }

        /// <summary>
        /// Check if a sub folder under a parent exists in cache by its name and create it in drive and in cache if neccessary
        /// </summary>
        /// <param name="parent"></param>
        /// <param name="folderName"></param>
        /// <returns></returns>
        private CacheFolder CheckToCreateDriveFolder(CacheFolder parent, string folderName)
        {
            //Check to create drive hierarchy in students folders
            var childFolder = parent.GetSubFolderByName(folderName);
            if (childFolder != null)
            {
                //Exists - return it
                return childFolder;
            }

            //Create drive folder
            var newFolder = new Google.Apis.Drive.v3.Data.File
            {
                MimeType = StudentsCache.GetFolderMimeType(),
                Name = folderName,
                Parents = new List<string>
                {
                    parent.FolderId
                }
            };

            var newDriveFolderRequest = driveService.Files.Create(newFolder);
            var newDriveFolder = newDriveFolderRequest.Execute();

            //Add the newly folder to the cache and return it
            var newCacheFolder = new CacheFolder(newDriveFolder.Id, folderName, parent.FolderId);
            parent.Folders.Add(newDriveFolder.Id, newCacheFolder);

            return newCacheFolder;
        }

        /// <summary>
        /// Create an empty presentation in drive under the given parent and update cache
        /// </summary>
        /// <param name="parent"></param>
        /// <param name="presentationName"></param>
        /// <returns></returns>
        private CachePresentation CreateEmptyPresentation(CacheFolder parent, string presentationName)
        {
            //
            var newPresentation = new Google.Apis.Drive.v3.Data.File
            {
                MimeType = StudentsCache.GetPresentationMimeType(),
                Name = presentationName,
                Parents = new List<string>
                {
                    parent.FolderId
                }
            };

            var newDriveFileRequest = driveService.Files.Create(newPresentation);
            var newDriveFile = newDriveFileRequest.Execute();

            //Add the presentation to the cache
            return parent.AddPresentation(newDriveFile.Id, presentationName);

        }

        /// <summary>
        /// Update sheet's cell to as a hyper link to a presentation
        /// </summary>
        /// <param name="spreadsheetId"></param>
        /// <param name="sheetName"></param>
        /// <param name="row"></param>
        /// <param name="presentationId"></param>
        private void UpdateSheetHyperlinkCell(string spreadsheetId, string sheetName, int row, string presentationId)
        {
            var valueRange = new ValueRange();
            var presentationHyperlink = new List<object>
                        {
                            string.Format(spreadsheetHyperlinkFormat, presentationId)
                        };

            valueRange.Values = new List<IList<Object>> { presentationHyperlink };
            var updateSheetRequest = sheetService.Spreadsheets.Values.Update(valueRange, spreadsheetId, string.Format(masterPlanSpreadsheetUpdateRangePattern, sheetName, row));
            updateSheetRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.USERENTERED;
            valueRange.MajorDimension = "ROWS";
            updateSheetRequest.Execute();
        }

        /// <summary>
        /// Loads a presentation as a drive file to compare time stamps
        /// </summary>
        /// <param name="presentationId"></param>
        /// <returns></returns>
        private Google.Apis.Drive.v3.Data.File LoadPresentationForTimeStampCheck(string presentationId)
        {
            var presentationFileRequest = driveService.Files.Get(presentationId);
            presentationFileRequest.Fields = "id, name, appProperties, modifiedTime";
            return presentationFileRequest.Execute();
        }

        /// <summary>
        /// Mark (with a time stamp) in drive that this presentation has been processed
        /// </summary>
        /// <param name="presentationFile"></param>
        private void MarkPresentationAsProcessed(Google.Apis.Drive.v3.Data.File presentationFile)
        {
               //Make sure normalize time is always after the modified time (15 seconds after)
               var normalizeTime = DateTime.Now.AddSeconds(15).ToString();
            if (presentationFile.AppProperties == null)
            {
                presentationFile.AppProperties = new Dictionary<string, string>();
            }
            if (!presentationFile.AppProperties.ContainsKey(APP_PROPERTY_NORMALIZE_TIME))
            {
                presentationFile.AppProperties.Add(APP_PROPERTY_NORMALIZE_TIME, normalizeTime);
            }
            else
            {
                presentationFile.AppProperties[APP_PROPERTY_NORMALIZE_TIME] = normalizeTime;
            }


            var updatePresentationFileRequest = driveService.Files.Update(presentationFile, presentationFile.Id);
            updatePresentationFileRequest.Fields = "appProperties";

            var id = presentationFile.Id;
            //Field cannot be updated - reset it before update - and the return the value back
            presentationFile.Id = null;

            updatePresentationFileRequest.Execute();

            presentationFile.Id = id;

        }

        /// <summary>
        /// Rename presentation
        /// </summary>
        /// <param name="presentationFile"></param>
        private void RenamePresentation(Google.Apis.Drive.v3.Data.File presentationFile, string newName)
        {
            var updatePresentationFileRequest = driveService.Files.Update(presentationFile, presentationFile.Id);
            presentationFile.Name = newName;

            var id = presentationFile.Id;

            //Field cannot be updated - reset it before update - and the return the value back
            presentationFile.Id = null;

            updatePresentationFileRequest.Fields = "name";
            updatePresentationFileRequest.Execute();

            //Restore the id
            presentationFile.Id = id;
        }

        /// <summary>
        /// Replaces double quotes with single quotes
        /// </summary>
        /// <param name="inputString"></param>
        /// <returns></returns>
        private string ReplaceDoubleQuotesWithSingle(string inputString)
        {
            if (inputString != null)
            {
                return inputString.Replace("\"\"", "\"");
            }

            return null;
        }

#endregion
    }

#endregion

    #region Class MyBatchRequest

    public class MyBatchRequest
    {
        #region Class Members

        SlidesService slidesService;
        BatchUpdatePresentationRequest batchUpdatePresentationRequest;
        readonly string presentationId;

        #endregion

        #region C'Tor/D'Tor
        public MyBatchRequest(SlidesService slidesService, string presentationId)
        {
            this.slidesService = slidesService;
            batchUpdatePresentationRequest = new BatchUpdatePresentationRequest
            {
                Requests = new List<Google.Apis.Slides.v1.Data.Request>()
            };
            this.presentationId = presentationId;
        }
        #endregion

        #region Methods
        /// <summary>
        /// Adds a CreateSlide request - for the "empty board" slide in the end
        /// </summary>
        public void AddCreateSlideRequest(int insertionIndex)
        {
            batchUpdatePresentationRequest.Requests.Add(new Google.Apis.Slides.v1.Data.Request()
            {
                CreateSlide = new CreateSlideRequest()
                {
                    SlideLayoutReference = new LayoutReference
                    {
                        LayoutId = ConfigurationManager.AppSettings["LayoutObjectId"]
                    },
                    InsertionIndex = insertionIndex
                }
            });
        }

        /// <summary>
        /// Adds a DeleteText request to an object - to delete its entire text
        /// </summary>
        public void AddDeleteTextRequest(string objectId, Google.Apis.Slides.v1.Data.Shape shape)
        {
            if (shape.Text == null)
            {
                return;
            }

            batchUpdatePresentationRequest.Requests.Add(new Google.Apis.Slides.v1.Data.Request()
            {
                DeleteText = new DeleteTextRequest()
                {
                    ObjectId = objectId,
                    TextRange = new Range
                    {
                        Type = "ALL"
                    }
                }
            });
        }

        /// <summary>
        /// Adds an InsertText request to an object
        /// </summary>
        /// <param name="objectId"></param>
        public void AddInsertTextRequest(string objectId, string text, int insertionIndex)
        {
            batchUpdatePresentationRequest.Requests.Add(new Google.Apis.Slides.v1.Data.Request()
            {
                InsertText = new InsertTextRequest()
                {
                    ObjectId = objectId,
                    Text = text,
                    InsertionIndex = insertionIndex
                }
            });
        }

        /// <summary>
        /// Adds an update text style request to a specific text with optional linking
        /// </summary>
        /// <param name="startIndex"></param>
        /// <param name="endIndex"></param>
        /// <param name="link"></param>
        public void AddUpdateTextStyleRequest(string objectId, string textStyleConfigKey, string textStyleFields, int startIndex, int endIndex, Google.Apis.Slides.v1.Data.Link link = null, bool underline = true)
        {
            var fields = String.Copy(textStyleFields);
            var textStyle = JsonConvert.DeserializeObject<Google.Apis.Slides.v1.Data.TextStyle>(ConfigurationManager.AppSettings[textStyleConfigKey]);

            if (link != null)
            {
                textStyle.Link = link;
                fields += ",link";
            }
            if (!underline)
            {
                textStyle.Underline = false;
                fields += ",underline";
            }

            batchUpdatePresentationRequest.Requests.Add(new Google.Apis.Slides.v1.Data.Request()
            {
                UpdateTextStyle = new UpdateTextStyleRequest()
                {
                    ObjectId = objectId,
                    Style = textStyle,
                    TextRange = new Range()
                    {
                        Type = "FIXED_RANGE",
                        StartIndex = startIndex,
                        EndIndex = endIndex
                    },
                    Fields = fields
                }
            });
        }

        /// <summary>
        /// Updates the entire text of the speaker notes with a paragraph style defined in config
        /// </summary>
        /// <param name="rtl">2 separate paragraph styles in config for ltr, rtl</param>
        public void AddUpdateParagraphStyleRequest(string objectId, bool rtl)
        {
            ParagraphStyle paragraphStyle;
            if (!rtl)
            {
                paragraphStyle = JsonConvert.DeserializeObject<ParagraphStyle>(ConfigurationManager.AppSettings["ParagraphStyleLTR"]);
            }
            else
            {
                paragraphStyle = JsonConvert.DeserializeObject<ParagraphStyle>(ConfigurationManager.AppSettings["ParagraphStyleRTL"]);
            }
            var fields = ConfigurationManager.AppSettings["ParagraphStyleFields"];

            batchUpdatePresentationRequest.Requests.Add(new Google.Apis.Slides.v1.Data.Request()
            {
                UpdateParagraphStyle = new UpdateParagraphStyleRequest()
                {

                    ObjectId = objectId,
                    Style = paragraphStyle,
                    TextRange = new Range()
                    {
                        Type = "ALL"
                    },
                    Fields = fields
                }
            });
        }

        /// <summary>
        /// Create a text box to hold the slide number
        /// </summary>
        /// <param name="pageObjectId"></param>
        /// <param name="size"></param>
        /// <param name="transform"></param>
        public void AddCreateShapeRequest(string pageObjectId, Size size, AffineTransform transform)
        {
            batchUpdatePresentationRequest.Requests.Add(new Google.Apis.Slides.v1.Data.Request()
            {
                CreateShape = new CreateShapeRequest()
                {
                    ShapeType = "TEXT_BOX",
                    ElementProperties = new PageElementProperties()
                    {
                        PageObjectId = pageObjectId,
                        Size = size,
                        Transform = transform
                    }
                }
            });
        }

        /// <summary>
        /// Align image to top/bottom
        /// </summary>
        /// <param name="image"></param>
        /// <param name="alignImage"></param>
        public void AddUpdatePageElementTransformRequest(PageElement pageElement, AlignImage alignImage)
        {
            double? newYPosition = 0;
            switch (alignImage)
            {
                case AlignImage.TOP:
                    newYPosition = Convert.ToDouble(ConfigurationManager.AppSettings["ImageAlignTopPosition"]);
                    break;

                case AlignImage.BOTTOM:
                    newYPosition = Convert.ToDouble(ConfigurationManager.AppSettings["ImageAlignBottomPosition"]) - (pageElement.Size.Height.Magnitude * pageElement.Transform.ScaleY);
                    break;
            }

            if (pageElement.Transform.TranslateY == newYPosition.Value)
            {
                //Already in place - nothing to move
                return;
            }

            batchUpdatePresentationRequest.Requests.Add(new Google.Apis.Slides.v1.Data.Request()
            {
                UpdatePageElementTransform = new UpdatePageElementTransformRequest()
                {
                    ObjectId = pageElement.ObjectId,
                    ApplyMode = "ABSOLUTE",
                    Transform = new AffineTransform()
                    {
                        ScaleX = pageElement.Transform.ScaleX,
                        ScaleY = pageElement.Transform.ScaleY,
                        TranslateX = pageElement.Transform.TranslateX,
                        TranslateY = newYPosition,
                        Unit = "EMU"
                    }
                }
            });
        }

        /// <summary>
        /// Deletes the specified object from the slide
        /// </summary>
        /// <param name="objectId"></param>
        public void AddDeleteObjectRequest(string objectId)
        {
            batchUpdatePresentationRequest.Requests.Add(new Google.Apis.Slides.v1.Data.Request()
            {
                DeleteObject = new DeleteObjectRequest()
                {
                    ObjectId = objectId
                }
            });
        }

        /// <summary>
        /// Replaces text globally in the entire presentation
        /// </summary>
        /// <param name="searchText"></param>
        /// <param name="replaceText"></param>
        public void AddReplaceAllTextRequest(string searchText, string replaceText)
        {
            batchUpdatePresentationRequest.Requests.Add(new Google.Apis.Slides.v1.Data.Request()
            {
                ReplaceAllText = new ReplaceAllTextRequest()
                {
                    ContainsText = new SubstringMatchCriteria()
                    {
                        Text = searchText,
                        MatchCase = false
                    },
                    ReplaceText = replaceText
                }
            });
        }

        /// <summary>
        /// Executes the requests added to the list
        /// </summary>
        /// <returns></returns>
        public BatchUpdatePresentationResponse Execute()
        {
            if (batchUpdatePresentationRequest.Requests != null && batchUpdatePresentationRequest.Requests.Count > 0)
            {
                var batchUpdateRequest = slidesService.Presentations.BatchUpdate(batchUpdatePresentationRequest, presentationId);
                var response = batchUpdateRequest.Execute();

                return response;
            }
            return null;
        }

        /// <summary>
        /// Clears the requests collection
        /// </summary>
        public void ClearRequests()
        {
            batchUpdatePresentationRequest.Requests = new List<Google.Apis.Slides.v1.Data.Request>();
        }
        #endregion
    }

    #endregion

    #region Class SlideParsedTextElement

    public class SlideParsedTextElement
    {
        #region Properties

        public string ObjectId { get; private set; }
        public int PageElementIndex { get; private set; }
        public string Text { get; private set; }
        public AffineTransform Transform { get; private set; }
        public Google.Apis.Slides.v1.Data.Shape Shape {get; private set; }

        #endregion

        #region C'Tor/Dtor
        public SlideParsedTextElement(string objectId, int pageElementIndex, string text, AffineTransform transform, Google.Apis.Slides.v1.Data.Shape shape)
        {
            ObjectId = objectId;
            PageElementIndex = pageElementIndex;
            Shape = shape;
            Text = text;
            Transform = transform;
        }
        #endregion
    }

    #endregion

    #region Class SlideParsedTextElements

    public class SlideParsedTextElements
    {
        #region Properties

        public SlideParsedTextElement Header { get; set; }
        public SlideParsedTextElement Footer { get; set; }
        public SlideParsedTextElement SlidePageId { get; set; }

        #endregion

        #region C'Tor/Dtor

        public SlideParsedTextElements()
        {
        }

        #endregion
    }

    #endregion

}