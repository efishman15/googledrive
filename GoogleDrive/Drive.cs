using Google.Apis.Auth.OAuth2;
using Google.Apis.Drive.v3;
using Google.Apis.Requests;
using Google.Apis.Services;
using Google.Apis.Slides.v1;
using Google.Apis.Slides.v1.Data;
using Google.Apis.Util.Store;
using Newtonsoft.Json;
using Newtonsoft.Json.Converters;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Threading;

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

        public void AddPresentation(string presentationId, string presentationName)
        {
            Presentations.Add(new CachePresentation(presentationId, presentationName));
            TotalPresentations++;
        }
        #endregion
    }
    #endregion

    #region Class Cache
    public class Cache
    {
        #region Properties
        public Dictionary<string, CacheFolder> Folders { get; private set; }
        public int TotalPresentations { get; set; }
        public string DateCreated { get; set; }
        #endregion

        #region C'Tor/Dtor
        public Cache()
        {
            Folders = new Dictionary<string, CacheFolder>();
        }
        #endregion

        #region Methods

        /// <summary>
        /// Adds a presentation to the folder in the tree
        /// </summary>
        /// <param name="folderId"></param>
        /// <param name="presentationId"></param>
        /// <param name="presentationName"></param>
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
        /// Get a folder in the tree by its folder id (recurssive)
        /// </summary>
        /// <param name="folderId"></param>
        /// <param name="folders"></param>
        /// <returns></returns>
        public CacheFolder GetFolder(string folderId, Dictionary<string, CacheFolder> folders)
        {
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

    #region Class Process Event Args

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

    #endregion

    #region Class Drive

    public class Drive
    {
        #region Class Members

        private DriveService driveService;
        private SlidesService slidesService;
        private static string[] Scopes = { DriveService.Scope.DriveReadonly, SlidesService.Scope.Presentations };
        private static string ApplicationName = "Google Drive";
        private JsonSerializer jsonSerializer;
        private string foldersFilter;
        private readonly int pathStartLevel;
        private readonly string pathSeparator;
        private readonly string folderNameSeparator;

        private readonly Link lastSlidelink;
        private readonly Link nextSlidelink;
        private readonly Link prevSlidelink;
        private readonly Link firstSlidelink;

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

        #endregion

        #region C'Tor/D'Tor
        public Drive()
        {
            #region Authentication to Google
            UserCredential credential;

            using (var stream =
                new FileStream("credentials.json", FileMode.Open, FileAccess.Read))
            {
                string credPath = "~/.credentials/token.json";
                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.Load(stream).Secrets,
                    Scopes,
                    ConfigurationManager.AppSettings["user"],
                    CancellationToken.None,
                    new FileDataStore(credPath, true)).Result;
            }

            // Create Drive, Slides API services.
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

            #endregion

            #region Try load presentations list from local cache

            jsonSerializer = new JsonSerializer();

            // deserialize JSON directly from a file
            if (File.Exists(ConfigurationManager.AppSettings["PresentationsListCache"]))
            {
                using (StreamReader file = File.OpenText(ConfigurationManager.AppSettings["PresentationsListCache"]))
                {
                    Cache = (Cache)jsonSerializer.Deserialize(file, typeof(Cache));
                }
            }
            else
            {
                Cache = new Cache();
            }

            pathStartLevel = Convert.ToInt32(ConfigurationManager.AppSettings["PathStartLevel"]);
            pathSeparator = ConfigurationManager.AppSettings["PathSeparator"];
            folderNameSeparator = ConfigurationManager.AppSettings["FolderNameSeparator"];

            #endregion

            #region Load Config Variables

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

            #endregion

        }
        #endregion

        #region Properties

        /// <summary>
        /// Returns the presentations list
        /// </summary>
        public Cache Cache { get; private set; }

        #endregion

        #region Methods

        /// <summary>
        /// Clears the cache
        /// </summary>
        public void ClearCache()
        {
            Cache = new Cache();
        }

        /// <summary>
        /// Build recursivelly a list of all presentations to work on
        /// </summary>
        /// <param name="rootFolderId"></param>
        public void BuildPresentationsList(string rootFolderId, bool isTop, CacheFolder parentFolder)
        {
            string filter = "'" + rootFolderId + "' in parents AND (mimeType = 'application/vnd.google-apps.folder') AND trashed=false";
            string pageToken = null;

            if (isTop)
            {
                foldersFilter = string.Empty;
            }
            do
            {
                var folderRequest = driveService.Files.List();
                folderRequest.Q = filter;
                folderRequest.Spaces = "drive";
                folderRequest.Fields = "nextPageToken, files(id)";
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

                    var newFolder = new CacheFolder(folder.Id, GetFolderName(folder.Id), parentFolder?.FolderId);
                    if (parentFolder == null)
                    {
                        newFolder.Level = 1;
                        Cache.Folders.Add(newFolder.FolderId, newFolder);
                    }
                    else
                    {
                        parentFolder.Folders.Add(newFolder.FolderId, newFolder);
                        newFolder.Level = parentFolder.Level + 1;
                    }

                    BuildPresentationsList(folder.Id, false, newFolder);
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
                filter = foldersFilter + " AND (mimeType = 'application/vnd.google-apps.presentation') AND trashed=false";
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
                        Cache.AddPresentationToFolder(file.Parents[0], file.Id, file.Name);
                    }
                    pageToken = fileResult.NextPageToken;
                } while (pageToken != null);
            }
        }

        /// <summary>
        /// Process the cache:
        /// Add path to folders starting at "PathStartLevel" in config
        /// </summary>
        public void BuildFoldersPath(Dictionary<string, CacheFolder> root, string parentPath)
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
                    foreach(var presentation in root[folderKey].Presentations)
                    {
                        presentation.FooterText = root[folderKey].Path;
                    }
                }
                BuildFoldersPath(root[folderKey].Folders, root[folderKey].Path);
            }
        }

        /// <summary>
        /// Save presentations list to local cache file
        /// </summary>
        public void SaveCache()
        {
            var outputFileName = ConfigurationManager.AppSettings["PresentationsListCache"];
            if (File.Exists(outputFileName))
            {
                File.Delete(outputFileName);
            }
            Cache.DateCreated = DateTime.Now.ToString();
            jsonSerializer.Converters.Add(new JavaScriptDateTimeConverter());
            jsonSerializer.NullValueHandling = NullValueHandling.Ignore;
            using (StreamWriter sw = new StreamWriter(outputFileName))
            using (JsonWriter writer = new JsonTextWriter(sw))
            {
                jsonSerializer.Serialize(writer, Cache);
            }

        }

        /// <summary>
        /// Gets folder name using folder id
        /// </summary>
        /// <param name="folderId"></param>
        /// <returns></returns>
        public string GetFolderName(string folderId)
        {
            var folderRequest = driveService.Files.Get(folderId);
            var folder = folderRequest.Execute();
            return folder.Name;
        }

        /// <summary>
        /// Process all the presentations in a root folder and its sub folders
        /// </summary>
        /// <param name="rootFolder"></param>
        public void ProcessFolderPresentations(CacheFolder rootFolder)
        {
            foreach (var cachePresentation in rootFolder.Presentations)
            {
                ProcessPresentation(cachePresentation);

            }
            //Process presentations in all subfolders
            foreach (var cachedFolderKey in rootFolder.Folders.Keys)
            {
                ProcessFolderPresentations(rootFolder.Folders[cachedFolderKey]);
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
        public void ProcessPresentation(CachePresentation cachePresentation)
        {
            try
            {
                #region Local variables

                string objectId;
                int currentStartIndex;

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
                    #region Spearker notes slide

                    currentStartIndex = 0;
                    var notesPage = presentation.Slides[i].SlideProperties.NotesPage;
                    if (notesPage.PageElements.Count != 2 ||
                        notesPage.PageElements[1].Shape == null ||
                        notesPage.PageElements[1].Shape.Text == null ||
                        notesPage.PageElements[1].Shape.Text.TextElements == null ||
                        notesPage.PageElements[1].Shape.Text.TextElements.Count != 9)
                    {
                        objectId = notesPage.PageElements[1].ObjectId;
                        myBatchRequest.AddDeleteTextRequest(objectId, notesPage.PageElements[1].Shape);

                        //Add buttons
                        myBatchRequest.AddInsertTextRequest(objectId, firstSlideText, currentStartIndex);
                        myBatchRequest.AddUpdateTextStyleRequest(objectId, "SpeakerNotesTextStyle", speakerNotesTextStyleFields, currentStartIndex, currentStartIndex + firstSlideText.Length - 1, firstSlidelink, false);
                        myBatchRequest.AddUpdateTextStyleRequest(objectId, "SpeakerNotesTextStyle", speakerNotesTextStyleFields, currentStartIndex + firstSlideText.Length - 1, currentStartIndex + firstSlideText.Length, null, false);
                        currentStartIndex += firstSlideText.Length;

                        //Prev
                        myBatchRequest.AddInsertTextRequest(objectId, prevSlideText, currentStartIndex);
                        myBatchRequest.AddUpdateTextStyleRequest(objectId, "SpeakerNotesTextStyle", speakerNotesTextStyleFields, currentStartIndex, currentStartIndex + prevSlideText.Length - 1, prevSlidelink, false);
                        myBatchRequest.AddUpdateTextStyleRequest(objectId, "SpeakerNotesTextStyle", speakerNotesTextStyleFields, currentStartIndex + prevSlideText.Length - 1, currentStartIndex + prevSlideText.Length, null, false);
                        currentStartIndex += prevSlideText.Length;

                        //Next
                        myBatchRequest.AddInsertTextRequest(objectId, nextSlideText, currentStartIndex);
                        myBatchRequest.AddUpdateTextStyleRequest(objectId, "SpeakerNotesTextStyle", speakerNotesTextStyleFields, currentStartIndex, currentStartIndex + nextSlideText.Length - 1, nextSlidelink, false);
                        myBatchRequest.AddUpdateTextStyleRequest(objectId, "SpeakerNotesTextStyle", speakerNotesTextStyleFields, currentStartIndex + nextSlideText.Length - 1, currentStartIndex + nextSlideText.Length, null, false);
                        currentStartIndex += nextSlideText.Length;

                        //Last
                        myBatchRequest.AddInsertTextRequest(objectId, lastSlideText, currentStartIndex);
                        myBatchRequest.AddUpdateTextStyleRequest(objectId, "SpeakerNotesTextStyle", speakerNotesTextStyleFields, currentStartIndex, currentStartIndex + lastSlideText.Length - 1, lastSlidelink, false);
                        myBatchRequest.AddUpdateTextStyleRequest(objectId, "SpeakerNotesTextStyle", speakerNotesTextStyleFields, currentStartIndex + lastSlideText.Length - 1, currentStartIndex + lastSlideText.Length, null, false);
                        currentStartIndex += lastSlideText.Length;

                        myBatchRequest.AddUpdateParagraphStyleRequest(objectId, false);
                    }

                    #endregion

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

                    var slideHeaderIndex = -1;
                    var slideFooterIndex = -1;
                    var slidePageIdIndex = -1;
                    var slideFooterCreated = false;
                    var slidePageIdCreated = false;

                    for (var j = 0; j < presentation.Slides[i].PageElements.Count; j++)
                    {
                        if (presentation.Slides[i].PageElements[j].Shape != null &&
                            presentation.Slides[i].PageElements[j].Shape.ShapeType == "TEXT_BOX")
                        {
                            //Found the header object
                            if (presentation.Slides[i].PageElements[j].Transform.ScaleX == slideHeaderTransform.ScaleX &&
                                presentation.Slides[i].PageElements[j].Transform.ScaleY == slideHeaderTransform.ScaleY &&
                                presentation.Slides[i].PageElements[j].Transform.TranslateX == slideHeaderTransform.TranslateX &&
                                presentation.Slides[i].PageElements[j].Transform.TranslateY == slideHeaderTransform.TranslateY)

                            {
                                slideHeaderIndex = j;
                            }

                            //Found the page id object
                            else if (presentation.Slides[i].PageElements[j].Transform.ScaleX == slideFooterTransform.ScaleX &&
                                     presentation.Slides[i].PageElements[j].Transform.ScaleY == slideFooterTransform.ScaleY &&
                                     presentation.Slides[i].PageElements[j].Transform.TranslateX == slideFooterTransform.TranslateX &&
                                     presentation.Slides[i].PageElements[j].Transform.TranslateY == slideFooterTransform.TranslateY)

                            {
                                slideFooterIndex = j;
                            }

                            //Found the page id object
                            else if (presentation.Slides[i].PageElements[j].Transform.ScaleX == slidePageIdTransform.ScaleX &&
                                     presentation.Slides[i].PageElements[j].Transform.ScaleY == slidePageIdTransform.ScaleY &&
                                     presentation.Slides[i].PageElements[j].Transform.TranslateX == slidePageIdTransform.TranslateX &&
                                     presentation.Slides[i].PageElements[j].Transform.TranslateY == slidePageIdTransform.TranslateY)

                            {
                                if (presentation.Slides[i].PageElements[j].Shape.Text == null ||
                                    presentation.Slides[i].PageElements[j].Shape.Text.TextElements == null ||
                                    presentation.Slides[i].PageElements[j].Shape.Text.TextElements.Count < 2 ||
                                    presentation.Slides[i].PageElements[j].Shape.Text.TextElements[1].TextRun == null)
                                {
                                    //An empty text box at the location of the page id - delete it
                                    myBatchRequest.AddDeleteObjectRequest(presentation.Slides[i].PageElements[j].ObjectId);
                                }                    
                                else
                                {
                                    //This is the page id text box
                                    slidePageIdIndex = j;
                                }
                            }
                        }
                    }

                    string desiredPageId = (i + 1).ToString();
                    string desiredFooter = cachePresentation.FooterText + "\n";

                    //Header
                    if (slideHeaderIndex >= 0 && 
                        presentation.Slides[i].PageElements[slideHeaderIndex].Shape != null &&
                        presentation.Slides[i].PageElements[slideHeaderIndex].Shape.Text != null &&
                        presentation.Slides[i].PageElements[slideHeaderIndex].Shape.Text.TextElements != null &&
                        presentation.Slides[i].PageElements[slideHeaderIndex].Shape.Text.TextElements.Count > 1 &&
                        presentation.Slides[i].PageElements[slideHeaderIndex].Shape.Text.TextElements[1].TextRun != null &&
                        presentation.Slides[i].PageElements[slideHeaderIndex].Shape.Text.TextElements[1].TextRun.Content != null)
                    {
                        //Header text box exists - check if text contains search text in config
                        if (presentation.Slides[i].PageElements[slideHeaderIndex].Shape.Text.TextElements[1].TextRun.Content.Contains(lookForTextInHeader))
                        {
                            PresentationError.Invoke(this, new SlideErrorEventArgs(new SlideError(presentation.PresentationId, presentation.Title, i + 1, "Header contains " + lookForTextInHeader)));
                        }
                    }                   

                    //Footer - exclude last 2 slides (homework + toc)
                    if (i < presentation.Slides.Count - 2)
                    {
                        if (slideFooterIndex >= 0)
                        {
                            //Footer text box exists - update only if different
                            if (presentation.Slides[i].PageElements[slideFooterIndex].Shape.Text.TextElements[1].TextRun.Content != desiredFooter)
                            {
                                myBatchRequest.AddDeleteTextRequest(presentation.Slides[i].PageElements[slideFooterIndex].ObjectId, presentation.Slides[i].PageElements[slideFooterIndex].Shape);
                                myBatchRequest.AddInsertTextRequest(presentation.Slides[i].PageElements[slideFooterIndex].ObjectId, desiredFooter, 0);
                            }
                        }
                        else
                        {
                            //Create a new text box to hold the slide footer
                            myBatchRequest.AddCreateShapeRequest(presentation.Slides[i].ObjectId, slideFooterSize, slideFooterTransform);
                            slideFooterCreated = true;
                        }
                    }

                    //Page Id
                    if (slidePageIdIndex >= 0)
                    {
                        //Page Id text box exists - update only if different
                        if (presentation.Slides[i].PageElements[slidePageIdIndex].Shape.Text.TextElements[1].TextRun.Content != desiredPageId)
                        {
                            myBatchRequest.AddDeleteTextRequest(presentation.Slides[i].PageElements[slidePageIdIndex].ObjectId, presentation.Slides[i].PageElements[slidePageIdIndex].Shape);
                            myBatchRequest.AddInsertTextRequest(presentation.Slides[i].PageElements[slidePageIdIndex].ObjectId, desiredPageId, 0);
                        }
                    }
                    else
                    {
                        //Create a new text box to hold the slide number
                        myBatchRequest.AddCreateShapeRequest(presentation.Slides[i].ObjectId, slidePageIdSize, slidePageIdTransform);
                        slidePageIdCreated = true;
                    }

                    var batchResponse = myBatchRequest.Execute();
                    myBatchRequest.ClearRequests();

                    var addTextBoxesBatchRequest = new MyBatchRequest(slidesService, cachePresentation.PresentationId);
                    var textBoxesAdded = false;
                    for (var k = 0; k < batchResponse.Replies.Count; k++)
                    {
                        if (batchResponse.Replies[k].CreateShape != null)
                        {
                            if (slideFooterCreated)
                            {
                                addTextBoxesBatchRequest.AddInsertTextRequest(batchResponse.Replies[k].CreateShape.ObjectId, desiredFooter, 0);
                                addTextBoxesBatchRequest.AddUpdateTextStyleRequest(batchResponse.Replies[k].CreateShape.ObjectId, "SlideFooterTextBoxTextStyle", slideFooterTextBoxTextStyleFields, 0, desiredFooter.Length, null);
                                addTextBoxesBatchRequest.AddUpdateParagraphStyleRequest(batchResponse.Replies[k].CreateShape.ObjectId, false);
                                slideFooterCreated = false;
                                textBoxesAdded = true;
                            }
                            if (slidePageIdCreated)
                            {
                                addTextBoxesBatchRequest.AddInsertTextRequest(batchResponse.Replies[k].CreateShape.ObjectId, desiredPageId, 0);
                                addTextBoxesBatchRequest.AddUpdateTextStyleRequest(batchResponse.Replies[k].CreateShape.ObjectId, "SlideIdTextBoxTextStyle", slideIdTextBoxTextStyleFields, 0, desiredPageId.Length, null);
                                addTextBoxesBatchRequest.AddUpdateParagraphStyleRequest(batchResponse.Replies[k].CreateShape.ObjectId, false);
                                slidePageIdCreated = false;
                                textBoxesAdded = true;
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
                        var link = new Link()
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

        #endregion

        #region Events

        public event EventHandler PresentationProcessed;
        public event EventHandler PresentationError;

        #endregion

        #region Private Methods

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
                Requests = new List<Request>()
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
            batchUpdatePresentationRequest.Requests.Add(new Request()
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
        public void AddDeleteTextRequest(string objectId, Shape shape)
        {
            if (shape.Text == null)
            {
                return;
            }

            batchUpdatePresentationRequest.Requests.Add(new Request()
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
            batchUpdatePresentationRequest.Requests.Add(new Request()
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
        public void AddUpdateTextStyleRequest(string objectId, string textStyleConfigKey, string textStyleFields, int startIndex, int endIndex, Link link = null, bool underline = true)
        {
            var fields = String.Copy(textStyleFields);
            var textStyle = JsonConvert.DeserializeObject<TextStyle>(ConfigurationManager.AppSettings[textStyleConfigKey]);

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

            batchUpdatePresentationRequest.Requests.Add(new Request()
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

            batchUpdatePresentationRequest.Requests.Add(new Request()
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
            batchUpdatePresentationRequest.Requests.Add(new Request()
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

            batchUpdatePresentationRequest.Requests.Add(new Request()
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
            batchUpdatePresentationRequest.Requests.Add(new Request()
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
            batchUpdatePresentationRequest.Requests.Add(new Request()
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
            batchUpdatePresentationRequest.Requests = new List<Request>();
        }
        #endregion
    }

    #endregion
        
}
