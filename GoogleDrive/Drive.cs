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
    #region Class Drive
    public class Drive
    {
        #region Class Members

        DriveService driveService;
        SlidesService slidesService;
        static string[] Scopes = { DriveService.Scope.DriveReadonly, SlidesService.Scope.Presentations };
        static string ApplicationName = "Google Drive";
        JsonSerializer jsonSerializer;

        public List<string> Presentations { get; }

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
                    Presentations = (List<string>)jsonSerializer.Deserialize(file, typeof(List<string>));
                }
            }
            else
            {
                Presentations = new List<string>();
            }
            #endregion

        }
        #endregion

        #region Methods

        /// <summary>
        /// Build recursivelly a list of all presentations to work on
        /// </summary>
        /// <param name="rootFolder"></param>
        public void BuildPresentationsList(string rootFolder)
        {
            string filter = "'" + rootFolder + "' in parents AND (mimeType = 'application/vnd.google-apps.folder' OR mimeType = 'application/vnd.google-apps.presentation') AND trashed=false";
            string pageToken = null;
            do
            {
                var request = driveService.Files.List();
                request.Q = filter;
                request.Spaces = "drive";
                request.Fields = "nextPageToken, files(id, name, mimeType)";
                request.PageToken = pageToken;
                var result = request.Execute();
                foreach (var file in result.Files)
                {
                    if (file.MimeType == "application/vnd.google-apps.presentation")
                    {
                        Presentations.Add(file.Id);
                    }
                    else
                    {
                        //This is a folder - continue recurssion
                        BuildPresentationsList(file.Id);
                    }
                }
                pageToken = result.NextPageToken;
            } while (pageToken != null);
        }

        /// <summary>
        /// Save presentations list to local cache file
        /// </summary>
        public void SavePresentationsList()
        {
            jsonSerializer.Converters.Add(new JavaScriptDateTimeConverter());
            jsonSerializer.NullValueHandling = NullValueHandling.Ignore;

            using (StreamWriter sw = new StreamWriter(ConfigurationManager.AppSettings["PresentationsListCache"]))
            using (JsonWriter writer = new JsonTextWriter(sw))
            {
                jsonSerializer.Serialize(writer, Presentations);
            }

        }

        /// <summary>
        /// Adjusts the presentation:
        /// 1) Adds an empty slide in the end, if it does not exist ("Empty board")
        /// 2) For each slide (except the last one "empty board"):
        ///     a) Delete existing speaker notes
        ///     b) Add Links to: "Prev Slide", "Next Slide" (to skip animated hints/solutions, "Last Slide" (empty board)
        ///     c) Adjust slide number text box
        /// 3) For the last slide: add "TOC": a link to each slide (except this last slide)
        /// </summary>
        /// <param name="presentationId"></param>
        public void ProcessPresentation(string presentationId)
        {
            var lastSlidelink = new Link() { RelativeLink = "LAST_SLIDE" };
            var nextSlidelink = new Link() { RelativeLink = "NEXT_SLIDE" };
            var prevSlidelink = new Link() { RelativeLink = "PREVIOUS_SLIDE" };
            var firstSlidelink = new Link() { RelativeLink = "FIRST_SLIDE" };

            var presentationRequest = slidesService.Presentations.Get(presentationId);
            var presentation = presentationRequest.Execute();

            if (presentation.Slides[presentation.Slides.Count-1].PageElements.Count > 2)
            {
                //Create empty slide as the last slide
                var createNewSlideBatchRequest = new MyBatchRequest(slidesService, presentationId);
                createNewSlideBatchRequest.AddCreateSlideRequest(presentation.Slides.Count);
                createNewSlideBatchRequest.Execute();

                //Read presentation with the newly created slide
                presentation = presentationRequest.Execute();
            }

            var firstSlideText = ConfigurationManager.AppSettings["FirstSlideText"] + "\t";
            var prevSlideText = ConfigurationManager.AppSettings["PrevSlideText"] + "\t";
            var nextSlideText = ConfigurationManager.AppSettings["NextSlideText"] + "\t";
            var lastSlideText = ConfigurationManager.AppSettings["LastSlideText"] + "\t";

            for (var i=0; i<presentation.Slides.Count-1; i++)
            {
                var myBatchRequest = new MyBatchRequest(slidesService, presentationId, presentation.Slides[i]);
                myBatchRequest.AddDeleteTextRequest();

                var currentStartIndex = 0;

                myBatchRequest.AddInsertTextRequest(firstSlideText, currentStartIndex);
                myBatchRequest.AddUpdateTextStyleRequest(currentStartIndex, currentStartIndex + firstSlideText.Length, firstSlidelink);
                currentStartIndex += firstSlideText.Length;

                myBatchRequest.AddInsertTextRequest(prevSlideText, currentStartIndex);
                myBatchRequest.AddUpdateTextStyleRequest(currentStartIndex, currentStartIndex + prevSlideText.Length, prevSlidelink);
                currentStartIndex += prevSlideText.Length;

                myBatchRequest.AddInsertTextRequest(nextSlideText, currentStartIndex);
                myBatchRequest.AddUpdateTextStyleRequest(currentStartIndex, currentStartIndex + nextSlideText.Length, nextSlidelink);
                currentStartIndex += nextSlideText.Length;

                myBatchRequest.AddInsertTextRequest(lastSlideText, currentStartIndex);
                myBatchRequest.AddUpdateTextStyleRequest(currentStartIndex, currentStartIndex + lastSlideText.Length, lastSlidelink);
                currentStartIndex += lastSlideText.Length;

                myBatchRequest.AddUpdateParagraphStyleRequest(false);

                myBatchRequest.Execute();

            }

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
        Page slide;
        string presentationId;

        #endregion

        #region C'Tor/D'Tor
        public MyBatchRequest(SlidesService slidesService, string presentationId, Page slide = null)
        {
            this.slidesService = slidesService;
            batchUpdatePresentationRequest = new BatchUpdatePresentationRequest();
            batchUpdatePresentationRequest.Requests = new List<Request>();
            this.presentationId = presentationId;
            this.slide = slide;
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
        public void AddDeleteTextRequest()
        {
            if (slide.SlideProperties.NotesPage.PageElements[1].Shape.Text == null)
            {
                return;
            }

            batchUpdatePresentationRequest.Requests.Add(new Request()
            {
                DeleteText = new DeleteTextRequest()
                {
                    ObjectId = slide.SlideProperties.NotesPage.PageElements[1].ObjectId,
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
        public void AddInsertTextRequest(string text, int insertionIndex)
        {
            batchUpdatePresentationRequest.Requests.Add(new Request()
            {
                InsertText = new InsertTextRequest()
                {
                    ObjectId = slide.SlideProperties.NotesPage.PageElements[1].ObjectId,
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
        public void AddUpdateTextStyleRequest(int startIndex, int endIndex, Link link)
        {
            var textStyle = JsonConvert.DeserializeObject<TextStyle>(ConfigurationManager.AppSettings["TextStyle"]);
            var fields = ConfigurationManager.AppSettings["TextStyleFields"];

            textStyle.Link = link;
            fields += ",link";

            batchUpdatePresentationRequest.Requests.Add(new Request()
            {
                UpdateTextStyle = new UpdateTextStyleRequest()
                {
                    ObjectId = slide.SlideProperties.NotesPage.PageElements[1].ObjectId,
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
        public void AddUpdateParagraphStyleRequest(bool rtl)
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
                    
                    ObjectId = slide.SlideProperties.NotesPage.PageElements[1].ObjectId,
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
        /// Executes the requests added to the list
        /// </summary>
        /// <returns></returns>
        public BatchUpdatePresentationResponse Execute()
        {
            if (batchUpdatePresentationRequest.Requests != null && batchUpdatePresentationRequest.Requests.Count > 0)
            {
                var batchUpdateRequest = slidesService.Presentations.BatchUpdate(batchUpdatePresentationRequest, presentationId);
                return batchUpdateRequest.Execute();
            }
            return null;
        }
        
        #endregion
    }
    #endregion
}


