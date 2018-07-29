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
        ///     d) if slide contains only a single image in the body - align it to top/bottom 
        /// 3) For the last slide: add "TOC": a link to each slide (except this last slide)
        /// </summary>
        /// <param name="presentationId"></param>
        public void ProcessPresentation(string presentationId)
        {
            #region Load variables

            string objectId;
            int currentStartIndex;

            var lastSlidelink = new Link() { RelativeLink = "LAST_SLIDE" };
            var nextSlidelink = new Link() { RelativeLink = "NEXT_SLIDE" };
            var prevSlidelink = new Link() { RelativeLink = "PREVIOUS_SLIDE" };
            var firstSlidelink = new Link() { RelativeLink = "FIRST_SLIDE" };

            var firstSlideText = ConfigurationManager.AppSettings["FirstSlideText"] + "\t";
            var prevSlideText = ConfigurationManager.AppSettings["PrevSlideText"] + "\t";
            var nextSlideText = ConfigurationManager.AppSettings["NextSlideText"] + "\t";
            var lastSlideText = ConfigurationManager.AppSettings["LastSlideText"] + "\t";

            var speakerNotesTextStyleFields = ConfigurationManager.AppSettings["SpeakerNotestTextStyleFields"];
            var slideIdTextBoxTextStyleFields = ConfigurationManager.AppSettings["SlideIdTextBoxTextStyleFields"];

            var slidePageIdSize = JsonConvert.DeserializeObject<Size>(ConfigurationManager.AppSettings["SlidePageIdSize"]);
            var slidePageIdTransform = JsonConvert.DeserializeObject<AffineTransform>(ConfigurationManager.AppSettings["SlidePageIdTransform"]);

            var alignImage = (AlignImage)Enum.Parse(typeof(AlignImage), ConfigurationManager.AppSettings["ImageAlign"]);

            #endregion

            #region Load Presentation

            var presentationRequest = slidesService.Presentations.Get(presentationId);
            var presentation = presentationRequest.Execute();

            #endregion

            #region Create Empty Slide (if neccessary)

            if (presentation.Slides[presentation.Slides.Count-1].PageElements.Count > 2)
            {
                //Create empty slide as the last slide
                var createNewSlideBatchRequest = new MyBatchRequest(slidesService, presentationId);
                createNewSlideBatchRequest.AddCreateSlideRequest(presentation.Slides.Count);
                createNewSlideBatchRequest.Execute();

                //Read presentation with the newly created slide
                presentation = presentationRequest.Execute();
            }

            #endregion

            #region Slides loop - processing all but last slide

            for (var i=0; i<presentation.Slides.Count-1; i++)
            {
                var myBatchRequest = new MyBatchRequest(slidesService, presentationId);

                #region Delete existing spearker notes from slide

                currentStartIndex = 0;
                objectId = presentation.Slides[i].SlideProperties.NotesPage.PageElements[1].ObjectId;
                myBatchRequest.AddDeleteTextRequest(objectId, presentation.Slides[i].SlideProperties.NotesPage.PageElements[1].Shape);

                #endregion

                #region Add First/Prev/Next/Last buttons

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

                #endregion

                #region Align Image

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

                #region Slide Id Text Box

                var slidePageIdIndex = -1;
                for (var j=0; j<presentation.Slides[i].PageElements.Count; j++)
                {
                    if (presentation.Slides[i].PageElements[j].Shape != null && 
                        presentation.Slides[i].PageElements[j].Shape.ShapeType == "TEXT_BOX" &&
                        presentation.Slides[i].PageElements[j].Transform.ScaleX == slidePageIdTransform.ScaleX &&
                        presentation.Slides[i].PageElements[j].Transform.ScaleY == slidePageIdTransform.ScaleY)
                    {
                        slidePageIdIndex = j;
                        break;
                    }
                }

                if (slidePageIdIndex >= 0)
                {
                    //Page Id text box exists
                    myBatchRequest.AddDeleteTextRequest(presentation.Slides[i].PageElements[slidePageIdIndex].ObjectId, presentation.Slides[i].PageElements[slidePageIdIndex].Shape);
                    myBatchRequest.AddInsertTextRequest(presentation.Slides[i].PageElements[slidePageIdIndex].ObjectId, (i+1).ToString(),0);
                }
                else
                {
                    //Create a new text box to hold the slide number
                    myBatchRequest.AddCreateShapeRequest(presentation.Slides[i].ObjectId, slidePageIdSize, slidePageIdTransform);
                }

                var batchResponse = myBatchRequest.Execute();
                if (batchResponse.Replies[batchResponse.Replies.Count-1].CreateShape != null)
                {
                    //Read presentation with the newly created text box for the slide id
                    //presentation = presentationRequest.Execute();
                    var addSlideIdTextBatchRequest = new MyBatchRequest(slidesService, presentationId);
                    addSlideIdTextBatchRequest.AddInsertTextRequest(batchResponse.Replies[batchResponse.Replies.Count - 1].CreateShape.ObjectId, (i + 1).ToString(), 0);
                    addSlideIdTextBatchRequest.AddUpdateTextStyleRequest(batchResponse.Replies[batchResponse.Replies.Count - 1].CreateShape.ObjectId, "SlideIdTextBoxTextStyle", slideIdTextBoxTextStyleFields,  0, (i + 1).ToString().Length, null);
                    addSlideIdTextBatchRequest.AddUpdateParagraphStyleRequest(batchResponse.Replies[batchResponse.Replies.Count - 1].CreateShape.ObjectId, false);
                    addSlideIdTextBatchRequest.Execute();
                }

                #endregion
            }

            #endregion

            #region Process Last Slide (TOC)

            var createTOCBatchRequest = new MyBatchRequest(slidesService, presentationId);
            objectId = presentation.Slides[presentation.Slides.Count-1].SlideProperties.NotesPage.PageElements[1].ObjectId;

            createTOCBatchRequest.AddDeleteTextRequest(objectId, presentation.Slides[presentation.Slides.Count - 1].SlideProperties.NotesPage.PageElements[1].Shape);

            currentStartIndex = 0;
            string currentPageIdString;
            for(var i=1; i<=presentation.Slides.Count-1; i++)
            {
                var link = new Link()
                {
                    SlideIndex = i-1
                };
                currentPageIdString = (i).ToString("00") + "\t";
                createTOCBatchRequest.AddInsertTextRequest(objectId, currentPageIdString, currentStartIndex);
                //Link - will not contain the tab ("\t")
                createTOCBatchRequest.AddUpdateTextStyleRequest(objectId, "SpeakerNotesTextStyle", speakerNotesTextStyleFields, currentStartIndex, currentStartIndex + currentPageIdString.Length - 1, link);
                currentStartIndex += currentPageIdString.Length;
            }
            createTOCBatchRequest.AddUpdateParagraphStyleRequest(objectId, true);
            createTOCBatchRequest.Execute();
            
            #endregion
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


