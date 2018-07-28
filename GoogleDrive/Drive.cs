using Google.Apis.Auth.OAuth2;
using Google.Apis.Drive.v3;
using Google.Apis.Services;
using Google.Apis.Slides.v1;
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
    public class Drive
    {
        #region Class Members

        DriveService driveService;
        static string[] Scopes = { DriveService.Scope.DriveReadonly, SlidesService.Scope.Presentations };
        static string ApplicationName = "Google Drive";

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

            // Create Drive API service.
            driveService = new DriveService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });

            #endregion

            #region Try load presentations list from local cache
            // deserialize JSON directly from a file
            if (File.Exists(ConfigurationManager.AppSettings["PresentationsListCache"]))
            {
                using (StreamReader file = File.OpenText(ConfigurationManager.AppSettings["PresentationsListCache"]))
                {
                    JsonSerializer serializer = new JsonSerializer();
                    Presentations = (List<string>)serializer.Deserialize(file, typeof(List<string>));
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
            JsonSerializer serializer = new JsonSerializer();
            serializer.Converters.Add(new JavaScriptDateTimeConverter());
            serializer.NullValueHandling = NullValueHandling.Ignore;

            using (StreamWriter sw = new StreamWriter(ConfigurationManager.AppSettings["PresentationsListCache"]))
            using (JsonWriter writer = new JsonTextWriter(sw))
            {
                serializer.Serialize(writer, Presentations);
            }

        }

        #endregion

    }
}


