using Google.Apis.Auth.OAuth2;
using Google.Apis.Drive.v3;
using Google.Apis.Drive.v3.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using OrbitEmailHelper;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace OrbitEmailHelper
{
    class Skeleton
    {
        static string[] Scopes = { DriveService.Scope.Drive };
        static string ApplicationName = "Drive API .NET Quickstart";
        static string TeamDriveName = "ORBiT Team Drive";
        static string RootFolderName = "ORBiT EFiles Folder";
        static string FlagFileName = "_addOnFinished.txt";   //this file is created by the gmail add on to signal the service that it is ok to process the folder
        static string TempFileName = "_addOnTemp.html";     //this is the temporary html file that is created before it is converted to pdf
        static System.Timers.Timer timer = new System.Timers.Timer();

        static void Main(string[] args)
        {
            //refer to https://developers.google.com/drive/api/v3/quickstart/dotnet to get this json file
            string credPath = @"c:\orbit-y-drive-service.json";
            GoogleCredential credential;
            using (var stream = new FileStream(credPath, FileMode.Open, FileAccess.Read))
            {
                credential = GoogleCredential.FromStream(stream).CreateScoped(Scopes);
            }

            DriveService service = new DriveService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });

            string teamDrivePageToken = null;
            string teamDriveID = null;
            do
            {
                TeamdrivesResource.ListRequest teamDriveList = service.Teamdrives.List();
                teamDriveList.Fields = "nextPageToken, teamDrives(kind, id, name)";
                teamDriveList.PageToken = teamDrivePageToken;
                var result = teamDriveList.Execute();
                var teamDrives = result.TeamDrives;

                if (teamDrives != null && teamDrives.Count > 0)
                {
                    foreach (TeamDrive drive in teamDrives)
                    {
                        if (drive.Name == TeamDriveName)
                        {
                            teamDriveID = drive.Id;
                            break;
                        }
                    }
                }
                teamDrivePageToken = result.NextPageToken;
            } while (teamDrivePageToken != null && teamDriveID == null);

            if (teamDriveID == null)
            {
                WriteLogEntry("Team drive not found", null);
                StopService();
                return;
            }

            string rootFolderPageToken = null;
            string rootFolderId = "";
            do
            {
                FilesResource.ListRequest rootFolderRequest = service.Files.List();
                rootFolderRequest.Fields = "nextPageToken, files(id, name, parents, mimeType)";
                rootFolderRequest.PageToken = rootFolderPageToken;
                rootFolderRequest.SupportsTeamDrives = true;
                rootFolderRequest.IncludeTeamDriveItems = true;
                rootFolderRequest.Corpora = "teamDrive";
                rootFolderRequest.TeamDriveId = teamDriveID;
                rootFolderRequest.Q = "parents='" + teamDriveID + "' and trashed=false and name='" + RootFolderName + "'";

                var result = rootFolderRequest.Execute();
                var files = result.Files;
                if (files != null && files.Count == 1)
                {
                    var file = files[0];
                    if (file.MimeType == "application/vnd.google-apps.folder")
                    {
                        rootFolderId = file.Id;
                        break;
                    }
                }
                rootFolderPageToken = result.NextPageToken;
            } while (rootFolderPageToken != null && rootFolderId == null);

            if (rootFolderId == "")
            {
                WriteLogEntry("Error finding Root Folder. This can be caused by a duplicate folder, no folder, or execution error etc.", null);
                StopService();
                return;
            }

            string recordFolderPageToken = null;
            List<Google.Apis.Drive.v3.Data.File> recordFolders = new List<Google.Apis.Drive.v3.Data.File>();
            do
            {
                FilesResource.ListRequest recordFoldersRequest = service.Files.List();
                recordFoldersRequest.Fields = "nextPageToken, files(id, name, parents, mimeType)";
                recordFoldersRequest.PageToken = recordFolderPageToken;
                recordFoldersRequest.SupportsTeamDrives = true;
                recordFoldersRequest.IncludeTeamDriveItems = true;
                recordFoldersRequest.Corpora = "teamDrive";
                recordFoldersRequest.TeamDriveId = teamDriveID;
                recordFoldersRequest.Q = "parents='" + rootFolderId + "' and trashed=false";

                var result = recordFoldersRequest.Execute();
                var files = result.Files;

                if (files != null && files.Count > 0)
                {
                    foreach (Google.Apis.Drive.v3.Data.File file in files)
                    {
                        if (file.MimeType == "application/vnd.google-apps.folder")
                        {
                            recordFolders.Add(file);
                        }
                        else
                        {
                            //file found. shouldnt happen - delete
                            TrashDriveObject(service, file.Id);
                        }
                    }
                }
                recordFolderPageToken = result.NextPageToken;
            } while (recordFolderPageToken != null);

            if (recordFolders.Count != 0)
            {
                foreach (Google.Apis.Drive.v3.Data.File recordFolder in recordFolders)
                {
                    string addOnEmail = "";
                    if (isAddOnFinished(service, teamDriveID, recordFolder.Id, ref addOnEmail))
                    {
                        string recordNumber = recordFolder.Name;
                        if (isValidRecordNumber(recordNumber) == false && isValidSubawardNumber(recordNumber) == false)
                        {
                            SendErrorEmail(addOnEmail, "Record number " + recordNumber + " does not exist in ORBiT.");
                            TrashDriveObject(service, recordFolder.Id);
                            continue;
                        }

                        string EFilesFolder = GetEFilesFolder(recordNumber);
                        if (EFilesFolder == "")
                        {
                            SendErrorEmail(addOnEmail, "Could not find an EFiles folder. Record number is " + recordNumber);
                            TrashDriveObject(service, recordFolder.Id);
                            continue;
                        }

                        string filePageToken = null;
                        List<string> lstSavedFiles = new List<string>();
                        do
                        {
                            FilesResource.ListRequest filesRequest = service.Files.List();
                            filesRequest.Fields = "nextPageToken, files(id, name, parents, mimeType)";
                            filesRequest.PageToken = filePageToken;
                            filesRequest.SupportsTeamDrives = true;
                            filesRequest.IncludeTeamDriveItems = true;
                            filesRequest.Corpora = "teamDrive";
                            filesRequest.TeamDriveId = teamDriveID;
                            filesRequest.Q = "parents='" + recordFolder.Id + "' and trashed=false";

                            var result = filesRequest.Execute();
                            var files = result.Files;
                            bool stopFlag = false;

                            foreach (Google.Apis.Drive.v3.Data.File file in files)
                            {
                                var fileInfo = service.Files.Get(file.Id);
                                if (file.MimeType == "application/vnd.google-apps.folder")
                                {
                                    //folder found - shouldnt happen - trash folder
                                    TrashDriveObject(service, file.Id);
                                }
                                else
                                {
                                    if (file.Name == FlagFileName || file.Name == TempFileName)
                                    {
                                        continue;
                                    }
                                    else
                                    {
                                        using (MemoryStream ms = new MemoryStream())
                                        {
                                            fileInfo.Download(ms);
                                            string filenameToSave = getUniqueFilename(EFilesFolder, file.Name);
                                            try
                                            {
                                                using (FileStream fs = new FileStream(EFilesFolder + filenameToSave, FileMode.Create, System.IO.FileAccess.Write))
                                                {
                                                    try
                                                    {
                                                        fs.Write(ms.GetBuffer(), 0, ms.GetBuffer().Length);
                                                    }
                                                    catch (Exception e)
                                                    {
                                                        string message = "Error saving file to Y Drive. Record Number: " + recordNumber + ", file name: " +
                                                            file.Name;
                                                        WriteLogEntry(message, e);

                                                        if (lstSavedFiles.Count > 0)
                                                        {
                                                            message += "<br /><br />";
                                                            message += "The following files were saved successfully:<br /><ul>";
                                                            for (int i = 0; i < lstSavedFiles.Count; i++)
                                                            {
                                                                message += "<li>" + lstSavedFiles[i] + "</li>";
                                                            }
                                                            message += "</ul>";
                                                        }

                                                        SendErrorEmail(addOnEmail, message);

                                                        stopFlag = true;
                                                        ms.Close();
                                                        break;
                                                    }

                                                    lstSavedFiles.Add(filenameToSave);
                                                }
                                                ms.Close();
                                            }
                                            catch (Exception e)
                                            {
                                                WriteLogEntry("Error writing to filestream. File name is " + filenameToSave, e);

                                                string message = "Error saving file to Y Drive. Record Number: " + recordNumber + ", file name: " +
                                                    file.Name;
                                                if (lstSavedFiles.Count > 0)
                                                {
                                                    message += "<br /><br />";
                                                    message += "The following files were saved successfully:<br /><ul>";
                                                    for (int i = 0; i < lstSavedFiles.Count; i++)
                                                    {
                                                        message += "<li>" + lstSavedFiles[i] + "</li>";
                                                    }
                                                    message += "</ul>";
                                                }

                                                SendErrorEmail(addOnEmail, message);

                                                stopFlag = true;
                                                break;
                                            }
                                        }
                                    }
                                }
                            }

                            if (stopFlag == true)
                            {
                                break;
                            }

                            filePageToken = result.NextPageToken;
                        } while (filePageToken != null);

                        TrashDriveObject(service, recordFolder.Id);
                    }
                }
            }
        }

        private static string getUniqueFilename(string EFilesFolder, string filename)
        {
            //custom logic
            return filename;
        }

        public static bool isValidRecordNumber(string recordNumber)
        {
            //custom validation
            return true;
        }

        public static bool isValidSubawardNumber(string recordNumber)
        {
            //custom validation
            return true;
        }

        public static string GetEFilesFolder(string recordNumber)
        {
            //custom logic
            return @"C:\testFolder\";
        }


        /// <summary>
        /// Writes a log entry to "Y Drive Email Service" log. 
        /// </summary>
        /// <param name="message">The main error message</param>
        /// <param name="e">The system generated exception, or null</param>
        public static void WriteLogEntry(string message, Exception e)
        {
            EventLog log = new EventLog();
            log.Source = "Y Drive Email Service";
            if (e != null)
            {
                log.WriteEntry("Error information: " + message + "\n\nException message: " + e.Message);
            }
            else
            {
                log.WriteEntry("Error information: " + message);
            }
        }


        /// <summary>
        /// Checks the drive folder to make sure the add-on is not still executing
        /// </summary>
        /// <param name="service"></param>
        /// <param name="teamDriveID"></param>
        /// <param name="recordFolderID"></param>
        /// <param name="sendToEmail"></param>
        /// <returns></returns>
        public static bool isAddOnFinished(DriveService service, string teamDriveID, string recordFolderID, ref string sendToEmail)
        {
            string filePageToken = null;
            List<string> lstSavedFiles = new List<string>();
            do
            {
                FilesResource.ListRequest filesRequest = service.Files.List();
                filesRequest.Fields = "nextPageToken, files(id, name, parents, mimeType)";
                filesRequest.PageToken = filePageToken;
                filesRequest.SupportsTeamDrives = true;
                filesRequest.IncludeTeamDriveItems = true;
                filesRequest.Corpora = "teamDrive";
                filesRequest.TeamDriveId = teamDriveID;
                filesRequest.Q = "parents='" + recordFolderID + "' and trashed=false";

                var result = filesRequest.Execute();
                var files = result.Files;

                foreach (Google.Apis.Drive.v3.Data.File file in files)
                {
                    var fileInfo = service.Files.Get(file.Id);
                    if (file.MimeType == "application/vnd.google-apps.folder")
                    {
                        //folder found - shouldnt happen. ignore here. deleted in file process
                    }
                    else
                    {
                        if (file.Name == FlagFileName)
                        {
                            using (MemoryStream ms = new MemoryStream())
                            {
                                fileInfo.Download(ms);
                                sendToEmail = Encoding.ASCII.GetString(ms.ToArray());
                                ms.Close();
                            }
                            return true;
                        }
                    }
                }
                filePageToken = result.NextPageToken;
            } while (filePageToken != null);

            return false;
        }

        public static bool isValidEmail(string email)
        {
            try
            {
                var addr = new System.Net.Mail.MailAddress(email);
                return addr.Address == email;
            }
            catch
            {
                return false;
            }
        }

        public static void SendErrorEmail(string addOnEmail, string message)
        {
            //custom logic
        }

        public static void TrashDriveObject(DriveService service, string objectID)
        {
            try
            {
                Google.Apis.Drive.v3.Data.File f = new Google.Apis.Drive.v3.Data.File();
                f.Trashed = true;
                var updateRequest = service.Files.Update(f, objectID);
                updateRequest.SupportsTeamDrives = true;
                updateRequest.Execute();

            }
            catch (Exception e)
            {
                WriteLogEntry("Error trashing object", e);
                StopService();
            }
        }

        private static void StopService()
        {
            //custom service logic
        }
    }
}
