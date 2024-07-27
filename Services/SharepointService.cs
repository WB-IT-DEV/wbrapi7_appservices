using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Security;
using System.Text;
using System.Threading.Tasks;
using Microsoft.VisualBasic;
using System.Web;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint;
using ClientOM = Microsoft.SharePoint.Client;

using System.Collections.Concurrent;
using System.Net.Http;
using System.Text.Json;
using System.Threading;
using Microsoft.AspNetCore.Mvc.Filters;
using wbrapi7_appservices.Repositories;
using System.Net;


namespace wbrapi7_appservices.Services
{
    public class SharepointService
    {
        public ClientContext clientContext { get; set; }
        // Private ServerSiteUrl As String = "https://h20bridge.sharepoint.com/sites/BudgetPlanning"
        // Private LibraryUrl As String = "Shared Documents/Spotfire_GIS_Data/"
        // Private UserName As String = "wbradmin@h2obridge.com"
        // Private Password As String = "W@ter$ridgeIT1$"

        private string UserName = ""; // = "hau.nguyen@h2obridge.com"
        private string Password = "";

        private Web WebClient { get; set; }
        DateTime LastRefreshWebClient;

        public SharepointService(string sUserName, string sPassword)
        {
            UserName = sUserName;
            Password = sPassword;
        }

        public bool Connect(string ServerSiteUrl)
        {
            try
            {
                var boolRefresh = false;

                if (LastRefreshWebClient == null)
                {
                    LastRefreshWebClient = DateTime.Now;
                    boolRefresh = true;
                }
                else
                {
                    TimeSpan min = DateTime.Now - LastRefreshWebClient;
                    boolRefresh = (min.TotalMinutes >= 3);
                    LastRefreshWebClient = DateTime.Now;

                }

                if (boolRefresh)
                {

                    var SPAPI = new AuthenticationManager();

                    var securePassword = new SecureString();

                    foreach (char c in Password)
                        securePassword.AppendChar(c);

                    //clientContext.Credentials = new SharePointOnlineCredentials(UserName, securePassword);
                    clientContext = SPAPI.GetContext(new Uri(ServerSiteUrl), UserName, securePassword); //ClientContext(ServerSiteUrl);

                    WebClient = clientContext.Web;
                }
                return true;
            }
            catch (Exception ex)
            {
                throw (ex);
            }
        }

        public string DownloadFiles(string ServerSiteUrl, string LibaryURL, bool bDeleteExitsFileInFolder = false)
        {
            try
            {
                string tempLocation = @"\\wbrd01\Deploy\GIS\";
                System.IO.DirectoryInfo di = new DirectoryInfo(tempLocation);

                Connect(ServerSiteUrl);

                if (bDeleteExitsFileInFolder)
                {
                    foreach (FileInfo file in di.GetFiles())
                        file.Delete();
                }

                FileCollection files = WebClient.GetFolderByServerRelativeUrl(LibaryURL).Files;

                clientContext.Load(files);
                clientContext.ExecuteQuery();
                if (clientContext.HasPendingRequest)
                    clientContext.ExecuteQuery();

                foreach (ClientOM.File file in files)
                {
                    try
                    {

                        Microsoft.SharePoint.Client.ClientResult<Stream> mstream = file.OpenBinaryStream();

                        clientContext.ExecuteQuery();
                        var filePath = tempLocation + file.Name;

                        try
                        {
                            if (System.IO.File.Exists(filePath))
                                System.IO.File.Delete(filePath);
                        }
                        catch (Exception ex)
                        {
                        }

                        using (var fileStream = new System.IO.FileStream(filePath, System.IO.FileMode.Create))
                        {
                            mstream.Value.CopyTo(fileStream);
                        }
                    }
                    catch (Exception fex)
                    {
                    }
                }

                return "";
            }
            catch (Exception ex)
            {
                throw (ex);
            }
        }

        private class CSharpImpl
        {
            [Obsolete("Please refactor calling code to use normal Visual Basic assignment")]
            public static T __Assign<T>(ref T target, T value)
            {
                target = value;
                return value;
            }
        }



        public bool UploadFilesLargeFile(string ServerSiteUrl,
                                         string LibaryURL,
                                         string sFolderName,
                                         string sFileName,
                                         byte[] FileContent,
                                         bool bEmptyFolder = false)
        {
            int fileChunkSizeInMB = 1;
            int blockSize = fileChunkSizeInMB * 1024 * 1024;
            byte[] buffer = new byte[blockSize - 1 + 1];
            byte[] lastBuffer = null;
            long fileoffset = 0;
            long totalBytesRead = 0;
            int bytesRead;
            bool first = true;
            bool last = false;
            Microsoft.SharePoint.Client.File uploadFile = null/* TODO Change to default(_) if this is not a reference type */;
            Guid uploadId = Guid.NewGuid();
            ClientResult<long> bytesUploaded = null/* TODO Change to default(_) if this is not a reference type */;
            int fileSize = FileContent.Length;
            try
            {

                if (Connect(ServerSiteUrl))
                {
                    var docs = WebClient.GetFolderByServerRelativeUrl(LibaryURL);

                    if (sFolderName != "")
                    {
                        try
                        {
                            docs.Folders.Add(sFolderName);
                        }
                        catch (Exception ex)
                        {
                        }
                    }

                    if (sFolderName != "")
                        docs = WebClient.GetFolderByServerRelativeUrl(LibaryURL + "/" + sFolderName + "/");
                    else
                        docs = WebClient.GetFolderByServerRelativeUrl(LibaryURL);


                    if (bEmptyFolder)
                        deleteFilesInfoSPFolder(ServerSiteUrl, LibaryURL, sFolderName);

                    try
                    {
                        clientContext.ExecuteQuery();
                        //return true;
                    }
                    catch (Exception ex)
                    {

                    }

                    if (sFileName != "")
                    {
                        MemoryStream fileStream = new MemoryStream(FileContent);
                        // fileStream.Write(FileContent, 0, FileContent.Length)
                        fileStream.Position = 0;

                        using (BinaryReader br = new BinaryReader(fileStream))
                        {
                            bytesRead = br.Read(buffer, 0, buffer.Length);

                            while (bytesRead > 0)
                            {
                                totalBytesRead = totalBytesRead + bytesRead;


                                if (totalBytesRead >= fileSize)
                                {
                                    last = true;
                                    lastBuffer = new byte[bytesRead - 1 + 1];
                                    Array.Copy(buffer, 0, lastBuffer, 0, bytesRead);
                                }

                                if (first)
                                {
                                    using (MemoryStream contentStream = new MemoryStream())
                                    {
                                        FileCreationInformation fileInfo = new FileCreationInformation();
                                        fileInfo.ContentStream = contentStream;
                                        fileInfo.Url = sFileName;
                                        fileInfo.Overwrite = true;

                                        uploadFile = docs.Files.Add(fileInfo);

                                        using (MemoryStream s = new MemoryStream(buffer))
                                        {
                                            bytesUploaded = uploadFile.StartUpload(uploadId, s);
                                            clientContext.ExecuteQuery();
                                            fileoffset = bytesUploaded.Value;
                                        }

                                        first = false;
                                    }
                                }
                                else if (last)
                                {
                                    using (MemoryStream s = new MemoryStream(lastBuffer))
                                    {
                                        uploadFile = uploadFile.FinishUpload(uploadId, fileoffset, s);
                                        clientContext.ExecuteQuery();
                                        return true;
                                    }
                                }
                                else
                                    using (MemoryStream s = new MemoryStream(buffer))
                                    {
                                        bytesUploaded = uploadFile.ContinueUpload(uploadId, fileoffset, s);
                                        clientContext.ExecuteQuery();
                                        fileoffset = bytesUploaded.Value;
                                    }


                                bytesRead = br.Read(buffer, 0, buffer.Length);
                            }
                        }
                    }
                }
                return true;
            }
            catch (Exception ex)
            {
                throw;
            }
        }

        public bool UploadFilesSmallSize(string ServerSiteUrl,
                                        string LibaryURL,
                                        string sFolderName,
                                        string sFileName,
                                        byte[] FileContent,
                                        bool bEmptyFolder = false)
        {
            try
            {

                if (Connect(ServerSiteUrl))
                {
                    var docs = WebClient.GetFolderByServerRelativeUrl(LibaryURL);

                    if (sFolderName != "")
                    {
                        try
                        {
                            docs.Folders.Add(sFolderName);
                        }
                        catch (Exception ex)
                        {
                        }
                    }

                    if (sFolderName != "")
                        docs = WebClient.GetFolderByServerRelativeUrl(LibaryURL + "/" +sFolderName + "/");
                    else
                        docs = WebClient.GetFolderByServerRelativeUrl(LibaryURL);


                    if (bEmptyFolder)
                        deleteFilesInfoSPFolder(ServerSiteUrl, LibaryURL, sFolderName);

                    if (sFileName != "")
                    {
                        FileCreationInformation newFile = new FileCreationInformation();
                        newFile.Content = FileContent;
                        // newFile.ContentStream = New FileStream()
                        newFile.Url = sFileName;
                        newFile.Overwrite = true;

                        Microsoft.SharePoint.Client.File uploadFile = docs.Files.Add(newFile);

                        clientContext.Load(uploadFile);
                    }

                    try
                    {
                        clientContext.ExecuteQuery();
                        return true;
                    }
                    catch (Exception ex)
                    {
                        if (ex.Message.Contains("already exists"))
                        {
                            return true;
                        }
                        else
                        {
                            return false;
                        }
                    }
                }
                return false;
            }
            catch (Exception ex)
            {
                throw;
            }
        }


        public bool UploadFiles(string ServerSiteUrl, string LibaryURL, string sFolderName, string sFileName, byte[] FileContent, bool bEmptyFolder = false)
        {
            try
            {
                if (FileContent == null)
                    return UploadFilesSmallSize(ServerSiteUrl, LibaryURL, sFolderName, sFileName, FileContent, bEmptyFolder);
                else if (FileContent.Length > (1024 * 1204))
                    return UploadFilesLargeFile(ServerSiteUrl, LibaryURL, sFolderName, sFileName, FileContent, bEmptyFolder);
                else
                    return UploadFilesSmallSize(ServerSiteUrl, LibaryURL, sFolderName, sFileName, FileContent, bEmptyFolder);

                // return false;
            }
            catch (Exception ex)
            {
                throw;
            }
        }

        public FileCollection getSharepointFile(string ServerSiteUrl, string sSourceURL, string sArkURL)
        {
            try
            {
                if (Connect(ServerSiteUrl))
                {
                    var srcFolder = WebClient.GetFolderByServerRelativeUrl(sSourceURL);
                    var arkFolder = WebClient.GetFolderByServerRelativeUrl(sArkURL);
                    clientContext.Load(srcFolder, f2 => f2.Files);
                    clientContext.ExecuteQuery();

                    // For Each file In srcFolder.Files
                    // file.CopyTo(sArkURL & file.Name, True)
                    // Next
                    // clientContext.ExecuteQuery()
                    return srcFolder.Files;
                }

                return null;
            }
            catch (Exception ex)
            {
                throw;
            }
        }


        public (Stream? fileStream, Microsoft.SharePoint.Client.File? file) getSharepointFileStream(string ServerSiteUrl, string sSourceURL, string sArkURL)
        {
            
            try
            {
                if (Connect(ServerSiteUrl))
                {
                    var srcFolder = WebClient.GetFolderByServerRelativeUrl(sSourceURL);
                    //clientContext.Load(srcFolder, f2 => f2.Files);

                    clientContext.Load(srcFolder.Files,
                   files => files.Include(
                       file => file.Name,
                       file => file.TimeLastModified
                   ).OrderByDescending(file => file.TimeLastModified));



                    clientContext.ExecuteQuery();

                    //if (srcFolder.Files.Count > 0)
                    //{
                    //    Microsoft.SharePoint.Client.File file = srcFolder.Files[0]; 
                    //    ClientResult<Stream> fileStreamResult = file.OpenBinaryStream();
                    //    clientContext.ExecuteQuery();

                    //    return (fileStreamResult.Value, file);
                    //}

                    if (srcFolder.Files.Count > 0)
                    {
                        foreach (Microsoft.SharePoint.Client.File file in srcFolder.Files)
                        {
                            //Microsoft.SharePoint.Client.File file = srcFolder.Files[0];
                            // Check if the file has an Excel extension
                            if (file.Name.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase))
                            {
                                ClientResult<Stream> fileStreamResult = file.OpenBinaryStream();
                                clientContext.Load(file);
                                clientContext.ExecuteQuery();
                                return (fileStreamResult.Value, file);
                            }

                        }

                        throw new ApplicationException("No Excel document found in the folder.");
                    }
                    else
                    {

                        throw new ApplicationException("No files found in the folder.");
                    }



                }

                return (null, null);
            }
            catch (Exception)
            {
                
                throw;
            }
        }



        public ClientContext getSharepointClientContext(string ServerSiteUrl)
        {
            try
            {

                if (Connect(ServerSiteUrl))
                    return clientContext;
                return null;
            }
            catch (Exception ex)
            {
                throw;
            }
        }

        public void deleteFilesInfoSPFolder(string ServerSiteUrl, string LibaryURL, string sFolderName)
        {
            var folder = WebClient.GetFolderByServerRelativeUrl(LibaryURL);
            try
            {
                if (sFolderName != "")
                    folder = WebClient.GetFolderByServerRelativeUrl(LibaryURL + sFolderName + "/");
                else
                    folder = WebClient.GetFolderByServerRelativeUrl(LibaryURL);

                clientContext.Load(folder, f2 => f2.Files);
                clientContext.ExecuteQuery();

                foreach (Microsoft.SharePoint.Client.File f in folder.Files)
                    f.DeleteObject();

                clientContext.ExecuteQuery();
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.Message);
            }
        }

        public void ITInventoryFileMove()
        {
            try
            {
            }
            catch (Exception ex)
            {
                throw;
            }
        }


        public string FilesDownload(string ServerSiteUrl,
                                    string LibaryURL,
                                    string DestinationFolder,
                                    ref string downloadFilesName,
                                    string ArchiveFolder = "")
        {
            try
            {
                string tempLocation = DestinationFolder;
                System.IO.DirectoryInfo di = new DirectoryInfo(tempLocation);
                downloadFilesName = "";
                Connect(ServerSiteUrl);

                if (!System.IO.Directory.Exists(tempLocation))
                    System.IO.Directory.CreateDirectory(tempLocation);

                FileCollection files = WebClient.GetFolderByServerRelativeUrl(LibaryURL).Files;

                clientContext.Load(files);
                clientContext.ExecuteQuery();
                if (clientContext.HasPendingRequest)
                    clientContext.ExecuteQuery();

                foreach (ClientOM.File file in files)
                {
                    try
                    {
                        //FileInformation fileInfo = ClientOM.File.OpenBinaryDirect(clientContext, file.ServerRelativeUrl);
                        Microsoft.SharePoint.Client.ClientResult<Stream> mstream = file.OpenBinaryStream();
                        clientContext.ExecuteQuery();
                        var filePath = tempLocation + file.Name;

                        try
                        {
                            if (System.IO.File.Exists(filePath))
                                System.IO.File.Delete(filePath);
                        }
                        catch (Exception ex)
                        {
                        }

                        using (var fileStream = new System.IO.FileStream(filePath, System.IO.FileMode.Create))
                        {
                            mstream.Value.CopyTo(fileStream);
                        }

                        if (downloadFilesName != "")
                            downloadFilesName = downloadFilesName + ";";
                        downloadFilesName = downloadFilesName + filePath;
                    }
                    catch (Exception fex)
                    {
                    }


                    try
                    {
                        file.MoveTo(ArchiveFolder + file.Name, Microsoft.SharePoint.Client.MoveOperations.Overwrite);
                        clientContext.ExecuteQuery();
                    }
                    catch (Exception ex)
                    {
                        Debug.WriteLine(ex.Message);
                    }
                }

                return "";
            }
            catch (Exception ex)
            {
                throw (ex);
            }
        }

        public string getBase64FileType(string fileExt)
        {
            if (fileExt.Length > 4)
            {
                string temp = fileExt.Substring(fileExt.Length - 4, 4);

                switch (temp)
                {
                    case ".png":
                        return "data:image/png;base64,";
                    case ".jpg":
                        return "data:image/jpeg;base64,";
                    case "html":
                        return "data:text/html;base64,";
                    case ".avi":
                        return "data:video/avi;base64,";
                    case ".mp4":
                        return "data:video/mp4;base64,";
                    default:
                        return "data:text/plain;base64,";
                }

            }
            return "data:text/plain;base64,";
        }

        //public string FilesDownloadSaveToDB(string ServerSiteUrl,
        //                                    string LibaryURL,
        //                                    IWBRDataRepository repository,
        //                                    string LinkTableName,
        //                                    int LinkPrimaryKey,
        //                                    System.DateTime ModifiedStartDate,
        //                                    System.DateTime ModifiedEndDate,
        //                                    int CheckDuplicate = 0)
        //{
        //    try
        //    {
        //        Connect(ServerSiteUrl);


        //        FileCollection files = WebClient.GetFolderByServerRelativeUrl(LibaryURL).Files;

        //        clientContext.Load(files);
        //        clientContext.ExecuteQuery();
        //        if (clientContext.HasPendingRequest)
        //            clientContext.ExecuteQuery();

        //        foreach (ClientOM.File file in files)
        //        {
        //            try
        //            {
        //                if (file.TimeLastModified >= ModifiedStartDate && file.TimeLastModified < ModifiedEndDate.AddDays(1))
        //                {

        //                    //FileInformation fileInfo = ClientOM.File.OpenBinaryDirect(clientContext, file.ServerRelativeUrl);
        //                    Microsoft.SharePoint.Client.ClientResult<Stream> mstream = file.OpenBinaryStream();
        //                    clientContext.ExecuteQuery();


        //                    byte[] bytes;
        //                    using (var memoryStream = new MemoryStream())
        //                    {
        //                        mstream.Value.CopyTo(memoryStream);
        //                        bytes = memoryStream.ToArray();
        //                    }
        //                    string base64 = Convert.ToBase64String(bytes);

        //                    var docFileDT = new Models.DocumentBase64ForCreateDTO
        //                    {
        //                        DocFileName = file.Name,
        //                        DocFileDesc = file.Name,
        //                        LinkTableName = LinkTableName,
        //                        LinkPrimaryKey = LinkPrimaryKey,
        //                        DocFileDataBase64 = this.getBase64FileType(file.Name) + base64,
        //                        CheckDuplicate = CheckDuplicate
        //                    };



        //                    repository.cispDocumentAddBase64(docFileDT);

        //                }
        //            }
        //            catch (Exception ex)
        //            {
        //                Debug.WriteLine(ex.Message);
        //            }

        //        }

        //        return "";
        //    }
        //    catch (Exception ex)
        //    {
        //        throw (ex);
        //    }
        //}
    }

    public class AuthenticationManager : IDisposable
    {
        private static readonly HttpClient httpClient = new HttpClient();
        private const string tokenEndpoint = "https://login.microsoftonline.com/common/oauth2/token";

        private const string defaultAADAppId = "6846cc4f-b2b1-4644-ae56-38b4003fb445";

        // Token cache handling
        private static readonly SemaphoreSlim semaphoreSlimTokens = new SemaphoreSlim(1);
        private AutoResetEvent tokenResetEvent = null;
        private readonly ConcurrentDictionary<string, string> tokenCache = new ConcurrentDictionary<string, string>();
        private bool disposedValue;

        internal class TokenWaitInfo
        {
            public RegisteredWaitHandle Handle = null;
        }

        public ClientContext GetContext(Uri web, string userPrincipalName, SecureString userPassword)
        {
            var context = new ClientContext(web);

            context.ExecutingWebRequest += (sender, e) =>
            {
                string accessToken = EnsureAccessTokenAsync(new Uri($"{web.Scheme}://{web.DnsSafeHost}"), userPrincipalName, new System.Net.NetworkCredential(string.Empty, userPassword).Password).GetAwaiter().GetResult();
                e.WebRequestExecutor.RequestHeaders["Authorization"] = "Bearer " + accessToken;
            };

            return context;
        }


        public async Task<string> EnsureAccessTokenAsync(Uri resourceUri, string userPrincipalName, string userPassword)
        {
            string accessTokenFromCache = TokenFromCache(resourceUri, tokenCache);
            if (accessTokenFromCache == null)
            {
                await semaphoreSlimTokens.WaitAsync().ConfigureAwait(false);
                try
                {
                    // No async methods are allowed in a lock section
                    string accessToken = await AcquireTokenAsync(resourceUri, userPrincipalName, userPassword).ConfigureAwait(false);
                    Console.WriteLine($"Successfully requested new access token resource {resourceUri.DnsSafeHost} for user {userPrincipalName}");
                    AddTokenToCache(resourceUri, tokenCache, accessToken);

                    // Register a thread to invalidate the access token once's it's expired
                    tokenResetEvent = new AutoResetEvent(false);
                    TokenWaitInfo wi = new TokenWaitInfo();
                    wi.Handle = ThreadPool.RegisterWaitForSingleObject(
                        tokenResetEvent,
                        async (state, timedOut) =>
                        {
                            if (!timedOut)
                            {
                                TokenWaitInfo internalWaitToken = (TokenWaitInfo)state;
                                if (internalWaitToken.Handle != null)
                                {
                                    internalWaitToken.Handle.Unregister(null);
                                }
                            }
                            else
                            {
                                try
                                {
                                    // Take a lock to ensure no other threads are updating the SharePoint Access token at this time
                                    await semaphoreSlimTokens.WaitAsync().ConfigureAwait(false);
                                    RemoveTokenFromCache(resourceUri, tokenCache);
                                    Console.WriteLine($"Cached token for resource {resourceUri.DnsSafeHost} and user {userPrincipalName} expired");
                                }
                                catch (Exception ex)
                                {
                                    Console.WriteLine($"Something went wrong during cache token invalidation: {ex.Message}");
                                    RemoveTokenFromCache(resourceUri, tokenCache);
                                }
                                finally
                                {
                                    semaphoreSlimTokens.Release();
                                }
                            }
                        },
                        wi,
                        (uint)CalculateThreadSleep(accessToken).TotalMilliseconds,
                        true
                    );

                    return accessToken;

                }
                finally
                {
                    semaphoreSlimTokens.Release();
                }
            }
            else
            {
                Console.WriteLine($"Returning token from cache for resource {resourceUri.DnsSafeHost} and user {userPrincipalName}");
                return accessTokenFromCache;
            }
        }

        private async Task<string> AcquireTokenAsync(Uri resourceUri, string username, string password)
        {
            string resource = $"{resourceUri.Scheme}://{resourceUri.DnsSafeHost}";

            var clientId = defaultAADAppId;
            var body = $"resource={resource}&client_id={clientId}&grant_type=password&username={HttpUtility.UrlEncode(username)}&password={HttpUtility.UrlEncode(password)}";
            using (var stringContent = new StringContent(body, Encoding.UTF8, "application/x-www-form-urlencoded"))
            {

                var result = await httpClient.PostAsync(tokenEndpoint, stringContent).ContinueWith((response) =>
                {
                    return response.Result.Content.ReadAsStringAsync().Result;
                }).ConfigureAwait(false);

                var tokenResult = JsonSerializer.Deserialize<JsonElement>(result);
                var token = tokenResult.GetProperty("access_token").GetString();
                return token;
            }
        }

        private static string TokenFromCache(Uri web, ConcurrentDictionary<string, string> tokenCache)
        {
            if (tokenCache.TryGetValue(web.DnsSafeHost, out string accessToken))
            {
                return accessToken;
            }

            return null;
        }

        private static void AddTokenToCache(Uri web, ConcurrentDictionary<string, string> tokenCache, string newAccessToken)
        {
            if (tokenCache.TryGetValue(web.DnsSafeHost, out string currentAccessToken))
            {
                tokenCache.TryUpdate(web.DnsSafeHost, newAccessToken, currentAccessToken);
            }
            else
            {
                tokenCache.TryAdd(web.DnsSafeHost, newAccessToken);
            }
        }

        private static void RemoveTokenFromCache(Uri web, ConcurrentDictionary<string, string> tokenCache)
        {
            tokenCache.TryRemove(web.DnsSafeHost, out string currentAccessToken);
        }

        private static TimeSpan CalculateThreadSleep(string accessToken)
        {
            var token = new System.IdentityModel.Tokens.Jwt.JwtSecurityToken(accessToken);
            var lease = GetAccessTokenLease(token.ValidTo);
            lease = TimeSpan.FromSeconds(lease.TotalSeconds - TimeSpan.FromMinutes(5).TotalSeconds > 0 ? lease.TotalSeconds - TimeSpan.FromMinutes(5).TotalSeconds : lease.TotalSeconds);
            return lease;
        }

        private static TimeSpan GetAccessTokenLease(DateTime expiresOn)
        {
            DateTime now = DateTime.UtcNow;
            DateTime expires = expiresOn.Kind == DateTimeKind.Utc ? expiresOn : TimeZoneInfo.ConvertTimeToUtc(expiresOn);
            TimeSpan lease = expires - now;
            return lease;
        }

        protected virtual void Dispose(bool disposing)
        {
            if (!disposedValue)
            {
                if (disposing)
                {
                    if (tokenResetEvent != null)
                    {
                        tokenResetEvent.Set();
                        tokenResetEvent.Dispose();
                    }
                }

                disposedValue = true;
            }
        }

        public void Dispose()
        {
            // Do not change this code. Put cleanup code in 'Dispose(bool disposing)' method
            Dispose(disposing: true);
            GC.SuppressFinalize(this);
        }

        
    }
}
