using Microsoft.IdentityModel.Clients.ActiveDirectory;
using Microsoft.Office365.Discovery;
using Microsoft.Office365.SharePoint.CoreServices;
using Microsoft.Office365.SharePoint.FileServices;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace O365_Win_Snippets
{
    class FilesOperations
    {
        private static SharePointClient _sharePointClient = null;

        /// <summary>
        /// Checks that an OutlookServicesClient object is available. 
        /// </summary>
        /// <returns>The OutlookServicesClient object. </returns>
        public static async Task<SharePointClient> GetSharePointClientAsync()
        {
            if (_sharePointClient != null && !String.IsNullOrEmpty(AuthenticationHelper.LastAuthority))
            {
                Debug.WriteLine("Got a SharePoint client for Files.");
                return _sharePointClient;
            }
            else
            {
                try
                {
                    //First, look for the authority used during the last authentication.
                    //If that value is not populated, use CommonAuthority.
                    string authority = null;

                    if (String.IsNullOrEmpty(AuthenticationHelper.LastAuthority))
                    {
                        authority = AuthenticationHelper.CommonAuthority;
                    }
                    else
                    {
                        authority = AuthenticationHelper.LastAuthority;
                    }

                    // Create an AuthenticationContext using this authority.
                    AuthenticationHelper._authenticationContext = new AuthenticationContext(authority);

                    // Set the value of _authenticationContext.UseCorporateNetwork to true so that you 
                    // can use this app inside a corporate intranet. If the value of UseCorporateNetwork 
                    // is true, you also need to add the Enterprise Authentication, Private Networks, and
                    // Shared User Certificates capabilities in the Package.appxmanifest file.
                    AuthenticationHelper._authenticationContext.UseCorporateNetwork = true;

                    //See the Discovery Service Sample (https://github.com/OfficeDev/Office365-Discovery-Service-Sample)
                    //for an approach that improves performance by storing the discovery service information in a cache.
                    DiscoveryClient discoveryClient = new DiscoveryClient(
                        async () => await AuthenticationHelper.GetTokenHelperAsync(AuthenticationHelper._authenticationContext, AuthenticationHelper.DiscoveryResourceId));

                    // Get the specified capability ("Calendar").
                    CapabilityDiscoveryResult result =
                        await discoveryClient.DiscoverCapabilityAsync("MyFiles");

                    var token = await AuthenticationHelper.GetTokenHelperAsync(AuthenticationHelper._authenticationContext, result.ServiceResourceId);
                    // Check the token
                    if (String.IsNullOrEmpty(token))
                    {
                        // User cancelled sign-in
                        return null;
                    }
                    else
                    {

                        _sharePointClient = new SharePointClient(
                            result.ServiceEndpointUri,
                            async () => await AuthenticationHelper.GetTokenHelperAsync(AuthenticationHelper._authenticationContext, result.ServiceResourceId));
                        Debug.WriteLine("Got a SharePoint client for Files.");
                        return _sharePointClient;
                    }
                }
                // The following is a list of exceptions you should consider handling in your app.
                // In the case of this sample, the exceptions are handled by returning null upstream. 
                catch (DiscoveryFailedException dfe)
                {
                    Debug.WriteLine(dfe.Message);
                }
                catch (ArgumentException ae)
                {
                    Debug.WriteLine(ae.Message);
                }

                AuthenticationHelper._authenticationContext.TokenCache.Clear();

                return null;
            }
        }

        //Files operations

        public static async Task<string> CreateFileAsync(string fileName, Stream fileContent)
        {
            try
            {
                var sharePointClient = await GetSharePointClientAsync();

                File newFile = new File
                {
                    Name = fileName
                };

                await sharePointClient.Files.AddItemAsync(newFile);
                await sharePointClient.Files.GetById(newFile.Id).ToFile().UploadAsync(fileContent);

                Debug.WriteLine("Created a file: " + newFile.Id);

                return newFile.Id;

            }
            catch { return null; }
        }

        public static async Task<bool> UpdateFileContentAsync(string Id, Stream fileContent)
        {
            try
            {
                var sharePointClient = await GetSharePointClientAsync();

                //update the file with the content 
                await sharePointClient.Files.GetById(Id).ToFile().UploadAsync(fileContent);

                Debug.WriteLine("Updated file content: " + Id);

                return true;

            }
            catch { return false; }
        }

        public static async Task<Stream> DownloadFileAsync(string Id)
        {

            try
            {
                var sharePointClient = await GetSharePointClientAsync();

                var stream = await sharePointClient.Files.GetById(Id).ToFile().DownloadAsync();

                Debug.WriteLine("Downloaded a file: " + Id);

                return stream;
            }
            catch { return null; }
        }

        public static async Task<bool> DeleteFileAsync(string Id)
        {
            try
            {
                var sharePointClient = await GetSharePointClientAsync();
                var file = await sharePointClient.Files.GetById(Id).ToFile().ExecuteAsync();
                await file.DeleteAsync();

                Debug.WriteLine("Deleted a file: " + Id);

                return true;
            }
            catch { return false; }
        }

        public static async Task<string> CopyFileAsync(string fileId, string destinationFolderId)
        {
            try
            {
                var sharePointClient = await GetSharePointClientAsync();

                var copiedFile = await sharePointClient.Files.GetById(fileId).ToFile().CopyAsync(destinationFolderId, null, null);

                Debug.WriteLine("Copied file to folder.");

                return copiedFile.Id;
            }
            catch
            {
                return null;
            }
        }

        public static async Task<string> RenameFileAsync(string fileId, string newName)
        {

            try
            {
                var sharePointClient = await GetSharePointClientAsync();

                var file = await sharePointClient.Files.GetById(fileId).ToFile().ExecuteAsync();

                file.Name = newName;
                await file.UpdateAsync();

                Debug.WriteLine("Renamed a file: " + fileId);

                return file.Name;
            }
            catch { return null; }
        }

        //Folders operations
        public static async Task<List<IItem>> GetFolderChildrenAsync(string folderId)
        {
            var sharePointClient = await GetSharePointClientAsync();
            try
            {
                var items = await sharePointClient.Files.GetById(folderId).ToFolder().Children.ExecuteAsync();

                Debug.WriteLine("First child of " + folderId + ": " + items.CurrentPage[0].Id);

                return items.CurrentPage.ToList();
            }
            catch { return null; }
        }

        public static async Task<Folder> CreateFolderAsync(string folderName, string parentFolderId)
        {
            try
            {
                var sharePointClient = await GetSharePointClientAsync();
                Folder newFolder = new Folder
                {
                    Name = folderName
                };

                await sharePointClient.Files.GetById(parentFolderId).ToFolder().Children.AddItemAsync(newFolder);
                var newItem = await sharePointClient.Files.GetById(newFolder.Id).ToFolder().ExecuteAsync();

                Debug.WriteLine("Created a folder: " + newItem.Id);

                return (Folder)newItem;
            }
            catch { return null; }
        }


        public static async Task<bool> DeleteFolderAsync(string folderId)
        {
            try
            {
                var sharePointClient = await GetSharePointClientAsync();
                var item = await sharePointClient.Files.GetById(folderId).ToFolder().ExecuteAsync();

                await item.DeleteAsync();

                return true;

            }
            catch { return false; }
        }

    }
}
