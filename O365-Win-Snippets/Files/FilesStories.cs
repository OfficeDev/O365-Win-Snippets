using Microsoft.Office365.SharePoint.FileServices;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace O365_Win_Snippets
{
    class FilesStories
    {
        private static readonly string STORY_DATA_IDENTIFIER = Guid.NewGuid().ToString();

        public static async Task<bool> TryGetSharePointClientAsync()
        {
            var sharepointClient = await FilesOperations.GetSharePointClientAsync();
            return sharepointClient != null;
        }

        //Files stories

        public static async Task<bool> TryCreateFileAsync()
        {
            // Grab a list of folder items
            var items = await FilesOperations.GetFolderChildrenAsync("root");
            if (items == null)
                return false;

            var origCount = items.Count;

            var createdFileId = await FilesOperations.CreateFileAsync(STORY_DATA_IDENTIFIER + "_" + Guid.NewGuid().ToString(), new MemoryStream(Encoding.UTF8.GetBytes("TryAddFileAsync")));
            if (createdFileId == null)
                return false;


            // Grab the files again
            items = await FilesOperations.GetFolderChildrenAsync("root");
            if (items == null)
                return false;

            // Number of files should have increased by 1
            if (items.Count != origCount + 1)
                return false;

            //Cleanup
            await FilesOperations.DeleteFileAsync(createdFileId);


            return true;

        }

        public static async Task<bool> TryUpdateFileContentAsync()
        {
            // Add a file & verify
            // Grab a list of files
            var items = await FilesOperations.GetFolderChildrenAsync("root");
            if (items == null)
                return false;

            var origCount = items.Count;

            // Create a file
            var createdFileId = await FilesOperations.CreateFileAsync(STORY_DATA_IDENTIFIER + "_" + Guid.NewGuid().ToString(), new MemoryStream(Encoding.UTF8.GetBytes("TryUpdateFileAsync")));
            if (createdFileId == null)
                return false;

            // Grab the files again
            items = await FilesOperations.GetFolderChildrenAsync("root");
            if (items == null)
                return false;

            // Number of files should have increased by 1
            if (items.Count != origCount + 1)
                return false;

            // Update the content

            string updatedContent = "Updated content";
            var updated = await FilesOperations.UpdateFileContentAsync(createdFileId, new MemoryStream(Encoding.UTF8.GetBytes(updatedContent)));

            // Download the file and compare with the updated content.

            using (var stream = await FilesOperations.DownloadFileAsync(createdFileId))
            {
                if (stream == null)
                    return false;

                StreamReader reader = new StreamReader(stream);
                var downloadedString = await reader.ReadToEndAsync();
                if (downloadedString != updatedContent)
                    return false;
            }

            //Cleanup
            await FilesOperations.DeleteFileAsync(createdFileId);

            return updated;

 
        }

        public static async Task<bool> TryDownloadFileAsync()
        {

            string fileContents = "TryDownloadFileAsync";

            // Create a file
            var createdFile = await FilesOperations.CreateFileAsync(STORY_DATA_IDENTIFIER + "_" + Guid.NewGuid().ToString(), new MemoryStream(Encoding.UTF8.GetBytes(fileContents)));
            if (createdFile == null)
                return false;

            // Download the file
            using (var stream = await FilesOperations.DownloadFileAsync(createdFile))
            {
                if (stream == null)
                    return false;

                StreamReader reader = new StreamReader(stream);
                var downloadedString = await reader.ReadToEndAsync();
                if (downloadedString != fileContents)
                    return false;
            }

            return true;
        }

        public static async Task<bool> TryDeleteFileAsync()
        {
            // Grab a list of files
            var items = await FilesOperations.GetFolderChildrenAsync("root");
            if (items == null)
                return false;

            var origCount = items.Count;

            // Create a file
            var createdFile = await FilesOperations.CreateFileAsync(STORY_DATA_IDENTIFIER + "_" + Guid.NewGuid().ToString(), new MemoryStream(Encoding.UTF8.GetBytes("CanAddFileAsync")));
            if (createdFile == null)
                return false;

            // Grab the files again
            items = await FilesOperations.GetFolderChildrenAsync("root");
            if (items == null)
                return false;

            // Number of files should have increased by 1
            if (items.Count != origCount + 1)
                return false;

            // Delete our test file
            await FilesOperations.DeleteFileAsync(createdFile);

            //Grab the files again
            items = await FilesOperations.GetFolderChildrenAsync("root");
            if (items == null)
                return false;

            // Number of files should be back at the original count
            if (items.Count != origCount)
                return false;

            return true;
        }

        //Folders stories

        public static async Task<bool> TryGetFolderChildrenAsync()
        {
            var items = await FilesOperations.GetFolderChildrenAsync("root");
            return items != null;
        }

        public static async Task<bool> TryCreateFolderAsync()
        {
            try
            {
                var folder = await FilesOperations.CreateFolderAsync(STORY_DATA_IDENTIFIER, "root");

                //Cleanup. Comment if you want to see the new folder under your root folder.
                await FilesOperations.DeleteFolderAsync(folder.Id);

                return folder != null;
            }

            catch { return false;  }
        }

        public static async Task<bool> TryDeleteFolderAsync()
        {
            try
            {
                var folder = await FilesOperations.CreateFolderAsync(STORY_DATA_IDENTIFIER, "root");


                var result = await FilesOperations.DeleteFolderAsync(folder.Id);
                if (!result)
                    return false;

                return true;
            }

            catch { return false; }

        }

        public static async Task<bool> TryCopyFileAsync()
        {
            try
            {
                // Grab the root folder.
                var items = await FilesOperations.GetFolderChildrenAsync("root");
                if (items == null)
                    return false;

                // Create a new file.
                var createdFileId = await FilesOperations.CreateFileAsync(STORY_DATA_IDENTIFIER + "_" + Guid.NewGuid().ToString(), new MemoryStream(Encoding.UTF8.GetBytes("TryAddFileAsync")));
                if (createdFileId == null)
                    return false;

                // Create a new folder in the root folder.
                var folder = await FilesOperations.CreateFolderAsync(STORY_DATA_IDENTIFIER, "root");

                // Copy the new file into the new folder.
                var copiedFileId = await FilesOperations.CopyFileAsync(createdFileId, folder.Id);

                // Clean up
                // Comment out if you want to see the file, the folder, and the copied file.
                await FilesOperations.DeleteFileAsync(createdFileId);
                await FilesOperations.DeleteFolderAsync(folder.Id);
                await FilesOperations.DeleteFileAsync(copiedFileId);

                return true;
            }

            catch { return false; }
        }

    }
}
