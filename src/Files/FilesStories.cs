// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.

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
            var sharepointClient = await FilesSnippets.GetSharePointClientAsync();
            return sharepointClient != null;
        }

        //Files stories

        public static async Task<bool> TryCreateFileAsync()
        {
            // Grab a list of folder items
            var items = await FilesSnippets.GetFolderChildrenAsync("root");
            if (items == null)
                return false;

            var origCount = items.Count;

            var createdFileId = await FilesSnippets.CreateFileAsync(STORY_DATA_IDENTIFIER + "_" + Guid.NewGuid().ToString(), new MemoryStream(Encoding.UTF8.GetBytes("TryAddFileAsync")));
            if (createdFileId == null)
                return false;


            // Grab the files again
            items = await FilesSnippets.GetFolderChildrenAsync("root");
            if (items == null)
                return false;

            // Number of files should have increased by 1
            if (items.Count != origCount + 1)
                return false;

            //Cleanup
            await FilesSnippets.DeleteFileAsync(createdFileId);


            return true;

        }

        public static async Task<bool> TryUpdateFileContentAsync()
        {
            // Add a file & verify
            // Grab a list of files
            var items = await FilesSnippets.GetFolderChildrenAsync("root");
            if (items == null)
                return false;

            var origCount = items.Count;

            // Create a file
            var createdFileId = await FilesSnippets.CreateFileAsync(STORY_DATA_IDENTIFIER + "_" + Guid.NewGuid().ToString(), new MemoryStream(Encoding.UTF8.GetBytes("TryUpdateFileAsync")));
            if (createdFileId == null)
                return false;

            // Grab the files again
            items = await FilesSnippets.GetFolderChildrenAsync("root");
            if (items == null)
                return false;

            // Number of files should have increased by 1
            if (items.Count != origCount + 1)
                return false;

            // Update the content

            string updatedContent = "Updated content";
            var updated = await FilesSnippets.UpdateFileContentAsync(createdFileId, new MemoryStream(Encoding.UTF8.GetBytes(updatedContent)));

            // Download the file and compare with the updated content.

            using (var stream = await FilesSnippets.DownloadFileAsync(createdFileId))
            {
                if (stream == null)
                    return false;

                StreamReader reader = new StreamReader(stream);
                var downloadedString = await reader.ReadToEndAsync();
                if (downloadedString != updatedContent)
                    return false;
            }

            //Cleanup
            await FilesSnippets.DeleteFileAsync(createdFileId);

            return updated;


        }

        public static async Task<bool> TryDownloadFileAsync()
        {

            string fileContents = "TryDownloadFileAsync";

            // Create a file
            var createdFile = await FilesSnippets.CreateFileAsync(STORY_DATA_IDENTIFIER + "_" + Guid.NewGuid().ToString(), new MemoryStream(Encoding.UTF8.GetBytes(fileContents)));
            if (createdFile == null)
                return false;

            // Download the file
            using (var stream = await FilesSnippets.DownloadFileAsync(createdFile))
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
            var items = await FilesSnippets.GetFolderChildrenAsync("root");
            if (items == null)
                return false;

            var origCount = items.Count;

            // Create a file
            var createdFile = await FilesSnippets.CreateFileAsync(STORY_DATA_IDENTIFIER + "_" + Guid.NewGuid().ToString(), new MemoryStream(Encoding.UTF8.GetBytes("CanAddFileAsync")));
            if (createdFile == null)
                return false;

            // Grab the files again
            items = await FilesSnippets.GetFolderChildrenAsync("root");
            if (items == null)
                return false;

            // Number of files should have increased by 1
            if (items.Count != origCount + 1)
                return false;

            // Delete our test file
            await FilesSnippets.DeleteFileAsync(createdFile);

            //Grab the files again
            items = await FilesSnippets.GetFolderChildrenAsync("root");
            if (items == null)
                return false;

            // Number of files should be back at the original count
            if (items.Count != origCount)
                return false;

            return true;
        }

        public static async Task<bool> TryCopyFileAsync()
        {

            // Grab the root folder.
            var items = await FilesSnippets.GetFolderChildrenAsync("root");
            if (items == null)
                return false;

            // Create a new file.
            var createdFileId = await FilesSnippets.CreateFileAsync(STORY_DATA_IDENTIFIER + "_" + Guid.NewGuid().ToString(), new MemoryStream(Encoding.UTF8.GetBytes("TryAddFileAsync")));
            if (createdFileId == null)
                return false;

            // Create a new folder in the root folder.
            var folder = await FilesSnippets.CreateFolderAsync(STORY_DATA_IDENTIFIER, "root");

            // Copy the new file into the new folder.
            var copiedFileId = await FilesSnippets.CopyFileAsync(createdFileId, folder.Id);

            // Clean up. 
            // Comment out if you want to see the file, the folder, and the copied file.
            await FilesSnippets.DeleteFileAsync(createdFileId);

            // Deleting the folder also deletes the file copied into it.
            await FilesSnippets.DeleteFolderAsync(folder.Id);


            return true;

        }

        public static async Task<bool> TryRenameFileAsync()
        {

            string newFileName = "updated name";

            // Create a file
            var createdFileId = await FilesSnippets.CreateFileAsync(STORY_DATA_IDENTIFIER + "_" + Guid.NewGuid().ToString(), new MemoryStream(Encoding.UTF8.GetBytes("TryUpdateFileAsync")));
            if (createdFileId == null)
                return false;

            var fileName = await FilesSnippets.RenameFileAsync(createdFileId, "updated name");

            if (fileName != newFileName)
                return false;

            //Cleanup

            await FilesSnippets.DeleteFileAsync(createdFileId);

            return true;


        }

        //Folders stories

        public static async Task<bool> TryGetFolderChildrenAsync()
        {
            var items = await FilesSnippets.GetFolderChildrenAsync("root");
            return items != null;
        }

        public static async Task<bool> TryCreateFolderAsync()
        {

            var folder = await FilesSnippets.CreateFolderAsync(STORY_DATA_IDENTIFIER, "root");

            //Cleanup. Comment if you want to see the new folder under your root folder.
            await FilesSnippets.DeleteFolderAsync(folder.Id);

            return folder != null;
        }

        public static async Task<bool> TryDeleteFolderAsync()
        {

            var folder = await FilesSnippets.CreateFolderAsync(STORY_DATA_IDENTIFIER, "root");


            var result = await FilesSnippets.DeleteFolderAsync(folder.Id);

            return result;

        }

    }
}

//********************************************************* 
// 
//O365-Win-Snippets, https://github.com/OfficeDev/O365-Win-Snippets
//
//Copyright (c) Microsoft Corporation
//All rights reserved. 
//
// MIT License:
// Permission is hereby granted, free of charge, to any person obtaining
// a copy of this software and associated documentation files (the
// ""Software""), to deal in the Software without restriction, including
// without limitation the rights to use, copy, modify, merge, publish,
// distribute, sublicense, and/or sell copies of the Software, and to
// permit persons to whom the Software is furnished to do so, subject to
// the following conditions:

// The above copyright notice and this permission notice shall be
// included in all copies or substantial portions of the Software.

// THE SOFTWARE IS PROVIDED ""AS IS"", WITHOUT WARRANTY OF ANY KIND,
// EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF
// MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND
// NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE
// LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION
// OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION
// WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
// 
//********************************************************* 