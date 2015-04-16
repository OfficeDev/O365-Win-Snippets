// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.

using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using System.Linq;
using Microsoft.Office365.OutlookServices;
using Microsoft.Office365.OutlookServices.Extensions;
using Microsoft.OData.Client;
using Microsoft.OData.Core;


namespace O365_Win_Snippets
{
    class EmailStories
    {
        private static readonly string STORY_DATA_IDENTIFIER = Guid.NewGuid().ToString();
        private static readonly string DEFAULT_MESSAGE_BODY = "This message was sent from the <a href='https://github.com/OfficeDev/O365-Win-Snippets' >Office 365 Windows Snippets project</a>";

        public static async Task<bool> TryGetOutlookClientAsync()
        {
            var exchangeClient = await EmailSnippets.GetOutlookClientAsync();
            return exchangeClient != null;
        }

        public static async Task<bool> TryGetInboxMessagesAsync()
        {
            var messages = await EmailSnippets.GetInboxMessagesAsync();

            return messages != null;
        }

        public static async Task<bool> TryGetMessagesAsync()
        {
            var messages = await EmailSnippets.GetMessagesAsync();
            return messages != null;
        }

        public static async Task<bool> TrySendMessageAsync()
        {

            try
            {

                 bool isSent= await EmailSnippets.SendMessageAsync(
                    STORY_DATA_IDENTIFIER,
                    DEFAULT_MESSAGE_BODY,
                    AuthenticationHelper.LoggedInUserEmail
                    );

                 return isSent;

            }

            catch { return false; }


        }

        public static async Task<bool> TryCreateDraftAsync()
        {
            try
            {
                // Create the draft message.
                var newMessageId = await EmailSnippets.CreateDraftAsync(
                        STORY_DATA_IDENTIFIER,
                        DEFAULT_MESSAGE_BODY,
                        AuthenticationHelper.LoggedInUserEmail
                    );

                if (newMessageId == null)
                    return false;

                //Cleanup
                await EmailSnippets.DeleteMessageAsync(newMessageId);

                    return true;
                
            }

            catch { return false; }
        }

        public static async Task<bool> TryReplyMessageAsync()
        {

            try
            {
                // Create a draft message and then send it. If you send the message without first creating a draft, you can't easily retrieve 
                // the message Id.

                var newMessageId = await EmailSnippets.CreateDraftAndSendAsync(
                        STORY_DATA_IDENTIFIER,
                        DEFAULT_MESSAGE_BODY,
                        AuthenticationHelper.LoggedInUserEmail
                    );

                if (newMessageId == null)
                    return false;

                // Find the sent message.
                var sentMessageId = await GetSentMessageIdAsync();
                if (String.IsNullOrEmpty(sentMessageId))
                    return false;

                // Reply to the message.
                bool isReplied = await EmailSnippets.ReplyMessageAsync(
                    sentMessageId,
                    DEFAULT_MESSAGE_BODY);

                return isReplied;

            }
            catch { return false; }
        }

        public static async Task<bool> TryReplyAllAsync()
        {

            try
            {
                // Create a draft message and then send it. If you send the message without first creating a draft, you can't easily retrieve 
                // the message Id.

                var newMessageId = await EmailSnippets.CreateDraftAndSendAsync(
                        STORY_DATA_IDENTIFIER,
                        DEFAULT_MESSAGE_BODY,
                        AuthenticationHelper.LoggedInUserEmail
                    );

                if (newMessageId == null)
                    return false;


                // Find the sent message.
                var sentMessageId = await GetSentMessageIdAsync();
                if (String.IsNullOrEmpty(sentMessageId))
                    return false;

                // Reply to the message.
                bool isReplied = await EmailSnippets.ReplyAllAsync(
                                sentMessageId,
                                DEFAULT_MESSAGE_BODY);

                return isReplied;

            }
            catch { return false; }
        }

        public static async Task<bool> TryForwardMessageAsync()
        {

            try
            {
                // Create a draft message and then send it. If you send the message without first creating a draft, you can't easily retrieve 
                // the message Id.

                var newMessageId = await EmailSnippets.CreateDraftAndSendAsync(
                        STORY_DATA_IDENTIFIER,
                        DEFAULT_MESSAGE_BODY,
                        AuthenticationHelper.LoggedInUserEmail
                    );

                if (newMessageId == null)
                    return false;

                // Find the sent message.
                var sentMessageId = await GetSentMessageIdAsync();
                if (String.IsNullOrEmpty(sentMessageId))
                    return false;

                // Reply to the message.
                bool isReplied = await EmailSnippets.ForwardMessageAsync(
                               sentMessageId,
                               DEFAULT_MESSAGE_BODY,
                               AuthenticationHelper.LoggedInUserEmail);

                return isReplied;

            }
            catch { return false; }
        }

        public static async Task<bool> TryUpdateMessageAsync()
        {

            try
            {
                // Create a draft message. If you send the message without first creating a draft, you can't easily retrieve the message Id.
                var newMessageId = await EmailSnippets.CreateDraftAsync(
                        STORY_DATA_IDENTIFIER,
                        DEFAULT_MESSAGE_BODY,
                        AuthenticationHelper.LoggedInUserEmail
                    );

                if (newMessageId == null)
                    return false;

                // Update the message.
                bool isUpdated = await EmailSnippets.UpdateMessageAsync(
                    newMessageId,
                    DEFAULT_MESSAGE_BODY);

                //Cleanup. Comment if you want to verify the update in your Drafts folder.
                await EmailSnippets.DeleteMessageAsync(newMessageId);

                return isUpdated;
            }
            catch { return false; }
        }

        public static async Task<bool> TryMoveMessageAsync()
        {

            try
            {
                // Create a draft message and then send it. If you send the message without first creating a draft, you can't easily retrieve 
                // the message Id.

                var newMessageId = await EmailSnippets.CreateDraftAndSendAsync(
                        STORY_DATA_IDENTIFIER,
                        DEFAULT_MESSAGE_BODY,
                        AuthenticationHelper.LoggedInUserEmail
                    );

                if (newMessageId == null)
                    return false;

                // Find the sent message.
                var sentMessageId = await GetSentMessageIdAsync();
                if (String.IsNullOrEmpty(sentMessageId))
                    return false;

                // Reply to the message.
                bool isReplied = await EmailSnippets.MoveMessageAsync(
                                sentMessageId,
                                "Inbox",
                                "Drafts");

                return isReplied;

            }
            catch { return false; }
        }

        public static async Task<bool> TryCopyMessageAsync()
        {

            try
            {
                // Create a draft message and then send it. If you send the message without first creating a draft, you can't easily retrieve 
                // the message Id.

                var newMessageId = await EmailSnippets.CreateDraftAndSendAsync(
                        STORY_DATA_IDENTIFIER,
                        DEFAULT_MESSAGE_BODY,
                        AuthenticationHelper.LoggedInUserEmail
                    );

                if (newMessageId == null)
                    return false;
                // Find the sent message.
                var sentMessageId = await GetSentMessageIdAsync();
                if (String.IsNullOrEmpty(sentMessageId))
                    return false;

                // Reply to the message.
                bool isReplied = await EmailSnippets.CopyMessageAsync(
                                sentMessageId,
                                "Inbox",
                                "Drafts");

                return isReplied;

            }
            catch { return false; }
        }

        public static async Task<bool> TryDeleteMessageAsync()
        {

            try
            {
                // Create a draft message. If you send the message without first creating a draft, you can't easily retrieve the message Id.
                var newMessageId = await EmailSnippets.CreateDraftAsync(
                        STORY_DATA_IDENTIFIER,
                        DEFAULT_MESSAGE_BODY,
                        AuthenticationHelper.LoggedInUserEmail
                    );

                if (newMessageId == null)
                    return false;

                // Delete the message.
                var isDeleted = await EmailSnippets.DeleteMessageAsync(newMessageId);

                return isDeleted;
            }
            catch { return false; }
        }

        public static async Task<bool> TryGetMailFoldersAsync()
        {
            try
            {
                // The example gets the Inbox and its siblings.
                var foldersResults = await EmailSnippets.GetMailFoldersAsync();

                foreach (var folder in foldersResults.CurrentPage)
                {
                    if ((folder.DisplayName == "Inbox")
                        || (folder.DisplayName == "Drafts")
                        || (folder.DisplayName == "DeletedItems")
                        || (folder.DisplayName == "SentItems"))
                        return true;
                }

                return false;

            }
            catch { return false; }
        }

        public static async Task<bool> TryCreateMailFolderAsync()
        {
            try
            {
                var folderId = await EmailSnippets.CreateMailFolderAsync("Inbox", "FolderToDelete");


                if (!string.IsNullOrEmpty(folderId))
                {
                    //Cleanup
                    await EmailSnippets.DeleteMailFolderAsync(folderId);

                    return true;
                }

                return false;
            }

            catch { return false; }
        }

        public static async Task<bool> TryUpdateMailFolderAsync()
        {
            try
            {
                var folderId = await EmailSnippets.CreateMailFolderAsync("Inbox", "FolderToUpdateAndDelete");


                if (!string.IsNullOrEmpty(folderId))
                {

                    bool isFolderUpdated = await EmailSnippets.UpdateMailFolderAsync(folderId, "FolderToDelete");

                    //Cleanup
                    await EmailSnippets.DeleteMailFolderAsync(folderId);

                    return isFolderUpdated;
                }

                return false;
            }

            catch { return false; }
        }

        public static async Task<bool> TryMoveMailFolderAsync()
        {
            try
            {
                var folderId = await EmailSnippets.CreateMailFolderAsync("Inbox", "FolderToDelete");


                if (!string.IsNullOrEmpty(folderId))
                {

                    bool isFolderMoved = await EmailSnippets.MoveMailFolderAsync(folderId, "Drafts");

                    //Cleanup
                    await EmailSnippets.DeleteMailFolderAsync(folderId);

                    return isFolderMoved;
                }

                return false;
            }

            catch { return false; }
        }

        public static async Task<bool> TryCopyMailFolderAsync()
        {
            try
            {
                var folderId = await EmailSnippets.CreateMailFolderAsync("Inbox", "FolderToCopyAndDelete");


                if (!string.IsNullOrEmpty(folderId))
                {

                    string copiedFolderId = await EmailSnippets.CopyMailFolderAsync(folderId, "Drafts");

                    if (!string.IsNullOrEmpty(copiedFolderId))
                    {

                        //Cleanup
                        await EmailSnippets.DeleteMailFolderAsync(folderId);
                        await EmailSnippets.DeleteMailFolderAsync(copiedFolderId);

                        return true;
                    }
                }

                return false;
            }

            catch { return false; }
        }

        public static async Task<bool> TryDeleteMailFolderAsync()
        {
            try
            {
                var folderId = await EmailSnippets.CreateMailFolderAsync("Inbox", "FolderToDelete");

                var isFolderDeleted = await EmailSnippets.DeleteMailFolderAsync(folderId);
                return isFolderDeleted;
            }
            catch { return false; }
        }

        private static async Task<string> GetSentMessageIdAsync()
        {
            // Search for a maximum of 10 times
            for (int i = 0; i < 10; i++)
            {
                var message = await EmailSnippets.GetMessagesAsync(STORY_DATA_IDENTIFIER
                                              , DateTimeOffset.Now.Subtract(new TimeSpan(0, 1, 0)));
                if (message != null)
                    return message.Id;

                // Delay before trying again. Increase this value if you connection to the server
                // is slow and causes false results. 
                await Task.Delay(200);

            }

            // Couldn't find the sent message
            return string.Empty;
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