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
            var exchangeClient = await EmailOperations.GetOutlookClientAsync();
            return exchangeClient != null;
        }

        public static async Task<bool> TryGetInboxMessagesAsync()
        {
            var messages = await EmailOperations.GetInboxMessagesAsync();

            return messages != null;
        }

        public static async Task<bool> TryGetMessagesAsync()
        {
            var messages = await EmailOperations.GetMessagesAsync();
            return messages != null;
        }

        public static async Task<bool> TrySendMessageAsync()
        {

            try
            {

                 bool isSent= await EmailOperations.SendMessageAsync(
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
                var newMessageId = await EmailOperations.CreateDraftAsync(
                        STORY_DATA_IDENTIFIER,
                        DEFAULT_MESSAGE_BODY,
                        AuthenticationHelper.LoggedInUserEmail
                    );

                if (newMessageId == null)
                    return false;

                //Cleanup
                await EmailOperations.DeleteMessageAsync(newMessageId);

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

                var newMessageId = await EmailOperations.CreateDraftAndSendAsync(
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
                bool isReplied = await EmailOperations.ReplyMessageAsync(
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

                var newMessageId = await EmailOperations.CreateDraftAndSendAsync(
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
                bool isReplied = await EmailOperations.ReplyAllAsync(
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

                var newMessageId = await EmailOperations.CreateDraftAndSendAsync(
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
                bool isReplied = await EmailOperations.ForwardMessageAsync(
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
                var newMessageId = await EmailOperations.CreateDraftAsync(
                        STORY_DATA_IDENTIFIER,
                        DEFAULT_MESSAGE_BODY,
                        AuthenticationHelper.LoggedInUserEmail
                    );

                if (newMessageId == null)
                    return false;

                // Update the message.
                bool isUpdated = await EmailOperations.UpdateMessageAsync(
                    newMessageId,
                    DEFAULT_MESSAGE_BODY);

                //Cleanup. Comment if you want to verify the update in your Drafts folder.
                await EmailOperations.DeleteMessageAsync(newMessageId);

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

                var newMessageId = await EmailOperations.CreateDraftAndSendAsync(
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
                bool isReplied = await EmailOperations.MoveMessageAsync(
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

                var newMessageId = await EmailOperations.CreateDraftAndSendAsync(
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
                bool isReplied = await EmailOperations.CopyMessageAsync(
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
                var newMessageId = await EmailOperations.CreateDraftAsync(
                        STORY_DATA_IDENTIFIER,
                        DEFAULT_MESSAGE_BODY,
                        AuthenticationHelper.LoggedInUserEmail
                    );

                if (newMessageId == null)
                    return false;

                // Delete the message.
                var isDeleted = await EmailOperations.DeleteMessageAsync(newMessageId);

                return isDeleted;
            }
            catch { return false; }
        }

        public static async Task<bool> TryGetMailFoldersAsync()
        {
            try
            {
                // The example gets the Inbox and its siblings.
                var foldersResults = await EmailOperations.GetMailFoldersAsync();

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
                var folderId = await EmailOperations.CreateMailFolderAsync("Inbox", "FolderToDelete");


                if (!string.IsNullOrEmpty(folderId))
                {
                    //Cleanup
                    await EmailOperations.DeleteMailFolderAsync(folderId);

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
                var folderId = await EmailOperations.CreateMailFolderAsync("Inbox", "FolderToUpdateAndDelete");


                if (!string.IsNullOrEmpty(folderId))
                {

                    bool isFolderUpdated = await EmailOperations.UpdateMailFolderAsync(folderId, "FolderToDelete");

                    //Cleanup
                    await EmailOperations.DeleteMailFolderAsync(folderId);

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
                var folderId = await EmailOperations.CreateMailFolderAsync("Inbox", "FolderToDelete");


                if (!string.IsNullOrEmpty(folderId))
                {

                    bool isFolderMoved = await EmailOperations.MoveMailFolderAsync(folderId, "Drafts");

                    //Cleanup
                    await EmailOperations.DeleteMailFolderAsync(folderId);

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
                var folderId = await EmailOperations.CreateMailFolderAsync("Inbox", "FolderToCopyAndDelete");


                if (!string.IsNullOrEmpty(folderId))
                {

                    string copiedFolderId = await EmailOperations.CopyMailFolderAsync(folderId, "Drafts");

                    if (!string.IsNullOrEmpty(copiedFolderId))
                    {

                        //Cleanup
                        await EmailOperations.DeleteMailFolderAsync(folderId);
                        await EmailOperations.DeleteMailFolderAsync(copiedFolderId);

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
                var folderId = await EmailOperations.CreateMailFolderAsync("Inbox", "FolderToDelete");

                var isFolderDeleted = await EmailOperations.DeleteMailFolderAsync(folderId);
                return isFolderDeleted;
            }
            catch { return false; }
        }

        private static async Task<string> GetSentMessageIdAsync()
        {
            // Search for a maximum of 10 times
            for (int i = 0; i < 10; i++)
            {
                var message = await EmailOperations.GetMessagesAsync(STORY_DATA_IDENTIFIER
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
