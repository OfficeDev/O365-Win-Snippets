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
                    "This message was sent from the <a href='https://github.com/OfficeDev/O365-Win-Snippets' >Office 365 Windows Snippets project</a>",
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
                        "This message was sent from the <a href='https://github.com/OfficeDev/O365-Win-Snippets' >Office 365 Windows Snippets project</a>",
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
                        "This message was sent from the <a href='https://github.com/OfficeDev/O365-Win-Snippets' >Office 365 Windows Snippets project</a>",
                        AuthenticationHelper.LoggedInUserEmail
                    );

                if (newMessageId == null)
                    return false;

                // Check the inbox until the sent message appears. You might need to increase the number of attempts
                // if it is taking especially long for the message to appear in the Inbox.
                for (int i = 0; i < 10; i++)
                {
                    var messages = await EmailOperations.GetMessagesAsync();

                    foreach (IMessage message in messages)
                    {
                        if (message.Subject == STORY_DATA_IDENTIFIER)
                        {
                            // Reply to the message.
                            bool isReplied = await EmailOperations.ReplyMessageAsync(
                                message.Id,
                                "This reply was sent from the <a href='https://github.com/OfficeDev/O365-Win-Snippets' >Office 365 Windows Snippets project</a>");
                            return isReplied;
                        }
                    }
                }

                return false;

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
                        "This message was sent from the <a href='https://github.com/OfficeDev/O365-Win-Snippets' >Office 365 Windows Snippets project</a>",
                        AuthenticationHelper.LoggedInUserEmail
                    );

                if (newMessageId == null)
                    return false;

                // Check the inbox until the sent message appears. You might need to increase the number of attempts
                // if it is taking especially long for the message to appear in the Inbox.
                for (int i = 0; i < 10; i++)
                {
                    var messages = await EmailOperations.GetMessagesAsync();

                    foreach (IMessage message in messages)
                    {
                        if (message.Subject == STORY_DATA_IDENTIFIER)
                        {
                            // Reply to the message.
                            bool isReplied = await EmailOperations.ReplyAllAsync(
                                message.Id,
                                "This reply all was sent from the <a href='https://github.com/OfficeDev/O365-Win-Snippets' >Office 365 Windows Snippets project</a>");
                            return isReplied;
                        }
                    }
                }

                return false;

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
                        "This message was sent from the <a href='https://github.com/OfficeDev/O365-Win-Snippets' >Office 365 Windows Snippets project</a>",
                        AuthenticationHelper.LoggedInUserEmail
                    );

                if (newMessageId == null)
                    return false;

                // Check the inbox until the sent message appears. You might need to increase the number of attempts
                // if it is taking especially long for the message to appear in the Inbox.
                for (int i = 0; i < 10; i++)
                {
                    var messages = await EmailOperations.GetMessagesAsync();

                    foreach (IMessage message in messages)
                    {
                        if (message.Subject == STORY_DATA_IDENTIFIER)
                        {
                            // Reply to the message.
                            bool isReplied = await EmailOperations.ForwardMessageAsync(
                                message.Id,
                                "This forward was sent from the <a href='https://github.com/OfficeDev/O365-Win-Snippets' >Office 365 Windows Snippets project</a>",
                                AuthenticationHelper.LoggedInUserEmail);
                            return isReplied;
                        }
                    }
                }

                return false;

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
                        "This message was sent from the <a href='https://github.com/OfficeDev/O365-Win-Snippets' >Office 365 Windows Snippets project</a>",
                        AuthenticationHelper.LoggedInUserEmail
                    );

                if (newMessageId == null)
                    return false;

                // Update the message.
                bool isUpdated = await EmailOperations.UpdateMessageAsync(
                    newMessageId,
                    "This message was updated by the <a href='https://github.com/OfficeDev/O365-Win-Snippets' >Office 365 Windows Snippets project</a>");

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
                        "This message was sent from the <a href='https://github.com/OfficeDev/O365-Win-Snippets' >Office 365 Windows Snippets project</a>",
                        AuthenticationHelper.LoggedInUserEmail
                    );

                if (newMessageId == null)
                    return false;

                // Check the inbox until the sent message appears. You might need to increase the number of attempts
                // if it is taking especially long for the message to appear in the Inbox.
                for (int i = 0; i < 10; i++)
                {
                    var messages = await EmailOperations.GetMessagesAsync();

                    foreach (IMessage message in messages)
                    {
                        if (message.Subject == STORY_DATA_IDENTIFIER)
                        {
                            // Reply to the message.
                            bool isReplied = await EmailOperations.MoveMessageAsync(
                                message.Id,
                                "Inbox",
                                "Drafts");
                            return isReplied;
                        }
                    }
                }

                return false;

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
                        "This message was sent from the <a href='https://github.com/OfficeDev/O365-Win-Snippets' >Office 365 Windows Snippets project</a>",
                        AuthenticationHelper.LoggedInUserEmail
                    );

                if (newMessageId == null)
                    return false;

                // Check the inbox until the sent message appears. You might need to increase the number of attempts
                // if it is taking especially long for the message to appear in the Inbox.
                for (int i = 0; i < 10; i++)
                {
                    var messages = await EmailOperations.GetMessagesAsync();

                    foreach (IMessage message in messages)
                    {
                        if (message.Subject == STORY_DATA_IDENTIFIER)
                        {
                            // Reply to the message.
                            bool isReplied = await EmailOperations.CopyMessageAsync(
                                message.Id,
                                "Inbox",
                                "Drafts");
                            return isReplied;
                        }
                    }
                }

                return false;

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
                        "This message was sent from the <a href='https://github.com/OfficeDev/O365-Win-Snippets' >Office 365 Windows Snippets project</a>",
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

    }
}
