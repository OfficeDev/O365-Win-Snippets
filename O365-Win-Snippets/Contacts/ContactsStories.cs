using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace O365_Win_Snippets
{
    public class ContactsStories
    {
        private static readonly string STORY_DATA_IDENTIFIER = Guid.NewGuid().ToString();

        public static async Task<bool> TryGetOutlookClientAsync()
        {
            var outlookClient = await ContactsSnippets.GetOutlookClientAsync();
            return outlookClient != null;
        }

        public static async Task<bool> TryGetContactsAsync()
        {
            var contacts = await ContactsSnippets.GetContactsPageAsync();

            return contacts != null;
        }

        public static async Task<bool> TryGetContactAsync()
        {
            var newContact = await ContactsSnippets.AddContactItemAsync(
                Guid.NewGuid().ToString(),
                Guid.NewGuid().ToString(),
                STORY_DATA_IDENTIFIER,
                Guid.NewGuid().ToString(),
                "a@b.com",
                Guid.NewGuid().ToString(),
                Guid.NewGuid().ToString());

            var contact = await ContactsSnippets.GetContactAsync(newContact.Id);

            //Cleanup

            await ContactsSnippets.DeleteContactAsync(newContact.Id);

            return contact != null;
        }

        public static async Task<bool> TryAddNewContactAsync()
        {
            var newContact = await ContactsSnippets.AddContactItemAsync(
                Guid.NewGuid().ToString(),
                Guid.NewGuid().ToString(),
                STORY_DATA_IDENTIFIER,
                Guid.NewGuid().ToString(),
                "a@b.com",
                Guid.NewGuid().ToString(),
                Guid.NewGuid().ToString());

            //Cleanup

            await ContactsSnippets.DeleteContactAsync(newContact.Id);

            return newContact != null;
        }

        public static async Task<bool> TryUpdateContactAsync()
        {
            var testContact = await ContactsSnippets.AddContactItemAsync(
                "FileAsValue",
                "FirstNameValue",
                STORY_DATA_IDENTIFIER,
                "JobTitleValue",
                "a@b.com",
                "WorkPhoneValue",
                "MobilePhoneValue");

            // Verify a valid contact id was returned
            if (testContact == null)
                return false;


            // Update our contact
            await ContactsSnippets.UpdateContactItemAsync(
                 testContact.Id,
                 "NewFileAsValue",
                "FirstNameValue",
                STORY_DATA_IDENTIFIER,
                "NewJobTitleValue",
                "a@b.com",
                "WorkPhoneValue",
                "MobilePhoneValue",
                null);

            var contactCheck = await ContactsSnippets.GetContactAsync(testContact.Id);
            if (contactCheck == null)
                return false;

            //Cleanup

            await ContactsSnippets.DeleteContactAsync(testContact.Id);

            return (contactCheck.FileAs == "NewFileAsValue" && contactCheck.JobTitle == "NewJobTitleValue");


        }

        public static async Task<bool> TryDeleteContactAsync()
        {
            var newContact = await ContactsSnippets.AddContactItemAsync(
                Guid.NewGuid().ToString(),
                Guid.NewGuid().ToString(),
                STORY_DATA_IDENTIFIER,
                Guid.NewGuid().ToString(),
                "a@b.com",
                Guid.NewGuid().ToString(),
                Guid.NewGuid().ToString());

            // Verify a valid contact id was returned
            if (newContact == null)
                return false;

            // Verify that count has increased by 1
            var contactCheck = await ContactsSnippets.GetContactAsync(newContact.Id);
            if (contactCheck == null)
                return false;

            // Delete contact we added
            var contactDeleted = await ContactsSnippets.DeleteContactAsync(newContact.Id);
            if (!contactDeleted)
                return false;


            contactCheck = await ContactsSnippets.GetContactAsync(newContact.Id);
            return (contactCheck == null);

        }

    }
}
