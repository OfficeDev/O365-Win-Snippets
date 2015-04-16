// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.

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