// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.

using Microsoft.IdentityModel.Clients.ActiveDirectory;
using Microsoft.Office365.Discovery;
using Microsoft.Office365.OutlookServices;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

//Snippets in this file:
//GetOutlookClientAsync
//GetContactsPageAsync
//GetContactAsync
//AddContactItemAsync
//UpdateContactItemAsync
//DeleteContactAsync


namespace O365_Win_Snippets
{
    public static class ContactsSnippets
    {

        private static OutlookServicesClient _outlookClient = null;

        /// <summary>
        /// Checks that an OutlookServicesClient object is available. 
        /// </summary>
        /// <returns>The OutlookServicesClient object. </returns>
        public static async Task<OutlookServicesClient> GetOutlookClientAsync()
        {

            if (_outlookClient != null && !String.IsNullOrEmpty(AuthenticationHelper.LastAuthority))
            {
                Debug.WriteLine("Got an Outlook client for Contacts.");
                return _outlookClient;
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
                        await discoveryClient.DiscoverCapabilityAsync("Contacts");

                    var token = await AuthenticationHelper.GetTokenHelperAsync(AuthenticationHelper._authenticationContext, result.ServiceResourceId);
                    // Check the token
                    if (String.IsNullOrEmpty(token))
                    {
                        // User cancelled sign-in
                        return null;
                    }
                    else
                    {

                        _outlookClient = new OutlookServicesClient(
                            result.ServiceEndpointUri,
                            async () => await AuthenticationHelper.GetTokenHelperAsync(AuthenticationHelper._authenticationContext, result.ServiceResourceId));
                        Debug.WriteLine("Got an Outlook client for Contacts");
                        return _outlookClient;
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

        /// <summary>
        /// Retrieve a page of contacts from the server.
        /// </summary>
        /// <returns>A list of contacts.</returns>
        public static async Task<List<IContact>> GetContactsPageAsync()
        {
            try
            {
                // Get exchangeclient
                var outlookClient = await GetOutlookClientAsync();

                // Get contacts
                var contactsResults = await outlookClient.Me.Contacts.ExecuteAsync();

                // You can access each contact as follows.
                if (contactsResults.CurrentPage.Count > 0)
                {
                    string contactId = contactsResults.CurrentPage[0].Id;

                    if ( contactsResults.CurrentPage.Count > 0)
                    {
                        Debug.WriteLine("First contact:" + contactId);
                    }
                }

                return contactsResults.CurrentPage.ToList();
            }

            catch { return null; }
        }

        public static async Task<IContact> GetContactAsync(string Id)
        {
            try
            {
                var exchangeClient = await GetOutlookClientAsync();
                var contact = await exchangeClient.Me.Contacts.GetById(Id).ExecuteAsync();

                Debug.WriteLine("Got contact:" + contact.Id);

                return contact;
            }
            catch { return null; }
        }

        /// <summary>
        /// Adds a new contact.
        /// </summary>
        public static async Task<IContact> AddContactItemAsync(
            string fileAs,
            string givenName,
            string surname,
            string jobTitle,
            string email,
            string workPhone,
            string mobilePhone
            )
        {
            Contact newContact = new Contact
            {
                FileAs = fileAs,
                GivenName = givenName,
                Surname = surname,
                JobTitle = jobTitle,
                MobilePhone1 = mobilePhone
            };

            newContact.BusinessPhones.Add(workPhone);


            // Note: Setting EmailAddress1 to a null or empty string will throw an exception that
            // states the email address is invalid and the contact cannot be added.
            // Setting EmailAddress1 to a string that does not resemble an email address will not
            // cause an exception to be thrown, but the value is not stored in EmailAddress1.
            if (!string.IsNullOrEmpty(email))
                newContact.EmailAddresses.Add(new EmailAddress() { Address = email });

            try
            {
                // Make sure we have a reference to the Exchange client
                var outlookClient = await GetOutlookClientAsync();

                // This results in a call to the service.
                await outlookClient.Me.Contacts.AddContactAsync(newContact);

                Debug.WriteLine("Added contact: " + newContact.Id);

                return newContact;
            }
            catch { return null; }
        }

        /// <summary>
        /// Updates an existing contact.
        /// </summary>
        public static async Task<IContact> UpdateContactItemAsync(string selectedContactId,
            string fileAs,
            string givenName,
            string surname,
            string jobTitle,
            string email,
            string workPhone,
            string mobilePhone,
            byte[] contactImage
           )
        {

            try
            {
                // Make sure we have a reference to the Exchange client
                var exchangeClient = await GetOutlookClientAsync();

                var contactToUpdate = await exchangeClient.Me.Contacts[selectedContactId].ExecuteAsync();

                contactToUpdate.FileAs = fileAs;
                contactToUpdate.GivenName = givenName;
                contactToUpdate.Surname = surname;
                contactToUpdate.JobTitle = jobTitle;

                contactToUpdate.MobilePhone1 = mobilePhone;

                // Note: Setting EmailAddress1 to a null or empty string will throw an exception that
                // states the email address is invalid and the contact cannot be added.
                // Setting EmailAddress1 to a string that does not resemble an email address will not
                // cause an exception to be thrown, but the value is not stored in EmailAddress1.

                if (!string.IsNullOrEmpty(email))
                {
                    contactToUpdate.EmailAddresses.Clear();
                    contactToUpdate.EmailAddresses.Add(new EmailAddress() { Address = email, Name = email });
                }

                // Update the contact in Exchange
                await contactToUpdate.UpdateAsync();

                Debug.WriteLine("Updated contact: " + contactToUpdate.Id);

                return contactToUpdate;

                // A note about Batch Updating
                // You can save multiple updates on the client and save them all at once (batch) by 
                // implementing the following pattern:
                // 1. Call UpdateAsync(true) for each contact you want to update. Setting the parameter dontSave to true 
                //    means that the updates are registered locally on the client, but won't be posted to the server.
                // 2. Call exchangeClient.Context.SaveChangesAsync() to post all contact updates you have saved locally  
                //    using the preceding UpdateAsync(true) call to the server, i.e., the user's Office 365 contacts list.
            }
            catch { return null; }
        }

        /// <summary>
        /// Deletes a contact.
        /// </summary>
        /// <param name="contactId">The contact to delete.</param>
        /// <returns>True if deleted;Otherwise, false.</returns>
        public static async Task<bool> DeleteContactAsync(string contactId)
        {
            try
            {
                // Make sure we have a reference to the Exchange client
                var outlookClient = await GetOutlookClientAsync();

                var contactToDelete = await outlookClient.Me.Contacts[contactId].ExecuteAsync();

                await contactToDelete.DeleteAsync();

                Debug.WriteLine("Deleted contact: " + contactToDelete.Id);

                return true;
            }
            catch { return false; }
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