// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.

using Microsoft.Azure.ActiveDirectory.GraphClient;
using Microsoft.IdentityModel.Clients.ActiveDirectory;
using Microsoft.Office365.Discovery;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;


namespace O365_Win_Snippets
{
    class UsersAndGroupsSnippets
    {
        /// <summary>
        /// Checks that a Graph client is available to the client.
        /// </summary>
        /// <returns>The Graph client.</returns>
        private static ActiveDirectoryClient _graphClient = null;

        public static async Task<ActiveDirectoryClient> GetGraphClientAsync()
        {
            //Check to see if this client has already been created. If so, return it. Otherwise, create a new one.
            if (_graphClient != null)
            {
                Debug.WriteLine("Got a Graph client for Users and Groups.");
                return _graphClient;
            }
            else
            {
                // Active Directory service endpoints
                const string AadServiceResourceId = "https://graph.windows.net/";
                Uri AadServiceEndpointUri = new Uri("https://graph.windows.net/");

                try
                {
                    //First, look for the authority used during the last authentication.
                    //If that value is not populated, use _commonAuthority.
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

                    var token = await AuthenticationHelper.GetTokenHelperAsync(AuthenticationHelper._authenticationContext, AadServiceResourceId);

                    // Check the token
                    if (String.IsNullOrEmpty(token))
                    {
                        // User cancelled sign-in
                        return null;
                    }
                    else
                    {
                        // Create our ActiveDirectory client.
                        _graphClient = new ActiveDirectoryClient(
                            new Uri(AadServiceEndpointUri, AuthenticationHelper.TenantId),
                            async () => await AuthenticationHelper.GetTokenHelperAsync(AuthenticationHelper._authenticationContext, AadServiceResourceId));

                        Debug.WriteLine("Got a Graph client for Users and Groups.");

                        return _graphClient;
                    }


                }

                catch (Exception)
                {
                    // Argument exception
                }
                AuthenticationHelper._authenticationContext.TokenCache.Clear();
                return null;
            }
        }

        public static async Task<List<IUser>> GetUsersAsync()
        {

            var client = await GetGraphClientAsync();

            var users = await client.Users.ExecuteAsync();

            Debug.WriteLine("First user in collection: " + users.CurrentPage[0].DisplayName);

            return users.CurrentPage.ToList();


        }

        public static async Task<ITenantDetail> GetTenantDetailsAsync()
        {

            var client = await GetGraphClientAsync();

            var tenantDetails = await client.TenantDetails.ExecuteAsync();

            Debug.WriteLine("Got tenant details.");

            return tenantDetails.CurrentPage.First();


        }

        public static async Task<List<IGroup>> GetGroupsAsync()
        {

            var client = await GetGraphClientAsync();

            var groups = await client.Groups.ExecuteAsync();

            if (groups.CurrentPage.Count == 0)
            {
                Debug.WriteLine("No groups.");
                return new List<IGroup>();
            }

            Debug.WriteLine("First group in collection: " + groups.CurrentPage[0].DisplayName);

            return groups.CurrentPage.ToList();

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