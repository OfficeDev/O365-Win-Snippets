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
    class UsersAndGroupsOperations
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
            try
            {
                var client = await GetGraphClientAsync();

                var users = await client.Users.ExecuteAsync();

                Debug.WriteLine("First user in collection: " + users.CurrentPage[0].DisplayName);

                return users.CurrentPage.ToList();
            }
            catch { return null; }

        }

        public static async Task<ITenantDetail> GetTenantDetailsAsync()
        {
            try
            {
                var client = await GetGraphClientAsync();

                var tenantDetails = await client.TenantDetails.ExecuteAsync();

                Debug.WriteLine("Got tenant details.");

                return tenantDetails.CurrentPage.First();
            }
            catch { return null; }

        }

        public static async Task<List<IGroup>> GetGroupsAsync()
        {
            try
            {
                var client = await GetGraphClientAsync();

                var groups = await client.Groups.ExecuteAsync();

                Debug.WriteLine("First group in collection: " + groups.CurrentPage[0].DisplayName);

                return groups.CurrentPage.ToList();
            }
            catch { return null; }

        }

    }
}
