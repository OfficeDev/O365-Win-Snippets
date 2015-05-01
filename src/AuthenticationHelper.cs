// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.

using Microsoft.IdentityModel.Clients.ActiveDirectory;
using Microsoft.Office365.Discovery;
using Microsoft.Office365.OutlookServices;
using System;
using System.Diagnostics;
using System.Net.Http;
using System.Linq;
using System.Threading.Tasks;
using Windows.Security.Authentication.Web;
using Windows.Storage;

namespace O365_Win_Snippets
{
    public static class AuthenticationHelper
    {
        // The ClientID is added as a resource in App.xaml when you register the app with Office 365. 
        // As a convenience, we load that value into a variable called ClientID. This way the variable 
        // will always be in sync with whatever client id is added to App.xaml.
        private static readonly string ClientID = App.Current.Resources["ida:ClientID"].ToString();
        private static Uri _returnUri = WebAuthenticationBroker.GetCurrentApplicationCallbackUri();


        // Properties used for communicating with your Windows Azure AD tenant.
        // The AuthorizationUri is added as a resource in App.xaml when you regiter the app with 
        // Office 365. As a convenience, we load that value into a variable called _commonAuthority, adding _common to this Url to signify
        // multi-tenancy. This way it will always be in sync with whatever value is added to App.xaml.
        public static readonly string CommonAuthority = App.Current.Resources["ida:AuthorizationUri"].ToString() + @"/Common";
        public static readonly Uri DiscoveryServiceEndpointUri = new Uri("https://api.office.com/discovery/v1.0/me/");
        public const string DiscoveryResourceId = "https://api.office.com/discovery/";


        //Store login settings
        public static ApplicationDataContainer _settings = ApplicationData.Current.LocalSettings;

        //Property for storing and returning the authority used by the last authentication.
        //This value is populated when the user connects to the service and made null when the user signs out.
        public static string LastAuthority
        {
            get
            {
                if (_settings.Values.ContainsKey("LastAuthority") && _settings.Values["LastAuthority"] != null)
                {
                    return _settings.Values["LastAuthority"].ToString();
                }
                else
                {
                    return string.Empty;
                }

            }

            set
            {
                _settings.Values["LastAuthority"] = value;
            }
        }

        //Property for storing the tenant id so that we can pass it to the ActiveDirectoryClient constructor.
        //This value is populated when the user connects to the service and made null when the user signs out.
        static internal string TenantId
        {
            get
            {
                if (_settings.Values.ContainsKey("TenantId") && _settings.Values["TenantId"] != null)
                {
                    return _settings.Values["TenantId"].ToString();
                }
                else
                {
                    return string.Empty;
                }

            }

            set
            {
                _settings.Values["TenantId"] = value;
            }
        }

        // Property for storing the logged-in user so that we can display user properties later.
        //This value is populated when the user connects to the service.
        static internal string LoggedInUser
        {
            get
            {
                if (_settings.Values.ContainsKey("LoggedInUser") && _settings.Values["LoggedInUser"] != null)
                {
                    return _settings.Values["LoggedInUser"].ToString();
                }
                else
                {
                    return string.Empty;
                }

            }

            set
            {
                _settings.Values["LoggedInUser"] = value;
            }
        }


        // Property for storing the logged-in user email address so that we can display user properties later.
        //This value is populated when the user connects to the service.
        public static string LoggedInUserEmail
        {
            get
            {
                if (_settings.Values.ContainsKey("LoggedInUserEmail") && _settings.Values["LoggedInUserEmail"] != null)
                {
                    return _settings.Values["LoggedInUserEmail"].ToString();
                }
                else
                {
                    return string.Empty;
                }

            }

            set
            {
                _settings.Values["LoggedInUserEmail"] = value;
            }
        }

        //Property for storing the authentication context.
        public static AuthenticationContext _authenticationContext { get; set; }


        /// <summary>
        /// Signs the user out of the service.
        /// </summary>
        public static void SignOut()
        {
            //Handle case where user signs out without first running any snippets.
            if (_authenticationContext == null)
            {
                _authenticationContext = new AuthenticationContext(LastAuthority);
            }

            _authenticationContext.TokenCache.Clear();

            //Clean up all existing clients
            //Clear stored values from last authentication.
            _settings.Values["TenantId"] = null;
            _settings.Values["LastAuthority"] = null;
            _settings.Values["LoggedInUser"] = null;
            _settings.Values["LoggedInUserEmail"] = null;

        }

        // Get an access token for the given context and resourceId. An attempt is first made to 
        // acquire the token silently. If that fails, then we try to acquire the token by prompting the user.
        public static async Task<string> GetTokenHelperAsync(AuthenticationContext context, string resourceId)
        {
            string accessToken = null;
            AuthenticationResult result = null;

            result = await context.AcquireTokenAsync(resourceId, ClientID, _returnUri);

            if (result.Status == AuthenticationStatus.Success)
            {
                accessToken = result.AccessToken;
                //Store values for logged-in user, tenant id, and authority, so that
                //they can be re-used if the user re-opens the app without disconnecting.
                _settings.Values["LoggedInUser"] = result.UserInfo.GivenName;
                _settings.Values["LoggedInUserEmail"] = result.UserInfo.DisplayableId;
                _settings.Values["TenantId"] = result.TenantId;
                _settings.Values["LastAuthority"] = context.Authority;

                return accessToken;
            }
            else
            {
                return null;
            }
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