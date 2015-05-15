// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.
using Microsoft.IdentityModel.Clients.ActiveDirectory;
using Microsoft.Office365.Discovery;
using Microsoft.Office365.OutlookServices;
using Microsoft.Office365.SharePoint.CoreServices;
using O365_APIs_Start_ASPNET_MVC.Utils;
using O365_APIs_Start_ASPNET_MVC.Models;
using System;
using System.Security.Claims;
using System.Threading.Tasks;
using System.Linq;

namespace O365_APIs_Start_ASPNET_MVC.Helpers
{
    // Provides a valid OutlookServices client that contains the bearer token for issuing requests against Calendar, Mail, and Contact resources.
    internal class AuthenticationHelper
    {
        internal static async Task<OutlookServicesClient> EnsureOutlookServicesClientCreatedAsync(string capabilityName)
        {

            var signInUserId = ClaimsPrincipal.Current.FindFirst(ClaimTypes.NameIdentifier).Value;
            var userObjectId = ClaimsPrincipal.Current.FindFirst("http://schemas.microsoft.com/identity/claims/objectidentifier").Value;

            AuthenticationContext authContext = new AuthenticationContext(SettingsHelper.Authority, new ADALTokenCache(signInUserId));

            try
            {
                DiscoveryClient discClient = new DiscoveryClient(SettingsHelper.DiscoveryServiceEndpointUri,
                    async () =>
                    {
                        var authResult = await authContext.AcquireTokenSilentAsync(SettingsHelper.DiscoveryServiceResourceId,
                                                                                   new ClientCredential(SettingsHelper.ClientId,
                                                                                                        SettingsHelper.ClientSecret),
                                                                                   new UserIdentifier(userObjectId,
                                                                                                      UserIdentifierType.UniqueId));

                        return authResult.AccessToken;
                    });

                var dcr = await discClient.DiscoverCapabilityAsync(capabilityName);

                return new OutlookServicesClient(dcr.ServiceEndpointUri,
                    async () =>
                    {
                        var authResult = await authContext.AcquireTokenSilentAsync(dcr.ServiceResourceId,
                                                                                   new ClientCredential(SettingsHelper.ClientId,
                                                                                                        SettingsHelper.ClientSecret),
                                                                                   new UserIdentifier(userObjectId,
                                                                                                      UserIdentifierType.UniqueId));

                        return authResult.AccessToken;
                    });
            }
            catch (AdalException exception)
            {
                //Handle token acquisition failure
                if (exception.ErrorCode == AdalError.FailedToAcquireTokenSilently)
                {
                    authContext.TokenCache.Clear();
                    throw exception;
                }
                return null;
            }
        }

        internal static async Task<SharePointClient> EnsureSharePointClientCreatedAsync(string capabilityName)
        {

            var signInUserId = ClaimsPrincipal.Current.FindFirst(ClaimTypes.NameIdentifier).Value;
            var userObjectId = ClaimsPrincipal.Current.FindFirst("http://schemas.microsoft.com/identity/claims/objectidentifier").Value;

            AuthenticationContext authContext = new AuthenticationContext(SettingsHelper.Authority, new ADALTokenCache(signInUserId));

            try
            {
                DiscoveryClient discClient = new DiscoveryClient(SettingsHelper.DiscoveryServiceEndpointUri,
                    async () =>
                    {
                        var authResult = await authContext.AcquireTokenSilentAsync(SettingsHelper.DiscoveryServiceResourceId,
                                                                                   new ClientCredential(SettingsHelper.ClientId,
                                                                                                        SettingsHelper.ClientSecret),
                                                                                   new UserIdentifier(userObjectId,
                                                                                                      UserIdentifierType.UniqueId));
                        return authResult.AccessToken;
                    });

                var dcr = await discClient.DiscoverCapabilityAsync(capabilityName);

                return new SharePointClient(dcr.ServiceEndpointUri,
                    async () =>
                    {
                        var authResult = await authContext.AcquireTokenSilentAsync(dcr.ServiceResourceId,
                                                                                   new ClientCredential(SettingsHelper.ClientId,
                                                                                                        SettingsHelper.ClientSecret),
                                                                                   new UserIdentifier(userObjectId,
                                                                                                      UserIdentifierType.UniqueId));

                        return authResult.AccessToken;
                    });
            }
            catch (AdalException exception)
            {
                //Partially handle token acquisition failure here and bubble it up to the controller
                if (exception.ErrorCode == AdalError.FailedToAcquireTokenSilently)
                {
                    authContext.TokenCache.Clear();
                    throw exception;
                }
                return null;
            }
        }
    }
}
//*********************************************************  
//  
//O365 APIs Starter Project for ASPNET MVC, https://github.com/OfficeDev/Office-365-APIs-Starter-Project-for-ASPNETMVC
// 
//Copyright (c) Microsoft Corporation 
//All rights reserved.  
// 
//MIT License: 
// 
//Permission is hereby granted, free of charge, to any person obtaining 
//a copy of this software and associated documentation files (the 
//""Software""), to deal in the Software without restriction, including 
//without limitation the rights to use, copy, modify, merge, publish, 
//distribute, sublicense, and/or sell copies of the Software, and to 
//permit persons to whom the Software is furnished to do so, subject to 
//the following conditions: 
// 
//The above copyright notice and this permission notice shall be 
//included in all copies or substantial portions of the Software. 
// 
//THE SOFTWARE IS PROVIDED ""AS IS"", WITHOUT WARRANTY OF ANY KIND, 
//EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF 
//MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND 
//NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE 
//LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION 
//OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION 
//WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE. 
//  
//********************************************************* 

