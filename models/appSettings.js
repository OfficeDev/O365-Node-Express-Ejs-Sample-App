/*
 *  Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.
 */


(function (appSettings) {
    
    // Set the Azure tenant for your Office 365 Developer site.
    appSettings.tenant = "contoso";
    
    // Configure the OAuth options to match your app.
    appSettings.oauthOptions = {
        clientId : "<your app clientId from AAD>"
        ,clientSecret : "<your app key from AAD>"
        ,tenantId : "<your tenantId GUID from AAD>"
        ,resource : "https://api.office.com/discovery/"

        // The redirectURL is set in AAD. For the following redirectURL  
        // "http://localhost:1337/auth/azureoauth/callback", 
        // the app needs to supply a matching middleware:
        //      app.get('/auth/azureoauth/callback', ...) 
        // to receive the auth results
    };
    
    appSettings.resources = {
        exchange : "https://outlook.office365.com/",
        onedrive : 'https://' + appSettings.tenant + '-my.sharepoint.com/',
        sharepoint : 'https://' + appSettings.tenant + '.sharepoint.com/',
        discovery : 'https://api.office.com/discovery/'
    }
    
    appSettings.apiEndpoints = {
        exchangeBaseUrl : "https://outlook.office365.com/api/v1.0/me",
        oneDriveBusinessBaseUrl : "https://" + appSettings.tenant + "-my.sharepoint.com/_api/v1.0/me",
        sharePointSiteBaseUrl : "https://" + appSettings.tenant + ".sharepoint.com/_api/web",
        discoveryServiceBaseUrl : "https://api.office.com/discovery/v1.0/me",
        accessTokenRequestUrl : "https://login.windows.net/common/oauth2/token"
    };
    
    appSettings.useFiddler = false;


})(module.exports);

// *********************************************************
//
// O365-Node-Express-Ejs-Sample-App, https://github.com/OfficeDev/O365-Node-Express-Ejs-Sample-App
//
// Copyright (c) Microsoft Corporation
// All rights reserved.
//
// MIT License:
// Permission is hereby granted, free of charge, to any person obtaining
// a copy of this software and associated documentation files (the
// "Software"), to deal in the Software without restriction, including
// without limitation the rights to use, copy, modify, merge, publish,
// distribute, sublicense, and/or sell copies of the Software, and to
// permit persons to whom the Software is furnished to do so, subject to
// the following conditions:
//
// The above copyright notice and this permission notice shall be
// included in all copies or substantial portions of the Software.
//
// THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND,
// EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF
// MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND
// NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE
// LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION
// OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION
// WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
//
// *********************************************************



