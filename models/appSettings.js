/*
 *  Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.
 */


(function (appSettings) {
    appSettings.oauthOptions = {
        clientId : "646a412d-c4a6-4a76-823a-eb0d70519dc5"
        ,clientSecret : "pi2o7cnLWpbJoH+CVNW3YXXkml/XysM7odVMEMeiCmQ="
        ,tenantId : "b7e72cd7-4df7-4e7e-b5ff-3310c56629e5"
        ,resource : "https://outlook.office365.com/" //"https%3A%2F%2Foutlook.office365.com%2F"
        //,redirectURL : "http://localhost:1337/auth/azureoauth/callback" // this is set in AAD for this app
    };

    appSettings.apiEndpoints = {
        exchangeBaseUrl : "https://outlook.office365.com/api/v1.0/me",
        oneDriveBusinessBaseUrl : "https://" + appSettings.oauthOptions.tenantId + "-my.sharepoint.com/_api/v1.0/me",
        publicSharePointBaseUrl : "https://" + appSettings.oauthOptions.tenantId + ".sharepoint.com/haymuto-public/_api/v1.0",
        teamSharePointBaseUrl : "https://" + appSettings.oauthOptions.tenantId + ".sharepoint.com/haymuto/_api/v1.0",
        discoveryServiceBaseUrl : "https://api.office.com/discovery/v1.0/me"

    };
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



