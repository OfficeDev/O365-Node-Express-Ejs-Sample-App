/*
 *  Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.
 */

// file.js

var request = require('request');
var appSettings = require('../models/appSettings.js');

// express app routes for files rest api  calls
module.exports = function (app, passport, utils) {
    
    // This middleware checks for and obtain, if necessary, access_token for
    // accessing OneDrive for Business service. 
    app.use('/files', function (req, res, next) {
        if (passport.user.hasToken(appSettings.resources.onedrive)) {
            // already has access token for onedrive, 
            // should also check for expiration, and other issues, ignore for now.
            // skip to the next middleware
            return next();
        } else {
            var data = 'grant_type=refresh_token' 
            + '&refresh_token=' + passport.user.refresh_token 
            + '&client_id=' + appSettings.oauthOptions.clientId 
            + '&client_secret=' + encodeURIComponent(appSettings.oauthOptions.clientSecret) 
            + '&resource=' + encodeURIComponent(appSettings.resources.onedrive);
            var opts = {
                url: appSettings.apiEndpoints.accessTokenRequestUrl,
                body: data,
                headers : { 'Content-Type' : 'application/x-www-form-urlencoded' }
            };
            require('request').post(opts, function (err, response, body) {
                if (err) {
                    return next(err)
                } else {
                    var token = JSON.parse(body);
                    passport.user.setToken(token);                    
                    return next();
                }
            })
        }
    })
    
    
    // Illustrating Files REST API call, using a drive example.
    // This call is called after the app.use('/files', ...) is executed.
    app.get('/files', function (req, res, next) {
        if (!passport.user.getToken(appSettings.resources.onedrive)) {
            return next('invalid token');
        } 
        var fileUrl = appSettings.apiEndpoints.oneDriveBusinessBaseUrl + '/drive';
        var opts = { auth: { 'bearer' : passport.user.getToken(appSettings.resources.onedrive).access_token } };
        
        // For debugging purposes
        if (appSettings.useFiddler) {
            opts.proxy =  'http://127.0.0.1:8888'; 
            opts.rejectUnauthorized = false;
        }
        
        require('request').get(fileUrl, opts, function (error, response, body) {
            if (error) {
                next(error);
            }
            else {
                data = { user: passport.user, result: JSON.parse(body) };
                //for (var key in data.result) { console.log("%s : %s", key, data.result[key])};
                res.render('file', { data: data });
            }
        });
    });

}

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

