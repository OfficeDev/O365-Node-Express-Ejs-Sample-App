/*
 *  Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.
 */

// discovery.js

var request = require('request');
var appSettings = require('../models/appSettings.js');

// express app routes for discovery services rest api  calls
module.exports = function (app, passport, utils) {
    // This middleware checks for and obtain, if necessary, access_token for
    // accessing Office 365 discovery service. 
    app.use('/discovery', function (req, res, next) {
        passport.getAccessToken(appSettings.resources.discovery, req, res, next);
    })

    app.get('/discovery', function (req, res, next) {
        if (!passport.user.getToken(appSettings.resources.discovery)) {
            return next('invalid token');
        }
        var fileUrl = appSettings.apiEndpoints.discoveryServiceBaseUrl + '/services';
        var opts = { auth: { 'bearer' : passport.user.getToken(appSettings.resources.discovery).access_token } };
        
        require('request').get(fileUrl, opts, function (error, response, body) {
            if (error) {
                next(error);
            }
            else {
                passport.user.setCapabilities(JSON.parse(body)['value']);
                data = { user: passport.user, capabilities: passport.user.capabilities };
                res.render('discovery', { data: data });
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

