/*
 *  Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.
 */

// site.js

var request = require('request');
var appSettings = require('../models/appSettings.js');

// express app routes for SharePoint  rest api  calls to get O365 site lists
module.exports = function (app, passport, utils) {
        
    // The following middleware checks for and obtain, if necessary, access_token for
    // accessing SharePoint site service. 
    app.use('/site', function (req, res, next) {
        passport.getAccessToken(appSettings.resources.sharepoint, req, res, next);
    })
    
    // Illustrating SharePoint REST API call to read List.
    // This call is called after the app.use('/list', ...) is executed.
    app.get('/site', function (req, res, next) {
        if (!passport.user.getToken(appSettings.resources.sharepoint)) {
            return next({ msg: 'invalid token' });
        }
        var fileUrl = appSettings.apiEndpoints.sharePointSiteBaseUrl + '/lists';
        var opts = {
            auth: { 'bearer' : passport.user.getToken(appSettings.resources.sharepoint).access_token },
            headers : { 'accept' : 'application/json; odata=verbose' },  // verbose resource representation differs from the non-verbose one. Use verbose to make life a little easier. :)
            secureProtocol: 'TLSv1_method'  // required of Shareoint site and OneDrive,
        };
        
        //For debugging purposes
        if (appSettings.useFiddler) {
            opts.proxy = 'http://127.0.0.1:8888';
            opts.rejectUnauthorized = false;
        }
        
        require('request').get(fileUrl, opts, function (error, response, body) {
            if (error) {
                next(error);
            }
            else if (response.statusCode != 200) {
                next({status : response.statusCode, msg : body });
            } else {
                data = { user: passport.user, result: JSON.parse(body) };
                res.render('site', { data: data });
            }
        });
    });
    
    // create an Office 365 SharePoint list 
    var newListUri; // cache the new list Uri for ensuing update and delete operations
    app.get('/site/list/create', function (req, res, next) {
        if (!passport.user.getToken(appSettings.resources.sharepoint)) {
            return next({ msg: 'invalid token' });
        }
        var fileUrl = appSettings.apiEndpoints.sharePointSiteBaseUrl + '/lists';
        var opts = {
            auth: { 'bearer' : passport.user.getToken(appSettings.resources.sharepoint).access_token },
            headers : {
                'accept' : 'application/json;odata=verbose', 
                'content-type' : 'application/json;odata=verbose' // verbose resource representation differs from the non-verbose one. Use verbose to make life a little easier. :)
            },
            body : JSON.stringify({
                '__metadata' : { 'type': 'SP.List' }, 
                'AllowContentTypes': true, 
                'BaseTemplate': 100, 
                'ContentTypesEnabled': true, 
                'Description': 'An example to add a SharePoint list', 
                'Title': 'New Test List'
            } ),
            secureProtocol: 'TLSv1_method'  // required of Shareoint site and OneDrive,
        };
        
        //For debugging purposes
        if (appSettings.useFiddler) {
            opts.proxy = 'http://127.0.0.1:8888';
            opts.rejectUnauthorized = false;
        }
        
        require('request').post(fileUrl, opts, function (error, response, body) {
            if (error) {
                next(error);
            }
            else if (response.statusCode > 299) {
                error = {status : response.statusCode, msg : body }
                next(error);
            } else {
                data = { user: passport.user, result: JSON.parse(body) };
                newListUri = data.result.d.__metadata.uri;
                //newListEtag = data.result.d.__metadata.etag;
                res.render('site', { data: data });
            }
        });

    })
    // update a ilst
    app.get('/site/list/update', function (req, res, next) {
        if (!passport.user.getToken(appSettings.resources.sharepoint)) {
            return next({ msg: 'invalid token' });
        }
        if (!newListUri) {
            return next({ msg: 'you must create a list before updating it.' })
        }
        var fileUrl = newListUri;
        var opts = {
            auth: { 'bearer' : passport.user.getToken(appSettings.resources.sharepoint).access_token },
            headers : {
                'accept' : 'application/json;odata=verbose', 
                'content-type' : 'application/json;odata=verbose',
                'if-match' : '*', 
                'X-HTTP-METHOD' : 'MERGE'
            },
            body : JSON.stringify({
                '__metadata' : { 'type': 'SP.List' }, 
                'Description': 'An Node.js example to add a SharePoint list' 
            }),
            secureProtocol: 'TLSv1_method'  // required of Shareoint site and OneDrive,
        };
        
        //For debugging purposes
        if (appSettings.useFiddler) {
            opts.proxy = 'http://127.0.0.1:8888';
            opts.rejectUnauthorized = false;
        }
        
        require('request').post(fileUrl, opts, function (error, response, body) {
            if (error) {
                next(error);
            }
            else if (response.statusCode > 299) {
                error = { status : response.statusCode, msg : body }
                next(error);
            } else {
                data = { user: passport.user, result: !body ? body : JSON.parse(body) };
                res.render('site', { data: data });
            }
        });

    })
    
    // Get the list
    app.get('/site/list/get', function (req, res, next) {
        if (!passport.user.getToken(appSettings.resources.sharepoint)) {
            return next({ msg: 'invalid token' });
        }
        var fileUrl = appSettings.apiEndpoints.sharePointSiteBaseUrl + "/lists/getbytitle('New Test List')";
        var opts = {
            auth: { 'bearer' : passport.user.getToken(appSettings.resources.sharepoint).access_token },
            headers : {
                'accept' : 'application/json;odata=verbose', 
                'content-type' : 'application/json;odata=verbose'
            },
            secureProtocol: 'TLSv1_method'  // required of Shareoint site and OneDrive,
        };
        
        //For debugging purposes
        if (appSettings.useFiddler) {
            opts.proxy = 'http://127.0.0.1:8888';
            opts.rejectUnauthorized = false;
        }
        
        require('request').get(fileUrl, opts, function (error, response, body) {
            if (error) {
                next(error);
            }
            else if (response.statusCode > 299) {
                error = { status : response.statusCode, msg : body }
                next(error);
            } else {
                data = { user: passport.user, result: !body ? body : JSON.parse(body) };
                newListUri = data.result.d.__metadata.uri;
                res.render('site', { data: data });
            }
        });

    })
    
    // Delete a list
    app.get('/site/list/delete', function (req, res, next) {
        if (!passport.user.getToken(appSettings.resources.sharepoint)) {
            return next({ msg: 'invalid token' });
        }
        if (!newListUri) {
            return next({ msg: 'you must create a list before deleting it.' })
        }
        var fileUrl = newListUri;
        var opts = {
            auth: { 'bearer' : passport.user.getToken(appSettings.resources.sharepoint).access_token },
            headers : {
                //'accept' : 'application/json;odata=verbose', 
                'if-match' : '*', 
                'X-HTTP-METHOD' : "DELETE"
            },
            secureProtocol: 'TLSv1_method'  // required of Shareoint site and OneDrive,
        };
        
        //For debugging purposes
        if (appSettings.useFiddler) {
            opts.proxy = 'http://127.0.0.1:8888';
            opts.rejectUnauthorized = false;
        }
        
        require('request').post(fileUrl, opts, function (error, response, body) {
            if (error) {
                next(error);
            }
            else if (response.statusCode > 299) {
                error = { status : response.statusCode, msg : body }
                next(error);
            } else {
                data = { user: passport.user, result: body.length==0 ? "" : JSON.parse(body) };
                res.render('site', { data: data });
            }
        });

    })

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

