/*
 *  Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.
 */

var url = require('url');
var appSettings = require('../models/appSettings.js');

// routes/index.js
module.exports = function (app, passport) {
    
    // =====================================
    // HOME PAGE (with login links) ========
    // =====================================
    app.get('/', function (req, res) {
        // load the index.jade or index.ejs file, 
        // depending on the view engine selected
        res.render('index', {title: 'O365-Node-Express-Ejs | Home'}); 
    });
    
    // =====================================
    // LOGIN ===============================
    // =====================================
    // show the login form
    app.get('/login', function (req, res) {
        // redirect the login request to Azure Active Directory oauth2
        res.redirect('/auth/azureOAuth');
    });
    
    // =====================================
    // Starts Azure authentication/authorization
    //
    app.get('/auth/azureOAuth', 
        passport.authenticate('azureoauth', { failureRedirect: '/' })
    );
    
    // =====================================
    // cache and handle access token and refresh token as returned from AAD. 
    // The accessToken 
    app.get('/auth/azureOAuth/callback', 
        passport.authenticate('azureoauth', { failureRedirect: '/' }),
        function (req, res) {
            console.log('user authenticated');
            //var options = url.parse(req.url, true);
            //var code = options.query.code;
            //var sessionState = options.query.session_state;
            //var username = passport.user.profile.username;
		// not to check if IsLoggedIn true
        res.render('apiTasks', {title: 'O365-Node-Express-Ejs | Tasks', user : req.user});
    });
    
    
    
    // =====================================
    // API task SECTION =====================
    // =====================================
    // we will want this protected so you have to be logged in to visit
    // we will use route middleware to verify this (the isLoggedIn function)
    app.get('/api/tasks', isLoggedIn, 
        function (req, res) {
        res.render('apiTasks', { title : 'O365NodeExpressEjs | Tasks',
            username : req.user.username // get the user out of session and pass to template
        });
    });
    
    // =====================================
    // LOGOUT ==============================
    // =====================================
    app.get('/logout', function (req, res) {
        req.logout();
        res.redirect('/');
    });

};

// route middleware to make sure a user is logged in
function isLoggedIn(req, res, next) {
    
    // if user is authenticated in the session, carry on 
    //if (req.isAuthenticated())
	if (req.user.username && req.user.profile.accessToken) // debug
        return next();
    console.log('user is not logged in.')
    // if they aren't redirect them to the home page
    res.redirect('/');
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

