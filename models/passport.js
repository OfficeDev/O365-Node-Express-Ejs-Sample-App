/*
 *  Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.
 */

var AzureOAuthStrategy = require('passport-azure-oauth').Strategy;
var User = require('./user.js');
var AppSettings = require('./appSettings.js');

module.exports=function (passport) {
    // =========================================================================
    // passport session setup ==================================================
    // =========================================================================
    // required for persistent login sessions
    // passport needs ability to serialize and unserialize users out of session
    
    // used to serialize the user for the session
    passport.serializeUser(function (user, done) {
        done(null, user);
    });
    
    // used to deserialize the user
    passport.deserializeUser(function (id, done) {
        User.findById(id, function (err, user) {
            done(err, user);
        });
    });
    

    // For information on the profile entries: see http://msdn.microsoft.com/en-us/library/azure/dn645542.aspx
    passport.use('azureoauth',  new AzureOAuthStrategy(
        AppSettings.oauthOptions, 
        function Verify(accessToken, refreshToken, params, profile, done) {
        User.validate(
            {
                'accessToken' : accessToken, 
                'refreshToken' : refreshToken, 
                'tokenParams': params, 
                'userProfile': profile
            }, 
            function (err, user) {
                passport.user = null;
                if (err)
                    return done(err);                
                if (!user)
                    return done('Cannot verify the user ' + profile.displayname + ', ' + profile.username);                                
                // all is well, return successful user
                passport.user = user;
                return done(null, user);
            });
        })
    );

};

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

