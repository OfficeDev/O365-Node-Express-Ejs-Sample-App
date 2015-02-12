/*
 *  Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.
 */

(function (user) {
    user.tokens = {};   
    user.refresh_token = "";
    user.username = "";
    user.displayName = "";

    user.validate = function (result, next) {
        if (!result) {
            return next('invalid user');
        } else if (!result.accessToken) {
            return next('invalid credentials');
        } else {
            user.displayname = result.userProfile.displayname;
            user.username = result.userProfile.username;
            
            user.accessToken = result.accessToken;
            user.refresh_token = result.refreshToken;

            result.tokenParams.refresh_token = result.refreshToken;
            user.setToken(result.tokenParams);
          
            return next(null, user);
        }
    };
    
    user.hasToken =  function(resourceUri) {
        if (user.tokens.hasOwnProperty(resourceUri)) {
            return true;
        } else {
            return false;
        }
    }
    user.getToken = function (resourceUri) {
        if (user.hasToken(resourceUri)) {
            return user.tokens[resourceUri];
        }
    }
    user.setToken = function (token) {
        if (!token.resource) {
            token.resource = "default";
        }
        user.tokens[token.resource] = token;
    }

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

