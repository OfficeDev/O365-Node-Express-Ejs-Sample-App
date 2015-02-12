// calendar.js
var request = require('request');
var appSettings = require('../models/appSettings.js');


// express app routes for calendar REST api calls
module.exports = function (app, passport, utils) {
    
    // get calendar events and displays the results in the raw odata format.
    app.get('/calendar', function (req, res, next) {
        var calendarUrl = appSettings.apiEndpoints.exchangeBaseUrl + "/events"; 
        request.get(calendarUrl,
            { auth : { 'bearer' : passport.user.getToken(appSettings.resources.exchange).access_token } },
            function (error, response, body) {
            if (error) next(error)
            else if (!body) {
                var msg = "";
                if (response.headers['x-ms-diagnostics']) msg = response.headers['x-ms-diagnostics'];
                error = { status: response.statusCode, message: msg , stack : "" };
                next(error);
            }
            else {
                data = { user: passport.user, events: JSON.parse(body)['value'] };
                res.render('calendar', { data: data });
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

