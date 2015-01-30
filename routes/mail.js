/*
 *  Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.
 */

// routes/mail.js

var url = require('url');
var request = require('request');
var appSettings = require('../models/appSettings.js');


module.exports = function (app, passport, utils) {

    // Get a messaget list in the user's Inbox using the O365 API,
    // displaying To, Subject and Preview for each message.
    app.get('/mail', function (req, res, next) {
        request.get(
            appSettings.apiEndpoints.exchangeBaseUrl + "/messages", 
            { auth : { 'bearer' : passport.user.accessToken } },
            function (error, response, body) {
                if (error) {
                    next(error);
                }
                else {
                    data = { user: passport.user, msgs: JSON.parse(body)['value'] };
                    res.render('mail', { data: data });
                }
            }
        );
    });
    
    // GET a given message and display content to the user,
    // displaying the message as-is in HTML.
    app.get('/mail/view', function (req, res, next) {
        var id = url.parse(req.url, true).query.id;
        request.get(
            appSettings.apiEndpoints.exchangeBaseUrl + "/messages/" + id, 
            { auth : { 'bearer' : passport.user.accessToken } },
            function (error, response, body) {
                if (error) {
                    next(error);
                }
                else {
                    var jsonBody = JSON.parse(body);
                    res.end(jsonBody.Body.Content);
                }
            }
        ); 
    })

    // delete a selected email message using the O365 API.
    app.get('/mail/delete', function (req, res, next) {
        var id = url.parse(req.url, true).query.id;
        request.del(
            appSettings.apiEndpoints.exchangeBaseUrl + "/messages/" + id, 
            { auth : { 'bearer' : passport.user.accessToken } },
            function (error, response, body) {
                if (error) {
                    next(error);
                }
                else {
                    res.redirect('/mail')
                }
            }
        );
    
    })
    
    // Pop up a message editor for th user to add comment to the original message.
    app.get('/mail/reply', function (req, res, next) {
        var id = url.parse(req.url, true).query.id;
        request.get(
            appSettings.apiEndpoints.exchangeBaseUrl + "/messages/" + id, 
            { auth : { 'bearer' : passport.user.accessToken } },
            function (error, response, body) {
                if (error) {
                    next(error);
                }
                else {
                    var jsonBody = JSON.parse(body);
                    var content = utils.htmlText(jsonBody.Body.Content);
                    res.render('mailReply', {
                        user : passport.user,
                        messageId : id,
                        recipients: jsonBody.Sender.EmailAddress.Address, 
                        subject: jsonBody.Subject, 
                        content: content
                    });              
                }
            }
        );
    });
    
    // Pop up a message composer for the user to creat and send a new mail message.
    app.get('/mail/new', function (req, res, next) {
        res.render('mailedit', { user: passport.user, recipients : "user@domain", subject: "Test", content : "" });
    })
    
    // send a new mail message to a specific recipient using O365 API. 
    // The request body must be a JSON string, not an JSON object.
    app.post('/mail/send', function (req, res, next) {
        var reqBody = {
            'Message' : {
                'Subject': req.body.subject,
                'Body': { 'ContentType': "Text", 'Content': req.body.message },
                'ToRecipients' : [{ 'EmailAddress': { 'Address' : req.body.to } }]
            },
            'SaveToSentItems' : 'false'
        };
        var reqHeaders = { 'content-type': 'application/json'};
        var reqUrl = appSettings.apiEndpoints.exchangeBaseUrl + "/sendmail";
        var reqAuth = { 'bearer': passport.user.accessToken };

        request.post(
            { url: reqUrl, headers: reqHeaders, body: JSON.stringify(reqBody), auth: reqAuth },        
            function (err, response, body) {
                if (err) { next(err); }
                else {
                    if (response.statusCode == 403) {
                        err.status = response.statusCode;
                        err.stack = body;
                        err.message = "Failed to send mail to " + req.body.to;
                        next(err);
                    }
                    else {
                        res.redirect('/mail' );
                    }
                    
                }
            }
        );
    })
    
    // reply a mail message using the O365 API. The app-submitted request body 
    // contains only the reply.The API will include the original message in the 
    // final request body before sending it over the wire.
    app.post('/mail/reply', function (req, res, next) {
        var messageId = req.body.messageId; 

        var reqBody = { 'Comment' : req.body.comment };
        var reqHeaders = { 'content-type': 'application/json' };
        var reqUrl = appSettings.apiEndpoints.exchangeBaseUrl + "/messages/" + messageId + "/reply";
        var reqAuth = { 'bearer': passport.user.accessToken };
        
        request.post(
            { url: reqUrl, headers: reqHeaders, body: JSON.stringify(reqBody), auth: reqAuth },        
            function (err, response, body) {
                if (err) { next(err); }
                else {
                    if (response.statusCode == 403) {
                        err.status = response.statusCode;
                        err.stack = body;
                        err.message = "Failed to reply mail to " + req.body.to;
                        next(err);
                    }
                    else {
                        res.redirect('/mail');
                    }                    
                }
            }
        );
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

