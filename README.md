# An Office 365 API sample app using Node, Express and Ejs

This simple Office 365 API sample app demonstrates how to program Office 365 REST API
in a Web application using [Nodejs](http://nodejs.org/), [Expressjs](http://expressjs.com/) 
and [Ejs](http://www.embeddedjs.com/). It is meant to provide a quick 
introduction, by way of a few tested examples, to getting started to explore  
Office 365 API features in an Express-based web application. Specifically, it covers how to

- [sign in to a user's Office 365 account](#sign-in), using the [passport-azure-oauth](https://www.npmjs.com/package/passport-azure-oauth) module,
- [get mail messages](#get-mails) from the user's inbox, using the [request](https://github.com/request/request) module,
- [view a specific mail message](#view-mail),
- [delete a specific mail message](#delete-mail),
- [send a new mail message](#send-mail) to a specified recipient,
- [reply a mail with comments](#reply-mail), 
- [get the user's calendar events](#get-calendar-events),
- [get the user's contacts](#get-contacts), 
- [inspect the user's files drive on OneDrive for Business](#get_files) and
- [make the app work](#make-app-work).

As a sample app to show programming of Office 365 APIs, no elaborate patterns or error handlings are attempted for node/express/ejs programming. 
 
For more information about Office 365 REST API, see the [API Docs](https://msdn.microsoft.com/en-us/office/office365/api/api-catalog).

<a name="sign-in">   
## Sign in to Office 365 
This demo uses the [_passport-azure-oauth_](https://www.npmjs.com/package/passport-azure-oauth) module for a user to sign in to his or her Office 365 account. 
For this to work, make sure that you have configured the _oauthOptions_, in the _appSettings.js_ file under the _/models_
sub directory of this application, to match the corresponding app settings you configured in the Azure Active Directory subscription.
For more information, see [Set up your Office 365 development environment](https://msdn.microsoft.com/en-us/office/office365/howto/setup-development-environment)

To sign in, a user selects the option on the app's home page to start a login request (`app.get('/login', ...)`), which is then rerouted
to the _passport-azure-oauth_ module to take the user through the Azure Active Directory
user authentication process (`app.get('/auth/azureOAuth', ...)`). Here, the user is asked to enter his or her Office 365 user credentials. 
Once the user is authenticated with permissions to use the app, the `passport-azure-oauth`  module returns initial access token, refresh token 
and other related information through the application's `redirect_uri`, 
which is set to be `http://localhost:1337/auth/azureOAuth/callback` in this app. To receive the results, the app supplies a callback 
function in the Express route of `app.get('/auth/azureOAuth/callback', callback)`. Notice that this Express route will work if a different 
host name and port are configured for the app. For details of an implementation, see the index.js file in the app's _routes_ directory.
 
The app caches the returned access token for use in subsequent HTTPS requests to access any Office 365 API functioinality. 
If the access token is expired, it can be refreshed using the refresh token, provided that the refresh token remains valid. Otherwise, the app will 
need to go through the sign-in process again.

<a name="get-mails">
## Get mail messages from the user's Inbox

To get mail messages from the signed-in user's Inbox, the app uses the _request_ module to submit a GET request on the Office 365 MAIL API 
resource as identified by this Url, `https://outlook.office365.com/api/v1.0/me/messages`. The access token must be specified as a bearer
ticket in the request's Authorization header. 

For this operation to work, the app must have been granted 
the `Mail.Read` scope. In this app, all the Mail API REST calls are implmented in the mail.js file in the _routes_ sub-directory.

The result contains a list of mail messages in JSON format.  

The HTTPS requests of Office 365 API resources are constitute client requests in a node.js app. They can be implemented in many different ways than that 
used in this app.

<a name="view-mail">
## View a specific mail message

Viewing a specific mail message involves a GET request against a Url of the `https://outlook.office365.com/api/v1.0/me/messages/{message-id}` format,
where `{message-id}` is a string value of the interested message Id. You must also supply a valid access token as part of the request.

<a name="delete-mail">
## Delete a specific mail message

Deleting a message involves submitting a DELETE request against a specific mail message. 

The app must have `Mail.Write` permission to delete a mail message.

<a name="send-mail">
## Send a new mail message to a specific recipient

Sending a new mail message is done by submitting a POST request on the `sendmail` resource, as identified by the `https://outlook.office365.com/api/v1.0/me/sendmail` 
Url. The request body contains the message specification, including the subject, message content, recipients and other related properties, 
in the JSON format.

For this to work, the app must have the `Mail.Send` scope.

<a name="reply-mail">
## Reply a mail message with comments
To reply a mail message sends a POST request on the `reply` resource. The app supplies to request body the comments that is 
to be added to the original message.

<a name="get-calendar-events">
## Get the user's calendar events

This involves a GET request against the calendar events resource (`https://outlook.office365.com/api/v1.0/me/calendar/events`).
The app must have the `Calendar.Read` or `Calendar.Write` permission.

<a name="get-contacts">
## Get the user's contacts

This involves a GET request against the calendar events resource (`https://outlook.office365.com/api/v1.0/me/contacts`).
The app must have the `Contacts.Read` or `Contacts.Write` permission.

<a name="get-files">
## Get the user's files on OneDrive for Business

This involves a GET request against the OneDrive for Business `drive` resource 
(`https://{tenant}-my.sharepoint.com/_api/v1.0/me/drive`). For an Office 365 developer site
with the domain name of `contoso.onmicrosoft.com`, the `{tenant}` value is `contoso`.
The app must have appropriate permissions to the Office 365 SharePoint Online service as configured
in Azure Active Directory.


<a name="make-app-work">
## Make the app work

- If you have not done so, install [node.js](http://nodejs.org/download/) to your working machine.

- Configure the app in your Azure Active Directory (AAD) subscription. For more information on how to do this in general, see the setup insturctions listed in
  [Office 365 APIs Starter project for Android](https://github.com/OfficeDev/O365-Android-Start).

- When granting the app permissions to other applications, make sure that no redudant permissions are selected. For example, 
  do not select both _Read users' mail_ and _Read and write access to users' mail_ for Office 365 Exchange Online because the former 
  is made redundant by the latter. Otherwise, you may get 403 error when trying to access the email service from the app.

- The _SIGN-ON URL_ value in AAD must match that assigned for the nodejs app. 
  for example if the node app is assigned an URL of `http://localhost:1337`, the AAD _SIGN-ON URL_ must have the same value. 

- To receive the Azure authentication/authoization results via `passport-azure-oauth` module,  the corresponding Express route 
  must have its path match the path of _REPLY URL_ in AAD. For example, if the _REPLY URL_ value is `http://localhost:1337/auth/azureOAuth/callback`, the app
  must enact a routing rule of the `app.get('/auth/azureOAuth/callback', callback)` format.
  
- To ensure all the node modules are included in the project, run the `npm install` command under the app's main directory, 
  where the _package.json_ file is located, from a shell window. 

  If using Visual Studio, right click the **npm** node in the **Solution Explorer** to select **Install Missing npm Packages** before running the app.

- To run the app in Visual Studio, hit F5. 

  To run the app in node shell, go to the app's bin directory and type `node www` to start the server. Then open a browser and enter "http://localhost:1337" in the address bar, assuming the default setting is preserved.

- Happy coding!



## Copyright
Copyright (c) Microsoft. All rights reserved.
