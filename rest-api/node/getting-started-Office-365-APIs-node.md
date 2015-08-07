# Getting Started with the Outlook Mail API and Node.js #

The purpose of this guide is to walk through the process of creating a simple Node.js app that retrieves messages in Office 365. The source code in this repository is what you should end up with if you follow the steps outlined here.


Let's dive right in! Create an empty directory where you want to create your new Node.js app. For the purposes of this guide I will assume the name of the directory is `node-tutorial`, but feel free to use any name you like. Using your favorite JavaScript editor, create a new file called `server.js`. Paste the following code into `server.js` and save it.

### Contents of the `.\server.js` file ###

    var http = require("http");
    var url = require("url");
    
    function start(route, handle) {
      function onRequest(request, response) {
	    var pathName = url.parse(request.url).pathname;
	    console.log("Request for " + pathName + " received.");
	    
	    route(handle, pathName, response, request);
      }
      
      http.createServer(onRequest).listen(8000);
      console.log("Server has started.");
    }
    
    exports.start = start;

If you're familiar with Node.js, this is nothing new for you. If you're new to it, this is basic code to allow Node to run a web server listening on port 8000. When requests come in, it sends them to the `route` function, which we need to implement!

Create a new file called `router.js`, and add the following code.

### Contents of the `.\router.js` file ###

    function route(handle, pathname, response, request) {
      console.log("About to route a request for " + pathname);
      if (typeof handle[pathname] === 'function') {
	    return handle[pathname](response, request);
      } else {
		    console.log("No request handler found for " + pathname);
		    response.writeHead(404 ,{"Content-Type": "text/plain"});
		    response.write("404 Not Found");
		    response.end();
	      }
    }
    
    exports.route = route;

This code looks up a function to call based on the requested path. It uses the `handle` array, which we haven't defined yet. Create a new file called `index.js`, and add the following code.

### Contents of the `.\index.js` file ###

    var server = require("./server");
    var router = require("./router");
    
    var handle = {};
    handle["/"] = home;
    
    server.start(router.route, handle);
    
    function home(response, request) {
      console.log("Request handler 'home' was called.");
      response.writeHead(200, {"Content-Type": "text/html"});
      response.write('<p>Hello world!</p>');
      response.end();
    }

At this point, you should have a working app. Open a command prompt to the directory where your files are located, and enter the following command.

    node index.js

You should get a confirmation saying `Server has started.` Open your browser and navigate to [http://localhost:8000](http://localhost:8000). You should see "Hello world!".

Now that we've confirmed that the app is working, we're ready to do some real work.

## Designing the app ##

Our app will be very simple. When a user visits the site, they will see a link to log in and view their email. Clicking that link will take them to the Azure login page where they can login with their Office 365 account and grant access to our app. Finally, they will be redirected back to our app, which will display a list of the most recent email in the user's inbox.

Let's begin by replacing the "Hello world!" message with a signon link. To do that, we'll modify the `home` function in `index.js`. Open this file in your favorite text editor. Update the `home` function to match the following.

### Updated `home` function ###

    function home(response, request) {
      console.log("Request handler 'home' was called.");
      response.writeHead(200, {"Content-Type": "text/html"});
      response.write('<p>Please <a href="#">sign in</a> with your Office 365 account.</p>');
      response.end();
    }

As you can see, our home page will be very simple. For now, the link doesn't do anything, but we'll fix that soon.

## Implementing OAuth2 ##

Our goal in this section is to make the link on our home page initiate the [OAuth2 Authorization Code Grant flow with Azure AD](https://msdn.microsoft.com/en-us/library/azure/dn645542.aspx). To make things easier, we'll use the [simple-oauth2 library](https://github.com/andreareginato/simple-oauth2) to handle our OAuth requests. At your command prompt, enter the following command.

    npm install simple-oauth2

Now the library is installed and ready to use. Create a new file called `authHelper.js`. We'll start here by defining a function to generate the login URL.

### Contents of the `.\authHelper.js` file ###

    var credentials = {
      clientID: "YOUR CLIENT ID HERE",
      clientSecret: "YOUR CLIENT SECRET HERE",
      site: "https://login.microsoftonline.com/common",
      authorizationPath: "/oauth2/authorize",
      tokenPath: "/oauth2/token"
    }
    var oauth2 = require("simple-oauth2")(credentials)
    
    function getAuthUrl() {
      var returnVal = oauth2.authCode.authorizeURL({
    	redirect_uri: "http://localhost:8000/authorize"
      });
      console.log("Generated auth url: " + returnVal);
      return returnVal;
    }
    
    exports.getAuthUrl = getAuthUrl;

The first thing we do here is define our client ID and secret. We also define a redirect URI as a hard-coded value. The values of `clientId` and `clientSecret` are just placeholders, so we need to generate valid values.

### Generate a client ID and secret ###

To get a client ID and secret, we need to [register the app](https://github.com/jasonjoh/office365-azure-guides/blob/master/RegisterAnAppInAzure.md). Use the following details to register.

#### Create parameters ####

- Name: node-tutorial
- Type: Web application and/or Web API

![](https://raw.githubusercontent.com/jasonjoh/node-tutorial/master/readme-images/azure-wizard1.PNG)
- Sign-on URL: http://localhost:8000
- App ID URL: https://your_Office365_domain/node-tutorial (Replace 'your_Office365_domain' with your actual Office 365 domain!)

![](https://raw.githubusercontent.com/jasonjoh/node-tutorial/master/readme-images/azure-wizard-2.PNG)

#### App configuration ####

- Keys: 1 year.
- Permissions to other applications: Office 365 Exchange Online, Delegated Permissions, "Read user's mail"

![](https://raw.githubusercontent.com/jasonjoh/node-tutorial/master/readme-images/azure-portal-3.PNG)

Once this is complete you should have a client ID and a secret. Replace the `<YOUR CLIENT ID>` and `<YOUR CLIENT SECRET>` placeholders with these values and save your changes.

### Back to coding ###

Now that we have actual values for the client ID and secret, let's put the `simple-oauth`library to work. Modify the `home` function in the `index.js` file to use the `getAuthUrl` function to fill in the link. You'll need to require the `authHelper` file to gain access to this function.

#### Updated contents of the `.\index.js` file ####

    var server = require("./server");
    var router = require("./router");
    var authHelper = require("./authHelper");
    
    var handle = {};
    handle["/"] = home;
    
    server.start(router.route, handle);
    
    function home(response, request) {
      console.log("Request handler 'home' was called.");
      response.writeHead(200, {"Content-Type": "text/html"});
      response.write('<p>Please <a href="' + authHelper.getAuthUrl() + '">sign in</a> with your Office 365 account.</p>');
      response.end();
    }

Save your changes and browse to [http://localhost:8000](http://localhost:8000). If you hover over the link, it should look like:

    https://login.microsoftonline.com/common/oauth2/authorize?client_id=<SOME GUID>&redirect_uri=http%3A%2F%2Flocalhost%3A8000%2Fauthorize&response_type=code

The `<SOME GUID>` portion should match your client ID. Click on the link and (assuming you are not already signed in to Office 365 in your browser), you should be presented with a sign in page:

![The Azure sign-in page.](https://raw.githubusercontent.com/jasonjoh/node-tutorial/master/readme-images/azure-sign-in.PNG)

Sign in with your Office 365 account. Your browser should redirect to back to our app, and you should see a lovely error:

    404 Not Found

The reason we're seeing the error is because we haven't implemented a route to handle the `/authorize` path we hard-coded as our redirect URI. Let's fix that error now.

### Exchanging the code for a token ###

First, let's add a route for the `/authorize` path to the `handle` array in `index.js`.

#### Updated handle array in `.\index.js`####

    var handle = {};
    handle["/"] = home;
    handle["/authorize"] = authorize;

The added line tells our router that when a GET request comes in for `/authorize`, invoke the `authorize` function. So to make this work, we need to implement that function. Add the following function to `index.js`.

#### `authorize` function ####

    var url = require("url");
    function authorize(response, request) {
      console.log("Request handler 'authorize' was called.");
      
      // The authorization code is passed as a query parameter
      var url_parts = url.parse(request.url, true);
      var code = url_parts.query.code;
      console.log("Code: " + code);
      response.writeHead(200, {"Content-Type": "text/html"});
      response.write('<p>Received auth code: ' + code + '</p>');
      response.end();
    }

Restart the Node server and refresh your browser (or repeat the sign-in process). Now instead of an error, you should see the value of the authorization code printed on the screen. We're getting closer, but that's still not very useful. Let's actually do something with that code.

Let's add another helper function to `authHelper.js` called `getTokenFromCode`.

#### `getTokenFromCode` in the `.\authHelper.js` file ####

    function getTokenFromCode(auth_code, resource, callback, response) {
      var token;
      oauth2.authCode.getToken({
	    code: auth_code,
	    redirect_uri: redirectUri,
	    resource: resource
	    }, function (error, result) {
	      if (error) {
		    console.log("Access token error: ", error.message);
		    callback(response, error, null);
	      }
	      else {
		    token = oauth2.accessToken.create(result);
		    console.log("Token created: ", token.token);
		    callback(response, null, token);
	      }
	    });
    }

    exports.getTokenFromCode = getTokenFromCode;

Let's make sure that works. Modify the `authorize` function in the `index.js` file to use this helper function and display the return value. Note that `getToken` function is asynchronous, so we need to implement a callback function to receive the results.

#### Updated `authorize` function in `.\index.js` ####

    function authorize(response, request) {
      console.log("Request handler 'authorize' was called.");
      
      // The authorization code is passed as a query parameter
      var url_parts = url.parse(request.url, true);
      var code = url_parts.query.code;
      console.log("Code: " + code);
      var token = authHelper.getTokenFromCode(code, 'https://outlook.office365.com/', tokenReceived, response);
    }

#### Callback function `tokenReceived` in `.\index.js` ####

    function tokenReceived(response, error, token) {
      if (error) {
	    console.log("Access token error: ", error.message);
	    response.writeHead(200, {"Content-Type": "text/html"});
	    response.write('<p>ERROR: ' + error + '</p>');
	    response.end();
      }
      else {
	    response.writeHead(200, {"Content-Type": "text/html"});
	    response.write('<p>Access token: ' + token.token.access_token + '</p>');
	    response.end();
      }
    }

If you save your changes, restart the server, and go through the sign-in process again, you should now see a long string of seemingly nonsensical characters. If everything's gone according to plan, that should be an access token. Copy the entire value and head over to http://jwt.calebb.net/. If you paste that value in, you should see a JSON representation of an access token. For details and alternative parsers, see [Validating your Office 365 Access Token](https://github.com/jasonjoh/office365-azure-guides/blob/master/ValidatingYourToken.md).

Once you're convinced that the token is what it should be, let's change our code to store the token in a session cookie instead of displaying it.

#### New version of `tokenReceived` function ####
    function tokenReceived(response, error, token) {
      if (error) {
	    console.log("Access token error: ", error.message);
	    response.writeHead(200, {"Content-Type": "text/html"});
	    response.write('<p>ERROR: ' + error + '</p>');
	    response.end();
      }
      else {
	    response.setHeader('Set-Cookie', ['node-tutorial-token =' + token.token.access_token + ';Max-Age=3600']);
	    response.writeHead(200, {"Content-Type": "text/html"});
	    response.write('<p>Access token saved in cookie.</p>');
	    response.end();
      }
    }

## Using the Mail API ##

Now that we can get an access token, we're in a good position to do something with the Mail API. Let's start by creating a `mail` route and function. Open the `index.js` file and update the `handle` array.

#### Updated handle array in `.\index.js`####

    var handle = {};
    handle["/"] = home;
    handle["/authorize"] = authorize;
    handle["/mail"] = mail;

#### `mail` function in `.\index.js`####

    function mail(response, request) {
      var cookieName = 'node-tutorial-token';
      var cookie = request.headers.cookie;
      if (cookie && cookie.indexOf(cookieName) !== -1) {
	    console.log("Cookie: ", cookie);
	    // Found our token, extract it from the cookie value
	    var start = cookie.indexOf(cookieName) + cookieName.length + 1;
	    var end = cookie.indexOf(';', start);
	    end = end === -1 ? cookie.length : end;
	    var token = cookie.substring(start, end);
	    console.log("Token found in cookie: " + token);
    
	    response.writeHead(200, {"Content-Type": "text/html"});
	    response.write('<p>Token retrieved from cookie: ' + token + '</p>');
	    response.end();
      }
      else {
	    response.writeHead(200, {"Content-Type": "text/html"});
	    response.write('<p> No token found in cookie!</p>');
	    response.end();
      }
    }

For now all this does is read the token back from the cookie and display it. Save your changes, restart the server, and go through the signon process again. You should see the token displayed. Now that we know we have access to the token in the `mail` function, we're ready to call the Mail API.

In order to use the Mail API, install the [node-outlook library](https://github.com/jasonjoh/node-outlook) from the command line.

    npm install node-outlook

The `node-outlook` library expects a callback function that it can use to request your access token. Let's implement that in `authHelper.js`.

#### Access token callbacks in `authHelper.js` ####
	var outlook = require("node-outlook");
    function getAccessToken(token) {
      var deferred = new outlook.Microsoft.Utility.Deferred();
	  deferred.resolve(token);
      return deferred;
    }
    
    function getAccessTokenFn(token) {
      return function() {
    	return getAccessToken(token);
      }
    }

	exports.getAccessTokenFn = getAccessTokenFn;

Now we can modify the `mail` function to use this library and retrieve email. First, require the `node-outlook` library by adding the following line to `index.js`.

    var outlook = require("node-outlook");

Then update the `mail` function to query the inbox.

#### New version of the `mail` function in `./index.js` ####

    function mail(response, request) {
      var cookieName = 'node-tutorial-token';
      var cookie = request.headers.cookie;
      if (cookie && cookie.indexOf(cookieName) !== -1) {
	    console.log("Cookie: ", cookie);
	    // Found our token, extract it from the cookie value
	    var start = cookie.indexOf(cookieName) + cookieName.length + 1;
	    var end = cookie.indexOf(';', start);
	    end = end === -1 ? cookie.length : end;
	    var token = cookie.substring(start, end);
	    console.log("Token found in cookie: " + token);
    
	    var outlookClient = new outlook.Microsoft.OutlookServices.Client('https://outlook.office365.com/api/v1.0', 
	      authHelper.getAccessTokenFn(token));
    
	    response.writeHead(200, {"Content-Type": "text/html"});
	    response.write('<div><span>Your inbox</span></div>');
	    response.write('<table><tr><th>From</th><th>Subject</th><th>Received</th></tr>');
    
	    outlookClient.me.messages.getMessages()
	    .orderBy('DateTimeReceived desc')
		.select('DateTimeReceived,From,Subject').fetchAll(10).then(function (result) {
	      result.forEach(function (message) {
		    var from = message.from ? message.from.emailAddress.name : "NONE";
		    response.write('<tr><td>' + from + 
		      '</td><td>' + message.subject +
		      '</td><td>' + message.dateTimeReceived.toString() + '</td></tr>');
	      });
  
	      response.write('</table>');
	      response.end();
    	},function (error) {
	      console.log(error);
	      response.write("<p>ERROR: " + error + "</p>");
	      response.end();
    	});
      }
      else {
	    response.writeHead(200, {"Content-Type": "text/html"});
	    response.write('<p> No token found in cookie!</p>');
	    response.end();
      }
    }


To summarize the new code in the `mail` function:

- It creates an `OutlookServices.Client` object, passing it the API endpoint, `https://outlook.office365.com/api/v1.0`, and a pointer to the access token callback we implemented earlier.
- It issues a GET request to the URL for inbox messages, with the following characteristics:
	- It uses the `OrderBy()` function with a value of `DateTimeReceived desc` to sort the results by DateTimeReceived.
	- It uses the `Select()` function to only request the `DateTimeReceived`, `From`, and `Subject` properties.
	- It uses the `fetchAll()` function with a value of `10` to limit the results to the first 10.
- It loops over the results and prints out the sender, the subject, and the date/time the message was received.

### Displaying the results ###

Save the changes and sign in to the app. You should now see a simple table of messages in your inbox.

![An HTML table displaying the contents of an inbox.](https://raw.githubusercontent.com/jasonjoh/node-tutorial/master/readme-images/inbox.PNG)

## Next Steps ##

Now that you've created a working sample, you may want to learn more about the [capabilities of the Mail API](https://msdn.microsoft.com/office/office365/APi/mail-rest-operations). If your sample isn't working, and you want to compare, you can download the end result of this tutorial from [GitHub](https://github.com/jasonjoh/node-tutorial).

## Copyright ##

Copyright (c) Microsoft. All rights reserved.

----------
Connect with me on Twitter [@JasonJohMSFT](https://twitter.com/JasonJohMSFT)

Follow the [Exchange Dev Blog](http://blogs.msdn.com/b/exchangedev/)
