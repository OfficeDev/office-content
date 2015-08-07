## Create the app ##

Let's dive right in! On your web server, create a new directory beneath the root directory called `php-tutorial`. For example, if you're using your development machine as your web server, the resulting URL would be `http://localhost/php-tutorial`. Within this directory create a `home.php` file and open it in your code editor of choice. This will be the home page for the app.

## Designing the app ##

Our app will be very simple. When a user visits the site, they will see a link to log in and view their email. Clicking that link will take them to the Azure login page where they can login with their Office 365 account and grant access to our app. Finally, they will be redirected back to our app, which will display a list of the most recent email in the user's inbox.

Let's begin by replacing the stock home page with a simpler one. Open the `./home.php` file and add the following code.

#### Contents of the `./home.php` file ####

    <?php
	  session_start();
      $loggedIn = false;
    ?>
    
    <html>
	  <head>
		<title>PHP Mail API Tutorial</title>
	  </head>
      <body>
    	<?php 
      	  if (!$loggedIn) {
    	?>
		  <!-- User not logged in, prompt for login -->
      	  <p>Please <a href="#">sign in</a> with your Office 365 account.</p>
    	<?php
  		  }
      	  else {
    	?>
		  <!-- User is logged in, do something here -->
      	  <p>Hello user!</p>
	    <?php
	      }
	    ?>
      </body>
    </html>

The link doesn't do anything yet, but we'll fix that soon.

## Implementing OAuth2 ##

Our goal in this section is to make the link on our home page initiate the [OAuth2 Authorization Code Grant flow with Azure AD](https://msdn.microsoft.com/en-us/library/azure/dn645542.aspx). 

Before we proceed, we need to register our app in Azure AD to obtain a client ID and secret. Head over to https://dev.outlook.com/appregistration to quickly get a client ID and secret. Assuming you are using `http://localhost` as your web server, use the following values for Step 2:

- **App Name:** php-tutorial
- **App Type:** Server-side Web app
- **Redirect URI:** http://localhost/php-tutorial/authorize.php
- **Home Page URL:** http://localhost/php-tutorial/home.php
- **Secret Valid For:** 1 year

Be sure to replace `http://localhost` with your correct web server address if you are using a different server.

![The Step 2 section of the App Registration Tool.](https://raw.githubusercontent.com/jasonjoh/php-tutorial/master/readme-images/registration-step-2.PNG)

In Step 3, select `Read mail`. If you plan on going beyond this tutorial and trying Calendar or Contacts API, go ahead and select additional permissions as well. For the purposes of this tutorial though, only `Read mail` is required.

![The Step 3 section of the App Registration Tool.](https://raw.githubusercontent.com/jasonjoh/php-tutorial/master/readme-images/registration-step-3.PNG)

After clicking the **Register App** button, copy your client ID and secret from the tool. We will use them in the next section.

### Back to coding ###

Create a new file to contain all of our OAuth functions called `oauth.php`. In this file we will define a helper class called `oAuthService`. Paste in the following code.

#### Contents of the `./oauth.php` file ####

	<?php
	  session_start();

	  class oAuthService {
	    private static $clientId = "YOUR CLIENT ID HERE";
	    private static $clientSecret = "YOUR CLIENT SECRET HERE";
	    private static $authority = "https://login.microsoftonline.com";
	    private static $authorizeUrl = '/common/oauth2/authorize?client_id=%1$s&redirect_uri=%2$s&response_type=code';
	    private static $tokenUrl = "/common/oauth2/token";
	    
	    public static function getLoginUrl($redirectUri) {
	      $loginUrl = self::$authority.sprintf(self::$authorizeUrl, self::$clientId, urlencode($redirectUri));
	      error_log("Generated login URL: ".$loginUrl);
	      return $loginUrl;
	    }
	  }
	?>

Paste in the client ID and secret you obtained from the App Registration Tool for the values of `$clientId` and `$clientSecret`.

The class exposes one function for now, `getLoginUrl`. This function will generate the URL to the Azure sign in page for the user to initiate the oAuth flow. Now we need to update the home page to call this function to provide a value for the `href` attribute of our sign in link.

#### Updated contents of the `./home.php` file ####

	<?php
	  session_start();
	  require('oauth.php');
	  
	  $loggedIn = false;
	?>
	
	<html>
	  <head>
  	    <title>PHP Mail API Tutorial</title>
	  </head>
	  <body>
	    <?php 
	      if (!$loggedIn) {
	    ?>
	      <!-- User not logged in, prompt for login -->
	      <p>Please <a href="<?php echo oAuthService::getLoginUrl('http://localhost/php-tutorial/authorize.php')?>">sign in</a> with your Office 365 account.</p>
	    <?php
	      }
	      else {
	    ?>
	      <!-- User is logged in, do something here -->
	      <p>Hello user!</p>
	    <?php    
	      }
	    ?>
	  </body>
	</html>

Save your work and copy the files to your web server. If you browse to `http://localhost/php-tutorial/home.php` and hover over the sign in link, it should look like this:

	https://login.microsoftonline.com/common/oauth2/authorize?client_id=<SOME GUID>&redirect_uri=http%3A%2F%2Flocalhost%2Fphp-tutorial%2Fauthorize.php&response_type=code

Clicking on the link will allow you to sign in and grant access to the app, but will then result in an error. Notice that your browser gets redirected to `http://localhost/php-tutorial/authorize.php`. That file doesn't exist yet. Don't worry, that's our next task.

Create the `authorize.php` file and add the following code.

#### Contents of the `./authorize.php` file ####

	<?php
	  $auth_code = $_GET['code'];
	?>
	
	<p>Auth code: <?php echo $auth_code ?></p>

This doesn't do anything but display the authorization code returned by Azure, but it will let us test that we can successfully log in. Now if you sign in to the app, you should end up with a page that shows the authorization code. Now let's do something with it.

### Exchanging the code for a token ###

Now let's add a new function to the `oAuthService` class to retrieve a token. Add the following function to the class in the `./oauth.php` file.

#### New `getTokenFromAuthCode` function in `./oauth.php` ####

	public static function getTokenFromAuthCode($authCode, $redirectUri) {
      // Build the form data to post to the OAuth2 token endpoint
      $token_request_data = array(
        "grant_type" => "authorization_code",
        "code" => $authCode,
        "redirect_uri" => $redirectUri,
        "resource" => "https://outlook.office365.com/",
        "client_id" => self::$clientId,
        "client_secret" => self::$clientSecret
      );
      
      // Calling http_build_query is important to get the data
      // formatted as Azure expects.
      $token_request_body = http_build_query($token_request_data);
      error_log("Request body: ".$token_request_body);
      
      $curl = curl_init(self::$authority.self::$tokenUrl);
      curl_setopt($curl, CURLOPT_RETURNTRANSFER, true);
      curl_setopt($curl, CURLOPT_POST, true);
      curl_setopt($curl, CURLOPT_POSTFIELDS, $token_request_body);
      
      $response = curl_exec($curl);
      error_log("curl_exec done.");
      
      $httpCode = curl_getinfo($curl, CURLINFO_HTTP_CODE);
      error_log("Request returned status ".$httpCode);
      if ($httpCode >= 400) {
        return array('errorNumber' => $httpCode,
                     'error' => 'Token request returned HTTP error '.$httpCode);
      }
      
      // Check error
      $curl_errno = curl_errno($curl);
      $curl_err = curl_error($curl);
      if ($curl_errno) {
        $msg = $curl_errno.": ".$curl_err;
        error_log("CURL returned an error: ".$msg);
        return array('errorNumber' => $curl_errno,
                     'error' => $msg);
      }
      
      curl_close($curl);
      
      // The response is a JSON payload, so decode it into
      // an array.
      $json_vals = json_decode($response, true);
      error_log("TOKEN RESPONSE:");
      foreach ($json_vals as $key=>$value) {
        error_log("  ".$key.": ".$value);
      }
      
      return $json_vals;
    }

This function uses cURL to issue the access token request to Azure. The first part of this function is building the payload of the request, then the rest is using cURL to issue a POST request to the Azure OAuth2 Token endpoint. Finally, the results are parsed into an array of values using `json_decode`.

Now replace the contents of the `./authorize.php` file with the following.

#### Updated contents of `./authorize.php` ####

	<?php
	  require_once('oauth.php');
	  $auth_code = $_GET['code'];
	  
	  $tokens = oAuthService::getTokenFromAuthCode($auth_code, 'http://localhost/php-tutorial/authorize.php');
	?>
	
	<p>Access Token: <?php echo $tokens['access_token'] ?></p>

Save your changes and restart the app. This time, after you sign in, you should see an access token displayed. Now let's update `./authorize.php` one more time to save the access token into a session variable and redirect back to the home page.

#### Updated contents of `./authorize.php` ####

	<?php
	  session_start();
	  require_once('oauth.php');
	  $auth_code = $_GET['code'];
	  
	  $tokens = oAuthService::getTokenFromAuthCode($auth_code, 'http://localhost/php-tutorial/authorize.php');
	  
	  if ($tokens['access_token']) {
	    $_SESSION['access_token'] = $tokens['access_token'];
	    
	    // Redirect back to home page
	    header("Location: http://localhost/php-tutorial/home.php");
	  }
	  else
	  {
	    echo "<p>ERROR: ".$tokens['error']."</p>";
	  }
	?>

Finally, let's update the `./home.php` file to check for the presence of the access token in the session and display it instead of the login link if present.

#### Updated contents of `./home.php` ####

	<?php
	  session_start();
	  require('oauth.php');
	  
	  $loggedIn = !is_null($_SESSION['access_token']);
	?>
	
	<html>
	  <head>
	    <title>PHP Mail API Tutorial</title>
	  </head>
	  <body>
	    <?php 
	      if (!$loggedIn) {
	    ?>
	      <!-- User not logged in, prompt for login -->
	      <p>Please <a href="<?php echo oAuthService::getLoginUrl('http://localhost/php-tutorial/authorize.php')?>">sign in</a> with your Office 365 account.</p>
	    <?php
	      }
	      else {
	    ?>
	      <!-- User is logged in, do something here -->
	      <p>Access token: <?php echo $_SESSION['access_token'] ?></p>
	    <?php    
	      }
	    ?>
	  </body>
	</html>

Now that we have an access token, we're ready to use the Mail API.

## Using the Mail API ##

Let's start by adding a new file to contain all of our Mail API functions called `outlook.php`. We'll start by creating a generic function `makeApiCall` that can be used to send REST requests. Paste in the following code.

#### Contents of `./outlook.php` ####

	<?php
	  class OutlookService {
	    public static function makeApiCall($access_token, $method, $url, $payload = NULL) {
	      // Generate the list of headers to always send.
	      $headers = array(
	        "User-Agent: php-tutorial/1.0",         // Sending a User-Agent header is a best practice.
	        "Authorization: Bearer ".$access_token, // Always need our auth token!
	        "Accept: application/json",             // Always accept JSON response.
	        "client-request-id: ".self::makeGuid(), // Stamp each new request with a new GUID.
	        "return-client-request-id: true"        // Tell the server to include our request-id GUID in the response.
	      );
	      
	      $curl = curl_init($url);
	      
	      switch(strtoupper($method)) {
	        case "GET":
	          // Nothing to do, GET is the default and needs no
	          // extra headers.
	          error_log("Doing GET");
	          break;
	        case "POST":
	          error_log("Doing POST");
	          // Add a Content-Type header (IMPORTANT!)
	          $headers[] = "Content-Type: application/json";
	          curl_setopt($curl, CURLOPT_POST, true);
	          curl_setopt($curl, CURLOPT_POSTFIELDS, $payload);
	          break;
	        case "PATCH":
	          error_log("Doing PATCH");
	          // Add a Content-Type header (IMPORTANT!)
	          $headers[] = "Content-Type: application/json";
	          curl_setopt($curl, CURLOPT_CUSTOMREQUEST, "PATCH");
	          curl_setopt($curl, CURLOPT_POSTFIELDS, $payload);
	          break;
	        case "DELETE":
	          error_log("Doing DELETE");
	          curl_setopt($curl, CURLOPT_CUSTOMREQUEST, "DELETE");
	          break;
	        default:
	          error_log("INVALID METHOD: ".$method);
	          exit;
	      }
	      
	      curl_setopt($curl, CURLOPT_RETURNTRANSFER, true);
	      curl_setopt($curl, CURLOPT_HTTPHEADER, $headers);
	      $response = curl_exec($curl);
	      error_log("curl_exec done.");
	      
	      $httpCode = curl_getinfo($curl, CURLINFO_HTTP_CODE);
	      error_log("Request returned status ".$httpCode);
	      
	      if ($httpCode >= 400) {
	        return array('errorNumber' => $httpCode,
	                     'error' => 'Request returned HTTP error '.$httpCode);
	      }
	      
	      $curl_errno = curl_errno($curl);
	      $curl_err = curl_error($curl);
	      
	      if ($curl_errno) {
	        $msg = $curl_errno.": ".$curl_err;
	        error_log("CURL returned an error: ".$msg);
	        curl_close($curl);
	        return array('errorNumber' => $curl_errno,
	                     'error' => $msg);
	      }
	      else {
	        error_log("Response: ".$response);
	        curl_close($curl);
	        return json_decode($response, true);
	      }
	    }
	    
	    // This function generates a random GUID.
	    public static function makeGuid(){
          if (function_exists('com_create_guid')) {
            error_log("Using 'com_create_guid'.");
            return strtolower(trim(com_create_guid(), '{}'));
          }
          else {
            error_log("Using custom GUID code.");
            $charid = strtolower(md5(uniqid(rand(), true)));
            $hyphen = chr(45);
            $uuid = substr($charid, 0, 8).$hyphen
                   .substr($charid, 8, 4).$hyphen
                   .substr($charid, 12, 4).$hyphen
                   .substr($charid, 16, 4).$hyphen
                   .substr($charid, 20, 12);
                 
            return $uuid;
          }
	    }
	  }
	?>

This function uses cURL to send the appropriate request to the specified endpoint, using the access token for authentication. We can use this function to call any of Outlook REST APIs. Let's add a new function to the `OutlookService` class to get the user's 10 most recent messages from the inbox.

In order to call our new `makeApiCall` function, we need an access token, a method, a URL, and an optional payload. We already have the access token, and from the [Mail API Reference](https://msdn.microsoft.com/office/office365/APi/mail-rest-operations#GetMessages), we know that the method to get messages is `GET` and that the URL to get messages is `https://outlook.office365.com/api/v1.0/me/messages`. Using that information, add a `getMessages` function in `outlook.php`.

#### New `getMessages` function in `./outlook.php` ####

	public static function getMessages($access_token) {
      $getMessagesUrl = self::$outlookApiUrl."/Me/Messages?"
                        ."\$select=Subject,DateTimeReceived,From"
                        ."&\$orderby=DateTimeReceived"
                        ."&\$top=10";
                        
      return self::makeApiCall($access_token, "GET", $getMessagesUrl);
    }

The function uses OData query parameters to do the following.

- Request that only the `Subject`, `DateTimeReceived`, and `From` fields for each message be returned. It's always a good idea to limit your result set to only those fields that you will use in your app.
- Sort the results by date and time each message was received.
- Limit the results to the first 10 items.

Update `./home.php` to call the `getMessages` function and display the results.

#### Updated contents of `./home.php` ####

	<?php
	  session_start();
	  require('oauth.php');
	  require('outlook.php');
	  
	  $loggedIn = !is_null($_SESSION['access_token']);
	?>
	
	<html>
	  <head>
	    <title>PHP Mail API Tutorial</title>
	  </head>
	  <body>
	    <?php 
	      if (!$loggedIn) {
	    ?>
	      <!-- User not logged in, prompt for login -->
	      <p>Please <a href="<?php echo oAuthService::getLoginUrl('http://localhost/php-tutorial/authorize.php')?>">sign in</a> with your Office 365 account.</p>
	    <?php
	      }
	      else {
	        $messages = OutlookService::getMessages($_SESSION['access_token']);
	    ?>
	      <!-- User is logged in, do something here -->
	      <p>Messages: <?php echo print_r($messages) ?></p>
	    <?php    
	      }
	    ?>
	  </body>
	</html>

If you restart the app now, you should get a very rough listing of the results array. Let's add a little HTML and PHP to display the results in a nicer way.

### Displaying the results ###

We'll add a basic HTML table to our home page, with columns for the subject, date and time received, and sender. We can then iterate over the results array and add rows to the table.

Update `./home.php` one final time to generate the table.

#### Updated contents of `./home.php` ####

	<?php
	  session_start();
	  require('oauth.php');
	  require('outlook.php');
	  
	  $loggedIn = !is_null($_SESSION['access_token']);
	?>
	
	<html>
	  <head>
		<title>PHP Mail API Tutorial</title>
	  </head>
	  <body>
	    <?php 
	      if (!$loggedIn) {
	    ?>
	      <!-- User not logged in, prompt for login -->
	      <p>Please <a href="<?php echo oAuthService::getLoginUrl('http://localhost/php-tutorial/authorize.php')?>">sign in</a> with your Office 365 account.</p>
	    <?php
	      }
	      else {
	        $messages = OutlookService::getMessages($_SESSION['access_token']);
	    ?>
	      <!-- User is logged in, do something here -->
	      <h2>Your messages</h2>
	      
	      <table>
	        <tr>
	          <th>From</th>
	          <th>Subject</th>
	          <th>Received</th>
	        </tr>
	        
	        <?php foreach($messages['value'] as $message) { ?>
	          <tr>
	            <td><?php echo $message['From']['EmailAddress']['Name'] ?></td>
	            <td><?php echo $message['Subject'] ?></td>
	            <td><?php echo $message['DateTimeReceived'] ?></td>
	          </tr>
	        <?php } ?>
	      </table>
	    <?php    
	      }
	    ?>
	  </body>
	</html>

Save your changes and run the app. You should now get a list of messages that looks something like this.

![The finished app displaying the user's inbox.](https://raw.githubusercontent.com/jasonjoh/php-tutorial/master/readme-images/inbox-listing.PNG)
