# Get started with Office 365 APIs in apps

The Office 365 APIs are REST services that provide access to Office 365 data, such as: mail, calendars, and contacts from Exchange Online; files and folders from SharePoint Online and OneDrive for Business; users and groups from Azure Active Directory.

##Get started with iOS development

<a name="bk_prerequisites"> </a>
##Before you start

**Note** These code samples use the [Office 365 SDK for iOS](https://github.com/OfficeDev/Office-365-SDK-for-iOS) to connect to Office 365.


Before you can create applications that access the Office 365 APIs, you'll need to set up your developer environment. This consists of three one-time tasks to make sure you've got the tools and environment to be successful:

1. Set up the iOS app development environment that you'll be using to create your apps. This includes installing and setting up the [CocoaPods](http://cocoapods.org) environment.

2. Get an Office 365 for business subscription, to access the Office 365 APIs.

3. Associate your Office 365 subscription with Azure Active Directory, so you can create and manage apps.

If you still need to complete any of these steps, take a look at [Set up your Office 365 development environment](..\howto\setup-development-environment.md) for detailed instructions on getting set up.

<a name="bk_createApp"> </a>
##Create your app and add dependencies
For this, you'll create an iOS project and use CocoaPods to add dependencies to the Office 365 SDK for iOS and the [Azure Active Directory Authentication Library (ADAL) for iOS](https://github.com/AzureAD/azure-activedirectory-library-for-objc).

###Create the SimpleMailApp project
1.	Open Xcode
2.	Choose **File > New > New**.
3.	Select the **Single View Application** template in the **iOS** project templates and click **Next**.
4.	Specify **SimpleMailApp** for the **Product Name**, select **Objective-C** for **Language**, **Universal** for **Devices**, a value for **Organization Identifier**, and then click **Next**.
5.	Select the location for your project, specify whether it should be under version control, and then click **Create**.
6.	Once the project has been created, close Xcode. 
 
The Cocoapods commands need to be run from the root of your project folder, so from Terminal, navigate to the project directory. If you did not change the default location when creating the project, the project will be located in the **SimpleMailApp** directory on the **Desktop**.

###Enable Cocoapods for the SimpleMailApp project
1.	Run the following command to initialize the Podfile for your project.

	```ShellSession
	pod init
	```

2.	Open the Podfile using the following command.

	```ShellSession
	Open podfile
	```

3.	Declare the dependencies for the Office 365 SDK for iOS and the ADAL SDK for iOS by adding the following definitions to the open Podfile.

	```ShellSession
	pod 'ADALiOS', '~> 1.2.0'
	pod 'Office365', '= 0.9.0'
	```

	These definitions should between the target and end statements, so the result will look like the following.

	```ShellSession
	target 'SimpleMailApp' do
	pod 'ADALiOS', '~> 1.2.0'
	pod 'Office365', '= 0.9.0'
	end
	```
     
4.	Close the Podfile.

**Note**  The Office 365 SDK for iOS is currently in developer preview, which means that it is subject to change prior to finalization, possibly even changes that would break code written using the SDK. The SDK uses the [Semantic Versioning](http://semver.org/) convention in which compatibility is specified using a three part version number: major.minor.patch. Until version 1.0 is reached, the minor version number will be incremented for breaking and other significant changes.

###Configure the dependencies for the SimpleMailApp project
To configure these dependencies and add them and the existing project to a new workspace, from Terminal, run the following command.

```ShellSession
    pod install
```

###Register your application with Azure Active Directory

If you already signed in and registered your app, you're all set, just follow the steps below.

Otherwise, find [instructions](https://msdn.microsoft.com/en-us/office/office365/howto/add-common-consent-manually) to do so in the Azure portal.

<a name="bk_codeYourApp"></a>
## Code your app 
Open the **SimpleMailApp** workspace in Xcode. 

###Authenticate with Azure AD and get the access token
An access token is required to access Office 365 APIs, so your application needs to implement the logic to retrieve and manage access tokens. The [Azure Active Directory authentication library (ADAL) for iOS and OSX](https://github.com/AzureAD/azure-activedirectory-library-for-objc) provides you with the ability to manage authentication in your application with just a few lines of code. The first thing you'll do is create a header file and class, AuthenticationManager that uses the ADAL for iOS and OSX to manage authentication for your app.

####To create the AuthenticationManager class
2. Right-click the SimpleMailApp project folder, select **New File**, and in the **iOS** section, click **Cocoa Touch Class**, and then click **Next**.
3. Specify **AuthenticationManager** as the **Class**, **NSObject** for **Subclass of:**, and then click **Next**.
4. Click **Create** to create the class and header files.

####To code the AuthenticationManager header file
1. Import the necessary Office 365 iOS SDK and ADAL SDK header files by adding the following code directives to AuthenticationManager.h.

	```objective-c
	#import <Foundation/Foundation.h>
	#import <ADALiOS/ADAuthenticationContext.h>
	#import <ADALiOS/ADAuthenticationSettings.h>
	#import <ADALiOS/ADLogger.h>
	#import <ADALiOS/ADInstanceDiscovery.h>
	#import <office365_odata_base/office365_odata_base.h>
	```
	
2. Declare a property for the **ADALDependencyResolver** object from the ADAL SDK which uses dependency injection to provide access to the authentication objects.

	```objective-c
	@property (readonly, nonatomic) ADALDependencyResolver *dependencyResolver;
	```
	
3. Specify the **AuthenticationManager** class as a singleton.

	```objective-c
	+(AuthenticationManager *)sharedInstance;
	```
	
4. Declare the methods for retrieving and clearing the access and refresh tokens.

	```objective-c
	//retrieve token
	-(void)acquireAuthTokenWithResourceId:(NSString *)resourceId completionHandler:(void (^)(BOOL authenticated))completionBlock;
	//clear token
	-(void)clearCredentials;
	```

####To code the AuthenticationManager class
1. Above the **@implementation** declaration, declare static variables for the redirect URI, client ID and the authority.

	```objective-c
	static NSString * const REDIRECT_URL_STRING = @"redirectUri";
	static NSString * const CLIENT_ID           = @"clientID";
	static NSString * const AUTHORITY           = @"https://login.microsoftonline.com/common";
	```

    Replace the *redirectUri* with the value you copied from step 8 when registering your app with Azure AD, and the *clientID* with the value you specified in step 13 of the same procedure.

2. In the ***interface** declaration, above **@implementation**, declare the following properties.

	```objective-c
	@interface AuthenticationManager ()
	@property (strong,    nonatomic) ADAuthenticationContext *authContext;
	@property (readwrite, nonatomic) ADALDependencyResolver  *dependencyResolver;
	@property (readonly, nonatomic) NSURL    *redirectURL;
	@property (readonly, nonatomic) NSString *authority;
	@property (readonly, nonatomic) NSString *clientId;
	@end
	```
3. Add code for the constructor to the implementation. 

	```objective-c
	-(instancetype)init
	{
	    self = [super init];
	    if (self) {
	        // These are settings that you need to set based on your
	        // client registration in Azure AD.
	        _redirectURL = [NSURL URLWithString:REDIRECT_URL_STRING];
	        _authority = AUTHORITY;
	        _clientId = CLIENT_ID;
	    }
	    return self;
	}
	```

4. Add the following code to use a single authentication manager for the application.

	```objective-c
	+(AuthenticationManager *)sharedInstance
	{
	    static AuthenticationManager *sharedInstance;
	    static dispatch_once_t onceToken;
	 
	    // Initialize the AuthenticationManager only once.
	    dispatch_once(&onceToken, ^{
	        sharedInstance = [[AuthenticationManager alloc] init];
	    });
	 
	    return sharedInstance;
	}
	```

5. Acquire access and refresh tokens from Azure AD for the user.

	```objective-c
	-(void)acquireAuthTokenWithResourceId:(NSString *)resourceId completionHandler:(void (^)(BOOL authenticated))completionBlock
	{
	        ADAuthenticationError *error;
	    self.authContext = [ADAuthenticationContext authenticationContextWithAuthority:self.authority error:&error];
	[self.authContext acquireTokenWithResource:resourceId
	                                      clientId:self.clientId
	                                   redirectUri:self.redirectURL
	                               completionBlock:^(ADAuthenticationResult *result) {
	                                   if (AD_SUCCEEDED != result.status) {
	                                       completionBlock(NO);
	                                   }
	                                   else {
	                                       NSUserDefaults *userDefaults = [NSUserDefaults standardUserDefaults];
	                                       [userDefaults setObject:result.tokenCacheStoreItem.userInformation.userId
	                                                        forKey:@"LogInUser"];
	                                       [userDefaults synchronize];
	                                       self.dependencyResolver = [[ADALDependencyResolver alloc] initWithContext:self.authContext
	                                                                                                      resourceId:resourceId
	                                                                                                        clientId:self.clientId
	                                                                                                     redirectUri:self.redirectURL];
	                                       completionBlock(YES);
	                                   }
	                               }];
	}
	```

    The first time the application runs, a request is sent to the URL specified for the AUTHORITY const (see step 1), which the redirects you to a login page where you can enter your credentials. If your login is successful, the response contains the access and refresh tokens. Subsequent times when the application runs, the authentication manager will use the access or refresh token for authenticating client requests, unless the token cache is cleared.

6. Finally, add code to log out the user by clearing the token cache and removing the application's cookies.

	```objective-c
	-(void)clearCredentials{
	    id<ADTokenCacheStoring> cache = [ADAuthenticationSettings sharedInstance].defaultTokenCacheStore;
	    ADAuthenticationError *error;
	 
	       if ([[cache allItemsWithError:&error] count] > 0)
	        [cache removeAllWithError:&error];
	        NSHTTPCookieStorage *cookieStore = [NSHTTPCookieStorage sharedHTTPCookieStorage];
	    for (NSHTTPCookie *cookie in cookieStore.cookies) {
	        [cookieStore deleteCookie:cookie];
	    }
	}
	```

###Connect to Office 365
Next you need to add class to your project to connect to Office 365, and use the Discovery service to retrieve the Exchange service endpoints.

####To create the Office365ClientFetcher class and header files
1. Right-click the SimpleMailApp project folder, , select **New File**, and in the **iOS** section, click **Cocoa Touch Class**, and then click **Next**.
2. Specify **Office365ClientFetcher** as the **Class**, **NSObject** for **Subclass of:**, and then click **Next**.
3. Click **Create** to create the class and header files.

####To code the Office365ClientFetcher header file
1. Import the necessary Office 365 iOS SDK header files by adding the following code directives to Office365ClientFetcher.h.

	```objective-c
	#import <Foundation/Foundation.h>
	#import <office365_odata_base/office365_odata_base.h>
	#import <office365_exchange_sdk/office365_exchange_sdk.h>
	#import "MSDiscoveryClient.h"
	#import <MSOutlookServicesClient.h>
	```

2. Declare the methods for fetching the Outlook and Discovery service clients.

	```objective-c
	-(void)fetchOutlookClient:(void (^)(MSOutlookServicesClient *outlookClient))callback;
	-(void)fetchDiscoveryClient:(void (^)(MSDiscoveryClient *discoveryClient))callback;
	```

####To code the Office365ClientFetcher class
1. Import the Office365ClientFetcher and AuthenticationManager header files.

	```objective-c
	#import "Office365ClientFetcher.h"
	#import "AuthenticationManager.h"
	```

2. Add the implementation code to the Office365ClientFetcher class.

	```objective-c
	-(void)fetchOutlookClient:(void (^)(MSOutlookServicesClient *outlookClient))callback
	{
	    // Get an instance of the authentication controller.
	    AuthenticationManager *authenticationManager = [AuthenticationManager sharedInstance];
	    [authenticationManager acquireAuthTokenWithResourceId:@"https://outlook.office365.com/"
	                                           completionHandler:^(BOOL authenticated) {
	        if(authenticated){
	            NSUserDefaults *userDefaults = [NSUserDefaults standardUserDefaults];
	            NSDictionary *serviceEndpoints = [userDefaults objectForKey:@"O365ServiceEndpoints"];
	            // Gets the MSOutlookServicesClient with the URL for the Mail service.
	            callback([[MSOutlookServicesClient alloc] initWithUrl:serviceEndpoints[@"Mail"]
	                                       dependencyResolver:authenticationManager.dependencyResolver]);
	        }
	        else{
	            //Display an alert in case of an error
	            dispatch_async(dispatch_get_main_queue(), ^{
	                NSLog(@"Error in the authentication");
	                UIAlertView *alert = [[UIAlertView alloc] initWithTitle:@"Error"
	                                                                message:@"Authentication failed. Check the log for errors."
	                                                               delegate:self
	                                                      cancelButtonTitle:@"OK"
	                                                      otherButtonTitles:nil];
	                [alert show];
	            });
	        }
	    }];
	}
	-(void)fetchDiscoveryClient:(void (^)(MSDiscoveryClient *discoveryClient))callback
	{
	    AuthenticationManager *authenticationManager = [AuthenticationManager sharedInstance];
	    [authenticationManager acquireAuthTokenWithResourceId:@"https://api.office.com/discovery/"
	                                        completionHandler:^(BOOL authenticated) {
	        if (authenticated) {
	            callback([[MSDiscoveryClient alloc] initWithUrl:@"https://api.office.com/discovery/v1.0/me/"
	                                         dependencyResolver:authenticationManager.dependencyResolver]);
	        }
	        else {
	            dispatch_async(dispatch_get_main_queue(), ^{
	                NSLog(@"Error in the authentication");
	                UIAlertView *alert = [[UIAlertView alloc] initWithTitle:@"Error"
	                                                                message:@"Authentication failed. This may be because the Internet connection is offline  or perhaps the credentials are incorrect. Check the log for errors and try again."
	                                                               delegate:self
	                                                      cancelButtonTitle:@"OK"
	                                                      otherButtonTitles:nil];
	                [alert show];
	            });
	        }
	    }];
	}
	```

###The View Controller
Next you need to call the methods to connect to the Office 365 services, triggering authentication, and then display the results in the UI. Here you're going to create a View Controller class containing UI controls using the same names as the View Controller and UI controls in the Office 365 iOS Connect sample. This enables you to download and use the storyboard from the Connect sample, allowing you to skip the steps normally required to connect the storyboard to a View Controller.

####To add the View Controller
1. Right-click the SimpleMailApp project, select **New File**, and in the **iOS** section, click **Cocoa Touch Class**, and then click **Next**.
2. Select **UIViewController** for **Subclass of:**, and specify **SendMailViewController** as the **Class** and then click **Next**.
3. Click **Create** to create the class and header files. 

####To code the View Controller
To start, import the necessary header files by adding the following code directives to SendMailViewController.m.

```objective-c
#import "SendMailViewController.h"
#import "Office365ClientFetcher.h"
#import "AuthenticationManager.h"
#import "MSDiscoveryServiceInfoCollectionFetcher.h"
```

Next, declare the following properties and methods in the **SendMailViewController** interface.

```objective-c
@interface SendMailViewController ()

@property (weak, nonatomic) IBOutlet UILabel *headerLabel;
@property (weak, nonatomic) IBOutlet UITextView *mainContentTextView;
@property (weak, nonatomic) IBOutlet UITextView *statusTextView;
@property (weak, nonatomic) IBOutlet UITextField *emailTextField;
@property (weak, nonatomic) IBOutlet UIBarButtonItem *disconnectBarButtonItem;
@property (weak, nonatomic) IBOutlet UIButton *sendMailButton;
@property (weak, nonatomic) IBOutlet UIActivityIndicatorView *activityIndicator;

@property (strong, nonatomic) Office365ClientFetcher *baseController;
@property (strong, nonatomic) NSMutableDictionary *serviceEndpointLookup;

-(IBAction)sendMailTapped:(id)sender;
-(IBAction)disconnectTapped:(id)sender;

@end
```

Now add the following code for the **SendMailViewController** implementation.

```objective-c
@implementation SendMailViewController

#pragma mark - Properties
-(Office365ClientFetcher *)baseController
{
    if (!_baseController) {
        _baseController = [[Office365ClientFetcher alloc] init];
    }

    return _baseController;
}

#pragma mark - Lifecycle Methods
- (void)viewDidLoad
{
    [super viewDidLoad];

    self.disconnectBarButtonItem.enabled = NO;
    self.sendMailButton.hidden = YES;
    self.emailTextField.hidden = YES;
    self.mainContentTextView.hidden = YES;
    self.headerLabel.hidden = YES;

    [self connectToOffice365];
}

#pragma mark - IBActions
//Send a mail message when the Send button is clicked
-(IBAction)sendMailTapped:(id)sender
{
    [self sendMailMessage];
}

// Clear the token cache and update the UI when the Disconnect button is tapped
-(IBAction)disconnectTapped:(id)sender
{
    self.disconnectBarButtonItem.enabled = NO;
    self.sendMailButton.hidden = YES;
    self.mainContentTextView.text = @"You're no longer connected to Office 365.";
    self.headerLabel.hidden = YES;
    self.emailTextField.hidden = YES;
    self.statusTextView.hidden = YES;

    // Clear the access and refresh tokens from the credential cache. You need to clear cookies
    // since ADAL uses information stored in the cookies to get a new access token.
    AuthenticationManager *authenticationManager = [AuthenticationManager sharedInstance];
    [authenticationManager clearCredentials];
}

#pragma mark - Helper Methods
-(void)connectToOffice365
{
    [self.baseController fetchDiscoveryClient:^(MSDiscoveryClient *discoveryClient) {
        MSDiscoveryServiceInfoCollectionFetcher *servicesInfoFetcher = [discoveryClient getservices];

        // Call the Discovery Service and get back an array of service endpoint information
        NSURLSessionTask *servicesTask = [servicesInfoFetcher readWithCallback:^(NSArray *serviceEndpoints, MSODataException *error) {
            if (serviceEndpoints) {

                self.serviceEndpointLookup = [[NSMutableDictionary alloc] init];

                for(MSDiscoveryServiceInfo *serviceEndpoint in serviceEndpoints) {
                    self.serviceEndpointLookup[serviceEndpoint.capability] = serviceEndpoint.serviceEndpointUri;
                }

                // Keep track of the service endpoints in the user defaults
                NSUserDefaults *userDefaults = [NSUserDefaults standardUserDefaults];

                [userDefaults setObject:self.serviceEndpointLookup
                                 forKey:@"O365ServiceEndpoints"];

                [userDefaults synchronize];

                dispatch_async(dispatch_get_main_queue(), ^{
                   
                    NSString *userEmail = [userDefaults stringForKey:@"LogInUser"];
                    NSArray *parts = [userEmail componentsSeparatedByString: @"@"];

                    self.headerLabel.text = [NSString stringWithFormat:@"Hi %@!", parts[0]];
                    self.headerLabel.hidden = NO;
                    self.mainContentTextView.hidden = NO;
                    self.emailTextField.text = userEmail;
                    self.statusTextView.text = @"";
                    self.disconnectBarButtonItem.enabled = YES;
                    self.sendMailButton.hidden = NO;
                    self.emailTextField.hidden = NO;
                });
            }
            else {
                dispatch_async(dispatch_get_main_queue(), ^{
                    NSLog(@"Error in the authentication: %@", error);

                    UIAlertView *alert = [[UIAlertView alloc] initWithTitle:@"Error"
                                                                    message:@"Authentication failed. This may be because the Internet connection is offline  or perhaps the credentials are incorrect. Check the log for errors and try again."
                                                                   delegate:nil
                                                          cancelButtonTitle:@"OK"
                                                          otherButtonTitles:nil];
                    [alert show];
                });
            }
        }];
        
        [servicesTask resume];
    }];
}

-(void)sendMailMessage
{
    MSOutlookServicesMessage *message = [self buildMessage];

    [self.baseController fetchOutlookClient:^(MSOutlookServicesClient *outlookClient) {
        dispatch_async(dispatch_get_main_queue(), ^{
            // Show the activity indicator
            [self.activityIndicator startAnimating];
        });

        MSOutlookServicesUserFetcher *userFetcher = [outlookClient getMe];
        MSOutlookServicesUserOperations *userOperations = [userFetcher operations];

        // Send the mail message. This results in a call to the service.
        NSURLSessionTask *task = [userOperations sendMailWithMessage:message
                                                     saveToSentItems:YES
                                                            callback:^(int returnValue, MSODataException *error) {
            NSString *statusText;

            if (error == nil) {
                statusText = @"Check your inbox, you have a new message. :)";
            }
            else {
                statusText = @"The email could not be sent. Check the log for errors.";
                NSLog(@"%@",[error localizedDescription]);
            }

            // Update the UI.
            dispatch_async(dispatch_get_main_queue(), ^{
                self.statusTextView.text = statusText;
                [self.activityIndicator stopAnimating];
            });
        }];

        [task resume];
    }];
}

//Compose the mail message
-(MSOutlookServicesMessage *)buildMessage
{
    // Create a new message. Set properties on the message.
    MSOutlookServicesMessage *message = [[MSOutlookServicesMessage alloc] init];
    message.Subject = @"Welcome to Office 365 development on iOS with the Office 365 Connect sample";

    // Get the recipient's email address.
    // See the helper method getRecipients to understand the usage.
    NSString *toEmail = self.emailTextField.text;

    MSOutlookServicesRecipient *recipient = [[MSOutlookServicesRecipient alloc] init];

    recipient.EmailAddress = [[MSOutlookServicesEmailAddress alloc] init];
    recipient.EmailAddress.Address = [toEmail stringByTrimmingCharactersInSet:[NSCharacterSet whitespaceCharacterSet]];

    message.ToRecipients = (NSMutableArray<MSOutlookServicesRecipient> *)[[NSMutableArray alloc] initWithObjects:recipient, nil];

    // Get the email text and put in the email body.
    NSString *filePath = [[NSBundle mainBundle] pathForResource:@"EmailBody" ofType:@"html" ];
    NSString *body = [NSString stringWithContentsOfFile:filePath encoding:NSUTF8StringEncoding error:nil];
    message.Body = [[MSOutlookServicesItemBody alloc] init];
    message.Body.ContentType = MSOutlookServices_BodyType_HTML;
    message.Body.Content = body;

    return message;
}

@end
```
####Main.storyboard
Since the View Controller you've created is named the same as the one in the [Office 365 iOS Connect sample](https://github.com/OfficeDev/O365-iOS-Connect), with the code specifying controls with the same names as the ones in the sample, you can use the sample's storyboard. To get the storyboard, [download](https://github.com/OfficeDev/O365-iOS-Connect/archive/master.zip) the sample, locate the **Main.storyboard** file in the O365-iOS-Connect-master\Sample\O365-iOS-Connect\Base.lproj folder of the download.

Copy this file to the SimpleMailApp\SimpleMailApp\Base.lproj folder in your project, overwriting the existing version of the file in that location.

###Email source
The final step is to create an HTML file that the sample uses to construct the body of the email. Add an HTML file to your project and name it EmailBody.html. You can customize it however you want, or use the following code for a basic test.

```html
<html>
    <body>
        <p>Sent from SimpleMailApp iOS sample</p>
    </body>
</html>
```

<a name="bk_test" />
##Testing the app
Build and run the application from Xcode. This launches the iOS Simulator, and you'll see a **Connect to Office 365** link. Click it, and you will be prompted to enter your credentials. Once you have successfully logged in, you will see a textbox with the email address for the account you logged in as and a **Send** link. Click **Send** to send an email message to the address specified in the textbox.
