# report-generator
Generate reports through CLI or CRON using PHP and MS Graph.

This application uses the client credential flow to authorize the application to anonymously access information about your tenant without having a user present. This means we can grab all sorts of useful information about the state of our tenant and compile reports, push actions such as creating new Planner boards, and send emails.

In this sample, we use the PHP client library to authorize our application, then grab data from Groups, Users, OneDrive, and SharePoint, and compile a summary email to the admin.

## Getting Started
To use this script, you will need to have:
1. An O365 subscription on a valid tenant
2. An application registered at [apps.dev.microsoft.com](apps.dev.microsoft.com)
3. Access to an admin account for the specified tenant

### Register your application
1. Go to [apps.dev.microsoft.com](apps.dev.microsoft.com) and click "Register Your App"
2. Select "Web application" (even though this will be a script, we will do a one-time authorization with an admin on behalf of your tenant)
3. Give your application a name and enter your email address
4. Put in a redirect URI. This can be anything since we don't need to get an access token from the authorization process. I just entered ```http://localhost```
5. Under Application Secrets, click "Generate New Password" and copy the value to the report-generator.php file
6. The Application Permissions section defines which privileges your application as a whole has, whereas delegated permissions define which privileges an individual user logged into your app has. Since there won't be a user present, we need to ask for application permissions. At a minimum, you'll want to select
    1. Directory.Read.All
    2. Group.Read.All
    3. Mail.Send
    4. User.Read.All
7. Click Save
8. Copy the Application ID into the report-generator.php file

### Set base variables
1. Copy the Application ID and Secret into the corresponding variables in the script
2. Enter the email address you would like to send the report to
3. Enter the tenant name or ID you wish to collect data on

### Grant access to your tenant's information
The first time you want to use your application, you will need to grant the application permission to your tenant as an administrator. This is easily achieved by going to 

```https://login.microsoftonline.com/{tenant}/adminconsent?client_id={client_id}&state=12345&redirect_uri=http://localhost:8000```

in your web browser, signing in as a tenant admin, and accepting the scopes we previously requested. This will not need to be done again.

## Running the script
Call ```php report-generator.php``` from the CLI
