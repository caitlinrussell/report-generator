# report-generator
Generate reports through CLI or CRON using PHP and MS Graph

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
 a. Directory.Read.All
 b. Group.Read.All
 c. Mail.Send
 d. User.Read.All
7. Click Save
8. Copy the Application ID into the report-generator.php file

### Set base variables
1. Copy the Application ID and Secret into the corresponding variables in the script
2. Enter the email address you would like to send the report to
3. Enter the tenant name or ID you wish to collect data on

## Running the script
Call ```php report-generator.php``` from the CLI
