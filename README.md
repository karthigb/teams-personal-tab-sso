# Tabs

Tabs are Teams-aware webpages embedded in Microsoft Teams. Personal tabs are scoped to a single user. They can be pinned to the left navigation bar for easy access.

## Prerequisites
-  [NodeJS](https://nodejs.org/en/)

-  [ngrok](https://ngrok.com/)

-  [M365 developer account](https://docs.microsoft.com/en-us/microsoftteams/platform/concepts/build-and-test/prepare-your-o365-tenant) or access to a Teams account with the appropriate permissions to install an app.

-  [Create an Azure App registration to support SSO and the User.Read Graph API](https://aka.ms/teams-toolkit-sso-appreg)

## ngrok

Teams needs to access your tab from a publically accessible URL. If you are running your app in localhost, you will need to use a tunneling service like ngrok. Run ngrok and point it to localhost.
  `ngrok http http://localhost:3000`

Note: It may be worth purchasing a basic subscription to ngrok so you can get a fixed subdomain ( see the --subdomain ngrok parameter)

**IMPORTANT**: If you don't have a paid subscription to ngrok, you will need to update your Azure app registration application ID URI and redirect URL ( See steps 5 and 13 [here](https://docs.microsoft.com/en-us/microsoftteams/platform/tabs/how-to/authentication/auth-aad-sso#steps) ) everytime you restart ngrok.

## Update the env files

- In the root directory, open the *.env* file and update the *REACT_APP_AZURE_APP_REGISTRATION_ID* variable with the Azure app registration client ID you created above.
- In the *api-server* directory, open the *.env* file and update the *CLIENT_ID* and *CLIENT_SECRET* variables with the client ID and secret from your Azure app registration.
- In the *.publish* directory, open the *.env* file and update the *CLIENT_ID* variable with the same client ID as above. If you are using ngrok, update the *baseUrl0* variable with your ngrok url.

## Build and Run

In the root directory, execute:

`npm install`

`npm start`

In the *api-server* directory, execute:

`npm install`

`npm start`

## Deploy to Teams

- Go to the Teams client and select the *Apps* icon. Press *Upload a custom app* and select the `Development.zip` from the *.publish* folder in your project.
  - [Upload a custom app](https://aka.ms/teams-toolkit-uploadapp) 
