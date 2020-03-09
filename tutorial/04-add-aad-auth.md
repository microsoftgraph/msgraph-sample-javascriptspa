<!-- markdownlint-disable MD002 MD041 -->

In this exercise you will extend the application from the previous exercise to support authentication with Azure AD. This is required to obtain the necessary OAuth access token to call the Microsoft Graph. In this step you will integrate the [Microsoft Authentication Library](https://github.com/AzureAD/microsoft-authentication-library-for-js) library into the application.

1. Create a new file in the root directory named `config.js` and add the following code.

    ```javascript
    const msalConfig = {
      auth: {
        clientId: 'YOUR_APP_ID_HERE',
        redirectUri: 'http://localhost:8080'
      },
      cache: {
        cacheLocation: "sessionStorage",
        storeAuthStateInCookie: false,
        forceRefresh: false
      }
    };

    const loginRequest = {
      scopes: [
        'openid',
        'profile',
        'user.read',
        'calendars.read'
      ]
    }
    ```

    Replace `YOUR_APP_ID_HERE` with the application ID from the Application Registration Portal.

    > [!IMPORTANT]
    > If you're using source control such as git, now would be a good time to exclude the `config.js` file from source control to avoid inadvertently leaking your app ID.

1. Open `auth.js` and add the following code to the beginning of the file.

    ```javascript
    // Create the main MSAL instance
    // configuration parameters are located in config.js
    const msalClient = new Msal.UserAgentApplication(msalConfig);

    if (msalClient.getAccount() && !msalClient.isCallback(window.location.hash)) {
      // avoid duplicate code execution on page load in case of iframe and Popup window.
      updatePage(msalClient.getAccount(), Views.home);
    }
    ```

## Implement sign-in

In this section you'll implement the `signIn` function and get an access token.

1. Add the following function to `auth.js` to get an access token.

    ```javascript
    function getAccessToken(scopes) {
      // First attempt to get token silently
      // This should work if the user is logged in and has already
      // granted consent to the requested permission scopes
      return msalClient.acquireTokenSilent(scopes)
        .catch(error => {
          console.log('Silent token acquisition failed. Acquiring token using popup.');
          // Fallback to interactive when silent acquisition fails
          return msalClient.acquireTokenPopup(scopes)
            .then(tokenResponse => {
            }).catch(error => {
              console.log(error);
            });
        });
    }
    ```

1. Replace the existing `signIn` function with the following.

    ```javascript
    function signIn() {
      // Login
      msalClient.loginPopup(loginRequest)
        .then(loginResponse => {
          // Login response contains an ID token, which
          // MSAL uses to create an account object
          console.log('id_token acquired at: ' + new Date().toString());
          if (msalClient.getAccount()) {
            updatePage(msalClient.getAccount(), Views.home);
          }
        }).catch(error => {
          console.log(error);
          updatePage(null, Views.error, {
            message: 'Error logging in',
            debug: error
          });
        });
    }
    ```

1. Replace the existing `signOut` function with the following.

    ```javascript
    function signOut() {
      msalClient.logout();
    }
    ```

Save your changes and refresh the page. After you sign in, you should end up back on the home page, but the UI should change to indicate that you are signed-in.

![A screenshot of the home page after signing in](./images/user-signed-in.png)

Click the user avatar in the top right corner to access the **Sign out** link. Clicking **Sign out** resets the session and returns you to the home page.

![A screenshot of the dropdown menu with the Sign out link](./images/sign-out-button.png)

## Storing and refreshing tokens

At this point your application has an access token, which is sent in the `Authorization` header of API calls. This is the token that allows the app to access the Microsoft Graph on the user's behalf.

However, this token is short-lived. The token expires an hour after it is issued. This is where the refresh token becomes useful. The refresh token allows the app to request a new access token without requiring the user to sign in again.

Because the app is using the MSAL library, you do not have to implement any token storage or refresh logic. MSAL caches the token in the browser session. The `acquireTokenSilent` method first checks the cached token, and if it is not expired, it returns it. If it is expired, it uses the cached refresh token to obtain a new one.
