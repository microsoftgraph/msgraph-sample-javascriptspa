
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

// <authInitSnippet>
// Create the main MSAL instance
// configuration parameters are located in config.js
const msalClient = new msal.PublicClientApplication(msalConfig);
// </authInitSnippet>

// <checkAuthSnippet>
// Check for an already logged-in user
const account = msalClient.getActiveAccount();
if (account) {
  initializeGraphClient(msalClient, account, msalRequest.scopes);
}
// </checkAuthSnippet>

// <signInSnippet>
async function signIn() {
    // Login
    try {
      // Use MSAL to login
      const authResult = await msalClient.loginPopup(msalRequest);
      console.log('id_token acquired at: ' + new Date().toString());
  
      msalClient.setActiveAccount(authResult.account);
  
      // Initialize the Graph client
      initializeGraphClient(msalClient, authResult.account, msalRequest.scopes);
  
      // Get the user's profile from Graph
      const user = await getUser();
      // Save the profile in session
      sessionStorage.setItem('graphUser', JSON.stringify(user));
      updatePage(Views.home);
    } catch (error) {
      console.log(error);
      updatePage(Views.error, {
        message: 'Error logging in',
        debug: error
      });
    }
  }
  // </signInSnippet>
  
  // <signOutSnippet>
function signOut() {
    sessionStorage.removeItem('graphUser');
    msalClient.logout();
  }
  // </signOutSnippet>