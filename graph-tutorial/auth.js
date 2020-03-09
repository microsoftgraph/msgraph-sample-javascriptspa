// Create the main MSAL instance
// configuration parameters are located in config.js
const msalClient = new Msal.UserAgentApplication(msalConfig);

if (msalClient.getAccount() && !msalClient.isCallback(window.location.hash)) {
  // avoid duplicate code execution on page load in case of iframe and Popup window.
  updatePage(msalClient.getAccount(), Views.home);
}

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

function signOut() {
  msalClient.logout();
}

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
