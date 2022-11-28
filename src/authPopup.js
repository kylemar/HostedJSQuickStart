// Create the main myMSALObj instance
// configuration parameters are located at authConfig.js
const myMSALObj = new msal.PublicClientApplication(msalConfig);

let username = "";

function selectAccount() {

    /**
     * See here for more info on account retrieval: 
     * https://github.com/AzureAD/microsoft-authentication-library-for-js/blob/dev/lib/msal-common/docs/Accounts.md
     */

    const currentAccounts = myMSALObj.getAllAccounts();
    if (currentAccounts.length === 0) {
        return;
    } else if (currentAccounts.length > 1) {
        // Add choose account code here
        console.warn("Multiple accounts detected.");
    } else if (currentAccounts.length === 1) {
        username = currentAccounts[0].username;
        setSignedIn();
        updateIdTokenTab(currentAccounts[0].idTokenClaims, null);
    }
}

function handleResponse(response) {

    /**
     * To see the full list of response object properties, visit:
     * https://github.com/AzureAD/microsoft-authentication-library-for-js/blob/dev/lib/msal-browser/docs/request-response-object.md#response
     */

    if (response !== null) {
        username = response.account.username;
        setSignedIn();
        updateIdTokenTab(response.idTokenClaims, response.idToken);
        updateTokenResponseTab(response);
    } else {
        selectAccount();
    }
}

function signIn() {

    /**
     * You can pass a custom request object below. This will override the initial configuration. For more information, visit:
     * https://github.com/AzureAD/microsoft-authentication-library-for-js/blob/dev/lib/msal-browser/docs/request-response-object.md#request
     */

     myMSALObj.loginPopup(loginRequest)
     .then(handleResponse)
     .catch(error => {
         console.error(error);
     });
}

function signOut() {

    /**
     * You can pass a custom request object below. This will override the initial configuration. For more information, visit:
     * https://github.com/AzureAD/microsoft-authentication-library-for-js/blob/dev/lib/msal-browser/docs/request-response-object.md#request
     */

    const logoutRequest = {
        account: myMSALObj.getAccountByUsername(username),
        mainWindowRedirectUri: window.location.href
    };

    myMSALObj.logoutPopup(logoutRequest);
    setSignedOut();
}

function getTokenPopup(request) {

    /**
     * See here for more info on account retrieval: 
     * https://github.com/AzureAD/microsoft-authentication-library-for-js/blob/dev/lib/msal-common/docs/Accounts.md
     */
    request.account = myMSALObj.getAccountByUsername(username);
    
    return myMSALObj.acquireTokenSilent(request)
        .catch(error => {
            console.warn("silent token acquisition fails. acquiring token using popup");
            if (error instanceof msal.InteractionRequiredAuthError) {
                // fallback to interaction when silent call fails
                return myMSALObj.acquireTokenPopup(request)
                    .then(tokenResponse => {
                        updateTokenResponseTab(response);
                        console.log(tokenResponse);
                        return tokenResponse;
                    }).catch(error => {
                        console.error(error);
                    });
            } else {
                console.warn(error);   
            }
    });
}

function getProfile() {
    const tokenRequest = {
        scopes: ["User.Read"],
        forceRefresh: false // Set this to "true" to skip a cached token and go to the server to get a new token
    };

    getTokenPopup(tokenRequest)
        .then(response => {
            callMSGraph("https://graph.microsoft.com/v1.0/me", response.accessToken, updateResultsTab);
        }).catch(error => {
            console.error(error);
        });
}
 
function getPeople() {
    const tokenRequest = {
        scopes: ["People.Read"],
        forceRefresh: false // Set this to "true" to skip a cached token and go to the server to get a new token
    };

    getTokenPopup(tokenRequest)
        .then(response => {
            callMSGraph("https://graph.microsoft.com/v1.0/me/people", response.accessToken, updatePeopleResultsTab);
        }).catch(error => {
            console.error(error);
        });
}

setAuthorityText();
selectAccount();
