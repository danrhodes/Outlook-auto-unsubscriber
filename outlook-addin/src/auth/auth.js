/**
 * Authentication module using MSAL.js for Microsoft 365
 */

// MSAL Configuration
// NOTE: You need to replace these values after creating your Azure AD app
const msalConfig = {
    auth: {
        clientId: "YOUR_CLIENT_ID_HERE", // Replace with your Azure AD App Client ID
        authority: "https://login.microsoftonline.com/common",
        redirectUri: window.location.origin + "/outlook-addin/src/taskpane/taskpane.html"
    },
    cache: {
        cacheLocation: "localStorage",
        storeAuthStateInCookie: true
    }
};

// Microsoft Graph API scopes
const loginRequest = {
    scopes: ["User.Read", "Mail.Read", "Mail.ReadBasic"]
};

const tokenRequest = {
    scopes: ["https://graph.microsoft.com/Mail.Read", "https://graph.microsoft.com/Mail.ReadBasic"]
};

// Initialize MSAL
let msalInstance;
let currentAccount = null;

function initializeMsal() {
    msalInstance = new msal.PublicClientApplication(msalConfig);

    // Handle redirect response
    return msalInstance.handleRedirectPromise()
        .then(response => {
            if (response) {
                currentAccount = response.account;
                console.log("Authentication successful:", currentAccount);
                return true;
            } else {
                // Check if there's an existing account
                const accounts = msalInstance.getAllAccounts();
                if (accounts.length > 0) {
                    currentAccount = accounts[0];
                    console.log("Found existing account:", currentAccount);
                    return true;
                }
            }
            return false;
        })
        .catch(error => {
            console.error("Authentication error:", error);
            throw error;
        });
}

// Sign in function
async function signIn() {
    try {
        const response = await msalInstance.loginPopup(loginRequest);
        currentAccount = response.account;
        console.log("Sign in successful:", currentAccount);
        return currentAccount;
    } catch (error) {
        console.error("Sign in error:", error);
        throw error;
    }
}

// Sign out function
function signOut() {
    const logoutRequest = {
        account: currentAccount
    };

    msalInstance.logoutPopup(logoutRequest)
        .then(() => {
            currentAccount = null;
            console.log("Sign out successful");
        })
        .catch(error => {
            console.error("Sign out error:", error);
        });
}

// Get access token for Microsoft Graph
async function getAccessToken() {
    if (!currentAccount) {
        throw new Error("No account signed in");
    }

    const request = {
        scopes: tokenRequest.scopes,
        account: currentAccount
    };

    try {
        // Try to acquire token silently
        const response = await msalInstance.acquireTokenSilent(request);
        return response.accessToken;
    } catch (error) {
        console.warn("Silent token acquisition failed, using popup:", error);

        // If silent acquisition fails, use popup
        try {
            const response = await msalInstance.acquireTokenPopup(request);
            return response.accessToken;
        } catch (popupError) {
            console.error("Token acquisition failed:", popupError);
            throw popupError;
        }
    }
}

// Check if user is authenticated
function isAuthenticated() {
    return currentAccount !== null;
}

// Get current user info
function getCurrentUser() {
    return currentAccount;
}

// Export functions
window.AuthModule = {
    initializeMsal,
    signIn,
    signOut,
    getAccessToken,
    isAuthenticated,
    getCurrentUser
};
