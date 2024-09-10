import * as msal from '@azure/msal-browser';
import { CLIENT_ID } from '../../env.js';


//TODO: Add client ID from Azure Key Vault???

/**
 * Configuration object for MSAL (Microsoft Authentication Library).
 * 
 * @property {Object} auth - Authentication parameters.
 * @property {string} auth.clientId - The client ID of the application registered in Azure AD.
 * @property {string} auth.authority - The authority URL, typically the Azure AD tenant or 'common' for multi-tenant applications.
 * @property {string} auth.redirectUri - The redirect URI where the authentication response is sent back to the application.
 */
const msalConfig = {
  auth: {
    clientId: CLIENT_ID, //TODO: 
    authority: 'https://login.microsoftonline.com/common', // Tenant ID or 'common'
    redirectUri: 'https://localhost:3000', // Redirect URI specified in Azure AD
  },
};

/**
 * An instance of the PublicClientApplication class from the MSAL (Microsoft Authentication Library).
 * This instance is configured with the provided msalConfig object and is used to handle authentication
 * and acquire tokens for Microsoft services.
 *
 * @type {msal.PublicClientApplication}
 */
const msalInstance = new msal.PublicClientApplication(msalConfig);

/**
 * Retrieves an access token for the authenticated user.
 * 
 * This function first attempts to acquire the token silently. If that fails,
 * it falls back to an interactive token acquisition via a popup.
 * 
 * @async
 * @function getAccessToken
 * @returns {Promise<string>} The access token for the authenticated user.
 * @throws Will throw an error if both silent and interactive token acquisition fail.
 */
async function getAccessToken() {
  const account = msalInstance.getAllAccounts()[0];
  console.log('account: ', account);
  try {
    const tokenResponse = await msalInstance.acquireTokenSilent({
      account: account,
      scopes: ['Files.ReadWrite.All', 'Sites.ReadWrite.All'],
    });
    return tokenResponse.accessToken;
  } catch (error) {
    // Fallback to interactive token acquisition
    const tokenResponse = await msalInstance.acquireTokenPopup({
      account: account,
      scopes: ['Files.ReadWrite.All', 'Sites.ReadWrite.All'],
    });
    return tokenResponse.accessToken;
  }
}

/**
 * Asynchronously handles the sign-in process using MSAL (Microsoft Authentication Library).
 * 
 * This function initializes the MSAL instance and attempts to sign in the user via a popup.
 * It requests permissions to read and write files and sites.
 * 
 * @async
 * @function signIn
 * @returns {Promise<void>} - A promise that resolves when the sign-in process is complete.
 * @throws {Error} - Throws an error if the sign-in process fails.
 */
async function signIn() {
  try {
    await msalInstance.initialize();
    const loginResponse = await msalInstance.loginPopup({
      scopes: ['Files.ReadWrite.All', 'Sites.ReadWrite.All'],
    });
    console.log('id_token acquired: ', loginResponse);
  } catch (error) {
    console.error(error);
  }
}


export { getAccessToken, signIn };