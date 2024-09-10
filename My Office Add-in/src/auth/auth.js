import * as msal from '@azure/msal-browser';
import { CLIENT_ID } from '../../env.js';


//TODO: Add client ID from Azure Key Vault???
const msalConfig = {
  auth: {
    clientId: CLIENT_ID, //TODO: 
    authority: 'https://login.microsoftonline.com/common', // Tenant ID or 'common'
    redirectUri: 'https://localhost:3000', // Redirect URI specified in Azure AD
  },
};

const msalInstance = new msal.PublicClientApplication(msalConfig);

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