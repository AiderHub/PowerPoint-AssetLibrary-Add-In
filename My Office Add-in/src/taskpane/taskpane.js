/**
 * Fetches SharePoint files using the access token.
 * 
 * @async
 * @function getSharePointFiles
 * @returns {Promise<void>} A promise that resolves when the SharePoint files are fetched and logged to the console.
 */

/**
 * Signs in the user and fetches SharePoint files.
 * 
 * @async
 * @function run
 * @returns {Promise<void>} A promise that resolves when the user is signed in and SharePoint files are fetched.
 */

/**
 * Office onReady event handler.
 * 
 * @async
 * @function Office.onReady
 * @param {Object} info - Information about the host application.
 * @param {string} info.host - The host application (e.g., PowerPoint).
 * @returns {Promise<void>} A promise that resolves when the user is signed in and SharePoint files are fetched if the host is PowerPoint.
 */
/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office */


import { getAccessToken, signIn } from '../auth/auth';

async function getSharePointFiles() {
  const accessToken = await getAccessToken();
  const sharepointSite = '<your-sharepoint-site>'; //TODO: Set url to Asset Library
  
  const response = await fetch(`${sharepointSite}/_api/web/lists`, {
    method: 'GET',
    headers: {
      Authorization: `Bearer ${accessToken}`,
      Accept: 'application/json;odata=verbose',
    },
  });
  
  const data = await response.json();
  console.log('SharePoint Lists:', data);
}

//TODO: Remove this after new Office.onReady is implemented
async function run() {
  await signIn();
  await getSharePointFiles();
}


Office.onReady((info) => {
  if (info.host === Office.HostType.PowerPoint) {
    // document.getElementById("sideload-msg").style.display = "none";
    // document.getElementById("app-body").style.display = "flex";
    // document.getElementById("run").onclick = run;
  }
});

// TODO: Implemet this when App is apporved by IT
// Office.onReady(async (info) => {
//   if (info.host === Office.HostType.PowerPoint) {
//     await signIn();
//     await getSharePointFiles();
//   }
// });
