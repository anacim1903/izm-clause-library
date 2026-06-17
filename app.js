'use strict';

const CLIENT_ID = '44d73efb-0980-447f-8629-04c970ac0e74';
const TENANT_ID = 'd322a5f5-4eaf-439f-a75a-d4c3822080db';
const SCOPES    = ['Files.ReadWrite'];

const FILE_URL  =
  'https://graph.microsoft.com/v1.0' +
  '/sites/izumedeirosadvogados.sharepoint.com:/sites/IzuMedeiros:' +
  '/drive/root:/AI/260616%20-%20IZM%20Clauses.xlsx:/content';

let msalInstance;
let allClauses = [];

Office.onReady(() => {
  msalInstance = new msal.PublicClientApplication({
    auth: {
      clientId:    CLIENT_ID,
      authority:   `https://login.microsoftonline.com/${TENANT_ID}`,
      redirectUri: window.location.href.split('?')[0].split('#')[0],
    },
    cache: { cacheLocation: 'sessionStorage', storeAuthStateInCookie: false },
  });
  msalInstance.handleRedirectPromise().then(init).catch(showError);
});

async function init() {
  const accounts = msalInstance.getAllAccounts();
  if (accounts.length > 0) {
    await loadClauses();
  } else {
    showStatus('Signing in…');
    try {
      await msalInstance.loginPopup({ scopes: SCOPES });
      await loadClauses();
    } catch (e) {
      showError(e);
    }
  }
}

async function getToken() {
  const account = msalInstance.getAllAccounts()[0];
  try {
    const r = await msalInstance.acquireTokenSilent({ scopes: SCOPES, account });
    return r.accessToken;
  } catch {
    const r = await msalInstance.acquireTokenPopup({ scopes: SCOPES, account });
    return r.accessToken;
  }
}

async function loadClauses() {
  showStatus('Loading clauses…');
  try {
    const token   = await getToken();
    const headers = { Authorization: `Bearer ${token}` };

    const fileRes = await fetch(FILE_URL,
