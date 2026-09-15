// ============================================================
//  Raimak LMS — Authentication (MSAL)
// ============================================================
const Auth = (() => {
  let msalInstance = null;
  let currentAccount = null;

  // ── Init MSAL ──────────────────────────────────────────────
  async function init() {
    msalInstance = new msal.PublicClientApplication({
      auth: {
        clientId: Config.azure.clientId,
        authority: `https://login.microsoftonline.com/${Config.azure.tenantId}`,
        redirectUri: Config.azure.redirectUri,
      },
      cache: {
        cacheLocation: "localStorage",
        storeAuthStateInCookie: true,
      },
    });

    const result = await msalInstance.handleRedirectPromise();

    // Set active account from redirect result for new users
    if (result && result.account) {
      currentAccount = result.account;
      msalInstance.setActiveAccount(result.account);
    } else {
      const accounts = msalInstance.getAllAccounts();
      if (accounts.length) {
        currentAccount = accounts[0];
        msalInstance.setActiveAccount(accounts[0]);
      }
    }

    return result;
  }

  // ── Sign In ────────────────────────────────────────────────
  async function signIn() {
    try {
      await msalInstance.loginRedirect({ scopes: Config.scopes });
    } catch (err) {
      console.error("Sign-in error:", err);
      UI.showToast("Sign-in failed. Please try again.", "error");
    }
  }

  // ── Sign Out ───────────────────────────────────────────────
  function signOut() {
    msalInstance.logoutRedirect({
      postLogoutRedirectUri: Config.azure.redirectUri,
    });
  }

  // ── Get Access Token ───────────────────────────────────────
  async function getToken() {
    const account =
      msalInstance.getActiveAccount() || msalInstance.getAllAccounts()[0];
    if (!account) return null;

    currentAccount = account;

    try {
      const result = await msalInstance.acquireTokenSilent({
        scopes: Config.scopes,
        account: currentAccount,
      });

      // Token came back but is empty — force re-consent
      if (!result.accessToken || result.accessToken.trim() === "") {
        await msalInstance.acquireTokenRedirect({
          scopes: Config.scopes,
          prompt: "consent", // forces the permissions consent screen
        });
        return null;
      }

      return result.accessToken;
    } catch (err) {
      if (err instanceof msal.InteractionRequiredAuthError) {
        await msalInstance.acquireTokenRedirect({
          scopes: Config.scopes,
          prompt: "consent", // forces the permissions consent screen
        });
        return null;
      }
      throw err;
    }
  }

  // ── Current User ───────────────────────────────────────────
  function getUser() {
    const account =
      msalInstance?.getActiveAccount() || msalInstance?.getAllAccounts()?.[0];
    if (!account) return null;
    currentAccount = account;
    return {
      name: account.name || account.username,
      email: account.username,
    };
  }

  function isSignedIn() {
    return !!(
      msalInstance?.getActiveAccount() || msalInstance?.getAllAccounts()?.length
    );
  }

  return { init, signIn, signOut, getToken, getUser, isSignedIn };
})();
