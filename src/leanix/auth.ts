import type { LeanIXAuthProvider } from "./types";

const unsupportedAuthMessage =
  "SAP LeanIX documentation found for this MVP documents API access through technical-user client credentials, not an official browser/Office add-in Authorization Code + PKCE flow. Add an officially registered browser OAuth flow here when SAP confirms one.";

export class BrowserOAuthPlaceholderAuthProvider implements LeanIXAuthProvider {
  async login(): Promise<void> {
    throw new Error(unsupportedAuthMessage);
  }

  async logout(): Promise<void> {
    return Promise.resolve();
  }

  async getAccessToken(): Promise<string> {
    throw new Error(unsupportedAuthMessage);
  }

  async isAuthenticated(): Promise<boolean> {
    return false;
  }
}
