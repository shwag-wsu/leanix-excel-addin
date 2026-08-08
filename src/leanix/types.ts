export interface LeanIXAuthProvider {
  login(): Promise<void>;
  logout(): Promise<void>;
  getAccessToken(): Promise<string>;
  isAuthenticated(): Promise<boolean>;
}

export type FlatJsonRow = Record<string, string | number | boolean | null>;
