export class LeanIXStorageClient {
  constructor(
    private baseUrl: string,
    private readonly getAccessToken: () => Promise<string>
  ) {}

  setBaseUrl(baseUrl: string): void {
    this.baseUrl = baseUrl;
  }

  async getObject(objectId: string): Promise<unknown> {
    const url = this.buildStorageUrl(objectId);
    const accessToken = await this.getAccessToken();

    let response: Response;
    try {
      response = await fetch(url, {
        headers: {
          Authorization: `Bearer ${accessToken}`,
          Accept: "application/json"
        }
      });
    } catch {
      throw new Error("Network error while contacting SAP LeanIX.");
    }

    if (response.status === 401) {
      throw new Error("LeanIX rejected the request. Please sign in again.");
    }

    if (response.status === 403) {
      throw new Error("Your LeanIX user does not have permission to read this object.");
    }

    if (response.status === 404) {
      throw new Error("The configured LeanIX storage object was not found.");
    }

    if (!response.ok) {
      throw new Error(`LeanIX returned HTTP ${response.status}.`);
    }

    try {
      return await response.json();
    } catch {
      throw new Error("LeanIX returned a response that is not valid JSON.");
    }
  }

  private buildStorageUrl(objectId: string): string {
    if (!this.baseUrl) {
      throw new Error("Configure the LeanIX base URL before retrieving storage objects.");
    }

    if (!objectId) {
      throw new Error("Enter a LeanIX Storage API object/resource path.");
    }

    if (/^https?:\/\//i.test(objectId)) {
      throw new Error("Use a Storage API path from your LeanIX OpenAPI Explorer, not a full URL.");
    }

    if (!objectId.startsWith("/services/")) {
      throw new Error("Use the official Storage API resource path from your LeanIX OpenAPI Explorer.");
    }

    return new URL(objectId, this.baseUrl).toString();
  }
}
