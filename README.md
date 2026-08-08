# LeanIX Storage Explorer

LeanIX Storage Explorer is a very small proof-of-concept Microsoft Excel task-pane add-in. It is a static Office.js app built with vanilla TypeScript and Vite, intended to be hosted from GitHub Pages.

## Current MVP Status

Implemented:

- Excel task-pane shell with Office.js initialization.
- A test button that writes `LeanIX Add-in Connected` to cell `A1`.
- A LeanIX authentication abstraction in `src/leanix/auth.ts`.
- A placeholder browser OAuth provider that intentionally does not invent endpoints, client IDs, or secrets.
- A LeanIX Custom Report bridge option that opens a LeanIX-hosted report in an Office dialog and receives JSON back.
- A Storage API client that sends `Authorization: Bearer <access token>` once a supported auth provider is plugged in.
- JSON display in the task pane.
- Flat JSON array export to an Excel table named `LeanIXStorageData`.
- GitHub Pages deployment workflow.

Blocked pending official SAP confirmation or tenant-specific OpenAPI details:

- Browser/Office add-in OAuth implementation.
- Exact Storage API object resource path.

## Prerequisites

- Node.js
- npm
- Microsoft Excel
- SAP LeanIX workspace access
- GitHub account

## Local Setup

```bash
git clone https://github.com/shwag-wsu/leanix-excel-addin.git
cd leanix-excel-addin
npm install
npm run dev
```

Vite serves the task pane locally. The production build assumes GitHub Pages will serve the app from:

```text
https://shwag-wsu.github.io/leanix-excel-addin/
```

If your repository name changes, update `base` in `vite.config.ts` and the URLs in `manifest.xml`.

The built task pane is served from:

```text
https://shwag-wsu.github.io/leanix-excel-addin/taskpane.html
```

## LeanIX Login

Official SAP documentation for SAP LeanIX custom report tooling says the custom report setup opens a browser OAuth login flow, stores credentials in a user-level configuration file, and supports re-authentication with:

```bash
npx lxr login
```

SAP also warns that the `lxr.json` configuration stores sensitive OAuth tokens and must not be committed.

For SAP LeanIX APIs generally, the official authentication documentation describes creating a Technical User, exchanging its API token for an OAuth access token with the `client_credentials` grant at:

```text
https://{SUBDOMAIN}.leanix.net/services/mtm/v1/oauth2/token
```

That technical-user flow is not appropriate for this GitHub Pages browser add-in because it would require putting a secret/API token in client JavaScript.

As of the documentation checked for this MVP, I found official support for user-based OAuth in SAP LeanIX custom report tools and MCP tooling, but not a documented, reusable Authorization Code + PKCE registration flow for arbitrary standalone browser or Office add-in clients. Therefore this project includes `LeanIXAuthProvider` and `BrowserOAuthPlaceholderAuthProvider`, but does not implement OAuth endpoints or copy `lxr` credentials.

## LeanIX Custom Report Bridge

The practical workaround is to host a small bridge page as a SAP LeanIX Custom Report. The Excel task pane stays on GitHub Pages, but it opens that LeanIX-hosted report with `Office.context.ui.displayDialogAsync`. The custom report runs inside LeanIX's supported user-authenticated context, retrieves the configured Storage API object, and sends only the resulting JSON back to Excel.

Recommended flow:

```text
Excel task pane
  -> opens LeanIX custom report bridge
  -> LeanIX-hosted report uses LeanIX-supported user context
  -> report retrieves one Storage API object
  -> report sends JSON back to Excel
  -> Excel displays JSON and writes a flat array to a worksheet table
```

Configure the bridge URL in `src/leanix/config.ts`:

```typescript
export const leanixConfig = {
  baseUrl: "https://YOUR-SUBDOMAIN.leanix.net",
  workspace: "YOUR-WORKSPACE",
  storageObjectId: "/services/...",
  customReportBridgeUrl: "https://YOUR-SUBDOMAIN.leanix.net/path/to/custom/report"
};
```

The Excel task pane calls the bridge with query parameters:

```text
?source=leanix-storage-explorer
&action=connect
&returnOrigin=https://shwag-wsu.github.io
```

or:

```text
?source=leanix-storage-explorer
&action=getObject
&objectId=/services/...
&returnOrigin=https://shwag-wsu.github.io
```

The LeanIX custom report should include Office.js when opened as an Office dialog, then call `Office.context.ui.messageParent(...)`.

For connection confirmation:

```typescript
Office.context.ui.messageParent(JSON.stringify({
  source: "leanix-storage-explorer-bridge",
  type: "connection",
  workspace: "YOUR-WORKSPACE"
}));
```

For a retrieved Storage API object:

```typescript
Office.context.ui.messageParent(JSON.stringify({
  source: "leanix-storage-explorer-bridge",
  type: "storageObject",
  payload: [
    {
      application: "App A",
      platform: "Azure",
      activeFrom: "2025-01-01"
    }
  ]
}));
```

For errors:

```typescript
Office.context.ui.messageParent(JSON.stringify({
  source: "leanix-storage-explorer-bridge",
  type: "error",
  message: "Unable to retrieve the storage object."
}));
```

Do not send LeanIX access tokens back to Excel unless SAP explicitly documents that as supported. The bridge should send retrieved data, not credentials.

Sources:

- SAP Help Portal, Set Up Your Custom Reports Project: https://help.sap.com/docs/leanix/ea/setting-up-your-custom-reports-project
- SAP Help Portal, Set Up Your Environment for Custom Reports with AI: https://help.sap.com/docs/leanix/ea/set-up-your-environment-for-custom-reports-with-ai
- SAP Help Portal, Authentication to SAP LeanIX Services: https://help.sap.com/docs/leanix/ea/authentication-to-sap-leanix-services
- SAP Help Portal, REST APIs Overview: https://help.sap.com/docs/leanix/ea/rest-apis-overview

## LeanIX Storage API

SAP's REST APIs overview lists a Storage API for managing and retrieving file resources for a workspace. The same REST API documentation directs users to the OpenAPI Explorer inside their SAP LeanIX workspace for exact endpoint reference.

Because the exact object retrieval endpoint is tenant/API-reference specific, this MVP does not guess a Storage API path. Configure the resource path from your workspace's OpenAPI Explorer in `src/leanix/config.ts`.

## Configuration

Edit non-secret settings in `src/leanix/config.ts`:

```typescript
export const leanixConfig = {
  baseUrl: "https://YOUR-SUBDOMAIN.leanix.net",
  workspace: "YOUR-WORKSPACE",
  storageObjectId: "/services/...",
  customReportBridgeUrl: "https://YOUR-SUBDOMAIN.leanix.net/path/to/custom/report"
};
```

Do not put tokens, API tokens, refresh tokens, client secrets, or anything from `lxr.json` in this file.

## GitHub Pages Deployment

1. Create a GitHub repository, for example `leanix-excel-addin`.
2. Confirm the GitHub Pages URLs in `manifest.xml` match your GitHub account and repository.
3. Confirm `vite.config.ts` uses `base: "/leanix-excel-addin/"`.
4. Push code to `main`.
5. Go to Settings -> Pages.
6. Select GitHub Actions as the deployment source if necessary.
7. Push to `main` and wait for the `Deploy GitHub Pages` workflow.
8. Confirm the generated HTTPS URL.

## Excel Sideloading

For Excel on the web:

1. Open Excel in Microsoft 365.
2. Go to Add-ins.
3. Choose More Add-ins or Upload My Add-in.
4. Upload `manifest.xml`.
5. Open the LeanIX tab command and launch the task pane.

For Excel Desktop, use Microsoft's sideloading flow for your platform and tenant policy. The key requirement is that `manifest.xml` points to the HTTPS GitHub Pages task pane URL.

## Testing

Phase 1 smoke test:

1. Sideload `manifest.xml`.
2. Open the task pane in Excel.
3. Click `Write Test Cell`.
4. Confirm cell `A1` contains `LeanIX Add-in Connected`.

After a supported LeanIX browser auth provider is implemented:

1. Configure `src/leanix/config.ts`.
2. If using the custom report bridge, set `customReportBridgeUrl`.
3. Click `Sign in to LeanIX`.
4. Retrieve one Storage API object.
5. Confirm the JSON appears in the task pane.
6. If the JSON is a flat array of objects, click `Write to Worksheet`.
7. Confirm an Excel table appears.

With the custom report bridge:

1. Deploy the LeanIX custom report bridge in SAP LeanIX.
2. Set `customReportBridgeUrl`.
3. Rebuild and deploy the GitHub Pages task pane.
4. Sideload `manifest.xml`.
5. Click `Sign in to LeanIX`.
6. Confirm the LeanIX-hosted report returns a connection message.
7. Retrieve one Storage API object through the bridge.
8. Confirm the JSON appears in the task pane.
9. If the JSON is a flat array of objects, click `Write to Worksheet`.
10. Confirm an Excel table appears.

## Security

- No LeanIX secrets belong in GitHub Pages.
- No technical-user token belongs in client JavaScript.
- Access tokens should be held only in memory when practical.
- Never commit tokens.
- Never log tokens.
- Do not copy credentials from `~/.leanix/lxr.json` or `%APPDATA%\leanix\lxr.json`.
- Do not reverse engineer `lxr`.
