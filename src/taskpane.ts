import "./styles.css";
import { writeJsonArrayToWorksheet, writeTestCell } from "./excel";
import { promptForCustomReportBridgeUrl } from "./leanix/bridge-url-prompt";
import { BrowserOAuthPlaceholderAuthProvider } from "./leanix/auth";
import { LeanIXCustomReportBridge } from "./leanix/report-bridge";
import { loadRuntimeConfig, saveCustomReportBridgeUrl } from "./leanix/runtime-config";
import { LeanIXStorageClient } from "./leanix/storage-client";

const authProvider = new BrowserOAuthPlaceholderAuthProvider();
let runtimeConfig = loadRuntimeConfig();
const storageClient = new LeanIXStorageClient(runtimeConfig.baseUrl, () =>
  authProvider.getAccessToken()
);
const reportBridge = new LeanIXCustomReportBridge(runtimeConfig.customReportBridgeUrl);

let retrievedJson: unknown;
let connectedWorkspace = "";
let isConnected = false;
let officeReady = false;

const elements = {
  authStatus: byId("auth-status"),
  workspaceStatus: byId("workspace-status"),
  signInButton: byId<HTMLButtonElement>("sign-in-button"),
  signOutButton: byId<HTMLButtonElement>("sign-out-button"),
  retrieveButton: byId<HTMLButtonElement>("retrieve-button"),
  writeButton: byId<HTMLButtonElement>("write-button"),
  testWriteButton: byId<HTMLButtonElement>("test-write-button"),
  objectIdInput: byId<HTMLInputElement>("object-id"),
  resultOutput: byId("result-output"),
  message: byId("message")
};

Office.onReady((info) => {
  officeReady = info.host === Office.HostType.Excel;
  elements.testWriteButton.disabled = !officeReady;
  elements.retrieveButton.disabled = false;
  elements.objectIdInput.value = runtimeConfig.storageObjectId;
  updateAuthUi();
  setMessage(
    officeReady
      ? "Excel is ready."
      : "Open this task pane in Excel to enable worksheet actions."
  );
});

elements.signInButton.addEventListener("click", async () => {
  await runAction(async () => {
    setMessage("Requesting LeanIX custom report bridge URL...");
    const customReportBridgeUrl = await promptForCustomReportBridgeUrl(
      runtimeConfig.customReportBridgeUrl
    );
    runtimeConfig = saveCustomReportBridgeUrl(customReportBridgeUrl);
    storageClient.setBaseUrl(runtimeConfig.baseUrl);
    reportBridge.setBridgeUrl(runtimeConfig.customReportBridgeUrl);

    setMessage("Opening LeanIX custom report bridge...");
    const connection = await reportBridge.connect();
    connectedWorkspace = connection.workspace ?? runtimeConfig.workspace;
    isConnected = true;
    await updateAuthUi();
    setMessage("Connected through the LeanIX custom report bridge.");
  });
});

elements.signOutButton.addEventListener("click", async () => {
  await runAction(async () => {
    await authProvider.logout();
    retrievedJson = undefined;
    connectedWorkspace = "";
    isConnected = false;
    elements.writeButton.disabled = true;
    await updateAuthUi();
    setMessage("Signed out.");
  });
});

elements.retrieveButton.addEventListener("click", async () => {
  await runAction(async () => {
    setMessage("Retrieving LeanIX storage object...");
    retrievedJson = runtimeConfig.customReportBridgeUrl
      ? await reportBridge.getObject(elements.objectIdInput.value.trim())
      : await storageClient.getObject(elements.objectIdInput.value.trim());
    elements.resultOutput.textContent = JSON.stringify(retrievedJson, null, 2);
    elements.writeButton.disabled = !officeReady;
    setMessage("Object retrieved.");
  });
});

elements.writeButton.addEventListener("click", async () => {
  await runAction(async () => {
    await writeJsonArrayToWorksheet(retrievedJson);
    setMessage("JSON data was written to an Excel table.");
  });
});

elements.testWriteButton.addEventListener("click", async () => {
  await runAction(async () => {
    await writeTestCell();
    setMessage("Wrote LeanIX Add-in Connected to A1.");
  });
});

async function updateAuthUi(): Promise<void> {
  const authenticated = isConnected || (await authProvider.isAuthenticated());
  elements.authStatus.textContent = authenticated ? "Connected to LeanIX" : "Not signed in";
  elements.workspaceStatus.textContent =
    authenticated && (connectedWorkspace || runtimeConfig.workspace)
      ? `Workspace: ${connectedWorkspace || runtimeConfig.workspace}`
      : runtimeConfig.baseUrl
        ? `Base URL: ${runtimeConfig.baseUrl}`
        : "Not connected";
  elements.signInButton.hidden = authenticated;
  elements.signOutButton.hidden = !authenticated;
}

async function runAction(action: () => Promise<void>): Promise<void> {
  try {
    await action();
  } catch (error) {
    setMessage(error instanceof Error ? error.message : "Something went wrong.");
  }
}

function setMessage(message: string): void {
  elements.message.textContent = message;
}

function byId<T extends HTMLElement = HTMLElement>(id: string): T {
  const element = document.getElementById(id);
  if (!element) {
    throw new Error(`Missing element: ${id}`);
  }

  return element as T;
}
