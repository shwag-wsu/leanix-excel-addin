import "./styles.css";
import { normalizeCustomReportBridgeUrl } from "./leanix/runtime-config";

const form = byId<HTMLFormElement>("bridge-url-form");
const input = byId<HTMLInputElement>("custom-report-bridge-url");
const cancelButton = byId<HTMLButtonElement>("cancel-button");
const message = byId("message");

Office.onReady(() => {
  input.value = new URLSearchParams(window.location.search).get("customReportBridgeUrl") ?? "";
  input.focus();
});

form.addEventListener("submit", (event) => {
  event.preventDefault();

  try {
    const customReportBridgeUrl = normalizeCustomReportBridgeUrl(input.value);
    Office.context.ui.messageParent(
      JSON.stringify({
        source: "leanix-storage-explorer-bridge-url-prompt",
        type: "submit",
        customReportBridgeUrl
      })
    );
  } catch (error) {
    message.textContent = error instanceof Error ? error.message : "Enter a valid URL.";
  }
});

cancelButton.addEventListener("click", () => {
  Office.context.ui.messageParent(
    JSON.stringify({
      source: "leanix-storage-explorer-bridge-url-prompt",
      type: "cancel"
    })
  );
});

function byId<T extends HTMLElement = HTMLElement>(id: string): T {
  const element = document.getElementById(id);
  if (!element) {
    throw new Error(`Missing element: ${id}`);
  }

  return element as T;
}
