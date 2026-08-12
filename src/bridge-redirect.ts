import "./styles.css";

const message = document.getElementById("message");
const target = new URLSearchParams(window.location.search).get("target");

try {
  if (!target) {
    throw new Error("Missing LeanIX redirect target.");
  }

  const targetUrl = new URL(target);
  if (targetUrl.protocol !== "https:" && targetUrl.protocol !== "http:") {
    throw new Error("LeanIX redirect target must start with http or https.");
  }

  window.location.replace(targetUrl.toString());
} catch (error) {
  if (message) {
    message.textContent =
      error instanceof Error ? error.message : "Unable to open LeanIX.";
  }
}
