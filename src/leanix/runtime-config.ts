import { leanixConfig } from "./config";

const STORAGE_KEY = "leanix-storage-explorer.runtimeConfig";

export interface LeanIXRuntimeConfig {
  baseUrl: string;
  workspace: string;
  storageObjectId: string;
  customReportBridgeUrl: string;
}

export function loadRuntimeConfig(): LeanIXRuntimeConfig {
  const savedConfig = readSavedConfig();
  const customReportBridgeUrl =
    savedConfig.customReportBridgeUrl || leanixConfig.customReportBridgeUrl;

  return {
    baseUrl:
      savedConfig.baseUrl ||
      deriveBaseUrlFromBridgeUrl(customReportBridgeUrl) ||
      leanixConfig.baseUrl,
    workspace: savedConfig.workspace || leanixConfig.workspace,
    storageObjectId: savedConfig.storageObjectId || leanixConfig.storageObjectId,
    customReportBridgeUrl
  };
}

export function saveCustomReportBridgeUrl(value: string): LeanIXRuntimeConfig {
  const customReportBridgeUrl = normalizeCustomReportBridgeUrl(value);
  const currentConfig = loadRuntimeConfig();
  const nextConfig = {
    ...currentConfig,
    baseUrl: deriveBaseUrlFromBridgeUrl(customReportBridgeUrl),
    customReportBridgeUrl
  };

  localStorage.setItem(STORAGE_KEY, JSON.stringify(nextConfig));
  return nextConfig;
}

export function normalizeCustomReportBridgeUrl(value: string): string {
  const trimmedValue = value.trim();
  if (!trimmedValue) {
    throw new Error("Enter the LeanIX custom report bridge URL.");
  }

  let url: URL;
  try {
    url = new URL(trimmedValue);
  } catch {
    throw new Error("Enter a valid LeanIX custom report bridge URL.");
  }

  if (url.protocol !== "https:" && url.protocol !== "http:") {
    throw new Error("The LeanIX custom report bridge URL must start with http or https.");
  }

  return url.toString();
}

export function deriveBaseUrlFromBridgeUrl(value: string): string {
  if (!value) {
    return "";
  }

  return new URL(normalizeCustomReportBridgeUrl(value)).origin;
}

function readSavedConfig(): Partial<LeanIXRuntimeConfig> {
  const rawValue = localStorage.getItem(STORAGE_KEY);
  if (!rawValue) {
    return {};
  }

  try {
    return JSON.parse(rawValue) as Partial<LeanIXRuntimeConfig>;
  } catch {
    return {};
  }
}
