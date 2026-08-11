interface BridgeUrlPromptMessage {
  source: "leanix-storage-explorer-bridge-url-prompt";
  type: "submit" | "cancel";
  customReportBridgeUrl?: string;
}

export async function promptForCustomReportBridgeUrl(currentValue: string): Promise<string> {
  const dialogUrl = new URL("bridge-url-prompt.html", window.location.href);
  if (currentValue) {
    dialogUrl.searchParams.set("customReportBridgeUrl", currentValue);
  }

  return new Promise((resolve, reject) => {
    Office.context.ui.displayDialogAsync(
      dialogUrl.toString(),
      { height: 42, width: 36, displayInIframe: true, promptBeforeOpen: false },
      (result) => {
        if (result.status !== Office.AsyncResultStatus.Succeeded) {
          reject(new Error(result.error.message));
          return;
        }

        const dialog = result.value;
        let settled = false;

        const settle = (callback: () => void) => {
          if (settled) {
            return;
          }

          settled = true;
          dialog.close();
          callback();
        };

        dialog.addEventHandler(
          Office.EventType.DialogMessageReceived,
          (args: { message: string; origin?: string } | { error: number }) => {
            try {
              if (!("message" in args)) {
                throw new Error(`LeanIX URL prompt returned error ${args.error}.`);
              }

              const message = parsePromptMessage(args.message);
              if (message.type === "cancel") {
                settle(() => reject(new Error("LeanIX sign-in was cancelled.")));
                return;
              }

              settle(() => resolve(message.customReportBridgeUrl ?? ""));
            } catch (error) {
              settle(() =>
                reject(error instanceof Error ? error : new Error("Invalid LeanIX URL prompt."))
              );
            }
          }
        );

        dialog.addEventHandler(Office.EventType.DialogEventReceived, () => {
          settle(() => reject(new Error("LeanIX sign-in was cancelled.")));
        });
      }
    );
  });
}

function parsePromptMessage(value: string): BridgeUrlPromptMessage {
  const message = JSON.parse(value) as Partial<BridgeUrlPromptMessage>;
  if (message.source !== "leanix-storage-explorer-bridge-url-prompt") {
    throw new Error("Received a message from an unexpected prompt.");
  }

  if (message.type === "cancel") {
    return { source: message.source, type: "cancel" };
  }

  if (message.type === "submit" && typeof message.customReportBridgeUrl === "string") {
    return {
      source: message.source,
      type: "submit",
      customReportBridgeUrl: message.customReportBridgeUrl
    };
  }

  throw new Error("Received an unsupported LeanIX URL prompt message.");
}
