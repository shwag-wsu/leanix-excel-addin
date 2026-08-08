export interface LeanIXBridgeConnection {
  workspace?: string;
}

type BridgeRequest =
  | { action: "connect" }
  | { action: "getObject"; objectId: string };

type BridgeMessage =
  | {
      source: "leanix-storage-explorer-bridge";
      type: "connection";
      workspace?: string;
    }
  | {
      source: "leanix-storage-explorer-bridge";
      type: "storageObject";
      payload: unknown;
    }
  | {
      source: "leanix-storage-explorer-bridge";
      type: "error";
      message: string;
    };

export class LeanIXCustomReportBridge {
  constructor(private readonly bridgeUrl: string) {}

  async connect(): Promise<LeanIXBridgeConnection> {
    const message = await this.openBridge({ action: "connect" });
    if (message.type !== "connection") {
      throw new Error("The LeanIX custom report did not return a connection confirmation.");
    }

    return { workspace: message.workspace };
  }

  async getObject(objectId: string): Promise<unknown> {
    if (!objectId) {
      throw new Error("Enter a LeanIX Storage API object/resource identifier.");
    }

    const message = await this.openBridge({ action: "getObject", objectId });
    if (message.type !== "storageObject") {
      throw new Error("The LeanIX custom report did not return a storage object.");
    }

    return message.payload;
  }

  private async openBridge(request: BridgeRequest): Promise<BridgeMessage> {
    if (!this.bridgeUrl) {
      throw new Error("Configure leanixConfig.customReportBridgeUrl before signing in.");
    }

    const url = new URL(this.bridgeUrl);
    url.searchParams.set("source", "leanix-storage-explorer");
    url.searchParams.set("action", request.action);
    url.searchParams.set("returnOrigin", window.location.origin);

    if (request.action === "getObject") {
      url.searchParams.set("objectId", request.objectId);
    }

    return new Promise((resolve, reject) => {
      Office.context.ui.displayDialogAsync(
        url.toString(),
        { height: 60, width: 50, displayInIframe: false },
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
                  throw new Error(`LeanIX bridge dialog returned error ${args.error}.`);
                }

                const message = parseBridgeMessage(args.message);
                if (message.type === "error") {
                  settle(() => reject(new Error(message.message)));
                  return;
                }

                settle(() => resolve(message));
              } catch (error) {
                settle(() =>
                  reject(error instanceof Error ? error : new Error("Invalid bridge message."))
                );
              }
            }
          );

          dialog.addEventHandler(Office.EventType.DialogEventReceived, () => {
            settle(() => reject(new Error("LeanIX sign-in window was closed.")));
          });
        }
      );
    });
  }
}

function parseBridgeMessage(value: string): BridgeMessage {
  const message = JSON.parse(value) as Partial<BridgeMessage>;
  if (message.source !== "leanix-storage-explorer-bridge") {
    throw new Error("Received a message from an unexpected source.");
  }

  if (message.type === "connection" || message.type === "storageObject") {
    return message as BridgeMessage;
  }

  if (message.type === "error" && typeof message.message === "string") {
    return message as BridgeMessage;
  }

  throw new Error("Received an unsupported LeanIX bridge message.");
}
