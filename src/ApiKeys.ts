
/**Retrieves the API key for OpenAI services.  Throws an error if no key is available. */
export async function getOpenAiApiKey() {
    const openaiApiKey = await getApiKey("OPENAI_API_KEY");
    if (openaiApiKey === null) {
        console.error("No openai api key available.");
        throw new Error("No openai api key available.");
    }
    return openaiApiKey;
}

async function getApiKey(key: string): Promise<string | null> {
    const keys = JSON.parse(localStorage.getItem('api-keys') ?? '{}');

    if (Object.hasOwn(keys, key)) {
        return keys[key];
    } else {
        await showImportApiKeysDialog();
        const keys = JSON.parse(localStorage.getItem('api-keys') ?? '{}');
        if (Object.hasOwn(keys, key)) {
            return keys[key];
        } else {
            return null;
        }
    }
}

function showImportApiKeysDialog() {
    return new Promise<void>((res, rej) => {
        try {
            Office.context.ui.displayDialogAsync(
                `${location.origin}/build/dialogs/ApiKeys/index.html`,
                {
                    displayInIframe: true,
                    width: 50,
                    height: 50,
                },
                function (asyncResult) {
                    const dialog = asyncResult.value;
                    dialog.addEventHandler(Office.EventType.DialogMessageReceived, (args: any) => {
                        console.log('MESSAGE RECIEVED', args);
                        if (args.message) {
                            const message = JSON.parse(args.message)
                            switch (message.type) {
                                case "dialogClosed":
                                    dialog.close();
                                    res();
                                    break;

                                case "apiKeys":
                                    addApiKeys(JSON.parse(message.content));
                                    break;

                                default:
                                    break;
                            }
                        }
                    });

                    dialog.addEventHandler(Office.EventType.DialogEventReceived, (arg: any) => {
                        if (arg.error) {
                            switch (arg.error) {
                                case 12006:
                                    res();
                                    break;
                            }
                        }
                    });
                });
        } catch (err) {
            rej(err);
        }
    });
}

function addApiKeys(keys: Record<string, string>) {
    const existingKeys = JSON.parse(localStorage.getItem('api-keys') ?? '{}');

    for (const [key, value] of Object.entries(keys)) {
        existingKeys[key] = value;
    }

    localStorage.setItem('api-keys', JSON.stringify(existingKeys));
}

export const requiredKeys = ["OPENAI_API_KEY"];

export async function parseKeyFile(data: string): Promise<Record<string, string>> {
    const lines = data.split('\n')
        .map(line => line.trim())
        .filter(line => !line.startsWith('#'));

    const variables: Record<string, string> = {};

    for (const line of lines) {
        const [key, value] = line.split('=', 2);
        variables[key] = value;
    }

    return variables;
}