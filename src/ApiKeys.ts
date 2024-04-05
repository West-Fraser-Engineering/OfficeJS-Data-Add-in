
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

function getRelativeUrlPath() {
    const index = location.pathname.indexOf('/build');
    return location.pathname.slice(0, index);
}

function showImportApiKeysDialog() {
    return new Promise<void>((res, rej) => {
        try {
            Office.context.ui.displayDialogAsync(
                `${location.origin + getRelativeUrlPath()}/build/dialogs/ApiKeys/index.html`,
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

export function parseKeyFile(data: string): Record<string, string> {
    const lines = data.split('\n')
        .map(line => line.trim())
        .filter(line => !line.startsWith('#')) // remove comments
        .filter(line => line.trim().length > 0) // remove empty lines

    const variables: Record<string, string> = {};

    for (let i = 0; i < lines.length; i++) {
        const line = lines[i];
        const splitIndex = line.indexOf('=');
        if (splitIndex >= 0) {
            const key = line.slice(0, splitIndex).trim();
            const value = line.slice(splitIndex + 1).trim();
            variables[key] = value;
        } else {
            throw new Error(`SyntaxError: Expected '=' on line ${i}.`);
        }
    }

    return variables;
}