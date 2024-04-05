/******/ (() => { // webpackBootstrap
/******/ 	"use strict";
var __webpack_exports__ = {};

;// CONCATENATED MODULE: ./src/ApiKeys.ts
/**Retrieves the API key for OpenAI services.  Throws an error if no key is available. */
async function getOpenAiApiKey() {
    const openaiApiKey = await getApiKey("OPENAI_API_KEY");
    if (openaiApiKey === null) {
        console.error("No openai api key available.");
        throw new Error("No openai api key available.");
    }
    return openaiApiKey;
}
async function getApiKey(key) {
    const keys = JSON.parse(localStorage.getItem('api-keys') ?? '{}');
    if (Object.hasOwn(keys, key)) {
        return keys[key];
    }
    else {
        await showImportApiKeysDialog();
        const keys = JSON.parse(localStorage.getItem('api-keys') ?? '{}');
        if (Object.hasOwn(keys, key)) {
            return keys[key];
        }
        else {
            return null;
        }
    }
}
function getRelativeUrlPath() {
    const index = location.pathname.indexOf('/build');
    return location.pathname.slice(0, index);
}
function showImportApiKeysDialog() {
    return new Promise((res, rej) => {
        try {
            Office.context.ui.displayDialogAsync(`${location.origin + getRelativeUrlPath()}/build/dialogs/ApiKeys/index.html`, {
                displayInIframe: true,
                width: 50,
                height: 50,
            }, function (asyncResult) {
                const dialog = asyncResult.value;
                dialog.addEventHandler(Office.EventType.DialogMessageReceived, (args) => {
                    console.log('MESSAGE RECIEVED', args);
                    if (args.message) {
                        const message = JSON.parse(args.message);
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
                dialog.addEventHandler(Office.EventType.DialogEventReceived, (arg) => {
                    if (arg.error) {
                        switch (arg.error) {
                            case 12006:
                                res();
                                break;
                        }
                    }
                });
            });
        }
        catch (err) {
            rej(err);
        }
    });
}
function addApiKeys(keys) {
    const existingKeys = JSON.parse(localStorage.getItem('api-keys') ?? '{}');
    for (const [key, value] of Object.entries(keys)) {
        existingKeys[key] = value;
    }
    localStorage.setItem('api-keys', JSON.stringify(existingKeys));
}
const requiredKeys = ["OPENAI_API_KEY"];
function parseKeyFile(data) {
    const lines = data.split('\n')
        .map(line => line.trim())
        .filter(line => !line.startsWith('#')) // remove comments
        .filter(line => line.trim().length > 0); // remove empty lines
    const variables = {};
    for (let i = 0; i < lines.length; i++) {
        const line = lines[i];
        const splitIndex = line.indexOf('=');
        if (splitIndex >= 0) {
            const key = line.slice(0, splitIndex).trim();
            const value = line.slice(splitIndex + 1).trim();
            variables[key] = value;
        }
        else {
            throw new Error(`SyntaxError: Expected '=' on line ${i}.`);
        }
    }
    return variables;
}

;// CONCATENATED MODULE: ./src/dialogs/ApiKeys/index.ts

document.addEventListener('DOMContentLoaded', main);
function main() {
    const fileInput = document.querySelector('#api-key-file-upload');
    const fileStatus = document.querySelector('#file-status');
    const closeButton = document.querySelector('#close-button');
    fileInput.addEventListener('change', () => {
        handleFileInputChange();
    });
    closeButton.addEventListener('click', () => {
        sendParentMessage('dialogClosed', '');
    });
    async function handleFileInputChange() {
        if (fileInput.files == null || fileInput.files.length == 0)
            return;
        const file = fileInput.files[0];
        const reader = new FileReader();
        const contents = await new Promise((res, rej) => {
            reader.addEventListener('load', e => {
                res(reader.result);
            });
            reader.addEventListener('error', e => {
                rej(reader.error);
            });
            reader.readAsText(file);
        });
        try {
            const keys = await parseKeyFile(contents);
            const missingKeys = [];
            for (const requiredKey of requiredKeys) {
                if (!Object.hasOwn(keys, requiredKey)) {
                    missingKeys.push(requiredKey);
                }
            }
            if (missingKeys.length == 0) {
                fileStatus.innerText = 'Loaded';
                fileStatus.style.color = 'green';
            }
            else {
                fileStatus.innerText = missingKeys.reduce((accumulator, item, index, array) => {
                    accumulator += `"${item}"`;
                    if (index < array.length - 1) {
                        accumulator += ", ";
                    }
                    return accumulator;
                }, 'Missing required keys: ');
                fileStatus.style.color = 'red';
            }
            sendParentMessage("apiKeys", JSON.stringify(keys));
        }
        catch (err) {
            fileStatus.innerText = err.message;
            fileStatus.style.color = 'red';
        }
    }
    function sendParentMessage(type, content) {
        Office.context.ui.messageParent(JSON.stringify({
            type,
            content
        }));
    }
}

/******/ })()
;