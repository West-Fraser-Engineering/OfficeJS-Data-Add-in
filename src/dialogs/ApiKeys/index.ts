import { parseKeyFile, requiredKeys } from "@src/ApiKeys";

document.addEventListener('DOMContentLoaded', main);

function main() {
    const fileInput = document.querySelector('#api-key-file-upload') as HTMLInputElement;
    const fileStatus = document.querySelector('#file-status') as HTMLParagraphElement;
    const closeButton = document.querySelector('#close-button') as HTMLButtonElement;

    fileInput.addEventListener('change', () => {
        handleFileInputChange();
    });

    closeButton.addEventListener('click', () => {
        sendParentMessage('dialogClosed', '');
    });

    async function handleFileInputChange() {
        if (fileInput.files == null || fileInput.files.length == 0) return;

        const file = fileInput.files[0];
        const reader = new FileReader();
        const contents = await new Promise<string>((res, rej) => {
            reader.addEventListener('load', e => {
                res(reader.result as string);
            });
            reader.addEventListener('error', e => {
                rej(reader.error);
            });
            reader.readAsText(file);
        });

        try {
            const keys = await parseKeyFile(contents);

            const missingKeys: string[] = [];

            for (const requiredKey of requiredKeys) {
                if (!Object.hasOwn(keys, requiredKey)) {
                    missingKeys.push(requiredKey)
                }
            }

            if (missingKeys.length == 0) {
                fileStatus.innerText = 'Loaded';
                fileStatus.style.color = 'green';
            } else {
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
        } catch (err: any) {
            fileStatus.innerText = err.message;
            fileStatus.style.color = 'red';
        }
    }

    function sendParentMessage(type: string, content: string) {
        Office.context.ui.messageParent(JSON.stringify({
            type,
            content
        }));
    }
}

