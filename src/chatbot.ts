import { getOpenAiApiKey } from "./ApiKeys";
import { runSql } from "./sql";
import { marked } from 'marked';

document.addEventListener('DOMContentLoaded', async () => {
    const chatsContainer = document.querySelector('#chats') as HTMLDivElement;
    const chatInputBox = document.querySelector('#chat-input-box') as HTMLInputElement;
    const tableSelect = document.querySelector('#table-selection') as HTMLSelectElement;

    enum UIState {
        WaitingForInput,
        WaitingForGptResponse,
    }

    let uiState: UIState = UIState.WaitingForInput;
    let messages: { role: string, content?: string, tool_call_id?: string }[] = []

    await waitForOfficeToInitialize();

    await Excel.run(async context => {
        const tables = context.workbook.tables.load("name");

        await context.sync();

        tables.onAdded.add(() => refreshTableSelect());
        tables.onDeleted.add(() => refreshTableSelect());
    });
    await refreshTableSelect();
    await resetConversationForTable(tableSelect.value);

    tableSelect.addEventListener('input', () => {
        resetConversationForTable(tableSelect.value);
    });

    chatInputBox.addEventListener('keydown', e => {
        switch (e.key) {
            case "Enter":
                handleMessageSend();
                break;
        }
    });

    async function handleMessageSend() {
        if (uiState !== UIState.WaitingForInput) return;

        let userMessage = chatInputBox.value.trim();
        chatInputBox.value = '';

        if (userMessage.length == 0) return;

        uiState = UIState.WaitingForGptResponse;
        try {
            messages.push({
                role: 'user',
                content: userMessage
            });
            addToChatsContainer(`You: ${userMessage}`);
            while (true) {
                console.log('CALLING CHATGPT WITH MESSAGES:', messages);
                const response = await fetch('https://api.openai.com/v1/chat/completions',
                    {
                        method: 'POST',
                        headers: {
                            "Content-Type": "application/json",
                            "Authorization": `Bearer ${await getOpenAiApiKey()}`
                        },
                        body: JSON.stringify({
                            model: 'gpt-4-0125-preview',
                            // model: 'gpt-3.5-turbo-1106',
                            messages,
                            tools: [
                                {
                                    type: "function",
                                    function: {
                                        name: "run_sql",
                                        description: "Executes an SQLite statement and returns the result",
                                        parameter: {
                                            type: "object",
                                            properties: {
                                                statement: {
                                                    type: "string",
                                                    description: "The SQLite statement to be executed.  For example, SELECT * WHERE age > 21;"
                                                }
                                            },
                                            required: ["statement"]
                                        }
                                    }
                                }
                            ],
                        })
                    });

                const json = await response.json();
                console.log('RESPONSE', json)
                const responseMessage = json.choices[0].message;
                messages.push(responseMessage);
                if (responseMessage.content) {
                    addToChatsContainer(`AI: ${responseMessage.content}`);
                    break;
                } else if (json.choices[0].finish_reason == "tool_calls") {
                    for (const tool_call of responseMessage.tool_calls) {
                        const argument_object = JSON.parse(tool_call.function.arguments);
                        console.log("TOOL CALL:", tool_call)
                        switch (tool_call.function.name) {
                            case "run_sql":
                                if (argument_object.statement !== undefined) {
                                    const result = await runSql(argument_object.statement);
                                    messages.push({
                                        role: 'tool',
                                        content: JSON.stringify(result),
                                        tool_call_id: tool_call.id
                                    });
                                } else {
                                    messages.push({
                                        role: 'tool',
                                        content: "Error: the required arguement \"statement\", which contains the SQL statement to run, was not provided.  Please try again, providing the \"statement\" argument.",
                                        tool_call_id: tool_call.id
                                    });
                                }
                                break;

                            default:
                                console.error("Unsupported tool.");
                                break;
                        }
                    }
                }
                console.log(messages);
            }
        } finally {
            uiState = UIState.WaitingForInput;
        }
    }

    function addToChatsContainer(message: string) {
        const div = document.createElement('div');
        div.innerHTML = marked.parse(message, { async: false }) as string;
        chatsContainer.append(div);
    }

    function refreshTableSelect() {
        return Excel.run(async context => {
            const tables = context.workbook.tables.load("name");
            await context.sync();

            for (const child of tableSelect.children) {
                child.remove();
            }

            const tableNames = new Set(tables.items.map(table => table.name).sort());

            for (const name of tableNames) {
                const o = document.createElement('option');
                o.value = name;
                o.innerText = name;
                tableSelect.append(o);
            }
        });
    }

    async function resetConversationForTable(tableName: string) {
        chatInputBox.innerHTML = '';

        const headers = await getTableHeaders(tableName);

        let system_prompt = 'Your purpose is to answer questions from the user about database tables.  The table you are acting on is named "' + tableName + '" and \
        its columns are ';

        for (let index = 0; index < headers.length; index++) {
            const header = headers[index];
            system_prompt += '"' + header + '"';
            if (index < headers.length - 1) {
                system_prompt += ', ';
            }
        }

        system_prompt += '.  Execute SQLite statements as necessary to gain information to answer the user\'s questions.';

        messages = [{
            role: 'system',
            content: system_prompt
        }];
    }




});

function waitForOfficeToInitialize() {
    return new Promise<void>((res, rej) => {
        Office.initialize = () => {
            res();
        };
    });
}

function getTableHeaders(tableName: string) {
    return new Promise<string[]>(async (res, rej) => {
        await Excel.run(async context => {
            try {
                const table = context.workbook.tables.getItem(tableName);
                const headerRange = table.getHeaderRowRange().load("values");
                await context.sync();

                res(headerRange.values[0].map(val => val.toString()));
            } catch (err) {
                rej(err);
            }
        });
    });
}

