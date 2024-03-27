import { getOpenAiApiKey } from "@src/ApiKeys";
import { generateSqlQueryFromNaturalLanguage, runSql } from "../sql";

// /**
//  * Executes SQL on the given range.
//  * @customfunction
//  * @param range 
//  * @param statement 
//  */
// async function sqlRange(range: any[][], statement: string): Promise<any[][]> {
//     try {
//         console.log(range, statement);

//         // Wait for SQL to be ready
//         while (!database) {
//             await Delay(1000);
//         }

//         // Create a table from the range
//         // Ensure the range is at least 2 rows x 1 column in size
//         if (range.length < 2 || range[0].length < 1) {
//             throw new Error("Invalid table.");
//         }
//         let headers = range[0].map((value, index) => {
//             return {
//                 name: value.toString() as string,
//                 type: JsTypesToSqliteTypesMap[typeof range[1][index]]
//             };
//         });
//         console.log(headers);


//         let sqlstr = headers.reduce((accumulator, item, index, array) => {
//             accumulator += item.name.replaceAll(" ", "_") + " " + item.type;
//             if (index < array.length - 1) {
//                 accumulator += ", ";
//             }
//             return accumulator;
//         }, "CREATE TABLE exceltbl (") + "); ";

//         for (let index = 1; index < range.length; index++) {
//             const row = range[index];
//             sqlstr += row.reduce((accumulator, item, index, array) => {
//                 const type = headers[index].type
//                 switch (type) {
//                     case "text":
//                         accumulator += `'${item}'`;
//                         break;

//                     case "double":
//                     case "int":
//                     case "bigint":
//                         accumulator += item;
//                         break;

//                     case "boolean":
//                         accumulator += item == true ? 'TRUE' : 'FALSE';
//                         break;

//                     case "null":
//                         accumulator += 'NULL';
//                         break;

//                     default:
//                         break;
//                 }
//                 if (index < array.length - 1) {
//                     accumulator += ", ";
//                 }
//                 return accumulator;
//             }, "INSERT INTO exceltbl VALUES (") + "); ";
//         }

//         console.log(sqlstr)

//         database.run(sqlstr);

//         // Execute the SQL statement on the table
//         const res = database.exec("SELECT * FROM exceltbl WHERE Column_A > 3");
//         console.log(res);
//         // // Prepare an sql statement
//         // const stmt = database.prepare("SELECT * FROM exceltbl WHERE Column_A > :aval");

//         // // Bind values to the parameters and fetch the results of the query
//         // const result = stmt.getAsObject({ ':aval': 0 });
//         // while (stmt.step()) console.log(stmt.get()); // Will print [0, 'hello']
//         // // console.log(result); // Will print {a:1, b:'world'}
//         // stmt.free();

//         // Delete the table
//         sqlstr = 'DROP TABLE exceltbl;';
//         database.run(sqlstr);
//         console.log('Table dropped');

//         let result = [res[0].columns.map((val: string) => val.replaceAll("_", " "))];
//         result = result.concat(res[0].values);

//         console.log('RESULT', result);
//         return result;

//         // Return the statement result
//         // return 0
//     } catch (err) {
//         console.error(err);
//         throw err;
//     }
// }

/**
 * Executes SQL on the table provided in the SQL statement.
 * @customfunction
 * @param tableName
 * @param query 
 */
async function queryTable(tableName: string, query: string): Promise<any[][]> {
    const sqlStatement = await generateSqlQueryFromNaturalLanguage(tableName, query);
    return sql(sqlStatement);
}
queryTable; // Stops 'X is declared but its value is never read' error.

/**
 * Executes SQL on the table provided in the SQL statement.
 * @customfunction
 * @param statement The SQL statement to execute.
 */
function sql(statement: string): Promise<any[][]> {
    return runSql(statement);
}
sql; // Stops 'X is declared but its value is never read' error.

/**Caches formula descriptions to their corresponding formula. */
const gptFormulaCache = new Map<string, string>();
/**
 * Builds an Excel formula from a natural-language prompt.  Copy and paste the result as
 * values to apply the formula.
 * @customfunction
 * @param description A natural-language description of the formula.
 */
async function makeFormula(description: string): Promise<string> {
    const cachedFormula = gptFormulaCache.get(description);
    if (cachedFormula !== undefined) {
        return cachedFormula;
    } else {
        const openaiApiKey = await getOpenAiApiKey();

        let systemPrompt = 'Your purpose is to generate Excel formulas based on \
        a description from the user.  Provide an Excel formula that satisfies the user\'s \
        description.  Return the formula only, and nothing else.  The formula must be \
        ready to copy and paste into an Excel cell (for example, it must start with "="). \
        If you cannot satisfy a user\'s request, return "ERROR: {description}".';

        const response = await fetch('https://api.openai.com/v1/chat/completions',
            {
                method: 'POST',
                headers: {
                    "Content-Type": "application/json",
                    "Authorization": `Bearer ${openaiApiKey}`
                },
                body: JSON.stringify({
                    model: 'gpt-3.5-turbo',
                    messages: [
                        {
                            role: 'system',
                            content: systemPrompt
                        },
                        {
                            role: 'user',
                            content: description
                        }
                    ]
                })
            });

        const json = await response.json();
        const text_response = json.choices[0]?.message.content as string;
        if (text_response.trim().startsWith("ERROR")) {
            alert(text_response);
            throw new Error(text_response);
        } else {
            gptFormulaCache.set(description, text_response);
            return text_response;
        }
    }
}
makeFormula; // Stops 'X is declared but its value is never read' error.

/**Caches asks to their corresponding responses. */
const askCache = new Map<string, string>();
/**
 * Asks the question of ChatGPT
 * @customfunction
 * @param promptPart A natural-language description of the formula.
 */
async function ask(promptPart: string[][][]): Promise<string> {
    let prompt = '';
    promptPart.forEach(range=> {
        range.forEach(row=> {
            row.forEach(item => {
                prompt += item + '\n\n';
            });
        });
    });

    const cachedFormula = askCache.get(prompt);
    if (cachedFormula !== undefined) {
        return cachedFormula;
    } else {
        const openaiApiKey = await getOpenAiApiKey();

        let systemPrompt = 'You are a helpful assistant.';

        const response = await fetch('https://api.openai.com/v1/chat/completions',
            {
                method: 'POST',
                headers: {
                    "Content-Type": "application/json",
                    "Authorization": `Bearer ${openaiApiKey}`
                },
                body: JSON.stringify({
                    model: 'gpt-3.5-turbo',
                    messages: [
                        {
                            role: 'system',
                            content: systemPrompt
                        },
                        {
                            role: 'user',
                            content: prompt
                        }
                    ]
                })
            });

        const json = await response.json();
        const text_response = json.choices[0]?.message.content as string;
        if (text_response.trim().startsWith("ERROR")) {
            alert(text_response);
            throw new Error(text_response);
        } else {
            askCache.set(prompt, text_response);
            return text_response;
        }
    }
}
ask;

/**
 * @customfunction
 * @param address 
 * @returns
 */
async function getRangeValue(address: string) {
    const context = new Excel.RequestContext();
    const range = context.workbook.worksheets.getActiveWorksheet().getRange(address);
    range.load("values");
    await context.sync();
    return range.values[0][0];
}


/**
 * Take a number as the input value and return a formatted number value as the output.
 * @customfunction
 * @param {number} value
 * @param {string} format (e.g. "0.00%")
 * @returns A formatted number value.
 */
function DEBUG_createFormattedNumber(value: number, format: string) {
    return {
        type: "FormattedNumber",
        basicValue: value,
        numberFormat: format
    }
}

/**
 * Returns a promise that takes `duration` seconds to resolve.
 * @customfunction DEBUG_LongPromiseReturn
 * @param {number} duration
 * @returns {Promise<string>} "Complete."
 */
function DEBUG_longPromiseReturn(duration: number): Promise<string> {
    return new Promise((res) => {
        setTimeout(() => {
            res("Complete.");
        }, duration * 1000);
    });
}

/**
 * Logs its input to the dev console.
 * @customfunction
 * @param {any} value
 * @returns
 */
function DEBUG_LogInput(value: any) {
    console.log(typeof value + ":", value);
}

/**
 * Opens a dialog box
 * @customfunction
 * @returns
 */
function DEBUG_OpenDialog() {
    // Office.context.ui.displayDialogAsync('/build/taskpane/taskpane.html', { displayInIframe: false, width: 200, height: 50 });
    console.log(window.location.host)
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
                            break;

                        case "apiKeys":
                            // addApiKeys(JSON.parse(message.content));
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
                            console.log('Dialog closed using system close button.');
                            break;

                        default:

                            break;
                    }
                }
            });
        });
}

/**
 * Shows an alert
 * @customfunction
 * @param {string} message
 * @returns
 */
function DEBUG_ShowAlert(message: string) {
    alert(message);
}

/**
 * Shows/hides the task pane
 * @customfunction
 * @param {boolean} visible
 * @returns
 */
function DEBUG_SetTaskPaneVisibility(visible: boolean) {
    if (visible) {
        Office.addin.showAsTaskpane();
    } else {
        Office.addin.hide();
    }
    return 0;
}







declare var CustomFunctions: any;