/******/ (() => { // webpackBootstrap
/******/ 	"use strict";
/******/ 	var __webpack_modules__ = ({

/***/ 150:
/***/ ((__unused_webpack_module, __webpack_exports__, __webpack_require__) => {

/* harmony export */ __webpack_require__.d(__webpack_exports__, {
/* harmony export */   u4: () => (/* binding */ getOpenAiApiKey)
/* harmony export */ });
/* unused harmony exports requiredKeys, parseKeyFile */
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
function showImportApiKeysDialog() {
    return new Promise((res, rej) => {
        try {
            Office.context.ui.displayDialogAsync(`${location.origin}/build/dialogs/ApiKeys/index.html`, {
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
const requiredKeys = (/* unused pure expression or super */ null && (["OPENAI_API_KEY"]));
async function parseKeyFile(data) {
    const lines = data.split('\n')
        .map(line => line.trim())
        .filter(line => !line.startsWith('#'));
    const variables = {};
    for (const line of lines) {
        const [key, value] = line.split('=', 2);
        variables[key] = value;
    }
    return variables;
}


/***/ }),

/***/ 144:
/***/ ((__unused_webpack_module, __webpack_exports__, __webpack_require__) => {


// EXPORTS
__webpack_require__.d(__webpack_exports__, {
  F2: () => (/* binding */ generateSqlQueryFromNaturalLanguage),
  Jh: () => (/* binding */ runSql)
});

// UNUSED EXPORTS: doesTableExist

// EXTERNAL MODULE: ./src/ApiKeys.ts
var ApiKeys = __webpack_require__(150);
;// CONCATENATED MODULE: ./src/utilities.ts
async function Delay(ms) {
    return new Promise(res => {
        setTimeout(res, ms);
    });
}

;// CONCATENATED MODULE: ./src/sql.ts


const JsTypesToSqliteTypesMap = {
    string: "text",
    number: "double",
    bigint: "bigint",
    boolean: "boolean",
    undefined: "null",
};
async function getDatabase() {
    // Wait for SQL to be ready
    while (!database) {
        await Delay(100);
    }
    return database;
}
let database = null;
(async () => {
    // const initSqlJs = require('sql.js');
    // or if you are in a browser:
    // @ts-ignore
    const initSqlJs = window.initSqlJs;
    // const initSqlJs = await import('https://cdnjs.cloudflare.com/ajax/libs/sql.js/1.10.2/sql-wasm.js')
    const SQL = await initSqlJs({
        // Required to load the wasm binary asynchronously. Of course, you can host it wherever you want
        // You can omit locateFile completely when running in node
        locateFile: (file) => `https://sql.js.org/dist/${file}`
    });
    // Create a database
    database = new SQL.Database();
    // NOTE: You can also use new SQL.Database(data) where
    // data is an Uint8Array representing an SQLite database file
})();
async function doesTableExist(tableName) {
    const database = await getDatabase();
    const query = `SELECT name FROM sqlite_master WHERE type = 'table' AND name = '${tableName}';`;
    const result = database.exec(query);
    return result[0]?.values.length > 0;
}
async function extractReferencedTableNameFromSql(statement) {
    const match = statement.match(/(?:from|join)\s+(\w+)/i);
    return (match && match[1]) ? match[1] : null;
}
async function importTableIntoSQL(database, tableName) {
    // We need to import the table into SQL
    await Excel.run(async (context) => {
        const table = context.workbook.tables.getItem(tableName);
        const headerRange = table.getHeaderRowRange().load("values");
        const bodyRange = table.getDataBodyRange().load("values");
        await context.sync();
        const eventResult = table.onChanged.add(async (args) => {
            // Delete the table
            sqlstr = `DROP TABLE ${tableName};`;
            database.run(sqlstr);
            console.log(`Table "${tableName}" dropped from SQL.`);
            await Excel.run(eventResult.context, async (context) => {
                eventResult.remove();
                await context.sync();
            });
        });
        const totalValues = headerRange.values.concat(bodyRange.values);
        console.log('Range', totalValues);
        // Create a table from the range
        // Ensure the range is at least 2 rows x 1 column in size
        if (totalValues.length < 2 || totalValues[0].length < 1) {
            throw new Error("Invalid table.");
        }
        let headers = totalValues[0].map((value, colIndex) => {
            let type = null;
            for (let rowIndex = 1; rowIndex < totalValues.length; rowIndex++) {
                const value = totalValues[rowIndex][colIndex];
                // Empty cells are ignored.
                if (value.toString().trim().length == 0) {
                    continue;
                }
                // Try to parse as number
                else if (!isNaN(parseFloat(value)) && isFinite(parseFloat(value))) {
                    type = "number";
                }
                else {
                    type = "string";
                    break;
                }
            }
            if (type == null) {
                type = "string";
            }
            return {
                name: value.toString(),
                type: JsTypesToSqliteTypesMap[type]
            };
        });
        console.log(headers);
        let sqlstr = headers.reduce((accumulator, item, index, array) => {
            accumulator += '"' + item.name + '" ' + item.type;
            if (index < array.length - 1) {
                accumulator += ", ";
            }
            return accumulator;
        }, `CREATE TABLE ${tableName} (`) + "); ";
        for (let index = 1; index < totalValues.length; index++) {
            const row = totalValues[index];
            sqlstr += row.reduce((accumulator, item, index, array) => {
                const type = headers[index].type;
                switch (type) {
                    case "text":
                        accumulator += `'${item}'`;
                        break;
                    case "double":
                    case "int":
                    case "bigint":
                        if (item.toString().trim().length == 0) {
                            accumulator += '0';
                        }
                        else {
                            accumulator += item;
                        }
                        break;
                    case "boolean":
                        accumulator += item == true ? 'TRUE' : 'FALSE';
                        break;
                    case "null":
                        accumulator += 'NULL';
                        break;
                    default:
                        break;
                }
                if (index < array.length - 1) {
                    accumulator += ", ";
                }
                return accumulator;
            }, `INSERT INTO ${tableName} VALUES (`) + "); ";
        }
        console.log(sqlstr);
        database.run(sqlstr);
    });
}
/**Caches GPT queries to their corresponding SQL statement. */
const gptQuerySqlCache = new Map();
async function generateSqlQueryFromNaturalLanguage(targetTableName, query) {
    let sqlStatement = '';
    await Excel.run(async (context) => {
        const table = context.workbook.tables.getItem(targetTableName);
        const headerRange = table.getHeaderRowRange().load("values");
        await context.sync();
        let cacheKey = targetTableName + '\n';
        for (let index = 0; index < headerRange.values[0].length; index++) {
            const header = headerRange.values[0][index];
            cacheKey += '"' + header + '"\n';
        }
        cacheKey += '\n';
        cacheKey += query;
        if (gptQuerySqlCache.has(cacheKey)) {
            sqlStatement = gptQuerySqlCache.get(cacheKey);
        }
        else {
            const openaiApiKey = await (0,ApiKeys/* getOpenAiApiKey */.u4)();
            let system_prompt = 'Your purpose is to generate SQLite statements.  You will \
        respond to high-level user requests with the appropriate SQL \
        statements.   The table you are acting on is named "' + targetTableName + '" and \
        its columns are ';
            for (let index = 0; index < headerRange.values[0].length; index++) {
                const header = headerRange.values[0][index];
                system_prompt += '"' + header + '"';
                if (index < headerRange.values[0].length - 1) {
                    system_prompt += ', ';
                }
            }
            system_prompt += '.  Provide an SQL statement appropriate for the user\'s \
            query.  Return SQL only, and nothing else.  If you cannot interpret a \
            user\'s request, return "ERROR: {description}".';
            const response = await fetch('https://api.openai.com/v1/chat/completions', {
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
                            content: system_prompt
                        },
                        {
                            role: 'user',
                            content: query
                        }
                    ]
                })
            });
            const json = await response.json();
            const text_response = json.choices[0]?.message.content;
            if (text_response.trim().startsWith("ERROR")) {
                alert(text_response);
                throw new Error(text_response);
            }
            else {
                console.log('AI-generated SQL statement:', text_response);
                sqlStatement = text_response;
                gptQuerySqlCache.set(cacheKey, sqlStatement);
            }
        }
    });
    return sqlStatement;
}
async function runSql(statement) {
    try {
        const tableName = await extractReferencedTableNameFromSql(statement);
        if (!tableName) {
            throw new Error("No table name detected.  Unsupported SQL statement.");
        }
        // Wait for SQL to be ready
        while (!database) {
            await Delay(1000);
        }
        // Check if the table exists
        if (await doesTableExist(tableName)) {
            // Nothing to do
        }
        else {
            await importTableIntoSQL(database, tableName);
        }
        // Execute the SQL statement on the table
        const res = database.exec(statement);
        console.log(res);
        // // Prepare an sql statement
        // const stmt = database.prepare("SELECT * FROM exceltbl WHERE Column_A > :aval");
        // // Bind values to the parameters and fetch the results of the query
        // const result = stmt.getAsObject({ ':aval': 0 });
        // while (stmt.step()) console.log(stmt.get()); // Will print [0, 'hello']
        // // console.log(result); // Will print {a:1, b:'world'}
        // stmt.free();
        // // Delete the table
        // let sqlstr = 'DROP TABLE exceltbl;';
        // database.run(sqlstr);
        // console.log('Table dropped');
        if (res.length > 0) {
            let result = [res[0].columns];
            result = result.concat(res[0].values);
            console.log('RESULT', result);
            return result;
        }
        else {
            let statement = `PRAGMA table_info(${tableName});`;
            // Execute the SQL statement on the table
            const res = database.exec(statement);
            console.log(res);
            if (res[0] && res[0].values) {
                const headers = res[0].values.map((row) => row[1]); // column names are in 2nd position
                console.log(headers);
                return [headers];
            }
            else {
                throw new Error("Cannot get table column names.");
            }
        }
        // Return the statement result
        // return 0
    }
    catch (err) {
        console.error(err);
        throw err;
    }
}


/***/ })

/******/ 	});
/************************************************************************/
/******/ 	// The module cache
/******/ 	var __webpack_module_cache__ = {};
/******/ 	
/******/ 	// The require function
/******/ 	function __webpack_require__(moduleId) {
/******/ 		// Check if module is in cache
/******/ 		var cachedModule = __webpack_module_cache__[moduleId];
/******/ 		if (cachedModule !== undefined) {
/******/ 			return cachedModule.exports;
/******/ 		}
/******/ 		// Create a new module (and put it into the cache)
/******/ 		var module = __webpack_module_cache__[moduleId] = {
/******/ 			// no module.id needed
/******/ 			// no module.loaded needed
/******/ 			exports: {}
/******/ 		};
/******/ 	
/******/ 		// Execute the module function
/******/ 		__webpack_modules__[moduleId](module, module.exports, __webpack_require__);
/******/ 	
/******/ 		// Return the exports of the module
/******/ 		return module.exports;
/******/ 	}
/******/ 	
/************************************************************************/
/******/ 	/* webpack/runtime/define property getters */
/******/ 	(() => {
/******/ 		// define getter functions for harmony exports
/******/ 		__webpack_require__.d = (exports, definition) => {
/******/ 			for(var key in definition) {
/******/ 				if(__webpack_require__.o(definition, key) && !__webpack_require__.o(exports, key)) {
/******/ 					Object.defineProperty(exports, key, { enumerable: true, get: definition[key] });
/******/ 				}
/******/ 			}
/******/ 		};
/******/ 	})();
/******/ 	
/******/ 	/* webpack/runtime/hasOwnProperty shorthand */
/******/ 	(() => {
/******/ 		__webpack_require__.o = (obj, prop) => (Object.prototype.hasOwnProperty.call(obj, prop))
/******/ 	})();
/******/ 	
/************************************************************************/
var __webpack_exports__ = {};
// This entry need to be wrapped in an IIFE because it need to be isolated against other modules in the chunk.
(() => {
/* harmony import */ var _src_ApiKeys__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(150);
/* harmony import */ var _sql__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(144);


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
async function queryTable(tableName, query) {
    const sqlStatement = await (0,_sql__WEBPACK_IMPORTED_MODULE_0__/* .generateSqlQueryFromNaturalLanguage */ .F2)(tableName, query);
    return sql(sqlStatement);
}
queryTable; // Stops 'X is declared but its value is never read' error.
/**
 * Executes SQL on the table provided in the SQL statement.
 * @customfunction
 * @param statement The SQL statement to execute.
 */
function sql(statement) {
    return (0,_sql__WEBPACK_IMPORTED_MODULE_0__/* .runSql */ .Jh)(statement);
}
sql; // Stops 'X is declared but its value is never read' error.
/**Caches formula descriptions to their corresponding formula. */
const gptFormulaCache = new Map();
/**
 * Builds an Excel formula from a natural-language prompt.  Copy and paste the result as
 * values to apply the formula.
 * @customfunction
 * @param description A natural-language description of the formula.
 */
async function makeFormula(description) {
    const cachedFormula = gptFormulaCache.get(description);
    if (cachedFormula !== undefined) {
        return cachedFormula;
    }
    else {
        const openaiApiKey = await (0,_src_ApiKeys__WEBPACK_IMPORTED_MODULE_1__/* .getOpenAiApiKey */ .u4)();
        let systemPrompt = 'Your purpose is to generate Excel formulas based on \
        a description from the user.  Provide an Excel formula that satisfies the user\'s \
        description.  Return the formula only, and nothing else.  The formula must be \
        ready to copy and paste into an Excel cell (for example, it must start with "="). \
        If you cannot satisfy a user\'s request, return "ERROR: {description}".';
        const response = await fetch('https://api.openai.com/v1/chat/completions', {
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
        const text_response = json.choices[0]?.message.content;
        if (text_response.trim().startsWith("ERROR")) {
            alert(text_response);
            throw new Error(text_response);
        }
        else {
            gptFormulaCache.set(description, text_response);
            return text_response;
        }
    }
}
makeFormula; // Stops 'X is declared but its value is never read' error.
/**Caches asks to their corresponding responses. */
const askCache = new Map();
/**
 * Asks the question of ChatGPT
 * @customfunction
 * @param promptPart A natural-language description of the formula.
 */
async function ask(promptPart) {
    let prompt = '';
    promptPart.forEach(range => {
        range.forEach(row => {
            row.forEach(item => {
                prompt += item + '\n\n';
            });
        });
    });
    const cachedFormula = askCache.get(prompt);
    if (cachedFormula !== undefined) {
        return cachedFormula;
    }
    else {
        const openaiApiKey = await (0,_src_ApiKeys__WEBPACK_IMPORTED_MODULE_1__/* .getOpenAiApiKey */ .u4)();
        let systemPrompt = 'You are a helpful assistant.';
        const response = await fetch('https://api.openai.com/v1/chat/completions', {
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
        const text_response = json.choices[0]?.message.content;
        if (text_response.trim().startsWith("ERROR")) {
            alert(text_response);
            throw new Error(text_response);
        }
        else {
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
async function getRangeValue(address) {
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
function DEBUG_createFormattedNumber(value, format) {
    return {
        type: "FormattedNumber",
        basicValue: value,
        numberFormat: format
    };
}
/**
 * Returns a promise that takes `duration` seconds to resolve.
 * @customfunction DEBUG_LongPromiseReturn
 * @param {number} duration
 * @returns {Promise<string>} "Complete."
 */
function DEBUG_longPromiseReturn(duration) {
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
function DEBUG_LogInput(value) {
    console.log(typeof value + ":", value);
}
/**
 * Opens a dialog box
 * @customfunction
 * @returns
 */
function DEBUG_OpenDialog() {
    // Office.context.ui.displayDialogAsync('/build/taskpane/taskpane.html', { displayInIframe: false, width: 200, height: 50 });
    console.log(window.location.host);
    Office.context.ui.displayDialogAsync(`${location.origin}/build/dialogs/ApiKeys/index.html`, {
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
                        break;
                    case "apiKeys":
                        // addApiKeys(JSON.parse(message.content));
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
function DEBUG_ShowAlert(message) {
    alert(message);
}
/**
 * Shows/hides the task pane
 * @customfunction
 * @param {boolean} visible
 * @returns
 */
function DEBUG_SetTaskPaneVisibility(visible) {
    if (visible) {
        Office.addin.showAsTaskpane();
    }
    else {
        Office.addin.hide();
    }
    return 0;
}
CustomFunctions.associate("QUERYTABLE", queryTable);
CustomFunctions.associate("SQL", sql);
CustomFunctions.associate("MAKEFORMULA", makeFormula);
CustomFunctions.associate("ASK", ask);
CustomFunctions.associate("GETRANGEVALUE", getRangeValue);
CustomFunctions.associate("DEBUG_CREATEFORMATTEDNUMBER", DEBUG_createFormattedNumber);
CustomFunctions.associate("DEBUG_LONGPROMISERETURN", DEBUG_longPromiseReturn);
CustomFunctions.associate("DEBUG_LOGINPUT", DEBUG_LogInput);
CustomFunctions.associate("DEBUG_OPENDIALOG", DEBUG_OpenDialog);
CustomFunctions.associate("DEBUG_SHOWALERT", DEBUG_ShowAlert);
CustomFunctions.associate("DEBUG_SETTASKPANEVISIBILITY", DEBUG_SetTaskPaneVisibility);

})();

/******/ })()
;