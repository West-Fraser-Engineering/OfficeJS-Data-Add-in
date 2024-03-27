import { getOpenAiApiKey } from "./ApiKeys";
import { Delay } from "./utilities";

const JsTypesToSqliteTypesMap: Record<string, string> = {
    string: "text",
    number: "double",
    bigint: "bigint",
    boolean: "boolean",
    undefined: "null",
}

async function getDatabase() {
    // Wait for SQL to be ready
    while (!database) {
        await Delay(100);
    }

    return database;
}

let database: any = null;
(async () => {
    // const initSqlJs = require('sql.js');
    // or if you are in a browser:
    // @ts-ignore
    const initSqlJs = window.initSqlJs;
    // const initSqlJs = await import('https://cdnjs.cloudflare.com/ajax/libs/sql.js/1.10.2/sql-wasm.js')

    const SQL = await initSqlJs({
        // Required to load the wasm binary asynchronously. Of course, you can host it wherever you want
        // You can omit locateFile completely when running in node
        locateFile: (file: any) => `https://sql.js.org/dist/${file}`
    });

    // Create a database
    database = new SQL.Database();
    // NOTE: You can also use new SQL.Database(data) where
    // data is an Uint8Array representing an SQLite database file
})();

export async function doesTableExist(tableName: string) {
    const database = await getDatabase();
    const query = `SELECT name FROM sqlite_master WHERE type = 'table' AND name = '${tableName}';`
    const result = database.exec(query);
    return result[0]?.values.length > 0;
}

async function extractReferencedTableNameFromSql(statement: string) {
    const match = statement.match(/(?:from|join)\s+(\w+)/i);
    return (match && match[1]) ? match[1] : null;
}


async function importTableIntoSQL(database: any, tableName: string) {
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
            })
        });

        const totalValues = headerRange.values.concat(bodyRange.values);

        console.log('Range', totalValues);

        // Create a table from the range
        // Ensure the range is at least 2 rows x 1 column in size
        if (totalValues.length < 2 || totalValues[0].length < 1) {
            throw new Error("Invalid table.");
        }
        let headers = totalValues[0].map((value, colIndex) => {

            let type: string | null = null;

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
                name: value.toString() as string,
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
            sqlstr += row.reduce((accumulator: string, item, index, array) => {
                const type = headers[index].type
                switch (type) {
                    case "text":
                        accumulator += `'${item}'`;
                        break;

                    case "double":
                    case "int":
                    case "bigint":
                        if (item.toString().trim().length == 0) {
                            accumulator += '0';
                        } else {
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

        console.log(sqlstr)

        database.run(sqlstr);
    });
}

/**Caches GPT queries to their corresponding SQL statement. */
const gptQuerySqlCache = new Map<string, string>();
export async function generateSqlQueryFromNaturalLanguage(
    targetTableName: string,
    query: string
): Promise<string> {
    let sqlStatement: string = '';
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
            sqlStatement = gptQuerySqlCache.get(cacheKey) as string;
        } else {
            const openaiApiKey = await getOpenAiApiKey();

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
            const text_response = json.choices[0]?.message.content as string;
            if (text_response.trim().startsWith("ERROR")) {
                alert(text_response);
                throw new Error(text_response);
            } else {
                console.log('AI-generated SQL statement:', text_response)
                sqlStatement = text_response;
                gptQuerySqlCache.set(cacheKey, sqlStatement);
            }
        }
    });

    return sqlStatement;
}

export async function runSql(statement: string): Promise<any[][]> {
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
        } else {
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
        } else {
            let statement = `PRAGMA table_info(${tableName});`
            // Execute the SQL statement on the table
            const res = database.exec(statement);
            console.log(res);
            if (res[0] && res[0].values) {
                const headers = res[0].values.map((row: any[]) => row[1]); // column names are in 2nd position
                console.log(headers)
                return [headers];
            }
            else {
                throw new Error("Cannot get table column names.");
            }
        }

        // Return the statement result
        // return 0
    } catch (err) {
        console.error(err);
        throw err;
    }
}