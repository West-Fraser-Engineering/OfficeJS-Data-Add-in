import { parseKeyFile } from "@src/ApiKeys";
import assert from "assert";

describe("KeyFileParse", () => {
    it("parses a key file (string) containing values for KEY_A and KEY_B", () => {
        const results = parseKeyFile(`
        
    KEY_A =   VAL1
    KEY_B=VAL2=23
        `);

        // console.log(results)

        assert(results["KEY_A"] == "VAL1");
        assert(results["KEY_B"] == "VAL2=23");


    });
});