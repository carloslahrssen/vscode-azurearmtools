// ----------------------------------------------------------------------------
// Copyright (c) Microsoft Corporation.  All rights reserved.
// ----------------------------------------------------------------------------

import * as assert from "assert";

import { FunctionDefinition } from "./FunctionDefinition";
import * as Json from "./JSON";
import * as language from "./Language";

/**
 * This class represents the definition of a user-defined namespace in a deployment template.
 */
export class NamespaceDefinition {
    /* Example:

            "functions": [
                {
                    "namespace": "contoso",
                    "members": {
                        "uniqueName": {
                            "parameters": [
                                {
                                    "name": "namePrefix",
                                    "type": "string"
                                }
                            ],
                            "output": {
                            "type": "string",
                            "value": "[concat(toLower(parameters('namePrefix')), uniqueString(resourceGroup().id))]"
                            }
                        }
                    }
                }
            ],
        */

    private _members: FunctionDefinition[] | undefined;

    constructor(private _value: Json.ObjectValue) {
        assert(_value);
    }

    public get name(): Json.StringValue | undefined {
        return Json.asStringValue(this._value.getPropertyValue("namespace"));
    }

    public get span(): language.Span {
        return this._value.span;
    }

    public get members(): FunctionDefinition[] {
        if (!this._members) {
            this._members = [];

            const members: Json.ObjectValue | undefined = Json.asObjectValue(this._value.getPropertyValue("members"));
            if (members) {
                for (let member of members.properties) {
                    let name: Json.StringValue = member.name;
                    let value = Json.asObjectValue(member.value);
                    if (name && member.value) {
                        let func = new FunctionDefinition(name, value);
                        this._members.push(func);
                    }
                }
            }
        }

        return this._members;
    }

    public getFunctionDefinition(functionName: string): FunctionDefinition | undefined {
        assert(!!functionName, "functionName cannot be null, undefined, or empty");
        let functionNameLC = functionName.toLowerCase();
        return this.members.find(fd => fd.name && fd.name.toString().toLowerCase() === functionNameLC);
    }
}
