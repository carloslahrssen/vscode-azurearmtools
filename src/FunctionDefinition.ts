// ----------------------------------------------------------------------------
// Copyright (c) Microsoft Corporation.  All rights reserved.
// ----------------------------------------------------------------------------

import * as assert from "assert";
import * as Json from "./JSON";
import { OutputDefinition } from "./OutputDefinition";

/**
 * This class represents the definition of a user-defined function in a deployment template.
 */
export class FunctionDefinition {
    private _output: OutputDefinition;

    constructor(private _name: Json.StringValue, private _value: Json.ObjectValue) {
        assert(_name);
        assert(_value);
    }

    public get name(): Json.StringValue {
        return this._name;
    }

    public get output(): OutputDefinition | undefined {
        if (!this._output) {
            let output = Json.asObjectValue(this._value.getPropertyValue("output"));
            if (output) {
                this._output = new OutputDefinition(output);
            }
        }

        return this._output;
    }

    // TODO: parameters
    // TODO: output

    //    public get description(): string {

    // if (this._description === undefined) {
    //     this._description = null;

    //     const parameterDefinition: Json.ObjectValue = Json.asObjectValue(this._property.value);
    //     if (parameterDefinition) {
    //         const metadata: Json.ObjectValue = Json.asObjectValue(parameterDefinition.getPropertyValue("metadata"));
    //         if (metadata) {
    //             const description: Json.StringValue = Json.asStringValue(metadata.getPropertyValue("description"));
    //             if (description) {
    //                 this._description = description.toString();
    //             }
    //         }
    //     }
    // }
    // return this._description;
    // }
}
