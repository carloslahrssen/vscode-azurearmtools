// ----------------------------------------------------------------------------
// Copyright (c) Microsoft Corporation.  All rights reserved.
// ----------------------------------------------------------------------------

import * as assert from "assert";
import * as Json from "./JSON";

/**
 * This class represents the definition of a user-defined output in a deployment template.
 */
export class OutputDefinition {
    constructor(private _value: Json.ObjectValue) {
        assert(_value);
    }

    public get outputType(): Json.StringValue | undefined {
        return Json.asStringValue(this._value.getPropertyValue("type"));
    }

    public get value(): Json.StringValue | undefined {
        return Json.asStringValue(this._value.getPropertyValue("value"));
    }
}
