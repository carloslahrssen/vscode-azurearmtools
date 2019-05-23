// ----------------------------------------------------------------------------
// Copyright (c) Microsoft Corporation.  All rights reserved.
// ----------------------------------------------------------------------------

import { DeploymentTemplate } from "./DeploymentTemplate";
import * as Json from "./JSON";
import { ParameterDefinition } from "./ParameterDefinition";

export interface ITemplateScope {
    parameterDefinitions: ParameterDefinition[];
    variableDefinitions: Json.Property[];
}

export class TemplateScope implements ITemplateScope {
    // At least one of these must be non-null
    // asdf private _deploymentTemplate: DeploymentTemplate | null;
    // private _scopeObjectValue: Json.ObjectValue | null;

    private _parameterDefinitions: ParameterDefinition[] | null;
    private _variableDefinitions: Json.Property[] | null;

    /**
     * Constructor
     * @param scopeObject The object containing the "parameters" and "variables" properties, or
     * the DeploymentTemplate at the top level (because if the deployment template is malformed, there
     * may be no top-level object value to use)asdf
     */
    constructor(private _scopeObjectValue: Json.ObjectValue | null) {
    }

    public get parameterDefinitions(): ParameterDefinition[] {
        if (!this._parameterDefinitions) {
            this._parameterDefinitions = [];

            if (this._scopeObjectValue) {
                const parameters: Json.ObjectValue = Json.asObjectValue(this._scopeObjectValue.getPropertyValue("parameters"));
                if (parameters) {
                    for (const parameter of parameters.properties) {
                        this._parameterDefinitions.push(new ParameterDefinition(parameter));
                    }
                }
            }
        }
        return this._parameterDefinitions;

    }

    public get variableDefinitions(): Json.Property[] {
        if (!this._variableDefinitions) {
            this._variableDefinitions = [];

            if (this._scopeObjectValue) {
                const variables: Json.ObjectValue = Json.asObjectValue(this._scopeObjectValue.getPropertyValue("variables"));
                if (variables) {
                    this._variableDefinitions = variables.properties;
                }
            }
        }
        return this._variableDefinitions;
    }
}

export interface ITemplateScopeContext {
    isInUserFunction(): boolean;
}

export class TemplateScopeContext {
    constructor(private _template: DeploymentTemplate, private _value: Json.Value) {
        //assert(_token.type === Json.TokenType.QuotedString, "Scope token must be the quoted string that contains the expression");
    }

    public isInUserFunction(): boolean {
        return false;
    }
}
