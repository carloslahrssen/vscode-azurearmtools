// ----------------------------------------------------------------------------
// Copyright (c) Microsoft Corporation.  All rights reserved.
// ----------------------------------------------------------------------------

import * as assert from "assert";
import { AzureRMAssets, FunctionsMetadata } from "./AzureRMAssets";
import { FunctionDefinition } from "./FunctionDefinition";
import { Histogram } from "./Histogram";
import * as Json from "./JSON";
import * as language from "./Language";
import { NamespaceDefinition } from "./NamespaceDefinition";
import { ParameterDefinition } from "./ParameterDefinition";
import { PositionContext } from "./PositionContext";
import * as Reference from "./Reference";
import { ITemplateScope, TemplateScope, TemplateScopeContext } from "./TemplateScope";
import * as TLE from "./TLE";
import * as Utilities from "./Utilities";

export class DeploymentTemplate {
    // Parse result for the template document as a whole
    private _jsonParseResult: Json.ParseResult;

    // The top-level parameters and variables (as opposed to those in user functions and deployment resources)
    private _topLevelScope: ITemplateScope | null;

    // A list of the tokens represented every quoted string in the template
    private _jsonQuotedStringTokens: Json.Token[];

    // A list of parse results for every quoted string in the template
    private _tleParseResults: TLE.ParseResult[];

    private _quotedStringToTleParseResultMap: { [key: string]: TLE.ParseResult };

    // The "functions" section (which is an array of namespace definitions)
    // https://docs.microsoft.com/en-us/azure/azure-resource-manager/resource-group-authoring-templates#functions
    private _namespaceDefinitions: NamespaceDefinition[];

    // All errors and warnings in the template
    private _errors: Promise<language.Issue[]>;
    private _warnings: language.Issue[];

    private _functionCounts: Histogram;
    private _schemaUri: string;

    /**
     * Create a new DeploymentTemplate object.
     *
     * @param _documentText The string text of the document.
     * @param _documentId A unique identifier for this document. Usually this will be a URI to the document.
     */
    constructor(private _documentText: string, private _documentId: string) {
        assert(_documentText !== null);
        assert(_documentText !== undefined);
        assert(_documentId);

        this._jsonParseResult = Json.parse(_documentText);
        this._topLevelScope = new TemplateScope(Json.asObjectValue(this.jsonParseResult.value));
    }

    public get topLevelScope(): ITemplateScope {
        return this._topLevelScope;
    }

    private get jsonQuotedStringTokens(): Json.Token[] {
        if (this._jsonQuotedStringTokens === undefined) {
            this._jsonQuotedStringTokens = [];

            for (const jsonToken of this._jsonParseResult.tokens) {
                if (jsonToken.type === Json.TokenType.QuotedString) {
                    this._jsonQuotedStringTokens.push(jsonToken);
                }
            }
        }
        return this._jsonQuotedStringTokens;
    }

    private get quotedStringToTleParseResultMap(): { [key: string]: TLE.ParseResult } {
        if (this._quotedStringToTleParseResultMap === undefined) {
            this._quotedStringToTleParseResultMap = {};

            let visitor = new StringVisitor(stringValue => {
                let jsonQuotedStringToken = this.getJSONTokenAtDocumentCharacterIndex(stringValue.span.startIndex);
                assert(!!jsonQuotedStringToken, "Expected token");
                assert(jsonQuotedStringToken.type === Json.TokenType.QuotedString, "Expected quoted string token");

                let tleParseResult: TLE.ParseResult = TLE.Parser.parse(
                    jsonQuotedStringToken.toString(),
                    new TemplateScopeContext(this, stringValue)); // asdf
                if (tleParseResult) { // asdf: can't map by string value, they might be in different scopes
                    this._quotedStringToTleParseResultMap[jsonQuotedStringToken.toString()] = tleParseResult;
                }
            });

            if (this.jsonParseResult.value) {
                this.jsonParseResult.value.accept(visitor);
            }

            for (let jsonQuotedStringToken of this.jsonQuotedStringTokens) {
                let parseResult = this._quotedStringToTleParseResultMap[jsonQuotedStringToken.toString()];
                assert(!!parseResult, "Missed a quoted string token");

                // let jsonValue = this.getJSONValueAtDocumentCharacterIndex(jsonQuotedStringToken.span.startIndex);
                // assert(jsonValue);
                // let tleParseResult: TLE.ParseResult = TLE.Parser.parse(jsonQuotedStringToken.toString(), new Scope(this, jsonValue)); // asdf
                // if (tleParseResult) {
                //     this._quotedStringToTleParseResultMap[jsonQuotedStringToken.toString()] = tleParseResult;
                // }
            }
        }
        return this._quotedStringToTleParseResultMap;
    }

    private get tleParseResults(): TLE.ParseResult[] {
        if (this._tleParseResults === undefined) {
            this._tleParseResults = [];

            // tslint:disable-next-line:forin no-for-in // Grandfathered in
            for (let quotedString in this.quotedStringToTleParseResultMap) {
                this._tleParseResults.push(this.quotedStringToTleParseResultMap[quotedString]);
            }
        }
        return this._tleParseResults;
    }

    /**
     * Get the document text as a string.
     */
    public get documentText(): string {
        return this._documentText;
    }

    /**
     * The unique identifier for this deployment template. Usually this will be a URI to the document.
     */
    public get documentId(): string {
        return this._documentId;
    }

    public get schemaUri(): string {
        if (this._schemaUri === undefined) {
            this._schemaUri = null;

            const value: Json.ObjectValue = Json.asObjectValue(this._jsonParseResult.value);
            if (value) {
                const schema: Json.Value = Json.asStringValue(value.getPropertyValue("$schema"));
                if (schema) {
                    this._schemaUri = schema.toString();
                }
            }
        }
        return this._schemaUri;
    }

    public hasValidSchemaUri(): boolean {
        return Utilities.isValidSchemaUri(this.schemaUri);
    }

    public get errors(): Promise<language.Issue[]> {
        if (this._errors === undefined) {
            this._errors = new Promise<language.Issue[]>(async (resolve, reject) => {
                try {
                    let functions: FunctionsMetadata = await AzureRMAssets.getFunctionsMetadata();
                    const parseErrors: language.Issue[] = [];
                    for (const jsonQuotedStringToken of this.jsonQuotedStringTokens) {
                        const jsonTokenStartIndex: number = jsonQuotedStringToken.span.startIndex;

                        const tleParseResult: TLE.ParseResult = this.getTLEParseResultFromJSONToken(jsonQuotedStringToken);
                        for (const error of tleParseResult.errors) {
                            parseErrors.push(error.translate(jsonTokenStartIndex));
                        }

                        const tleExpression: TLE.Value = tleParseResult.expression;
                        //asdf
                        // const tleUndefinedParameterAndVariableVisitor =
                        //     TLE.UndefinedParameterAndVariableVisitor.visit(
                        //         tleExpression,
                        //         this,
                        //         new Scope(jsonQuotedStringToken)); //asdf
                        // for (const error of tleUndefinedParameterAndVariableVisitor.errors) {
                        //     parseErrors.push(error.translate(jsonTokenStartIndex));
                        // }

                        const tleUnrecognizedFunctionVisitor = TLE.UnrecognizedFunctionVisitor.visit(this, tleExpression, functions);
                        for (const error of tleUnrecognizedFunctionVisitor.errors) {
                            parseErrors.push(error.translate(jsonTokenStartIndex));
                        }

                        const tleIncorrectArgumentCountVisitor = TLE.IncorrectFunctionArgumentCountVisitor.visit(tleExpression, functions);
                        for (const error of tleIncorrectArgumentCountVisitor.errors) {
                            parseErrors.push(error.translate(jsonTokenStartIndex));
                        }

                        const tleUndefinedVariablePropertyVisitor = TLE.UndefinedVariablePropertyVisitor.visit(tleExpression, this);
                        for (const error of tleUndefinedVariablePropertyVisitor.errors) {
                            parseErrors.push(error.translate(jsonTokenStartIndex));
                        }
                    }

                    if (this.jsonParseResult) {
                        const deploymentTemplateObject: Json.ObjectValue = Json.asObjectValue(this.jsonParseResult.value);
                        if (deploymentTemplateObject) {
                            const variablesObject: Json.ObjectValue = Json.asObjectValue(deploymentTemplateObject.getPropertyValue("variables"));
                            if (variablesObject) {
                                //asdf
                                // const referenceInVariablesFinder = new ReferenceInVariableDefinitionJSONVisitor(
                                //     this,
                                //     new Scope(variablesObject)); //asdf
                                // variablesObject.accept(referenceInVariablesFinder);

                                // for (const referenceSpan of referenceInVariablesFinder.referenceSpans) {
                                //     parseErrors.push(
                                //         new language.Issue(referenceSpan, "reference() cannot be invoked inside of a variable definition."));
                                // }
                            }
                        }
                    }

                    resolve(parseErrors);
                } catch (err) {
                    reject(err);
                }
            });
        }

        return this._errors;
    }

    public get warnings(): language.Issue[] {
        // asdf do for all subscopes (user funcs, sub deployments)
        if (this._warnings === undefined) {
            this._warnings = [];

            for (const parameterDefinition of this.parameterDefinitions) {
                const parameterReferences: Reference.List =
                    this.findReferences(Reference.ReferenceKind.Parameter, parameterDefinition.name.toString());
                if (parameterReferences.length === 1) {
                    this._warnings.push(
                        new language.Issue(parameterDefinition.name.span, `The parameter '${parameterDefinition.name.toString()}' is never used.`));
                }
            }

            for (const variableDefinition of this.variableDefinitions) {
                const variableReferences: Reference.List = this.findReferences(Reference.ReferenceKind.Variable, variableDefinition.name.toString());
                if (variableReferences.length === 1) {
                    this._warnings.push(
                        new language.Issue(variableDefinition.name.span, `The variable '${variableDefinition.name.toString()}' is never used.`));
                }
            }

            //asdf
            // for (const nsDefinition of this.namespaceDefinitions) {
            //     this._warnings.push(new language.Issue(nsDefinition.name.span, `namespace ${nsDefinition.name.toString()}`));

            //     for (const funcDefinition of nsDefinition.members) {
            //         this._warnings.push(new language.Issue(
            //             funcDefinition.name.span,
            //             `func ${funcDefinition.name.toString()}`)
            //         );
            //     }
            // }
        }
        return this._warnings;
    }

    public get functionCounts(): Histogram {
        if (this._functionCounts === undefined) {
            this._functionCounts = new Histogram();

            for (let tleParseResult of this.tleParseResults) {
                let tleFunctionCountVisitor = TLE.FunctionCountVisitor.visit(tleParseResult.expression);
                this._functionCounts.add(tleFunctionCountVisitor.functionCounts);
            }
        }
        return this._functionCounts;
    }

    public get jsonParseResult(): Json.ParseResult {
        return this._jsonParseResult;
    }

    /**
     * Get the number of lines that are in the file.
     */
    public get lineCount(): number {
        return this._jsonParseResult.lineLengths.length;
    }

    /**
     * Get the maximum column index for the provided line. For the last line in the file,
     * the maximum column index is equal to the line length. For every other line in the file,
     * the maximum column index is less than the line length.
     */
    public getMaxColumnIndex(lineIndex: number): number {
        return this._jsonParseResult.getMaxColumnIndex(lineIndex);
    }

    /**
     * Get the maximum document character index for this deployment template.
     */
    public get maxCharacterIndex(): number {
        return this._jsonParseResult.maxCharacterIndex;
    }

    // public get parameterDefinitions(): ParameterDefinition[] {
    //     return this._topLevelScope.parameterDefinitions;
    // }

    // public get variableDefinitions(): Json.Property[] {
    //     return this._topLevelScope.variableDefinitions;
    // }

    public get namespaceDefinitions(): NamespaceDefinition[] {
        if (!this._namespaceDefinitions) {
            this._namespaceDefinitions = [];

            const value: Json.ObjectValue = Json.asObjectValue(this._jsonParseResult.value);
            if (value) {
                const namespaces: Json.ArrayValue = Json.asArrayValue(value.getPropertyValue("functions"));
                if (namespaces) {
                    for (const member of namespaces.elements) {
                        let valueObject = Json.asObjectValue(member);
                        if (valueObject) {
                            this._namespaceDefinitions.push(new NamespaceDefinition(valueObject));
                        }
                    }
                }
            }
        }
        return this._namespaceDefinitions;
    }

    public getParameterDefinition(parameterName: string): ParameterDefinition | null {
        assert(parameterName, "parameterName cannot be null, undefined, or empty");

        const unquotedParameterName = Utilities.unquote(parameterName);
        let parameterNameLC = unquotedParameterName.toLowerCase();

        assert(!!this.parameterDefinitions);
        let result = this.parameterDefinitions.find(pd => pd.name && pd.name.toString().toLowerCase() === parameterNameLC);
        return result ? result : null;
    }

    public getNamespaceDefinition(namespaceName: string): NamespaceDefinition | undefined {
        assert(!!namespaceName, "namespaceName cannot be null, undefined, or empty");
        let namespaceNameLC = namespaceName.toLowerCase();
        return this.namespaceDefinitions.find(nd => nd.name && nd.name.toString().toLowerCase() === namespaceNameLC);
    }

    public getFunctionDefinition(namespaceName: string, functionName: string): FunctionDefinition | null {
        assert(!!functionName, "functionName cannot be null, undefined, or empty");
        let nd = this.getNamespaceDefinition(namespaceName);
        if (nd) {
            let result = nd.getFunctionDefinition(functionName);
            return result ? result : null;
        }

        return null;
    }

    public getVariableDefinition(variableName: string): Json.Property | undefined {
        assert(variableName, "variableName cannot be null, undefined, or empty");

        const unquotedVariableName = Utilities.unquote(variableName);
        const variableNameLC = unquotedVariableName.toLowerCase();

        assert(!!this.variableDefinitions);
        let result = this.variableDefinitions.find(vd => vd.name && vd.name.toString().toLowerCase() === variableNameLC);
        return result ? result : null;
    }

    // TODO: findNamespaceDefinitionsWithPrefix

    public getVariableDefinitionFromFunction(tleFunction: TLE.FunctionValue): Json.Property {
        let result: Json.Property = null;

        if (tleFunction && tleFunction.isBuiltin("variables")) { // asdf
            const variableName: TLE.StringValue = TLE.asStringValue(tleFunction.argumentExpressions[0]);
            if (variableName) {
                result = this.getVariableDefinition(variableName.toString());
            }
        }

        return result;
    }

    public findParameterDefinitionsWithPrefix(parameterNamePrefix: string): ParameterDefinition[] {
        assert(parameterNamePrefix !== null, "parameterNamePrefix cannot be null");
        assert(parameterNamePrefix !== undefined, "parameterNamePrefix cannot be undefined");

        let result: ParameterDefinition[] = [];

        if (parameterNamePrefix !== "") {
            let lowerCasedPrefix = parameterNamePrefix.toLowerCase();
            for (let parameterDefinition of this.parameterDefinitions) {
                if (parameterDefinition.name.toString().toLowerCase().startsWith(lowerCasedPrefix)) {
                    result.push(parameterDefinition);
                }
            }
        } else {
            result = this.parameterDefinitions;
        }

        return result;
    }

    public findVariableDefinitionsWithPrefix(variableNamePrefix: string): Json.Property[] {
        assert(variableNamePrefix !== null, "variableNamePrefix cannot be null");
        assert(variableNamePrefix !== undefined, "variableNamePrefix cannot be undefined");

        let result: Json.Property[];
        if (variableNamePrefix) {
            result = [];

            const lowerCasedPrefix = variableNamePrefix.toLowerCase();
            for (const variableDefinition of this.variableDefinitions) {
                if (variableDefinition.name.toString().toLowerCase().startsWith(lowerCasedPrefix)) {
                    result.push(variableDefinition);
                }
            }
        } else {
            result = this.variableDefinitions;
        }

        return result;
    }

    public getDocumentCharacterIndex(documentLineIndex: number, documentColumnIndex: number): number {
        return this._jsonParseResult.getCharacterIndex(documentLineIndex, documentColumnIndex);
    }

    public getDocumentPosition(documentCharacterIndex: number): language.Position {
        return this._jsonParseResult.getPositionFromCharacterIndex(documentCharacterIndex);
    }

    public getJSONTokenAtDocumentCharacterIndex(documentCharacterIndex: number): Json.Token {
        return this._jsonParseResult.getTokenAtCharacterIndex(documentCharacterIndex);
    }

    public getJSONValueAtDocumentCharacterIndex(documentCharacterIndex: number): Json.Value {
        return this._jsonParseResult.getValueAtCharacterIndex(documentCharacterIndex);
    }

    public getContextFromDocumentLineAndColumnIndexes(documentLineIndex: number, documentColumnIndex: number): PositionContext {
        return PositionContext.fromDocumentLineAndColumnIndexes(this, documentLineIndex, documentColumnIndex);
    }

    public getContextFromDocumentCharacterIndex(documentCharacterIndex: number): PositionContext {
        return PositionContext.fromDocumentCharacterIndex(this, documentCharacterIndex);
    }

    public getTLEParseResultFromJSONToken(jsonToken: Json.Token): TLE.ParseResult {
        return jsonToken ? this.getTLEParseResultFromString(jsonToken.toString()) : null;
    }

    public getTLEParseResultFromJSONStringValue(jsonStringValue: Json.StringValue): TLE.ParseResult {
        return jsonStringValue ? this.getTLEParseResultFromString(`"${jsonStringValue.toString()}"`) : null;
    }

    private getTLEParseResultFromString(value: string): TLE.ParseResult {
        let result: TLE.ParseResult = null;
        if (value) {
            result = this.quotedStringToTleParseResultMap[value];
            if (result === undefined) {
                result = null;
            }
        }
        return result;
    }

    public findReferences(referenceType: Reference.ReferenceKind, referenceName: string): Reference.List {
        const result: Reference.List = new Reference.List(referenceType);

        if (referenceName) {
            switch (referenceType) {
                case Reference.ReferenceKind.Parameter:
                    const parameterDefinition: ParameterDefinition = this.getParameterDefinition(referenceName);
                    if (parameterDefinition) {
                        result.add(parameterDefinition.name.unquotedSpan);
                    }
                    break;

                case Reference.ReferenceKind.Variable:
                    const variableDefinition: Json.Property = this.getVariableDefinition(referenceName);
                    if (variableDefinition) {
                        result.add(variableDefinition.name.unquotedSpan);
                    }
                    break;

                default:
                    assert.fail(`Unrecognized Reference.Kind: ${referenceType}`);
                    break;
            }

            for (const jsonStringToken of this.jsonQuotedStringTokens) {
                const tleParseResult: TLE.ParseResult = this.getTLEParseResultFromJSONToken(jsonStringToken);
                if (tleParseResult && tleParseResult.expression) {
                    const visitor = TLE.FindReferencesVisitor.visit(tleParseResult.expression, referenceType, referenceName);
                    result.addAll(visitor.references.translate(jsonStringToken.span.startIndex));
                }
            }
        }

        return result;
    }
}

export class ReferenceInVariableDefinitionJSONVisitor extends Json.Visitor {
    private _referenceSpans: language.Span[] = [];

    constructor(private _deploymentTemplate: DeploymentTemplate) {
        super();

        assert(_deploymentTemplate);
    }

    public get referenceSpans(): language.Span[] {
        return this._referenceSpans;
    }

    public visitStringValue(value: Json.StringValue): void {
        assert(value, "Cannot visit a null or undefined Json.StringValue.");

        const tleParseResult: TLE.ParseResult = this._deploymentTemplate.getTLEParseResultFromJSONStringValue(value);
        if (tleParseResult && tleParseResult.expression) {
            const tleVisitor = new ReferenceInVariableDefinitionTLEVisitor();
            tleParseResult.expression.accept(tleVisitor);

            const jsonValueStartIndex: number = value.startIndex;
            for (const tleReferenceSpan of tleVisitor.referenceSpans) {
                this._referenceSpans.push(tleReferenceSpan.translate(jsonValueStartIndex));
            }
        }
    }
}

class ReferenceInVariableDefinitionTLEVisitor extends TLE.Visitor {
    private _referenceSpans: language.Span[] = [];

    public get referenceSpans(): language.Span[] {
        return this._referenceSpans;
    }

    public visitFunction(functionValue: TLE.FunctionValue): void {
        if (functionValue && functionValue.doesNameMatch("", "reference")) {
            this._referenceSpans.push(functionValue.nameToken.span);
        }

        super.visitFunction(functionValue);
    }
}

class StringVisitor extends Json.Visitor {
    constructor(private _visitStringValue: (stringValue: Json.StringValue) => void) {
        super();
    }

    public visitStringValue(stringValue: Json.StringValue): void {
        this._visitStringValue(stringValue);
    }
}
