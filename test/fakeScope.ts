// ----------------------------------------------------------------------------
// Copyright (c) Microsoft Corporation.  All rights reserved.
// ----------------------------------------------------------------------------

import { TLE } from "../extension.bundle";
import { ITemplateScopeContext } from "../src/TemplateScope";

export class FakeScope implements ITemplateScopeContext {
    public isInUserFunction(): boolean {
        return false;
    }
}

export function parse(stringValue: string): TLE.ParseResult {
    return TLE.Parser.parse(stringValue, new FakeScope());
}
