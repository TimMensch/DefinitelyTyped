// Type definitions for xlsx-populate 1.19
// Project: https://github.com/dtjohnson/xlsx-populate
// Definitions by: Tim Mensch <https://github.com/TimMensch>
// Definitions: https://github.com/DefinitelyTyped/DefinitelyTyped
// TypeScript Version: 2.2

/// <reference types="node" />

declare type Style = any;
declare type CellType = string | boolean | number | Date | undefined;
declare type StringArray2D = string[][];
declare type DataTypes = string | number[] | ArrayBuffer | Uint8Array | Buffer;

interface IAddressOptions {
    includeSheetName?: boolean;
    anchored?: boolean;
}

declare class FormulaError {
    /**
     * Get the error code.
     */
    error(): string;
    /**
     * #DIV/0! error.
     */
    static DIV0: FormulaError;
    /**
     * #N/A error.
     */
    static NA: FormulaError;
    /**
     * #NAME? error.
     */
    static NAME: FormulaError;
    /**
     * #NULL! error.
     */
    static NULL: FormulaError;
    /**
     * #NUM! error.
     */
    static NUM: FormulaError;
    /**
     * #REF! error.
     */
    static REF: FormulaError;
    /**
     * #VALUE! error.
     */
    static VALUE: FormulaError;
}

declare interface Base {
    address(opts?: IAddressOptions): string;
    sheet(): Sheet;
    workbook(): Workbook;
}

declare interface Color {
    rgb?: string;
    theme?: number;
    tint?: number;
}
declare interface IHyperlinkOptions {
    hyperLink?: string | Cell;
    tooltip?: string;
    email?: string;
    emailSubject?: string;
}

declare interface Cell extends Base {
    active(): boolean;
    active(active: boolean): Cell;
    column(): Column;
    clear(): Cell;
    columnName(): number;
    columnNumber(): number;
    find(pattern: string | RegExp, replacement?: string | replacer): boolean;
    formula(): string;
    formula(formula: string): Cell;
    hyperlink(): string | undefined;
    hyperlink(hyperlink: string | undefined | IHyperlinkOptions): Cell;
    dataValidation(): object | undefined;
    dataValidation(dataValidation: object | undefined): Cell;
    tap<T>(callback: (cell: Cell) => void): Cell;
    thru<T>(callback: (cell: Cell) => T): T;
    rangeTo(cell: Cell | string): Range;
    relativeCell(rowOffset: number, columnOffset: number): Cell;
    row(): Row;
    rowNumber(): number;
    style(name: string): Style;
    style(names: string[]): { [styleName: string]: Style };
    style(name: string, value: Style): Cell;
    style(style: Style): Cell;

    style(name: string, values: StringArray2D): Range;
    value(): CellType;
    value(value: CellType): Cell;
    value(valueArray: CellType[][]): Range;
    addHorizontalPageBreak(): Cell;
}

declare interface Column extends Base {
    cell(rowNumber: number): Cell;
    columnName(): string;
    columnNumber(): number;
    hidden(): boolean;
    hidden(hidden: boolean): Column;
    width(): undefined | number;
    width(width: number): Column;
    addPageBreak(): Column;
    style(name: string): Style;
    style(names: string[]): { [styleName: string]: Style };
    style(name: string, value: Style): Cell;
    style(style: Style): Cell;
}

declare interface Row extends Base {
    cell(columnNameOrNumber: number | string): Cell;
    height(): undefined | number;
    height(height: number): Row;
    hidden(): boolean;
    hidden(hidden: boolean): Row;
    rowNumber(): number;
    addPageBreak(): Row;

    style(name: string): Style;
    style(names: string[]): { [styleName: string]: Style };
    style(name: string, value: Style): Cell;
    style(style: Style): Cell;
}

declare interface PageBreaks {
    count: number;
    list: any[];
    add(id: number): PageBreaks;
    remove(index: number): PageBreaks;
}

declare interface Range {
    address(opts?: IAddressOptions): string;
    cell(ri: number, ci: number): Cell;
    autoFilter(): Range;
    cells(): Cell[][];
    clear(): Range;
    endCell(): Cell;
    forEach(callback: (cell: Cell, ri: number, ci: number, range: Range) => void): Range;
    formula(): string | undefined;
    formula(formula: string | undefined): Range;
    map<T>(callback: (cell: Cell, ri: number, ci: number, range: Range) => T): T[][];
    merged(): boolean;
    merged(merged: boolean): Range;
    dataValidation(): object | undefined;
    dataValidation(dataValidation: object | undefined): Range;
    reduce(
        callback: (acc: any, cell: Cell, ri: number, ci: number, range: Range) => void,
        initialValue: any
    ): any;
    sheet(): Sheet;
    startCell(): Cell;
    style(name: string): Style[][];
    style(names: string[]): { [styleName: string]: Style };
    style(name: string, value: Style): Range;
    tap<T>(callback: (range: Range) => void): Range;
    thru<T>(callback: (range: Range) => T): T;
    value(): CellType[][];
    value(
        p:
            | CellType
            | CellType[][]
            | ((cell: Cell, ri: number, ci: number, range: Range) => any)
    ): Range;
    workbook(): Workbook;
}

type replacer = (substring: string, ...args: any[]) => string;

declare interface Sheet {
    active(): boolean;
    active(active: boolean): Sheet;
    activeCell(): Cell;
    activeCell(cell: Cell): Sheet;
    activeCell(rowNumber: number, columnNameOrNumber: number | string): Sheet;
    cell(address: string): Cell;
    cell(rowNumber: number, columnNameOrNumber: number | string): Cell;
    column(columnNameOrNumber: string | number): Column;
    definedName(name: string): undefined | string | Cell | Range | Row | Column;
    definedName(name: string, refersTo: string | Cell | Range | Row | Column): Workbook;
    delete(): Workbook;
    find(pattern: string | RegExp, replacement?: string | replacer): Cell[];
    gridLinesVisible(selected?: boolean): Sheet;
    hidden(): boolean | string;
    hidden(hidden: boolean | string): Sheet;
    move(indexOrBeforeSheet: number | string | Sheet): Sheet;
    name(): string;
    name(name: string): Sheet;
    range(address: string): Range;
    range(startCell: string | Cell, endCell: string | Cell): Range;
    range(
        startRowNumber: number,
        startColumnNameOrNumber: string | number,
        endRowNumber: number,
        endColumnNameOrNumber: string | number
    ): Range;
    autoFilter(range?: Range): Sheet;
    row(rowNumber: number): Row;
    tabColor(): undefined | Color;
    tabColor(color: Color | string | number): void;
    tabSelected(): boolean;
    tabSelected(selected: boolean): Sheet;
    usedRange(): Range | undefined;
    workbook(): Workbook;
    pageBreaks(): PageBreaks;
    verticalPageBreaks(): PageBreaks;
    horizontalPageBreaks(): PageBreaks;
    hyperlink(address: string): string | undefined;
    hyperlink(address: string, hyperlink: string, internal?: boolean): Sheet;
    hyperlink(address: string, optsOrCell: Cell | IHyperlinkOptions): Sheet;
    printOptions(attributeName: string): boolean;
    printOptions(attributeName: string, attributeEnabled: boolean): Sheet;
    printGridLines(): boolean;
    printGridLines(enabled: boolean): Sheet;
    pageMargins(attributeName: string): number;
    pageMargins(
        attributeName: string,
        attributeStringValue: string | number | undefined
    ): Sheet;
    pageMarginsPreset(): string;
    pageMarginsPreset(presetName: undefined | string): Sheet;
    pageMarginsPreset(presetName: string, presetAttributes: any): Sheet;
}

declare type OutputTypes =
    | "base64"
    | "binarystring"
    | "uint8array"
    | "arraybuffer"
    | "blob"
    | "nodebuffer";

declare interface IOOptions {
    password?: string;
}

declare interface IOutputOptions extends IOOptions {
    type?: OutputTypes;
}

declare interface ICoreProperties {
    [key: string]: any;
    // ??
}

declare interface Workbook {
    activeSheet(): Sheet;
    activeSheet(sheet: Sheet): Workbook;
    addSheet(name: string, indexOrBeforeSheet?: number | string | Sheet): Sheet;
    definedName(name: string): undefined | string | Cell | Range | Row | Column;
    definedName(name: string, refersTo?: string | Cell | Range | Row | Column): Workbook;
    find(pattern: string | RegExp, replacement?: string | replacer): boolean;
    moveSheet(
        sheet: Sheet | string | number,
        indexOrBeforeSheet: Sheet | string | number
    ): Workbook;
    outputAsync(typeOrOpts?: OutputTypes | IOutputOptions): Promise<DataTypes>;
    sheet(sheetNameOrIndex: string | number): Sheet | undefined;
    sheets(): Sheet[];
    property(name: string): any;
    property(names: string[]): { [key: string]: any };
    property(name: string, value: any): Workbook;
    property(properties: { [key: string]: any }): Workbook;
    properties(): ICoreProperties;
    toFileAsync(path: string, opts?: IOOptions): Promise<void>;
    cloneSheet(
        from: Sheet,
        name: string,
        indexOrBeforeSheet: Sheet | string | number
    ): Sheet;
}

export = XlsxPopulate;
export as namespace XlsxPopulate;
declare namespace XlsxPopulate {
    let MIME_TYPE: string;
    const FormulaError: FormulaError;
    const dateToNumber: (date: Date) => number;
    const fromBlankAsync: () => Promise<Workbook>;
    const fromDataAsync: (
        data: DataTypes | Promise<DataTypes>,
        opts?: IOOptions
    ) => Promise<Workbook>;
    const fromFileAsync: (path: string, opts?: IOOptions) => Promise<Workbook>;
    const numberToDate: (n: number) => Date;
}
