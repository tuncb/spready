import {
  formatWorkbookNumberDisplay,
  getSheetColumnCount,
  getSheetRowCount,
  isFormulaInput,
  parseCellReference,
  type WorkbookCellStyle,
  type FormulaErrorCode,
  type WorkbookState,
  type WorkbookSheet,
} from "./workbook-core";

export type CellKey = string;

type CellAddress = {
  sheetId: string;
  rowIndex: number;
  columnIndex: number;
};

type FormulaReferenceAddress = {
  sheetName?: string;
  rowIndex: number;
  columnIndex: number;
};

type BlankValue = {
  type: "blank";
};

type BooleanValue = {
  type: "boolean";
  value: boolean;
};

type ErrorValue = {
  type: "error";
  errorCode: FormulaErrorCode;
};

type NumberValue = {
  type: "number";
  value: number;
};

type TextValue = {
  type: "text";
  value: string;
};

type ScalarFormulaValue = BlankValue | BooleanValue | ErrorValue | NumberValue | TextValue;

type MaterializedRangeArea = CellAddress[][];

type RectangularRangeArea = {
  endColumn: number;
  endRow: number;
  sheetId: string;
  startColumn: number;
  startRow: number;
  type: "rectangular";
};

type RangeArea = MaterializedRangeArea | RectangularRangeArea;

type RangeValue = {
  type: "range";
  areas: RangeArea[];
  cells: RangeArea;
};

type FormulaValue = RangeValue | ScalarFormulaValue;

export interface CellEvaluation {
  input: string;
  display: string;
  isFormula: boolean;
  value: ScalarFormulaValue;
  errorCode?: FormulaErrorCode;
  dependencies: CellKey[];
}

export interface SheetEvaluationSnapshot {
  sheetId: string;
  workbookVersion: number;
  hasVolatileFunctions: boolean;
  cells: Map<CellKey, CellEvaluation>;
  dependents: Map<CellKey, Set<CellKey>>;
  precedents: Map<CellKey, Set<CellKey>>;
}

export interface FormulaEvaluationOptions {
  metrics?: FormulaEvaluationMetrics;
  now?: Date;
  parseCache?: FormulaParseCache;
  seedSnapshots?: Map<string, SheetEvaluationSnapshot>;
}

export interface FormulaEvaluationMetrics {
  cellsEvaluated: number;
  dependencyKeysRecorded: number;
  formulasParsed: number;
  rangeCellsMaterialized: number;
}

export type FormulaParseCache = Map<string, FormulaAst | "PARSE_ERROR">;

type FormulaToken =
  | { type: "comma" }
  | { type: "error"; errorCode: FormulaErrorCode }
  | { type: "bang" }
  | { type: "identifier"; value: string }
  | { type: "leftParen" }
  | { type: "number"; value: number }
  | {
      type: "operator";
      value: "+" | "-" | "*" | "/" | "^" | "&" | "=" | "<>" | "<" | "<=" | ">" | ">=" | ":" | "%";
    }
  | { type: "rightParen" }
  | { type: "sheetName"; value: string }
  | { type: "space" }
  | { type: "structuredReference"; value: string }
  | { type: "text"; value: string };

type FormulaAst =
  | {
      type: "binary";
      operator: "+" | "-" | "*" | "/" | "^" | "&" | "=" | "<>" | "<" | "<=" | ">" | ">=";
      left: FormulaAst;
      right: FormulaAst;
    }
  | { type: "function"; name: string; args: FormulaAst[] }
  | { type: "intersection"; left: FormulaAst; right: FormulaAst }
  | { type: "literal"; value: ScalarFormulaValue }
  | { type: "name"; name: string }
  | {
      type: "percent";
      operand: FormulaAst;
    }
  | { type: "range"; start: FormulaReferenceAddress; end: FormulaReferenceAddress }
  | { type: "reference"; sheetName?: string; rowIndex: number; columnIndex: number }
  | { type: "structuredReference"; tableName?: string; specifier: string }
  | { type: "union"; left: FormulaAst; right: FormulaAst }
  | { type: "unary"; operator: "+" | "-"; operand: FormulaAst };

type FunctionArgumentValue = {
  fromRange: boolean;
  values: ScalarFormulaValue[];
};

type FormulaFunctionHandler = (args: FormulaAst[], dependencies: Set<CellKey>) => FormulaValue;

type StructuredReferenceSelection = {
  columnEnd?: string;
  columnStart?: string;
  currentRow: boolean;
  dataOnly: boolean;
};

const BLANK_VALUE: BlankValue = {
  type: "blank",
};

const EMPTY_DEPENDENCIES: CellKey[] = [];
const MILLISECONDS_PER_DAY = 24 * 60 * 60 * 1000;
const EXCEL_1900_EPOCH_UTC = Date.UTC(1899, 11, 31);
const EXCEL_1900_LEAP_BUG_SERIAL = 60;
const EXCEL_MIN_SERIAL = 1;
const EXCEL_MAX_SERIAL = 2958465;
const NUMBER_LITERAL_PATTERN = /^[+-]?(?:\d+(?:\.\d+)?|\.\d+)(?:[Ee][+-]?\d+)?$/;
const ERROR_LITERALS: ReadonlyArray<[string, FormulaErrorCode]> = [
  ["#DIV/0!", "DIV0"],
  ["#NAME?", "NAME"],
  ["#NULL!", "NULL"],
  ["#VALUE!", "VALUE"],
  ["#REF!", "REF"],
  ["#NUM!", "NUM"],
  ["#N/A", "NA"],
];

export function createCellKey(sheetId: string, rowIndex: number, columnIndex: number): CellKey {
  return `${sheetId}:${rowIndex}:${columnIndex}`;
}

export function createFormulaEvaluationMetrics(): FormulaEvaluationMetrics {
  return {
    cellsEvaluated: 0,
    dependencyKeysRecorded: 0,
    formulasParsed: 0,
    rangeCellsMaterialized: 0,
  };
}

function cloneSeedEvaluationSnapshot(
  snapshot: SheetEvaluationSnapshot,
  workbookVersion: number,
): SheetEvaluationSnapshot {
  return {
    sheetId: snapshot.sheetId,
    workbookVersion,
    hasVolatileFunctions: snapshot.hasVolatileFunctions,
    cells: new Map(snapshot.cells),
    dependents: cloneCellKeySetMap(snapshot.dependents),
    precedents: cloneCellKeySetMap(snapshot.precedents),
  };
}

function cloneCellKeySetMap(map: Map<CellKey, Set<CellKey>>): Map<CellKey, Set<CellKey>> {
  return new Map([...map].map(([key, values]) => [key, new Set(values)]));
}

export function getCellEvaluation(
  snapshot: SheetEvaluationSnapshot,
  rowIndex: number,
  columnIndex: number,
): CellEvaluation {
  const evaluation = snapshot.cells.get(createCellKey(snapshot.sheetId, rowIndex, columnIndex));

  if (!evaluation) {
    throw new Error(`Cell ${rowIndex}:${columnIndex} is missing from the evaluation snapshot.`);
  }

  return evaluation;
}

export function compareCellEvaluationSortValues(
  left: CellEvaluation,
  right: CellEvaluation,
): number {
  const leftBlank = left.value.type === "blank";
  const rightBlank = right.value.type === "blank";

  if (leftBlank || rightBlank) {
    if (leftBlank === rightBlank) {
      return 0;
    }

    return leftBlank ? 1 : -1;
  }

  if (left.value.type === "number" && right.value.type === "number") {
    return left.value.value - right.value.value;
  }

  if (left.value.type === "boolean" && right.value.type === "boolean") {
    return Number(left.value.value) - Number(right.value.value);
  }

  if (left.value.type === "text" && right.value.type === "text") {
    return compareDisplayText(left.value.value, right.value.value);
  }

  if (left.value.type === "error" && right.value.type === "error") {
    return compareDisplayText(
      getErrorDisplay(left.value.errorCode),
      getErrorDisplay(right.value.errorCode),
    );
  }

  return compareDisplayText(left.display, right.display);
}

export function tokenizeFormula(input: string): FormulaToken[] {
  const expression = input.startsWith("=") ? input.slice(1) : input;
  const tokens: FormulaToken[] = [];
  let index = 0;

  while (index < expression.length) {
    const character = expression[index];

    if (/\s/.test(character)) {
      while (index < expression.length && /\s/.test(expression[index])) {
        index += 1;
      }

      tokens.push({ type: "space" });
      continue;
    }

    if (character === '"') {
      const { nextIndex, value } = readStringLiteral(expression, index);

      tokens.push({
        type: "text",
        value,
      });
      index = nextIndex;
      continue;
    }

    if (character === "'") {
      const { nextIndex, value } = readQuotedSheetName(expression, index);

      tokens.push({
        type: "sheetName",
        value,
      });
      index = nextIndex;
      continue;
    }

    if (character === "#") {
      const matchedError = ERROR_LITERALS.find(([literal]) =>
        expression
          .slice(index, index + literal.length)
          .toUpperCase()
          .startsWith(literal),
      );

      if (!matchedError) {
        throw new Error(`Unexpected token "${character}" in formula.`);
      }

      tokens.push({
        type: "error",
        errorCode: matchedError[1],
      });
      index += matchedError[0].length;
      continue;
    }

    const twoCharacterOperator = expression.slice(index, index + 2);

    if (
      twoCharacterOperator === "<>" ||
      twoCharacterOperator === "<=" ||
      twoCharacterOperator === ">="
    ) {
      tokens.push({
        type: "operator",
        value: twoCharacterOperator,
      });
      index += 2;
      continue;
    }

    if (
      character === "+" ||
      character === "-" ||
      character === "*" ||
      character === "/" ||
      character === "^" ||
      character === "&" ||
      character === "=" ||
      character === "<" ||
      character === ">" ||
      character === ":" ||
      character === "%"
    ) {
      tokens.push({
        type: "operator",
        value: character,
      });
      index += 1;
      continue;
    }

    if (character === ",") {
      tokens.push({ type: "comma" });
      index += 1;
      continue;
    }

    if (character === "!") {
      tokens.push({ type: "bang" });
      index += 1;
      continue;
    }

    if (character === "(") {
      tokens.push({ type: "leftParen" });
      index += 1;
      continue;
    }

    if (character === ")") {
      tokens.push({ type: "rightParen" });
      index += 1;
      continue;
    }

    if (character === "[") {
      const { nextIndex, value } = readStructuredReferenceSpecifier(expression, index);

      tokens.push({
        type: "structuredReference",
        value,
      });
      index = nextIndex;
      continue;
    }

    if (/[A-Za-z_\\]/.test(character)) {
      const identifierMatch = /^[A-Za-z_\\][A-Za-z0-9_.]*/.exec(expression.slice(index));

      if (!identifierMatch) {
        throw new Error(`Unexpected token "${character}" in formula.`);
      }

      tokens.push({
        type: "identifier",
        value: identifierMatch[0],
      });
      index += identifierMatch[0].length;
      continue;
    }

    const numberMatch = /^(?:\d+(?:\.\d+)?|\.\d+)(?:[Ee][+-]?\d+)?/.exec(expression.slice(index));

    if (numberMatch) {
      tokens.push({
        type: "number",
        value: Number(numberMatch[0]),
      });
      index += numberMatch[0].length;
      continue;
    }

    throw new Error(`Unexpected token "${character}" in formula.`);
  }

  return tokens;
}

export function parseFormula(input: string): FormulaAst {
  const tokens = tokenizeFormula(input);
  let index = 0;

  if (tokens.length === 0) {
    throw new Error("Formula cannot be empty.");
  }

  function skipSpaces() {
    while (tokens[index]?.type === "space") {
      index += 1;
    }
  }

  function getNonSpaceTokenIndex(startIndex: number): number {
    let tokenIndex = startIndex;

    while (tokens[tokenIndex]?.type === "space") {
      tokenIndex += 1;
    }

    return tokenIndex;
  }

  function canStartReferenceTerm(token: FormulaToken | undefined): boolean {
    return (
      token?.type === "identifier" ||
      token?.type === "leftParen" ||
      token?.type === "sheetName" ||
      token?.type === "structuredReference"
    );
  }

  function parseExpression(allowUnion = true): FormulaAst {
    let node = parseComparison();

    if (!allowUnion) {
      return node;
    }

    skipSpaces();

    while (tokens[index]?.type === "comma") {
      index += 1;
      const right = parseComparison();

      if (!isReferenceExpressionAst(node) || !isReferenceExpressionAst(right)) {
        throw new Error("Formula union operands must be references.");
      }

      node = {
        type: "union",
        left: node,
        right,
      };
      skipSpaces();
    }

    return node;
  }

  function parseComparison(): FormulaAst {
    let node = parseConcatenation();
    skipSpaces();
    let token = tokens[index];

    while (
      token?.type === "operator" &&
      (token.value === "=" ||
        token.value === "<>" ||
        token.value === "<" ||
        token.value === "<=" ||
        token.value === ">" ||
        token.value === ">=")
    ) {
      const operator = token.value;

      index += 1;
      node = {
        type: "binary",
        operator,
        left: node,
        right: parseConcatenation(),
      };
      skipSpaces();
      token = tokens[index];
    }

    return node;
  }

  function parseConcatenation(): FormulaAst {
    let node = parseAdditive();
    skipSpaces();
    let token = tokens[index];

    while (token?.type === "operator" && token.value === "&") {
      index += 1;
      node = {
        type: "binary",
        operator: "&",
        left: node,
        right: parseAdditive(),
      };
      skipSpaces();
      token = tokens[index];
    }

    return node;
  }

  function parseAdditive(): FormulaAst {
    let node = parseMultiplicative();
    skipSpaces();
    let token = tokens[index];

    while (token?.type === "operator" && (token.value === "+" || token.value === "-")) {
      const operator = token.value;

      index += 1;
      node = {
        type: "binary",
        operator,
        left: node,
        right: parseMultiplicative(),
      };
      skipSpaces();
      token = tokens[index];
    }

    return node;
  }

  function parseMultiplicative(): FormulaAst {
    let node = parsePower();
    skipSpaces();
    let token = tokens[index];

    while (token?.type === "operator" && (token.value === "*" || token.value === "/")) {
      const operator = token.value;

      index += 1;
      node = {
        type: "binary",
        operator,
        left: node,
        right: parsePower(),
      };
      skipSpaces();
      token = tokens[index];
    }

    return node;
  }

  function parsePower(): FormulaAst {
    let node = parsePercent();
    skipSpaces();
    let token = tokens[index];

    while (token?.type === "operator" && token.value === "^") {
      index += 1;
      node = {
        type: "binary",
        operator: "^",
        left: node,
        right: parsePercent(),
      };
      skipSpaces();
      token = tokens[index];
    }

    return node;
  }

  function parsePercent(): FormulaAst {
    let node = parseUnary();
    skipSpaces();
    let token = tokens[index];

    while (token?.type === "operator" && token.value === "%") {
      index += 1;
      node = {
        type: "percent",
        operand: node,
      };
      skipSpaces();
      token = tokens[index];
    }

    return node;
  }

  function parseUnary(): FormulaAst {
    skipSpaces();
    const token = tokens[index];

    if (token?.type === "operator" && (token.value === "+" || token.value === "-")) {
      index += 1;
      return {
        type: "unary",
        operator: token.value,
        operand: parseUnary(),
      };
    }

    return parseRange();
  }

  function parseRange(): FormulaAst {
    let node = parseRangeTerm();

    while (tokens[index]?.type === "space") {
      const rightTokenIndex = getNonSpaceTokenIndex(index);

      if (!canStartReferenceTerm(tokens[rightTokenIndex])) {
        index = rightTokenIndex;
        break;
      }

      index = rightTokenIndex;
      const right = parseRangeTerm();

      if (!isReferenceExpressionAst(node) || !isReferenceExpressionAst(right)) {
        throw new Error("Formula intersection operands must be references.");
      }

      node = {
        type: "intersection",
        left: node,
        right,
      };
    }

    return node;
  }

  function parseRangeTerm(): FormulaAst {
    const node = parsePrimary();
    const operatorIndex = getNonSpaceTokenIndex(index);
    const token = tokens[operatorIndex];

    if (token?.type !== "operator" || token.value !== ":") {
      return node;
    }

    if (node.type !== "reference") {
      throw new Error("Formula range start must be a cell reference.");
    }

    index = operatorIndex + 1;

    const endNode = parsePrimary();

    if (endNode.type !== "reference") {
      throw new Error("Formula range end must be a cell reference.");
    }

    return {
      type: "range",
      start: {
        sheetName: node.sheetName,
        rowIndex: node.rowIndex,
        columnIndex: node.columnIndex,
      },
      end: {
        sheetName: endNode.sheetName ?? node.sheetName,
        rowIndex: endNode.rowIndex,
        columnIndex: endNode.columnIndex,
      },
    };
  }

  function parsePrimary(): FormulaAst {
    skipSpaces();
    const token = tokens[index];

    if (!token) {
      throw new Error("Formula ended unexpectedly.");
    }

    if (token.type === "number") {
      index += 1;
      return {
        type: "literal",
        value: {
          type: "number",
          value: token.value,
        },
      };
    }

    if (token.type === "text") {
      index += 1;
      return {
        type: "literal",
        value: {
          type: "text",
          value: token.value,
        },
      };
    }

    if (token.type === "error") {
      index += 1;
      return {
        type: "literal",
        value: createErrorValue(token.errorCode),
      };
    }

    if (token.type === "identifier") {
      index += 1;

      const structuredReferenceToken = tokens[index];

      if (structuredReferenceToken?.type === "structuredReference") {
        index += 1;

        return {
          type: "structuredReference",
          tableName: token.value,
          specifier: structuredReferenceToken.value,
        };
      }

      if (tokens[index]?.type === "bang") {
        return parseSheetQualifiedReference(token.value);
      }

      if (tokens[index]?.type === "leftParen") {
        return parseFunctionCall(token.value);
      }

      const normalizedIdentifier = token.value.toUpperCase();

      if (normalizedIdentifier === "TRUE" || normalizedIdentifier === "FALSE") {
        return {
          type: "literal",
          value: {
            type: "boolean",
            value: normalizedIdentifier === "TRUE",
          },
        };
      }

      if (isCellReferenceIdentifier(token.value)) {
        const parsedReference = parseCellReference(token.value);

        return {
          type: "reference",
          rowIndex: parsedReference.rowIndex,
          columnIndex: parsedReference.columnIndex,
        };
      }

      return {
        type: "name",
        name: token.value,
      };
    }

    if (token.type === "structuredReference") {
      index += 1;

      return {
        type: "structuredReference",
        specifier: token.value,
      };
    }

    if (token.type === "sheetName") {
      index += 1;

      const structuredReferenceToken = tokens[index];

      if (structuredReferenceToken?.type === "structuredReference") {
        index += 1;

        return {
          type: "structuredReference",
          tableName: token.value,
          specifier: structuredReferenceToken.value,
        };
      }

      return parseSheetQualifiedReference(token.value);
    }

    if (token.type === "leftParen") {
      index += 1;
      const expression = parseExpression(true);
      skipSpaces();

      if (tokens[index]?.type !== "rightParen") {
        throw new Error("Formula is missing a closing parenthesis.");
      }

      index += 1;
      return expression;
    }

    throw new Error("Formula contains an unexpected token.");
  }

  function parseFunctionCall(name: string): FormulaAst {
    if (tokens[index]?.type !== "leftParen") {
      throw new Error("Formula function call is missing an opening parenthesis.");
    }

    index += 1;

    const args: FormulaAst[] = [];
    skipSpaces();

    if (tokens[index]?.type !== "rightParen") {
      let shouldContinue = true;

      while (shouldContinue) {
        args.push(parseExpression(false));
        skipSpaces();

        if (tokens[index]?.type === "comma") {
          index += 1;
          skipSpaces();
          continue;
        }

        shouldContinue = false;
      }
    }

    if (tokens[index]?.type !== "rightParen") {
      throw new Error("Formula function call is missing a closing parenthesis.");
    }

    index += 1;

    return {
      type: "function",
      name,
      args,
    };
  }

  function parseSheetQualifiedReference(sheetName: string): FormulaAst {
    if (tokens[index]?.type !== "bang") {
      throw new Error("Formula sheet reference is missing !.");
    }

    index += 1;

    const referenceToken = tokens[index];

    if (referenceToken?.type !== "identifier" || !isCellReferenceIdentifier(referenceToken.value)) {
      throw new Error("Formula sheet reference must point to a cell reference.");
    }

    index += 1;

    const parsedReference = parseCellReference(referenceToken.value);

    return {
      type: "reference",
      sheetName,
      rowIndex: parsedReference.rowIndex,
      columnIndex: parsedReference.columnIndex,
    };
  }

  const ast = parseExpression();
  skipSpaces();

  if (index !== tokens.length) {
    throw new Error("Formula contains unexpected trailing tokens.");
  }

  return ast;

  function isReferenceExpressionAst(ast: FormulaAst): boolean {
    return (
      ast.type === "intersection" ||
      ast.type === "range" ||
      ast.type === "reference" ||
      ast.type === "structuredReference" ||
      ast.type === "union"
    );
  }
}

export function evaluateSheet(
  sheet: WorkbookSheet,
  workbookVersion: number,
  options: FormulaEvaluationOptions = {},
): SheetEvaluationSnapshot {
  return evaluateWorkbookSheet(
    {
      sheets: [sheet],
      tables: [],
    },
    sheet.id,
    workbookVersion,
    options,
  );
}

export function evaluateWorkbookSheet(
  workbook: Pick<WorkbookState, "sheets"> & Partial<Pick<WorkbookState, "tables">>,
  sheetId: string,
  workbookVersion: number,
  options: FormulaEvaluationOptions = {},
): SheetEvaluationSnapshot {
  return (
    evaluateWorkbook(workbook, workbookVersion, options).get(sheetId) ??
    createMissingSheetSnapshot(sheetId, workbookVersion)
  );
}

export function evaluateWorkbook(
  workbook: Pick<WorkbookState, "sheets"> & Partial<Pick<WorkbookState, "tables">>,
  workbookVersion: number,
  options: FormulaEvaluationOptions = {},
): Map<string, SheetEvaluationSnapshot> {
  const metrics = options.metrics;
  const parseCache = options.parseCache;
  const sheetById = new Map(workbook.sheets.map((sheet) => [sheet.id, sheet]));
  const sheetIdByName = new Map(
    workbook.sheets.map((sheet) => [getSheetNameKey(sheet.name), sheet.id]),
  );
  const workbookTables = workbook.tables ?? [];
  const tableByName = new Map(
    workbookTables.map((table) => [getStructuredReferenceNameKey(table.name), table]),
  );
  const snapshots = new Map<string, SheetEvaluationSnapshot>(
    workbook.sheets.map((sheet) => {
      const seedSnapshot = options.seedSnapshots?.get(sheet.id);

      return [
        sheet.id,
        seedSnapshot
          ? cloneSeedEvaluationSnapshot(seedSnapshot, workbookVersion)
          : {
              sheetId: sheet.id,
              workbookVersion,
              hasVolatileFunctions: false,
              cells: new Map<CellKey, CellEvaluation>(),
              dependents: new Map<CellKey, Set<CellKey>>(),
              precedents: new Map<CellKey, Set<CellKey>>(),
            },
      ];
    }),
  );
  const evaluationStack: CellKey[] = [];
  const cycleCellKeys = new Set<CellKey>();
  const evaluationNow = options.now ? new Date(options.now.getTime()) : new Date();
  const formulaCellStack: CellAddress[] = [];
  const formulaDependencyReuseStack: boolean[] = [];
  const reusedDependencyCellKeys = new Set<CellKey>();
  let hasVolatileFunctions = false;

  function getInput(sheetId: string, rowIndex: number, columnIndex: number): string {
    const sheet = sheetById.get(sheetId);

    return sheet?.cells[rowIndex]?.[columnIndex] ?? "";
  }

  function getCellStyle(address: CellAddress): WorkbookCellStyle | undefined {
    const sheet = sheetById.get(address.sheetId);

    return sheet?.cellStyles[createSheetCellStyleKey(address.rowIndex, address.columnIndex)];
  }

  function recordDependencies(
    snapshot: SheetEvaluationSnapshot,
    cellKey: CellKey,
    dependencies: readonly CellKey[],
  ) {
    if (dependencies.length === 0) {
      return;
    }

    const precedentSet = new Set(dependencies);

    snapshot.precedents.set(cellKey, precedentSet);
    if (metrics) {
      metrics.dependencyKeysRecorded += precedentSet.size;
    }

    for (const dependencyKey of precedentSet) {
      const dependencySnapshot = snapshots.get(getCellKeySheetId(dependencyKey));

      if (!dependencySnapshot) {
        continue;
      }

      const dependentSet = dependencySnapshot.dependents.get(dependencyKey) ?? new Set<CellKey>();

      dependentSet.add(cellKey);
      dependencySnapshot.dependents.set(dependencyKey, dependentSet);
    }
  }

  function evaluateCell(address: CellAddress): CellEvaluation {
    const cellKey = createCellKey(address.sheetId, address.rowIndex, address.columnIndex);
    const snapshot = getSnapshot(address.sheetId);
    const cachedEvaluation = snapshot.cells.get(cellKey);

    if (cachedEvaluation) {
      return cachedEvaluation;
    }

    const stackIndex = evaluationStack.indexOf(cellKey);

    if (stackIndex >= 0) {
      for (const cycleKey of evaluationStack.slice(stackIndex)) {
        cycleCellKeys.add(cycleKey);
      }

      return createErrorEvaluation(
        getInput(address.sheetId, address.rowIndex, address.columnIndex),
        "CYCLE",
        EMPTY_DEPENDENCIES,
      );
    }

    evaluationStack.push(cellKey);
    if (metrics) {
      metrics.cellsEvaluated += 1;
    }

    try {
      const input = getInput(address.sheetId, address.rowIndex, address.columnIndex);
      let evaluation: CellEvaluation;

      if (!isFormulaInput(input)) {
        const value = parseRawCellValue(input);

        evaluation = {
          input,
          display: getDisplayForInputValue(input, value, getCellStyle(address)),
          isFormula: false,
          value,
          dependencies: EMPTY_DEPENDENCIES,
        };
      } else {
        evaluation = evaluateFormulaCell(input, cellKey, address);
      }

      snapshot.cells.set(cellKey, evaluation);
      if (!reusedDependencyCellKeys.has(cellKey)) {
        recordDependencies(snapshot, cellKey, evaluation.dependencies);
      }
      return evaluation;
    } finally {
      evaluationStack.pop();
    }
  }

  function evaluateFormulaCell(
    input: string,
    cellKey: CellKey,
    address: CellAddress,
  ): CellEvaluation {
    const ast = getFormulaAst(input);

    if (ast === "PARSE_ERROR") {
      return createErrorEvaluation(input, "PARSE", EMPTY_DEPENDENCIES);
    }

    const reusableDependencies = getSnapshot(address.sheetId).precedents.get(cellKey);
    const reuseDependencies = reusableDependencies !== undefined;
    const dependencies = new Set<CellKey>();
    let value: ScalarFormulaValue;

    formulaCellStack.push(address);
    formulaDependencyReuseStack.push(reuseDependencies);

    try {
      value = scalarizeFormulaValue(evaluateAst(ast, dependencies));
    } finally {
      formulaDependencyReuseStack.pop();
      formulaCellStack.pop();
    }

    const normalizedDependencies = reuseDependencies
      ? [...reusableDependencies]
      : [...dependencies];
    const errorCode = cycleCellKeys.has(cellKey)
      ? "CYCLE"
      : value.type === "error"
        ? value.errorCode
        : undefined;

    if (reuseDependencies) {
      reusedDependencyCellKeys.add(cellKey);
    }

    if (errorCode) {
      return createErrorEvaluation(input, errorCode, normalizedDependencies);
    }

    return {
      input,
      display: getDisplayForValue(value, getCellStyle(address)),
      isFormula: true,
      value,
      dependencies: normalizedDependencies,
    };
  }

  function shouldCollectFormulaDependencies(): boolean {
    return formulaDependencyReuseStack[formulaDependencyReuseStack.length - 1] !== true;
  }

  function getFormulaAst(input: string): FormulaAst | "PARSE_ERROR" {
    const cachedAst = parseCache?.get(input);

    if (cachedAst) {
      return cachedAst;
    }

    try {
      if (metrics) {
        metrics.formulasParsed += 1;
      }
      const ast = parseFormula(input);

      parseCache?.set(input, ast);
      return ast;
    } catch {
      parseCache?.set(input, "PARSE_ERROR");
      return "PARSE_ERROR";
    }
  }

  function evaluateAst(ast: FormulaAst, dependencies: Set<CellKey>): FormulaValue {
    switch (ast.type) {
      case "function":
        return evaluateFunctionCall(ast.name, ast.args, dependencies);
      case "intersection":
        return evaluateIntersection(ast, dependencies);
      case "literal":
        return ast.value;
      case "name":
        return createErrorValue("NAME");
      case "percent": {
        const numericOperand = coerceToNumber(evaluateAst(ast.operand, dependencies));

        if (numericOperand.type === "error") {
          return numericOperand;
        }

        return {
          type: "number",
          value: numericOperand.value / 100,
        };
      }
      case "range":
        return createRangeValue(
          resolveFormulaAddress(ast.start),
          resolveFormulaAddress(ast.end),
          dependencies,
        );
      case "reference": {
        const address = resolveFormulaAddress(ast);

        return createRangeValue(address, address, dependencies);
      }
      case "structuredReference":
        return evaluateStructuredReference(ast, dependencies);
      case "unary": {
        const numericOperand = coerceToNumber(evaluateAst(ast.operand, dependencies));

        if (numericOperand.type === "error") {
          return numericOperand;
        }

        return {
          type: "number",
          value: ast.operator === "-" ? -numericOperand.value : numericOperand.value,
        };
      }
      case "union":
        return evaluateUnion(ast, dependencies);
      case "binary":
        return evaluateBinaryOperation(ast, dependencies);
    }
  }

  function evaluateBinaryOperation(
    ast: Extract<FormulaAst, { type: "binary" }>,
    dependencies: Set<CellKey>,
  ): FormulaValue {
    if (ast.operator === "&") {
      const leftText = coerceToText(evaluateAst(ast.left, dependencies));

      if (leftText.type === "error") {
        return leftText;
      }

      const rightText = coerceToText(evaluateAst(ast.right, dependencies));

      if (rightText.type === "error") {
        return rightText;
      }

      return {
        type: "text",
        value: `${leftText.value}${rightText.value}`,
      };
    }

    if (
      ast.operator === "=" ||
      ast.operator === "<>" ||
      ast.operator === "<" ||
      ast.operator === "<=" ||
      ast.operator === ">" ||
      ast.operator === ">="
    ) {
      return compareFormulaValues(
        evaluateAst(ast.left, dependencies),
        evaluateAst(ast.right, dependencies),
        ast.operator,
      );
    }

    const leftValue = coerceToNumber(evaluateAst(ast.left, dependencies));

    if (leftValue.type === "error") {
      return leftValue;
    }

    const rightValue = coerceToNumber(evaluateAst(ast.right, dependencies));

    if (rightValue.type === "error") {
      return rightValue;
    }

    switch (ast.operator) {
      case "+":
        return {
          type: "number",
          value: leftValue.value + rightValue.value,
        };
      case "-":
        return {
          type: "number",
          value: leftValue.value - rightValue.value,
        };
      case "*":
        return {
          type: "number",
          value: leftValue.value * rightValue.value,
        };
      case "/":
        if (rightValue.value === 0) {
          return createErrorValue("DIV0");
        }

        return {
          type: "number",
          value: leftValue.value / rightValue.value,
        };
      case "^":
        return {
          type: "number",
          value: leftValue.value ** rightValue.value,
        };
    }
  }

  function evaluateIntersection(
    ast: Extract<FormulaAst, { type: "intersection" }>,
    dependencies: Set<CellKey>,
  ): RangeValue | ErrorValue {
    const leftValue = evaluateAst(ast.left, new Set<CellKey>());

    if (isErrorValue(leftValue)) {
      return leftValue;
    }

    if (leftValue.type !== "range") {
      return createErrorValue("VALUE");
    }

    const rightValue = evaluateAst(ast.right, new Set<CellKey>());

    if (isErrorValue(rightValue)) {
      return rightValue;
    }

    if (rightValue.type !== "range") {
      return createErrorValue("VALUE");
    }

    return intersectRangeValues(leftValue, rightValue, dependencies);
  }

  function evaluateUnion(
    ast: Extract<FormulaAst, { type: "union" }>,
    dependencies: Set<CellKey>,
  ): RangeValue | ErrorValue {
    const leftValue = evaluateAst(ast.left, new Set<CellKey>());

    if (isErrorValue(leftValue)) {
      return leftValue;
    }

    if (leftValue.type !== "range") {
      return createErrorValue("VALUE");
    }

    const rightValue = evaluateAst(ast.right, new Set<CellKey>());

    if (isErrorValue(rightValue)) {
      return rightValue;
    }

    if (rightValue.type !== "range") {
      return createErrorValue("VALUE");
    }

    return createUnionRangeValue(leftValue, rightValue, dependencies);
  }

  function createRangeValue(
    start: CellAddress | ErrorValue,
    end: CellAddress | ErrorValue,
    dependencies: Set<CellKey>,
  ): RangeValue | ErrorValue {
    if (isErrorValue(start)) {
      return start;
    }

    if (isErrorValue(end)) {
      return end;
    }

    if (start.sheetId !== end.sheetId) {
      return createErrorValue("REF");
    }

    const sheet = sheetById.get(start.sheetId);

    if (!sheet) {
      return createErrorValue("REF");
    }

    const rowCount = getSheetRowCount(sheet);
    const columnCount = getSheetColumnCount(sheet);

    if (
      start.rowIndex < 0 ||
      start.rowIndex >= rowCount ||
      start.columnIndex < 0 ||
      start.columnIndex >= columnCount ||
      end.rowIndex < 0 ||
      end.rowIndex >= rowCount ||
      end.columnIndex < 0 ||
      end.columnIndex >= columnCount
    ) {
      return createErrorValue("REF");
    }

    const startRow = Math.min(start.rowIndex, end.rowIndex);
    const endRow = Math.max(start.rowIndex, end.rowIndex);
    const startColumn = Math.min(start.columnIndex, end.columnIndex);
    const endColumn = Math.max(start.columnIndex, end.columnIndex);
    const cellsInRange = createRectangularRangeArea(
      start.sheetId,
      startRow,
      endRow,
      startColumn,
      endColumn,
    );

    recordRangeDependencies([cellsInRange], dependencies);

    return {
      type: "range",
      areas: [cellsInRange],
      cells: cellsInRange,
    };
  }

  function createRectangularRangeArea(
    sheetId: string,
    startRow: number,
    endRow: number,
    startColumn: number,
    endColumn: number,
  ): RectangularRangeArea {
    return {
      endColumn,
      endRow,
      sheetId,
      startColumn,
      startRow,
      type: "rectangular",
    };
  }

  function createSingleCellRangeValue(cellAddress: CellAddress): RangeValue {
    const area = createRectangularRangeArea(
      cellAddress.sheetId,
      cellAddress.rowIndex,
      cellAddress.rowIndex,
      cellAddress.columnIndex,
      cellAddress.columnIndex,
    );

    return {
      type: "range",
      areas: [area],
      cells: area,
    };
  }

  function isRectangularRangeArea(area: RangeArea): area is RectangularRangeArea {
    return !Array.isArray(area);
  }

  function getRangeAreaHeight(area: RangeArea): number {
    return isRectangularRangeArea(area) ? area.endRow - area.startRow + 1 : area.length;
  }

  function getRangeAreaWidth(area: RangeArea): number {
    return isRectangularRangeArea(area)
      ? area.endColumn - area.startColumn + 1
      : (area[0]?.length ?? 0);
  }

  function getRangeAreaCell(
    area: RangeArea,
    rowOffset: number,
    columnOffset: number,
  ): CellAddress | undefined {
    if (isRectangularRangeArea(area)) {
      const rowIndex = area.startRow + rowOffset;
      const columnIndex = area.startColumn + columnOffset;

      if (
        rowIndex < area.startRow ||
        rowIndex > area.endRow ||
        columnIndex < area.startColumn ||
        columnIndex > area.endColumn
      ) {
        return undefined;
      }

      return {
        sheetId: area.sheetId,
        rowIndex,
        columnIndex,
      };
    }

    return area[rowOffset]?.[columnOffset];
  }

  function* iterateRangeAreaCells(area: RangeArea): Generator<CellAddress> {
    if (isRectangularRangeArea(area)) {
      for (let rowIndex = area.startRow; rowIndex <= area.endRow; rowIndex += 1) {
        for (let columnIndex = area.startColumn; columnIndex <= area.endColumn; columnIndex += 1) {
          yield {
            sheetId: area.sheetId,
            rowIndex,
            columnIndex,
          };
        }
      }

      return;
    }

    for (const row of area) {
      for (const cellAddress of row) {
        yield cellAddress;
      }
    }
  }

  function materializeRangeArea(area: RangeArea): MaterializedRangeArea {
    if (!isRectangularRangeArea(area)) {
      return area;
    }

    const rowCount = getRangeAreaHeight(area);
    const columnCount = getRangeAreaWidth(area);

    if (metrics) {
      metrics.rangeCellsMaterialized += rowCount * columnCount;
    }

    return Array.from({ length: rowCount }, (_rowValue, rowOffset) =>
      Array.from({ length: columnCount }, (_columnValue, columnOffset) => ({
        sheetId: area.sheetId,
        rowIndex: area.startRow + rowOffset,
        columnIndex: area.startColumn + columnOffset,
      })),
    );
  }

  function intersectRangeValues(
    left: RangeValue,
    right: RangeValue,
    dependencies: Set<CellKey>,
  ): RangeValue | ErrorValue {
    const intersectedAreas: RangeArea[] = [];

    for (const leftArea of left.areas) {
      for (const rightArea of right.areas) {
        const rightCellKeys = new Set(
          [...iterateRangeAreaCells(rightArea)].map((cellAddress) =>
            createCellKey(cellAddress.sheetId, cellAddress.rowIndex, cellAddress.columnIndex),
          ),
        );
        const intersectedArea = materializeRangeArea(leftArea)
          .map((row) => {
            return row.filter((cellAddress) => {
              return rightCellKeys.has(
                createCellKey(cellAddress.sheetId, cellAddress.rowIndex, cellAddress.columnIndex),
              );
            });
          })
          .filter((row) => row.length > 0);

        if (intersectedArea.length > 0) {
          intersectedAreas.push(intersectedArea);
        }
      }
    }

    if (intersectedAreas.length === 0) {
      return createErrorValue("NULL");
    }

    recordRangeDependencies(intersectedAreas, dependencies);

    return {
      type: "range",
      areas: intersectedAreas,
      cells: intersectedAreas[0],
    };
  }

  function createUnionRangeValue(
    left: RangeValue,
    right: RangeValue,
    dependencies: Set<CellKey>,
  ): RangeValue {
    const areas = [...left.areas, ...right.areas];

    recordRangeDependencies(areas, dependencies);

    return {
      type: "range",
      areas,
      cells: areas[0],
    };
  }

  function evaluateStructuredReference(
    ast: Extract<FormulaAst, { type: "structuredReference" }>,
    dependencies: Set<CellKey>,
  ): RangeValue | ErrorValue {
    const selection = parseStructuredReferenceSelection(ast.specifier);

    if (isErrorValue(selection)) {
      return selection;
    }

    const currentCell = getCurrentFormulaCell();

    if (isErrorValue(currentCell)) {
      return currentCell;
    }

    const table =
      ast.tableName === undefined
        ? workbookTables.find((entry) => workbookTableContainsCell(entry, currentCell))
        : tableByName.get(getStructuredReferenceNameKey(ast.tableName));

    if (!table) {
      return createErrorValue("REF");
    }

    const sheet = sheetById.get(table.range.sheetId);

    if (!sheet) {
      return createErrorValue("REF");
    }

    const headerOffset = table.hasHeaderRow ? 1 : 0;
    const bodyStartRow = table.range.startRow + headerOffset;
    const bodyRowCount = table.range.rowCount - headerOffset;
    const bodyEndRow = bodyStartRow + bodyRowCount - 1;

    if (bodyRowCount <= 0) {
      return createErrorValue("REF");
    }

    const startColumn =
      selection.columnStart === undefined
        ? table.range.startColumn
        : resolveStructuredReferenceColumn(sheet, table, selection.columnStart);

    if (isErrorValue(startColumn)) {
      return startColumn;
    }

    const endColumn =
      selection.columnEnd === undefined
        ? startColumn
        : resolveStructuredReferenceColumn(sheet, table, selection.columnEnd);

    if (isErrorValue(endColumn)) {
      return endColumn;
    }

    const currentRow = currentCell.rowIndex;
    const startRow = selection.currentRow ? currentRow : bodyStartRow;
    const endRow = selection.currentRow ? currentRow : bodyEndRow;

    if (selection.currentRow && (currentRow < bodyStartRow || currentRow > bodyEndRow)) {
      return createErrorValue("REF");
    }

    return createRangeValue(
      {
        sheetId: table.range.sheetId,
        rowIndex: startRow,
        columnIndex: startColumn,
      },
      {
        sheetId: table.range.sheetId,
        rowIndex: endRow,
        columnIndex: endColumn,
      },
      dependencies,
    );
  }

  function recordRangeDependencies(areas: readonly RangeArea[], dependencies: Set<CellKey>) {
    if (!shouldCollectFormulaDependencies()) {
      return;
    }

    for (const area of areas) {
      for (const cellAddress of iterateRangeAreaCells(area)) {
        dependencies.add(
          createCellKey(cellAddress.sheetId, cellAddress.rowIndex, cellAddress.columnIndex),
        );
      }
    }
  }

  function isMultiAreaRange(value: RangeValue): boolean {
    return value.areas.length > 1;
  }

  function getOnlyRangeCell(value: RangeValue): CellAddress | undefined {
    let onlyCell: CellAddress | undefined;

    for (const area of value.areas) {
      for (const cellAddress of iterateRangeAreaCells(area)) {
        if (onlyCell) {
          return undefined;
        }

        onlyCell = cellAddress;
      }
    }

    return onlyCell;
  }

  function scalarizeFormulaValue(value: FormulaValue): ScalarFormulaValue {
    if (value.type !== "range") {
      return value;
    }

    const referencedCell = getOnlyRangeCell(value);

    if (!referencedCell) {
      return createErrorValue("VALUE");
    }

    return evaluateCell(referencedCell).value;
  }

  function coerceToNumber(value: FormulaValue): NumberValue | ErrorValue {
    const scalarValue = scalarizeFormulaValue(value);

    switch (scalarValue.type) {
      case "blank":
        return {
          type: "number",
          value: 0,
        };
      case "boolean":
        return {
          type: "number",
          value: scalarValue.value ? 1 : 0,
        };
      case "error":
        return scalarValue;
      case "number":
        return scalarValue;
      case "text": {
        const parsedNumeric = parseNumericLiteral(scalarValue.value);

        if (parsedNumeric === undefined) {
          return createErrorValue("VALUE");
        }

        return {
          type: "number",
          value: parsedNumeric,
        };
      }
    }
  }

  function coerceToText(value: FormulaValue): TextValue | ErrorValue {
    const scalarValue = scalarizeFormulaValue(value);

    switch (scalarValue.type) {
      case "blank":
        return {
          type: "text",
          value: "",
        };
      case "boolean":
        return {
          type: "text",
          value: scalarValue.value ? "TRUE" : "FALSE",
        };
      case "error":
        return scalarValue;
      case "number":
        return {
          type: "text",
          value: formatNumericDisplay(scalarValue.value),
        };
      case "text":
        return scalarValue;
    }
  }

  function compareFormulaValues(
    left: FormulaValue,
    right: FormulaValue,
    operator: "=" | "<>" | "<" | "<=" | ">" | ">=",
  ): BooleanValue | ErrorValue {
    const normalizedOperands = normalizeComparableOperands(left, right);

    if ("errorCode" in normalizedOperands) {
      return normalizedOperands;
    }

    const comparison = compareScalarValues(normalizedOperands.left, normalizedOperands.right);

    switch (operator) {
      case "=":
        return {
          type: "boolean",
          value: comparison === 0,
        };
      case "<>":
        return {
          type: "boolean",
          value: comparison !== 0,
        };
      case "<":
        return {
          type: "boolean",
          value: comparison < 0,
        };
      case "<=":
        return {
          type: "boolean",
          value: comparison <= 0,
        };
      case ">":
        return {
          type: "boolean",
          value: comparison > 0,
        };
      case ">=":
        return {
          type: "boolean",
          value: comparison >= 0,
        };
    }
  }

  function normalizeComparableOperands(
    left: FormulaValue,
    right: FormulaValue,
  ): { left: ScalarFormulaValue; right: ScalarFormulaValue } | ErrorValue {
    const leftScalar = scalarizeFormulaValue(left);

    if (leftScalar.type === "error") {
      return leftScalar;
    }

    const rightScalar = scalarizeFormulaValue(right);

    if (rightScalar.type === "error") {
      return rightScalar;
    }

    const normalizedLeft =
      leftScalar.type === "blank" ? coerceBlankForComparison(rightScalar) : leftScalar;
    const normalizedRight =
      rightScalar.type === "blank" ? coerceBlankForComparison(leftScalar) : rightScalar;

    return {
      left: normalizedLeft,
      right: normalizedRight,
    };
  }

  function coerceBlankForComparison(other: ScalarFormulaValue): ScalarFormulaValue {
    switch (other.type) {
      case "blank":
        return BLANK_VALUE;
      case "boolean":
        return {
          type: "boolean",
          value: false,
        };
      case "error":
        return other;
      case "number":
        return {
          type: "number",
          value: 0,
        };
      case "text":
        return {
          type: "text",
          value: "",
        };
    }
  }

  function compareScalarValues(left: ScalarFormulaValue, right: ScalarFormulaValue): number {
    if (left.type === "boolean" && right.type === "boolean") {
      return Number(left.value) - Number(right.value);
    }

    const leftNumber = toComparableNumber(left);
    const rightNumber = toComparableNumber(right);

    if (leftNumber !== undefined && rightNumber !== undefined) {
      if (leftNumber === rightNumber) {
        return 0;
      }

      return leftNumber < rightNumber ? -1 : 1;
    }

    const leftText = getComparableText(left);
    const rightText = getComparableText(right);

    if (leftText === rightText) {
      return 0;
    }

    return leftText < rightText ? -1 : 1;
  }

  function toComparableNumber(value: ScalarFormulaValue): number | undefined {
    switch (value.type) {
      case "blank":
        return 0;
      case "boolean":
        return value.value ? 1 : 0;
      case "error":
        return undefined;
      case "number":
        return value.value;
      case "text":
        return parseNumericLiteral(value.value);
    }
  }

  function getComparableText(value: ScalarFormulaValue): string {
    switch (value.type) {
      case "blank":
        return "";
      case "boolean":
        return value.value ? "TRUE" : "FALSE";
      case "error":
        return getErrorDisplay(value.errorCode);
      case "number":
        return formatNumericDisplay(value.value);
      case "text":
        return value.value.toUpperCase();
    }
  }

  function coerceToBoolean(value: FormulaValue): BooleanValue | ErrorValue {
    const scalarValue = scalarizeFormulaValue(value);

    switch (scalarValue.type) {
      case "blank":
        return {
          type: "boolean",
          value: false,
        };
      case "boolean":
        return scalarValue;
      case "error":
        return scalarValue;
      case "number":
        return {
          type: "boolean",
          value: scalarValue.value !== 0,
        };
      case "text": {
        const normalizedText = scalarValue.value.trim().toUpperCase();

        if (normalizedText === "TRUE" || normalizedText === "FALSE") {
          return {
            type: "boolean",
            value: normalizedText === "TRUE",
          };
        }

        const parsedNumeric = parseNumericLiteral(normalizedText);

        if (parsedNumeric !== undefined) {
          return {
            type: "boolean",
            value: parsedNumeric !== 0,
          };
        }

        return createErrorValue("VALUE");
      }
    }
  }

  function getScalarArgument(arg: FormulaAst, dependencies: Set<CellKey>): ScalarFormulaValue {
    return scalarizeFormulaValue(evaluateAst(arg, dependencies));
  }

  function getFunctionArgumentValue(
    arg: FormulaAst,
    dependencies: Set<CellKey>,
  ): FunctionArgumentValue | ErrorValue {
    const value = evaluateAst(arg, dependencies);

    if (value.type !== "range") {
      const scalarValue = scalarizeFormulaValue(value);

      if (scalarValue.type === "error") {
        return scalarValue;
      }

      return {
        fromRange: false,
        values: [scalarValue],
      };
    }

    const flattenedValues = flattenRangeCells(value);

    if (isErrorValue(flattenedValues)) {
      return flattenedValues;
    }

    return {
      fromRange: true,
      values: flattenedValues,
    };
  }

  function getFunctionArgumentValues(
    args: FormulaAst[],
    dependencies: Set<CellKey>,
  ): FunctionArgumentValue[] | ErrorValue {
    const argumentValues: FunctionArgumentValue[] = [];

    for (const arg of args) {
      const argumentValue = getFunctionArgumentValue(arg, dependencies);

      if (isErrorValue(argumentValue)) {
        return argumentValue;
      }

      argumentValues.push(argumentValue);
    }

    return argumentValues;
  }

  function flattenRangeCells(rangeValue: RangeValue): ScalarFormulaValue[] | ErrorValue {
    const flattenedValues: ScalarFormulaValue[] = [];

    for (const area of rangeValue.areas) {
      for (const cellAddress of iterateRangeAreaCells(area)) {
        const cellValue = evaluateCell(cellAddress).value;

        if (cellValue.type === "error") {
          return cellValue;
        }

        flattenedValues.push(cellValue);
      }
    }

    return flattenedValues;
  }

  function getCurrentFormulaCell(): CellAddress | ErrorValue {
    const currentCell = formulaCellStack[formulaCellStack.length - 1];

    if (!currentCell) {
      return createErrorValue("VALUE");
    }

    return currentCell;
  }

  function getRangeArgument(arg: FormulaAst, dependencies: Set<CellKey>): RangeValue | ErrorValue {
    const value = evaluateAst(arg, dependencies);

    if (value.type !== "range") {
      return createErrorValue("VALUE");
    }

    if (isMultiAreaRange(value)) {
      return createErrorValue("VALUE");
    }

    return value;
  }

  function getFirstRangeCell(rangeValue: RangeValue): CellAddress | ErrorValue {
    const firstCell = getRangeAreaCell(rangeValue.cells, 0, 0);

    if (!firstCell) {
      return createErrorValue("REF");
    }

    return firstCell;
  }

  function getVectorAddresses(rangeValue: RangeValue): CellAddress[] | ErrorValue {
    const height = getRangeAreaHeight(rangeValue.cells);
    const width = getRangeAreaWidth(rangeValue.cells);

    if (height === 0 || width === 0) {
      return createErrorValue("REF");
    }

    if (height === 1) {
      return Array.from({ length: width }, (_value, columnOffset) => {
        const cell = getRangeAreaCell(rangeValue.cells, 0, columnOffset);

        if (!cell) {
          throw new Error("Range vector cell was unexpectedly missing.");
        }

        return cell;
      });
    }

    if (width === 1) {
      return Array.from({ length: height }, (_value, rowOffset) => {
        const cell = getRangeAreaCell(rangeValue.cells, rowOffset, 0);

        if (!cell) {
          throw new Error("Range vector cell was unexpectedly missing.");
        }

        return cell;
      });
    }

    return createErrorValue("VALUE");
  }

  function expectArgumentCount(
    args: FormulaAst[],
    minCount: number,
    maxCount = minCount,
  ): ErrorValue | undefined {
    if (args.length < minCount || args.length > maxCount) {
      return createErrorValue("VALUE");
    }

    return undefined;
  }

  function getDirectAggregateNumber(
    value: ScalarFormulaValue,
  ): NumberValue | ErrorValue | undefined {
    switch (value.type) {
      case "blank":
        return undefined;
      case "boolean":
        return {
          type: "number",
          value: value.value ? 1 : 0,
        };
      case "error":
        return value;
      case "number":
        return value;
      case "text": {
        const parsedNumeric = parseNumericLiteral(value.value);

        if (parsedNumeric === undefined) {
          return createErrorValue("VALUE");
        }

        return {
          type: "number",
          value: parsedNumeric,
        };
      }
    }
  }

  function collectAggregateNumbers(
    args: FormulaAst[],
    dependencies: Set<CellKey>,
  ): number[] | ErrorValue {
    const argumentValues = getFunctionArgumentValues(args, dependencies);

    if (isErrorValue(argumentValues)) {
      return argumentValues;
    }

    const numericValues: number[] = [];

    for (const argumentValue of argumentValues) {
      if (argumentValue.fromRange) {
        for (const value of argumentValue.values) {
          if (value.type === "number") {
            numericValues.push(value.value);
          }
        }

        continue;
      }

      const directNumber = getDirectAggregateNumber(argumentValue.values[0]);

      if (directNumber?.type === "error") {
        return directNumber;
      }

      if (directNumber) {
        numericValues.push(directNumber.value);
      }
    }

    return numericValues;
  }

  function evaluateFunctionCall(
    name: string,
    args: FormulaAst[],
    dependencies: Set<CellKey>,
  ): FormulaValue {
    const handler = functionRegistry.get(name.toUpperCase());

    if (!handler) {
      return createErrorValue("NAME");
    }

    return handler(args, dependencies);
  }

  function evaluateSum(args: FormulaAst[], dependencies: Set<CellKey>): FormulaValue {
    const numericValues = collectAggregateNumbers(args, dependencies);

    if (isErrorValue(numericValues)) {
      return numericValues;
    }

    return {
      type: "number",
      value: numericValues.reduce((total, value) => total + value, 0),
    };
  }

  function evaluateProduct(args: FormulaAst[], dependencies: Set<CellKey>): FormulaValue {
    const numericValues = collectAggregateNumbers(args, dependencies);

    if (isErrorValue(numericValues)) {
      return numericValues;
    }

    if (numericValues.length === 0) {
      return {
        type: "number",
        value: 0,
      };
    }

    return {
      type: "number",
      value: numericValues.reduce((total, value) => total * value, 1),
    };
  }

  function evaluateMin(args: FormulaAst[], dependencies: Set<CellKey>): FormulaValue {
    const numericValues = collectAggregateNumbers(args, dependencies);

    if (isErrorValue(numericValues)) {
      return numericValues;
    }

    return {
      type: "number",
      value: numericValues.length === 0 ? 0 : Math.min(...numericValues),
    };
  }

  function evaluateMax(args: FormulaAst[], dependencies: Set<CellKey>): FormulaValue {
    const numericValues = collectAggregateNumbers(args, dependencies);

    if (isErrorValue(numericValues)) {
      return numericValues;
    }

    return {
      type: "number",
      value: numericValues.length === 0 ? 0 : Math.max(...numericValues),
    };
  }

  function evaluateAverage(args: FormulaAst[], dependencies: Set<CellKey>): FormulaValue {
    const numericValues = collectAggregateNumbers(args, dependencies);

    if (isErrorValue(numericValues)) {
      return numericValues;
    }

    if (numericValues.length === 0) {
      return createErrorValue("DIV0");
    }

    return {
      type: "number",
      value: numericValues.reduce((total, value) => total + value, 0) / numericValues.length,
    };
  }

  function evaluateCount(args: FormulaAst[], dependencies: Set<CellKey>): FormulaValue {
    const argumentValues = getFunctionArgumentValues(args, dependencies);

    if (isErrorValue(argumentValues)) {
      return argumentValues;
    }

    let count = 0;

    for (const argumentValue of argumentValues) {
      if (argumentValue.fromRange) {
        for (const value of argumentValue.values) {
          if (value.type === "number") {
            count += 1;
          }
        }

        continue;
      }

      const directNumber = getDirectAggregateNumber(argumentValue.values[0]);

      if (directNumber?.type === "error") {
        return directNumber;
      }

      if (directNumber) {
        count += 1;
      }
    }

    return {
      type: "number",
      value: count,
    };
  }

  function evaluateCountA(args: FormulaAst[], dependencies: Set<CellKey>): FormulaValue {
    const argumentValues = getFunctionArgumentValues(args, dependencies);

    if (isErrorValue(argumentValues)) {
      return argumentValues;
    }

    let count = 0;

    for (const argumentValue of argumentValues) {
      for (const value of argumentValue.values) {
        if (value.type !== "blank") {
          count += 1;
        }
      }
    }

    return {
      type: "number",
      value: count,
    };
  }

  function evaluateAbs(args: FormulaAst[], dependencies: Set<CellKey>): FormulaValue {
    const argumentError = expectArgumentCount(args, 1);

    if (argumentError) {
      return argumentError;
    }

    const numericValue = coerceToNumber(evaluateAst(args[0], dependencies));

    if (numericValue.type === "error") {
      return numericValue;
    }

    return {
      type: "number",
      value: Math.abs(numericValue.value),
    };
  }

  function evaluateRound(args: FormulaAst[], dependencies: Set<CellKey>): FormulaValue {
    const argumentError = expectArgumentCount(args, 2);

    if (argumentError) {
      return argumentError;
    }

    const value = coerceToNumber(evaluateAst(args[0], dependencies));

    if (value.type === "error") {
      return value;
    }

    const digits = coerceToNumber(evaluateAst(args[1], dependencies));

    if (digits.type === "error") {
      return digits;
    }

    const precision = Math.trunc(digits.value);
    const factor = 10 ** precision;

    return {
      type: "number",
      value: Math.round(value.value * factor) / factor,
    };
  }

  function evaluateInt(args: FormulaAst[], dependencies: Set<CellKey>): FormulaValue {
    const argumentError = expectArgumentCount(args, 1);

    if (argumentError) {
      return argumentError;
    }

    const value = coerceToNumber(evaluateAst(args[0], dependencies));

    if (value.type === "error") {
      return value;
    }

    return {
      type: "number",
      value: Math.floor(value.value),
    };
  }

  function evaluateMod(args: FormulaAst[], dependencies: Set<CellKey>): FormulaValue {
    const argumentError = expectArgumentCount(args, 2);

    if (argumentError) {
      return argumentError;
    }

    const value = coerceToNumber(evaluateAst(args[0], dependencies));

    if (value.type === "error") {
      return value;
    }

    const divisor = coerceToNumber(evaluateAst(args[1], dependencies));

    if (divisor.type === "error") {
      return divisor;
    }

    if (divisor.value === 0) {
      return createErrorValue("DIV0");
    }

    let result = value.value % divisor.value;

    if (result !== 0 && result < 0 !== divisor.value < 0) {
      result += divisor.value;
    }

    return {
      type: "number",
      value: result,
    };
  }

  function evaluatePower(args: FormulaAst[], dependencies: Set<CellKey>): FormulaValue {
    const argumentError = expectArgumentCount(args, 2);

    if (argumentError) {
      return argumentError;
    }

    const base = coerceToNumber(evaluateAst(args[0], dependencies));

    if (base.type === "error") {
      return base;
    }

    const exponent = coerceToNumber(evaluateAst(args[1], dependencies));

    if (exponent.type === "error") {
      return exponent;
    }

    const result = base.value ** exponent.value;

    if (Number.isNaN(result)) {
      return createErrorValue("NUM");
    }

    return {
      type: "number",
      value: result,
    };
  }

  function evaluateSqrt(args: FormulaAst[], dependencies: Set<CellKey>): FormulaValue {
    const argumentError = expectArgumentCount(args, 1);

    if (argumentError) {
      return argumentError;
    }

    const value = coerceToNumber(evaluateAst(args[0], dependencies));

    if (value.type === "error") {
      return value;
    }

    if (value.value < 0) {
      return createErrorValue("NUM");
    }

    return {
      type: "number",
      value: Math.sqrt(value.value),
    };
  }

  function evaluateLn(args: FormulaAst[], dependencies: Set<CellKey>): FormulaValue {
    const argumentError = expectArgumentCount(args, 1);

    if (argumentError) {
      return argumentError;
    }

    const value = coerceToNumber(evaluateAst(args[0], dependencies));

    if (value.type === "error") {
      return value;
    }

    if (value.value <= 0) {
      return createErrorValue("NUM");
    }

    return {
      type: "number",
      value: Math.log(value.value),
    };
  }

  function evaluateTrue(args: FormulaAst[]): FormulaValue {
    const argumentError = expectArgumentCount(args, 0);

    if (argumentError) {
      return argumentError;
    }

    return {
      type: "boolean",
      value: true,
    };
  }

  function evaluateFalse(args: FormulaAst[]): FormulaValue {
    const argumentError = expectArgumentCount(args, 0);

    if (argumentError) {
      return argumentError;
    }

    return {
      type: "boolean",
      value: false,
    };
  }

  function evaluateAnd(args: FormulaAst[], dependencies: Set<CellKey>): FormulaValue {
    if (args.length === 0) {
      return createErrorValue("VALUE");
    }

    const argumentValues = getFunctionArgumentValues(args, dependencies);

    if (isErrorValue(argumentValues)) {
      return argumentValues;
    }

    let result = true;

    for (const argumentValue of argumentValues) {
      for (const value of argumentValue.values) {
        const booleanValue = coerceToBoolean(value);

        if (booleanValue.type === "error") {
          return booleanValue;
        }

        result = result && booleanValue.value;
      }
    }

    return {
      type: "boolean",
      value: result,
    };
  }

  function evaluateOr(args: FormulaAst[], dependencies: Set<CellKey>): FormulaValue {
    if (args.length === 0) {
      return createErrorValue("VALUE");
    }

    const argumentValues = getFunctionArgumentValues(args, dependencies);

    if (isErrorValue(argumentValues)) {
      return argumentValues;
    }

    let result = false;

    for (const argumentValue of argumentValues) {
      for (const value of argumentValue.values) {
        const booleanValue = coerceToBoolean(value);

        if (booleanValue.type === "error") {
          return booleanValue;
        }

        result = result || booleanValue.value;
      }
    }

    return {
      type: "boolean",
      value: result,
    };
  }

  function evaluateNot(args: FormulaAst[], dependencies: Set<CellKey>): FormulaValue {
    const argumentError = expectArgumentCount(args, 1);

    if (argumentError) {
      return argumentError;
    }

    const booleanValue = coerceToBoolean(evaluateAst(args[0], dependencies));

    if (booleanValue.type === "error") {
      return booleanValue;
    }

    return {
      type: "boolean",
      value: !booleanValue.value,
    };
  }

  function evaluateIf(args: FormulaAst[], dependencies: Set<CellKey>): FormulaValue {
    const argumentError = expectArgumentCount(args, 2, 3);

    if (argumentError) {
      return argumentError;
    }

    const testValue = coerceToBoolean(evaluateAst(args[0], dependencies));

    if (testValue.type === "error") {
      return testValue;
    }

    if (testValue.value) {
      return evaluateAst(args[1], dependencies);
    }

    if (args[2]) {
      return evaluateAst(args[2], dependencies);
    }

    return {
      type: "boolean",
      value: false,
    };
  }

  function evaluateIfError(args: FormulaAst[], dependencies: Set<CellKey>): FormulaValue {
    const argumentError = expectArgumentCount(args, 2);

    if (argumentError) {
      return argumentError;
    }

    const firstValue = scalarizeFormulaValue(evaluateAst(args[0], dependencies));

    if (firstValue.type === "error") {
      return evaluateAst(args[1], dependencies);
    }

    return firstValue;
  }

  function evaluateLen(args: FormulaAst[], dependencies: Set<CellKey>): FormulaValue {
    const argumentError = expectArgumentCount(args, 1);

    if (argumentError) {
      return argumentError;
    }

    const textValue = coerceToText(evaluateAst(args[0], dependencies));

    if (textValue.type === "error") {
      return textValue;
    }

    return {
      type: "number",
      value: textValue.value.length,
    };
  }

  function evaluateLeft(args: FormulaAst[], dependencies: Set<CellKey>): FormulaValue {
    const argumentError = expectArgumentCount(args, 1, 2);

    if (argumentError) {
      return argumentError;
    }

    const textValue = coerceToText(evaluateAst(args[0], dependencies));

    if (textValue.type === "error") {
      return textValue;
    }

    const charCount = args[1]
      ? coerceToNumber(evaluateAst(args[1], dependencies))
      : ({ type: "number", value: 1 } satisfies NumberValue);

    if (charCount.type === "error") {
      return charCount;
    }

    const normalizedCount = Math.trunc(charCount.value);

    if (normalizedCount < 0) {
      return createErrorValue("VALUE");
    }

    return {
      type: "text",
      value: textValue.value.slice(0, normalizedCount),
    };
  }

  function evaluateRight(args: FormulaAst[], dependencies: Set<CellKey>): FormulaValue {
    const argumentError = expectArgumentCount(args, 1, 2);

    if (argumentError) {
      return argumentError;
    }

    const textValue = coerceToText(evaluateAst(args[0], dependencies));

    if (textValue.type === "error") {
      return textValue;
    }

    const charCount = args[1]
      ? coerceToNumber(evaluateAst(args[1], dependencies))
      : ({ type: "number", value: 1 } satisfies NumberValue);

    if (charCount.type === "error") {
      return charCount;
    }

    const normalizedCount = Math.trunc(charCount.value);

    if (normalizedCount < 0) {
      return createErrorValue("VALUE");
    }

    return {
      type: "text",
      value:
        normalizedCount === 0
          ? ""
          : textValue.value.slice(-normalizedCount || textValue.value.length),
    };
  }

  function evaluateMid(args: FormulaAst[], dependencies: Set<CellKey>): FormulaValue {
    const argumentError = expectArgumentCount(args, 3);

    if (argumentError) {
      return argumentError;
    }

    const textValue = coerceToText(evaluateAst(args[0], dependencies));

    if (textValue.type === "error") {
      return textValue;
    }

    const startValue = coerceToNumber(evaluateAst(args[1], dependencies));

    if (startValue.type === "error") {
      return startValue;
    }

    const lengthValue = coerceToNumber(evaluateAst(args[2], dependencies));

    if (lengthValue.type === "error") {
      return lengthValue;
    }

    const startIndex = Math.trunc(startValue.value);
    const length = Math.trunc(lengthValue.value);

    if (startIndex < 1 || length < 0) {
      return createErrorValue("VALUE");
    }

    return {
      type: "text",
      value: textValue.value.slice(startIndex - 1, startIndex - 1 + length),
    };
  }

  function evaluateTrim(args: FormulaAst[], dependencies: Set<CellKey>): FormulaValue {
    const argumentError = expectArgumentCount(args, 1);

    if (argumentError) {
      return argumentError;
    }

    const textValue = coerceToText(evaluateAst(args[0], dependencies));

    if (textValue.type === "error") {
      return textValue;
    }

    return {
      type: "text",
      value: textValue.value.trim().replace(/ +/g, " "),
    };
  }

  function evaluateLower(args: FormulaAst[], dependencies: Set<CellKey>): FormulaValue {
    const argumentError = expectArgumentCount(args, 1);

    if (argumentError) {
      return argumentError;
    }

    const textValue = coerceToText(evaluateAst(args[0], dependencies));

    if (textValue.type === "error") {
      return textValue;
    }

    return {
      type: "text",
      value: textValue.value.toLowerCase(),
    };
  }

  function evaluateUpper(args: FormulaAst[], dependencies: Set<CellKey>): FormulaValue {
    const argumentError = expectArgumentCount(args, 1);

    if (argumentError) {
      return argumentError;
    }

    const textValue = coerceToText(evaluateAst(args[0], dependencies));

    if (textValue.type === "error") {
      return textValue;
    }

    return {
      type: "text",
      value: textValue.value.toUpperCase(),
    };
  }

  function evaluateConcat(args: FormulaAst[], dependencies: Set<CellKey>): FormulaValue {
    const argumentValues = getFunctionArgumentValues(args, dependencies);

    if (isErrorValue(argumentValues)) {
      return argumentValues;
    }

    let text = "";

    for (const argumentValue of argumentValues) {
      for (const value of argumentValue.values) {
        const textValue = coerceToText(value);

        if (textValue.type === "error") {
          return textValue;
        }

        text += textValue.value;
      }
    }

    return {
      type: "text",
      value: text,
    };
  }

  function evaluateTextJoin(args: FormulaAst[], dependencies: Set<CellKey>): FormulaValue {
    const argumentError = expectArgumentCount(args, 3, Number.POSITIVE_INFINITY);

    if (argumentError) {
      return argumentError;
    }

    const delimiter = coerceToText(evaluateAst(args[0], dependencies));

    if (delimiter.type === "error") {
      return delimiter;
    }

    const ignoreEmpty = coerceToBoolean(evaluateAst(args[1], dependencies));

    if (ignoreEmpty.type === "error") {
      return ignoreEmpty;
    }

    const parts: string[] = [];

    for (const arg of args.slice(2)) {
      const argumentValue = getFunctionArgumentValue(arg, dependencies);

      if (isErrorValue(argumentValue)) {
        return argumentValue;
      }

      for (const value of argumentValue.values) {
        const textValue = coerceToText(value);

        if (textValue.type === "error") {
          return textValue;
        }

        if (ignoreEmpty.value && textValue.value.length === 0) {
          continue;
        }

        parts.push(textValue.value);
      }
    }

    return {
      type: "text",
      value: parts.join(delimiter.value),
    };
  }

  function evaluateValue(args: FormulaAst[], dependencies: Set<CellKey>): FormulaValue {
    const argumentError = expectArgumentCount(args, 1);

    if (argumentError) {
      return argumentError;
    }

    const scalarValue = getScalarArgument(args[0], dependencies);

    switch (scalarValue.type) {
      case "blank":
        return {
          type: "number",
          value: 0,
        };
      case "boolean":
        return createErrorValue("VALUE");
      case "error":
        return scalarValue;
      case "number":
        return scalarValue;
      case "text": {
        const parsedNumeric = parseNumericLiteral(scalarValue.value.trim());

        if (parsedNumeric === undefined) {
          return createErrorValue("VALUE");
        }

        return {
          type: "number",
          value: parsedNumeric,
        };
      }
    }
  }

  function evaluateToday(args: FormulaAst[]): FormulaValue {
    const argumentError = expectArgumentCount(args, 0);

    if (argumentError) {
      return argumentError;
    }

    markVolatileFunction();

    const serial = createExcelDateSerial(
      evaluationNow.getFullYear(),
      evaluationNow.getMonth() + 1,
      evaluationNow.getDate(),
    );

    if (serial === undefined) {
      return createErrorValue("NUM");
    }

    return {
      type: "number",
      value: serial,
    };
  }

  function evaluateNow(args: FormulaAst[]): FormulaValue {
    const today = evaluateToday(args);

    if (today.type === "error") {
      return today;
    }

    if (today.type !== "number") {
      return createErrorValue("VALUE");
    }

    const timeFraction =
      (evaluationNow.getHours() * 60 * 60 * 1000 +
        evaluationNow.getMinutes() * 60 * 1000 +
        evaluationNow.getSeconds() * 1000 +
        evaluationNow.getMilliseconds()) /
      MILLISECONDS_PER_DAY;

    return {
      type: "number",
      value: today.value + timeFraction,
    };
  }

  function evaluateDate(args: FormulaAst[], dependencies: Set<CellKey>): FormulaValue {
    const argumentError = expectArgumentCount(args, 3);

    if (argumentError) {
      return argumentError;
    }

    const year = coerceToNumber(evaluateAst(args[0], dependencies));

    if (year.type === "error") {
      return year;
    }

    const month = coerceToNumber(evaluateAst(args[1], dependencies));

    if (month.type === "error") {
      return month;
    }

    const day = coerceToNumber(evaluateAst(args[2], dependencies));

    if (day.type === "error") {
      return day;
    }

    const serial = createExcelDateSerial(
      Math.trunc(year.value),
      Math.trunc(month.value),
      Math.trunc(day.value),
    );

    if (serial === undefined) {
      return createErrorValue("NUM");
    }

    return {
      type: "number",
      value: serial,
    };
  }

  function evaluateYear(args: FormulaAst[], dependencies: Set<CellKey>): FormulaValue {
    return evaluateDatePart(args, dependencies, "year");
  }

  function evaluateMonth(args: FormulaAst[], dependencies: Set<CellKey>): FormulaValue {
    return evaluateDatePart(args, dependencies, "month");
  }

  function evaluateDay(args: FormulaAst[], dependencies: Set<CellKey>): FormulaValue {
    return evaluateDatePart(args, dependencies, "day");
  }

  function evaluateDatePart(
    args: FormulaAst[],
    dependencies: Set<CellKey>,
    part: "day" | "month" | "year",
  ): FormulaValue {
    const argumentError = expectArgumentCount(args, 1);

    if (argumentError) {
      return argumentError;
    }

    const value = coerceToNumber(evaluateAst(args[0], dependencies));

    if (value.type === "error") {
      return value;
    }

    const dateParts = getExcelDateParts(value.value);

    if (isErrorValue(dateParts)) {
      return dateParts;
    }

    return {
      type: "number",
      value: dateParts[part],
    };
  }

  function evaluateChoose(args: FormulaAst[], dependencies: Set<CellKey>): FormulaValue {
    if (args.length < 2) {
      return createErrorValue("VALUE");
    }

    const indexValue = coerceToNumber(evaluateAst(args[0], dependencies));

    if (indexValue.type === "error") {
      return indexValue;
    }

    const choiceIndex = Math.trunc(indexValue.value);

    if (choiceIndex < 1 || choiceIndex >= args.length) {
      return createErrorValue("VALUE");
    }

    return evaluateAst(args[choiceIndex], dependencies);
  }

  function evaluateRow(args: FormulaAst[], dependencies: Set<CellKey>): FormulaValue {
    const argumentError = expectArgumentCount(args, 0, 1);

    if (argumentError) {
      return argumentError;
    }

    if (args.length === 0) {
      const currentCell = getCurrentFormulaCell();

      if (isErrorValue(currentCell)) {
        return currentCell;
      }

      return {
        type: "number",
        value: currentCell.rowIndex + 1,
      };
    }

    const rangeValue = getRangeArgument(args[0], dependencies);

    if (isErrorValue(rangeValue)) {
      return rangeValue;
    }

    const firstCell = getFirstRangeCell(rangeValue);

    if (isErrorValue(firstCell)) {
      return firstCell;
    }

    return {
      type: "number",
      value: firstCell.rowIndex + 1,
    };
  }

  function evaluateColumn(args: FormulaAst[], dependencies: Set<CellKey>): FormulaValue {
    const argumentError = expectArgumentCount(args, 0, 1);

    if (argumentError) {
      return argumentError;
    }

    if (args.length === 0) {
      const currentCell = getCurrentFormulaCell();

      if (isErrorValue(currentCell)) {
        return currentCell;
      }

      return {
        type: "number",
        value: currentCell.columnIndex + 1,
      };
    }

    const rangeValue = getRangeArgument(args[0], dependencies);

    if (isErrorValue(rangeValue)) {
      return rangeValue;
    }

    const firstCell = getFirstRangeCell(rangeValue);

    if (isErrorValue(firstCell)) {
      return firstCell;
    }

    return {
      type: "number",
      value: firstCell.columnIndex + 1,
    };
  }

  function evaluateIndex(args: FormulaAst[], dependencies: Set<CellKey>): FormulaValue {
    const argumentError = expectArgumentCount(args, 2, 3);

    if (argumentError) {
      return argumentError;
    }

    const arrayValue = getRangeArgument(args[0], dependencies);

    if (isErrorValue(arrayValue)) {
      return arrayValue;
    }

    const rowValue = coerceToNumber(evaluateAst(args[1], dependencies));

    if (rowValue.type === "error") {
      return rowValue;
    }

    const rowNumber = Math.trunc(rowValue.value);

    if (rowNumber < 1) {
      return createErrorValue("REF");
    }

    const height = getRangeAreaHeight(arrayValue.cells);
    const width = getRangeAreaWidth(arrayValue.cells);
    let resolvedRow = rowNumber;
    let resolvedColumn = 1;

    if (args[2]) {
      const columnValue = coerceToNumber(evaluateAst(args[2], dependencies));

      if (columnValue.type === "error") {
        return columnValue;
      }

      resolvedColumn = Math.trunc(columnValue.value);

      if (resolvedColumn < 1) {
        return createErrorValue("REF");
      }
    } else if (width === 1) {
      resolvedColumn = 1;
    } else if (height === 1) {
      resolvedColumn = rowNumber;
      resolvedRow = 1;
    } else {
      return createErrorValue("VALUE");
    }

    if (resolvedRow > height || resolvedColumn > width) {
      return createErrorValue("REF");
    }

    const targetCell = getRangeAreaCell(arrayValue.cells, resolvedRow - 1, resolvedColumn - 1);

    if (!targetCell) {
      return createErrorValue("REF");
    }

    return createSingleCellRangeValue(targetCell);
  }

  function evaluateMatch(args: FormulaAst[], dependencies: Set<CellKey>): FormulaValue {
    const argumentError = expectArgumentCount(args, 2, 3);

    if (argumentError) {
      return argumentError;
    }

    const lookupValue = getScalarArgument(args[0], dependencies);

    if (lookupValue.type === "error") {
      return lookupValue;
    }

    const lookupRange = getRangeArgument(args[1], dependencies);

    if (isErrorValue(lookupRange)) {
      return lookupRange;
    }

    const lookupVector = getVectorAddresses(lookupRange);

    if (isErrorValue(lookupVector)) {
      return lookupVector;
    }

    const matchTypeValue = args[2]
      ? coerceToNumber(evaluateAst(args[2], dependencies))
      : ({ type: "number", value: 0 } satisfies NumberValue);

    if (matchTypeValue.type === "error") {
      return matchTypeValue;
    }

    const matchType = Math.trunc(matchTypeValue.value);

    if (matchType !== 0 && matchType !== 1 && matchType !== -1) {
      return createErrorValue("VALUE");
    }

    let bestIndex = -1;
    let bestValue: ScalarFormulaValue | undefined;

    for (let index = 0; index < lookupVector.length; index += 1) {
      const cellValue = evaluateCell(lookupVector[index]).value;

      if (cellValue.type === "error") {
        return cellValue;
      }

      const comparison = compareScalarValues(cellValue, lookupValue);

      if (matchType === 0) {
        if (comparison === 0) {
          return {
            type: "number",
            value: index + 1,
          };
        }

        continue;
      }

      if (matchType === 1) {
        if (comparison > 0) {
          continue;
        }

        if (!bestValue || compareScalarValues(cellValue, bestValue) > 0) {
          bestIndex = index;
          bestValue = cellValue;
        }

        continue;
      }

      if (comparison < 0) {
        continue;
      }

      if (!bestValue || compareScalarValues(cellValue, bestValue) < 0) {
        bestIndex = index;
        bestValue = cellValue;
      }
    }

    if (bestIndex < 0) {
      return createErrorValue("NA");
    }

    return {
      type: "number",
      value: bestIndex + 1,
    };
  }

  function evaluateXLookup(args: FormulaAst[], dependencies: Set<CellKey>): FormulaValue {
    const argumentError = expectArgumentCount(args, 3, 4);

    if (argumentError) {
      return argumentError;
    }

    const lookupValue = getScalarArgument(args[0], dependencies);

    if (lookupValue.type === "error") {
      return lookupValue;
    }

    const lookupRange = getRangeArgument(args[1], dependencies);

    if (isErrorValue(lookupRange)) {
      return lookupRange;
    }

    const returnRange = getRangeArgument(args[2], dependencies);

    if (isErrorValue(returnRange)) {
      return returnRange;
    }

    const lookupVector = getVectorAddresses(lookupRange);
    const returnVector = getVectorAddresses(returnRange);

    if (isErrorValue(lookupVector)) {
      return lookupVector;
    }

    if (isErrorValue(returnVector)) {
      return returnVector;
    }

    if (lookupVector.length !== returnVector.length) {
      return createErrorValue("VALUE");
    }

    for (let index = 0; index < lookupVector.length; index += 1) {
      const candidateValue = evaluateCell(lookupVector[index]).value;

      if (candidateValue.type === "error") {
        return candidateValue;
      }

      if (compareScalarValues(candidateValue, lookupValue) === 0) {
        return evaluateCell(returnVector[index]).value;
      }
    }

    if (args[3]) {
      return evaluateAst(args[3], dependencies);
    }

    return createErrorValue("NA");
  }

  function evaluateVLookup(args: FormulaAst[], dependencies: Set<CellKey>): FormulaValue {
    const argumentError = expectArgumentCount(args, 3, 4);

    if (argumentError) {
      return argumentError;
    }

    const lookupValue = getScalarArgument(args[0], dependencies);

    if (lookupValue.type === "error") {
      return lookupValue;
    }

    const tableRange = getRangeArgument(args[1], dependencies);

    if (isErrorValue(tableRange)) {
      return tableRange;
    }

    const tableHeight = getRangeAreaHeight(tableRange.cells);
    const tableWidth = getRangeAreaWidth(tableRange.cells);

    if (tableHeight === 0 || tableWidth === 0) {
      return createErrorValue("REF");
    }

    const columnValue = coerceToNumber(evaluateAst(args[2], dependencies));

    if (columnValue.type === "error") {
      return columnValue;
    }

    const returnColumnIndex = Math.trunc(columnValue.value);
    if (returnColumnIndex < 1) {
      return createErrorValue("VALUE");
    }

    if (returnColumnIndex > tableWidth) {
      return createErrorValue("REF");
    }

    const rangeLookup = args[3]
      ? coerceToBoolean(evaluateAst(args[3], dependencies))
      : ({ type: "boolean", value: false } satisfies BooleanValue);

    if (rangeLookup.type === "error") {
      return rangeLookup;
    }

    let bestRowIndex = -1;
    let bestValue: ScalarFormulaValue | undefined;

    for (let rowIndex = 0; rowIndex < tableHeight; rowIndex += 1) {
      const lookupCell = getRangeAreaCell(tableRange.cells, rowIndex, 0);

      if (!lookupCell) {
        return createErrorValue("REF");
      }

      const candidateValue = evaluateCell(lookupCell).value;

      if (candidateValue.type === "error") {
        return candidateValue;
      }

      const comparison = compareScalarValues(candidateValue, lookupValue);

      if (!rangeLookup.value) {
        if (comparison === 0) {
          const targetCell = getRangeAreaCell(tableRange.cells, rowIndex, returnColumnIndex - 1);

          return targetCell ? evaluateCell(targetCell).value : createErrorValue("REF");
        }

        continue;
      }

      if (comparison > 0) {
        continue;
      }

      if (!bestValue || compareScalarValues(candidateValue, bestValue) > 0) {
        bestRowIndex = rowIndex;
        bestValue = candidateValue;
      }
    }

    if (bestRowIndex < 0) {
      return createErrorValue("NA");
    }

    const targetCell = getRangeAreaCell(tableRange.cells, bestRowIndex, returnColumnIndex - 1);

    return targetCell ? evaluateCell(targetCell).value : createErrorValue("REF");
  }

  const functionRegistry = new Map<string, FormulaFunctionHandler>([
    ["SUM", evaluateSum],
    ["PRODUCT", evaluateProduct],
    ["MIN", evaluateMin],
    ["MAX", evaluateMax],
    ["AVERAGE", evaluateAverage],
    ["COUNT", evaluateCount],
    ["COUNTA", evaluateCountA],
    ["ABS", evaluateAbs],
    ["ROUND", evaluateRound],
    ["INT", evaluateInt],
    ["MOD", evaluateMod],
    ["POWER", evaluatePower],
    ["SQRT", evaluateSqrt],
    ["LN", evaluateLn],
    ["TRUE", evaluateTrue],
    ["FALSE", evaluateFalse],
    ["AND", evaluateAnd],
    ["OR", evaluateOr],
    ["NOT", evaluateNot],
    ["IF", evaluateIf],
    ["IFERROR", evaluateIfError],
    ["LEN", evaluateLen],
    ["LEFT", evaluateLeft],
    ["RIGHT", evaluateRight],
    ["MID", evaluateMid],
    ["TRIM", evaluateTrim],
    ["LOWER", evaluateLower],
    ["UPPER", evaluateUpper],
    ["CONCAT", evaluateConcat],
    ["TEXTJOIN", evaluateTextJoin],
    ["VALUE", evaluateValue],
    ["TODAY", evaluateToday],
    ["NOW", evaluateNow],
    ["DATE", evaluateDate],
    ["YEAR", evaluateYear],
    ["MONTH", evaluateMonth],
    ["DAY", evaluateDay],
    ["CHOOSE", evaluateChoose],
    ["ROW", evaluateRow],
    ["COLUMN", evaluateColumn],
    ["INDEX", evaluateIndex],
    ["MATCH", evaluateMatch],
    ["XLOOKUP", evaluateXLookup],
    ["VLOOKUP", evaluateVLookup],
  ]);

  for (const sheet of workbook.sheets) {
    const rowCount = getSheetRowCount(sheet);
    const columnCount = getSheetColumnCount(sheet);

    for (let rowIndex = 0; rowIndex < rowCount; rowIndex += 1) {
      for (let columnIndex = 0; columnIndex < columnCount; columnIndex += 1) {
        evaluateCell({
          sheetId: sheet.id,
          rowIndex,
          columnIndex,
        });
      }
    }
  }

  if (hasVolatileFunctions) {
    for (const snapshot of snapshots.values()) {
      snapshot.hasVolatileFunctions = true;
    }
  }

  return snapshots;

  function getSnapshot(sheetId: string): SheetEvaluationSnapshot {
    const snapshot = snapshots.get(sheetId);

    if (!snapshot) {
      throw new Error(`Sheet "${sheetId}" is missing from the evaluation snapshot.`);
    }

    return snapshot;
  }

  function resolveFormulaSheetId(sheetName?: string): string | ErrorValue {
    if (sheetName === undefined) {
      const currentCell = getCurrentFormulaCell();

      return isErrorValue(currentCell) ? currentCell : currentCell.sheetId;
    }

    return sheetIdByName.get(getSheetNameKey(sheetName)) ?? createErrorValue("REF");
  }

  function resolveFormulaAddress(reference: FormulaReferenceAddress): CellAddress | ErrorValue {
    const sheetId = resolveFormulaSheetId(reference.sheetName);

    if (isErrorValue(sheetId)) {
      return sheetId;
    }

    return {
      sheetId,
      rowIndex: reference.rowIndex,
      columnIndex: reference.columnIndex,
    };
  }

  function markVolatileFunction() {
    hasVolatileFunctions = true;
  }
}

function createMissingSheetSnapshot(
  sheetId: string,
  workbookVersion: number,
): SheetEvaluationSnapshot {
  return {
    sheetId,
    workbookVersion,
    hasVolatileFunctions: false,
    cells: new Map<CellKey, CellEvaluation>(),
    dependents: new Map<CellKey, Set<CellKey>>(),
    precedents: new Map<CellKey, Set<CellKey>>(),
  };
}

function readStringLiteral(
  expression: string,
  startIndex: number,
): { value: string; nextIndex: number } {
  let value = "";
  let index = startIndex + 1;

  while (index < expression.length) {
    const character = expression[index];

    if (character === '"') {
      if (expression[index + 1] === '"') {
        value += '"';
        index += 2;
        continue;
      }

      return {
        value,
        nextIndex: index + 1,
      };
    }

    value += character;
    index += 1;
  }

  throw new Error("Formula text literal is missing a closing quote.");
}

function readQuotedSheetName(
  expression: string,
  startIndex: number,
): { value: string; nextIndex: number } {
  let value = "";
  let index = startIndex + 1;

  while (index < expression.length) {
    const character = expression[index];

    if (character === "'") {
      if (expression[index + 1] === "'") {
        value += "'";
        index += 2;
        continue;
      }

      return {
        value,
        nextIndex: index + 1,
      };
    }

    value += character;
    index += 1;
  }

  throw new Error("Formula sheet name is missing a closing quote.");
}

function readStructuredReferenceSpecifier(
  expression: string,
  startIndex: number,
): { value: string; nextIndex: number } {
  let depth = 0;
  let index = startIndex;

  while (index < expression.length) {
    const character = expression[index];

    if (character === "[") {
      depth += 1;
    } else if (character === "]") {
      depth -= 1;

      if (depth === 0) {
        return {
          value: expression.slice(startIndex, index + 1),
          nextIndex: index + 1,
        };
      }
    }

    if (depth < 0) {
      break;
    }

    index += 1;
  }

  throw new Error("Formula structured reference is missing a closing bracket.");
}

function createErrorEvaluation(
  input: string,
  errorCode: FormulaErrorCode,
  dependencies: readonly CellKey[],
): CellEvaluation {
  return {
    input,
    display: getErrorDisplay(errorCode),
    isFormula: isFormulaInput(input),
    value: createErrorValue(errorCode),
    errorCode,
    dependencies: [...dependencies],
  };
}

function createErrorValue(errorCode: FormulaErrorCode): ErrorValue {
  return {
    type: "error",
    errorCode,
  };
}

function isErrorValue(value: unknown): value is ErrorValue {
  return typeof value === "object" && value !== null && "type" in value && value.type === "error";
}

function getDisplayForInputValue(
  input: string,
  value: ScalarFormulaValue,
  style?: WorkbookCellStyle,
): string {
  if (value.type === "number" && style?.numberFormat) {
    return formatWorkbookNumberDisplay(value.value, style.numberFormat);
  }

  return input;
}

function getDisplayForValue(value: ScalarFormulaValue, style?: WorkbookCellStyle): string {
  switch (value.type) {
    case "blank":
      return "";
    case "boolean":
      return value.value ? "TRUE" : "FALSE";
    case "error":
      return getErrorDisplay(value.errorCode);
    case "number":
      return formatWorkbookNumberDisplay(value.value, style?.numberFormat);
    case "text":
      return value.value;
  }
}

function formatNumericDisplay(value: number): string {
  const normalizedValue = Object.is(value, -0) ? 0 : value;

  return String(normalizedValue);
}

type ExcelDateParts = {
  year: number;
  month: number;
  day: number;
};

function createExcelDateSerial(year: number, month: number, day: number): number | undefined {
  if (!Number.isFinite(year) || !Number.isFinite(month) || !Number.isFinite(day)) {
    return undefined;
  }

  const normalizedYear = year >= 0 && year <= 1899 ? year + 1900 : year;

  if (normalizedYear < 1900 || normalizedYear > 9999) {
    return undefined;
  }

  if (normalizedYear === 1900 && month === 2 && day === 29) {
    return EXCEL_1900_LEAP_BUG_SERIAL;
  }

  const dateUtc = Date.UTC(normalizedYear, month - 1, day);

  if (!Number.isFinite(dateUtc)) {
    return undefined;
  }

  const date = new Date(dateUtc);
  const resultYear = date.getUTCFullYear();

  if (resultYear < 1900 || resultYear > 9999) {
    return undefined;
  }

  const serial = getExcelDateSerialFromUtcDate(date);

  if (serial < 0 || serial > EXCEL_MAX_SERIAL) {
    return undefined;
  }

  return serial;
}

function getExcelDateParts(serial: number): ExcelDateParts | ErrorValue {
  if (!Number.isFinite(serial)) {
    return createErrorValue("NUM");
  }

  const wholeDaySerial = Math.floor(serial);

  if (wholeDaySerial < EXCEL_MIN_SERIAL || wholeDaySerial > EXCEL_MAX_SERIAL) {
    return createErrorValue("NUM");
  }

  if (wholeDaySerial === EXCEL_1900_LEAP_BUG_SERIAL) {
    return {
      year: 1900,
      month: 2,
      day: 29,
    };
  }

  const adjustedSerial =
    wholeDaySerial > EXCEL_1900_LEAP_BUG_SERIAL ? wholeDaySerial - 1 : wholeDaySerial;
  const date = new Date(EXCEL_1900_EPOCH_UTC + adjustedSerial * MILLISECONDS_PER_DAY);
  const year = date.getUTCFullYear();

  if (year < 1900 || year > 9999) {
    return createErrorValue("NUM");
  }

  return {
    year,
    month: date.getUTCMonth() + 1,
    day: date.getUTCDate(),
  };
}

function getExcelDateSerialFromUtcDate(date: Date): number {
  const dateUtc = Date.UTC(date.getUTCFullYear(), date.getUTCMonth(), date.getUTCDate());
  const serial = Math.floor((dateUtc - EXCEL_1900_EPOCH_UTC) / MILLISECONDS_PER_DAY);

  return serial >= EXCEL_1900_LEAP_BUG_SERIAL ? serial + 1 : serial;
}

function getErrorDisplay(errorCode: FormulaErrorCode): string {
  switch (errorCode) {
    case "PARSE":
      return "#ERROR!";
    case "REF":
      return "#REF!";
    case "DIV0":
      return "#DIV/0!";
    case "VALUE":
      return "#VALUE!";
    case "CYCLE":
      return "#CYCLE!";
    case "NAME":
      return "#NAME?";
    case "NUM":
      return "#NUM!";
    case "NA":
      return "#N/A";
    case "NULL":
      return "#NULL!";
  }
}

function isCellReferenceIdentifier(value: string): boolean {
  return /^[A-Za-z]+[1-9][0-9]*$/.test(value);
}

function getSheetNameKey(value: string): string {
  return value.trim().toLowerCase();
}

function getStructuredReferenceNameKey(value: string): string {
  return value.trim().toLowerCase();
}

function parseStructuredReferenceSelection(
  specifier: string,
): StructuredReferenceSelection | ErrorValue {
  const inner = unwrapStructuredReferencePart(specifier);

  if (inner === undefined) {
    return createErrorValue("REF");
  }

  const selection: StructuredReferenceSelection = {
    currentRow: false,
    dataOnly: false,
  };
  const parts = splitStructuredReferenceTopLevel(inner, ",");

  for (const rawPart of parts) {
    const part = normalizeStructuredReferencePart(rawPart);

    if (part.length === 0) {
      return createErrorValue("REF");
    }

    if (part.toUpperCase() === "#DATA") {
      selection.dataOnly = true;
      continue;
    }

    if (part.toUpperCase() === "#THIS ROW" || part === "@") {
      selection.currentRow = true;
      continue;
    }

    const columnRange = parseStructuredReferenceColumnRange(part);

    if (!columnRange) {
      return createErrorValue("REF");
    }

    selection.currentRow = selection.currentRow || columnRange.currentRow;
    selection.columnStart = columnRange.columnStart;
    selection.columnEnd = columnRange.columnEnd;
  }

  return selection;
}

function parseStructuredReferenceColumnRange(
  value: string,
): { columnStart: string; columnEnd?: string; currentRow: boolean } | undefined {
  const parts = splitStructuredReferenceTopLevel(value, ":");

  if (parts.length < 1 || parts.length > 2) {
    return undefined;
  }

  const start = parseStructuredReferenceColumnPart(parts[0]);
  const end = parts[1] === undefined ? undefined : parseStructuredReferenceColumnPart(parts[1]);

  if (!start || (parts[1] !== undefined && !end)) {
    return undefined;
  }

  return {
    columnStart: start.columnName,
    columnEnd: end?.columnName,
    currentRow: start.currentRow || (end?.currentRow ?? false),
  };
}

function parseStructuredReferenceColumnPart(
  value: string,
): { columnName: string; currentRow: boolean } | undefined {
  let part = value.trim();
  let currentRow = false;

  if (part.startsWith("@")) {
    currentRow = true;
    part = part.slice(1).trim();
  }

  const unwrapped = unwrapStructuredReferencePart(part);

  if (unwrapped !== undefined) {
    part = unwrapped.trim();
  }

  if (part.startsWith("@")) {
    currentRow = true;
    part = part.slice(1).trim();
  }

  if (part.length === 0 || part.startsWith("#")) {
    return undefined;
  }

  return {
    columnName: part,
    currentRow,
  };
}

function normalizeStructuredReferencePart(value: string): string {
  const trimmed = value.trim();
  const unwrapped = unwrapStructuredReferencePart(trimmed);

  return unwrapped === undefined ? trimmed : unwrapped.trim();
}

function unwrapStructuredReferencePart(value: string): string | undefined {
  const trimmed = value.trim();

  if (!trimmed.startsWith("[") || !trimmed.endsWith("]")) {
    return undefined;
  }

  let depth = 0;

  for (let index = 0; index < trimmed.length; index += 1) {
    const character = trimmed[index];

    if (character === "[") {
      depth += 1;
    } else if (character === "]") {
      depth -= 1;

      if (depth === 0 && index !== trimmed.length - 1) {
        return undefined;
      }
    }

    if (depth < 0) {
      return undefined;
    }
  }

  return depth === 0 ? trimmed.slice(1, -1) : undefined;
}

function splitStructuredReferenceTopLevel(value: string, separator: "," | ":"): string[] {
  const parts: string[] = [];
  let depth = 0;
  let startIndex = 0;

  for (let index = 0; index < value.length; index += 1) {
    const character = value[index];

    if (character === "[") {
      depth += 1;
    } else if (character === "]") {
      depth -= 1;
    } else if (character === separator && depth === 0) {
      parts.push(value.slice(startIndex, index));
      startIndex = index + 1;
    }
  }

  parts.push(value.slice(startIndex));
  return parts;
}

function workbookTableContainsCell(
  table: WorkbookState["tables"][number],
  cell: CellAddress,
): boolean {
  return (
    table.range.sheetId === cell.sheetId &&
    cell.rowIndex >= table.range.startRow &&
    cell.rowIndex < table.range.startRow + table.range.rowCount &&
    cell.columnIndex >= table.range.startColumn &&
    cell.columnIndex < table.range.startColumn + table.range.columnCount
  );
}

function resolveStructuredReferenceColumn(
  sheet: WorkbookSheet,
  table: WorkbookState["tables"][number],
  columnName: string,
): number | ErrorValue {
  const matches: number[] = [];
  const expectedKey = getStructuredReferenceNameKey(columnName);

  for (let columnOffset = 0; columnOffset < table.range.columnCount; columnOffset += 1) {
    const columnIndex = table.range.startColumn + columnOffset;
    const tableColumnName = getStructuredReferenceTableColumnName(sheet, table, columnOffset);

    if (getStructuredReferenceNameKey(tableColumnName) === expectedKey) {
      matches.push(columnIndex);
    }
  }

  return matches.length === 1 ? matches[0] : createErrorValue("REF");
}

function getStructuredReferenceTableColumnName(
  sheet: WorkbookSheet,
  table: WorkbookState["tables"][number],
  columnOffset: number,
): string {
  if (!table.hasHeaderRow) {
    return `Column${columnOffset + 1}`;
  }

  const headerValue =
    sheet.cells[table.range.startRow]?.[table.range.startColumn + columnOffset]?.trim() ?? "";

  return headerValue.length > 0 ? headerValue : `Column${columnOffset + 1}`;
}

export function getCellKeySheetId(cellKey: CellKey): string {
  const columnSeparator = cellKey.lastIndexOf(":");
  const rowSeparator = cellKey.lastIndexOf(":", columnSeparator - 1);

  return rowSeparator < 0 ? "" : cellKey.slice(0, rowSeparator);
}

function createSheetCellStyleKey(rowIndex: number, columnIndex: number): string {
  return `${rowIndex}:${columnIndex}`;
}

function compareDisplayText(left: string, right: string): number {
  return left.localeCompare(right, undefined, {
    numeric: true,
    sensitivity: "base",
  });
}

function parseNumericLiteral(input: string): number | undefined {
  if (!NUMBER_LITERAL_PATTERN.test(input)) {
    return undefined;
  }

  return Number(input);
}

function parseRawCellValue(input: string): ScalarFormulaValue {
  if (input.length === 0) {
    return BLANK_VALUE;
  }

  const parsedNumeric = parseNumericLiteral(input);

  if (parsedNumeric !== undefined) {
    return {
      type: "number",
      value: parsedNumeric,
    };
  }

  return {
    type: "text",
    value: input,
  };
}
