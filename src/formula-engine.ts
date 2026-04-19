import {
  getSheetColumnCount,
  getSheetRowCount,
  isFormulaInput,
  parseCellReference,
  type FormulaErrorCode,
  type WorkbookSheet,
} from "./workbook-core";

export type CellKey = `${number}:${number}`;

type CellAddress = {
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

type ScalarFormulaValue =
  | BlankValue
  | BooleanValue
  | ErrorValue
  | NumberValue
  | TextValue;

type RangeValue = {
  type: "range";
  cells: CellAddress[][];
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
  cells: Map<CellKey, CellEvaluation>;
  dependents: Map<CellKey, Set<CellKey>>;
  precedents: Map<CellKey, Set<CellKey>>;
}

type FormulaToken =
  | { type: "comma" }
  | { type: "error"; errorCode: FormulaErrorCode }
  | { type: "identifier"; value: string }
  | { type: "leftParen" }
  | { type: "number"; value: number }
  | {
      type: "operator";
      value:
        | "+"
        | "-"
        | "*"
        | "/"
        | "^"
        | "&"
        | "="
        | "<>"
        | "<"
        | "<="
        | ">"
        | ">="
        | ":"
        | "%";
    }
  | { type: "rightParen" }
  | { type: "text"; value: string };

type FormulaAst =
  | {
      type: "binary";
      operator:
        | "+"
        | "-"
        | "*"
        | "/"
        | "^"
        | "&"
        | "="
        | "<>"
        | "<"
        | "<="
        | ">"
        | ">=";
      left: FormulaAst;
      right: FormulaAst;
    }
  | { type: "function"; name: string; args: FormulaAst[] }
  | { type: "literal"; value: ScalarFormulaValue }
  | { type: "name"; name: string }
  | {
      type: "percent";
      operand: FormulaAst;
    }
  | { type: "range"; start: CellAddress; end: CellAddress }
  | { type: "reference"; rowIndex: number; columnIndex: number }
  | { type: "unary"; operator: "+" | "-"; operand: FormulaAst };

type FunctionArgumentValue = {
  fromRange: boolean;
  values: ScalarFormulaValue[];
};

type FormulaFunctionHandler = (
  args: FormulaAst[],
  dependencies: Set<CellKey>,
) => FormulaValue;

const BLANK_VALUE: BlankValue = {
  type: "blank",
};

const EMPTY_DEPENDENCIES: CellKey[] = [];
const NUMBER_LITERAL_PATTERN =
  /^[+-]?(?:\d+(?:\.\d+)?|\.\d+)(?:[Ee][+-]?\d+)?$/;
const ERROR_LITERALS: ReadonlyArray<[string, FormulaErrorCode]> = [
  ["#DIV/0!", "DIV0"],
  ["#NAME?", "NAME"],
  ["#NULL!", "NULL"],
  ["#VALUE!", "VALUE"],
  ["#REF!", "REF"],
  ["#NUM!", "NUM"],
  ["#N/A", "NA"],
];

export function createCellKey(rowIndex: number, columnIndex: number): CellKey {
  return `${rowIndex}:${columnIndex}`;
}

export function getCellEvaluation(
  snapshot: SheetEvaluationSnapshot,
  rowIndex: number,
  columnIndex: number,
): CellEvaluation {
  const evaluation = snapshot.cells.get(createCellKey(rowIndex, columnIndex));

  if (!evaluation) {
    throw new Error(
      `Cell ${rowIndex}:${columnIndex} is missing from the evaluation snapshot.`,
    );
  }

  return evaluation;
}

export function tokenizeFormula(input: string): FormulaToken[] {
  const expression = input.startsWith("=") ? input.slice(1) : input;
  const tokens: FormulaToken[] = [];
  let index = 0;

  while (index < expression.length) {
    const character = expression[index];

    if (/\s/.test(character)) {
      index += 1;
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

    if (/[A-Za-z_\\]/.test(character)) {
      const identifierMatch = /^[A-Za-z_\\][A-Za-z0-9_.]*/.exec(
        expression.slice(index),
      );

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

    const numberMatch = /^(?:\d+(?:\.\d+)?|\.\d+)(?:[Ee][+-]?\d+)?/.exec(
      expression.slice(index),
    );

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

  function parseExpression(): FormulaAst {
    return parseComparison();
  }

  function parseComparison(): FormulaAst {
    let node = parseConcatenation();
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
      token = tokens[index];
    }

    return node;
  }

  function parseConcatenation(): FormulaAst {
    let node = parseAdditive();
    let token = tokens[index];

    while (token?.type === "operator" && token.value === "&") {
      index += 1;
      node = {
        type: "binary",
        operator: "&",
        left: node,
        right: parseAdditive(),
      };
      token = tokens[index];
    }

    return node;
  }

  function parseAdditive(): FormulaAst {
    let node = parseMultiplicative();
    let token = tokens[index];

    while (
      token?.type === "operator" &&
      (token.value === "+" || token.value === "-")
    ) {
      const operator = token.value;

      index += 1;
      node = {
        type: "binary",
        operator,
        left: node,
        right: parseMultiplicative(),
      };
      token = tokens[index];
    }

    return node;
  }

  function parseMultiplicative(): FormulaAst {
    let node = parsePower();
    let token = tokens[index];

    while (
      token?.type === "operator" &&
      (token.value === "*" || token.value === "/")
    ) {
      const operator = token.value;

      index += 1;
      node = {
        type: "binary",
        operator,
        left: node,
        right: parsePower(),
      };
      token = tokens[index];
    }

    return node;
  }

  function parsePower(): FormulaAst {
    let node = parsePercent();
    let token = tokens[index];

    while (token?.type === "operator" && token.value === "^") {
      index += 1;
      node = {
        type: "binary",
        operator: "^",
        left: node,
        right: parsePercent(),
      };
      token = tokens[index];
    }

    return node;
  }

  function parsePercent(): FormulaAst {
    let node = parseUnary();
    let token = tokens[index];

    while (token?.type === "operator" && token.value === "%") {
      index += 1;
      node = {
        type: "percent",
        operand: node,
      };
      token = tokens[index];
    }

    return node;
  }

  function parseUnary(): FormulaAst {
    const token = tokens[index];

    if (
      token?.type === "operator" &&
      (token.value === "+" || token.value === "-")
    ) {
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
    const node = parsePrimary();
    const token = tokens[index];

    if (token?.type !== "operator" || token.value !== ":") {
      return node;
    }

    if (node.type !== "reference") {
      throw new Error("Formula range start must be a cell reference.");
    }

    index += 1;

    const endNode = parsePrimary();

    if (endNode.type !== "reference") {
      throw new Error("Formula range end must be a cell reference.");
    }

    return {
      type: "range",
      start: {
        rowIndex: node.rowIndex,
        columnIndex: node.columnIndex,
      },
      end: {
        rowIndex: endNode.rowIndex,
        columnIndex: endNode.columnIndex,
      },
    };
  }

  function parsePrimary(): FormulaAst {
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

    if (token.type === "leftParen") {
      index += 1;
      const expression = parseExpression();

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
      throw new Error(
        "Formula function call is missing an opening parenthesis.",
      );
    }

    index += 1;

    const args: FormulaAst[] = [];

    if (tokens[index]?.type !== "rightParen") {
      let shouldContinue = true;

      while (shouldContinue) {
        args.push(parseExpression());

        if (tokens[index]?.type === "comma") {
          index += 1;
          continue;
        }

        shouldContinue = false;
      }
    }

    if (tokens[index]?.type !== "rightParen") {
      throw new Error(
        "Formula function call is missing a closing parenthesis.",
      );
    }

    index += 1;

    return {
      type: "function",
      name,
      args,
    };
  }

  const ast = parseExpression();

  if (index !== tokens.length) {
    throw new Error("Formula contains unexpected trailing tokens.");
  }

  return ast;
}

export function evaluateSheet(
  sheet: WorkbookSheet,
  workbookVersion: number,
): SheetEvaluationSnapshot {
  const rowCount = getSheetRowCount(sheet);
  const columnCount = getSheetColumnCount(sheet);
  const cells = new Map<CellKey, CellEvaluation>();
  const dependents = new Map<CellKey, Set<CellKey>>();
  const precedents = new Map<CellKey, Set<CellKey>>();
  const evaluationStack: CellKey[] = [];
  const cycleCellKeys = new Set<CellKey>();
  const formulaCellStack: CellAddress[] = [];

  function getInput(rowIndex: number, columnIndex: number): string {
    return sheet.cells[rowIndex]?.[columnIndex] ?? "";
  }

  function recordDependencies(
    cellKey: CellKey,
    dependencies: readonly CellKey[],
  ) {
    if (dependencies.length === 0) {
      return;
    }

    const precedentSet = new Set(dependencies);

    precedents.set(cellKey, precedentSet);

    for (const dependencyKey of precedentSet) {
      const dependentSet = dependents.get(dependencyKey) ?? new Set<CellKey>();

      dependentSet.add(cellKey);
      dependents.set(dependencyKey, dependentSet);
    }
  }

  function evaluateCell(rowIndex: number, columnIndex: number): CellEvaluation {
    const cellKey = createCellKey(rowIndex, columnIndex);
    const cachedEvaluation = cells.get(cellKey);

    if (cachedEvaluation) {
      return cachedEvaluation;
    }

    const stackIndex = evaluationStack.indexOf(cellKey);

    if (stackIndex >= 0) {
      for (const cycleKey of evaluationStack.slice(stackIndex)) {
        cycleCellKeys.add(cycleKey);
      }

      return createErrorEvaluation(
        getInput(rowIndex, columnIndex),
        "CYCLE",
        EMPTY_DEPENDENCIES,
      );
    }

    evaluationStack.push(cellKey);

    try {
      const input = getInput(rowIndex, columnIndex);
      let evaluation: CellEvaluation;

      if (!isFormulaInput(input)) {
        const value = parseRawCellValue(input);

        evaluation = {
          input,
          display: input,
          isFormula: false,
          value,
          dependencies: EMPTY_DEPENDENCIES,
        };
      } else {
        evaluation = evaluateFormulaCell(input, cellKey, rowIndex, columnIndex);
      }

      cells.set(cellKey, evaluation);
      recordDependencies(cellKey, evaluation.dependencies);
      return evaluation;
    } finally {
      evaluationStack.pop();
    }
  }

  function evaluateFormulaCell(
    input: string,
    cellKey: CellKey,
    rowIndex: number,
    columnIndex: number,
  ): CellEvaluation {
    let ast: FormulaAst;

    try {
      ast = parseFormula(input);
    } catch {
      return createErrorEvaluation(input, "PARSE", EMPTY_DEPENDENCIES);
    }

    const dependencies = new Set<CellKey>();
    let value: ScalarFormulaValue;

    formulaCellStack.push({ rowIndex, columnIndex });

    try {
      value = scalarizeFormulaValue(evaluateAst(ast, dependencies));
    } finally {
      formulaCellStack.pop();
    }

    const normalizedDependencies = [...dependencies];
    const errorCode = cycleCellKeys.has(cellKey)
      ? "CYCLE"
      : value.type === "error"
        ? value.errorCode
        : undefined;

    if (errorCode) {
      return createErrorEvaluation(input, errorCode, normalizedDependencies);
    }

    return {
      input,
      display: getDisplayForValue(value),
      isFormula: true,
      value,
      dependencies: normalizedDependencies,
    };
  }

  function evaluateAst(
    ast: FormulaAst,
    dependencies: Set<CellKey>,
  ): FormulaValue {
    switch (ast.type) {
      case "function":
        return evaluateFunctionCall(ast.name, ast.args, dependencies);
      case "literal":
        return ast.value;
      case "name":
        return createErrorValue("NAME");
      case "percent": {
        const numericOperand = coerceToNumber(
          evaluateAst(ast.operand, dependencies),
        );

        if (numericOperand.type === "error") {
          return numericOperand;
        }

        return {
          type: "number",
          value: numericOperand.value / 100,
        };
      }
      case "range":
        return createRangeValue(ast.start, ast.end, dependencies);
      case "reference":
        return createRangeValue(
          {
            rowIndex: ast.rowIndex,
            columnIndex: ast.columnIndex,
          },
          {
            rowIndex: ast.rowIndex,
            columnIndex: ast.columnIndex,
          },
          dependencies,
        );
      case "unary": {
        const numericOperand = coerceToNumber(
          evaluateAst(ast.operand, dependencies),
        );

        if (numericOperand.type === "error") {
          return numericOperand;
        }

        return {
          type: "number",
          value:
            ast.operator === "-" ? -numericOperand.value : numericOperand.value,
        };
      }
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

  function createRangeValue(
    start: CellAddress,
    end: CellAddress,
    dependencies: Set<CellKey>,
  ): RangeValue | ErrorValue {
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
    const cellsInRange = Array.from(
      { length: endRow - startRow + 1 },
      (_, rowOffset) =>
        Array.from(
          { length: endColumn - startColumn + 1 },
          (_, columnOffset) => {
            const cellAddress = {
              rowIndex: startRow + rowOffset,
              columnIndex: startColumn + columnOffset,
            };

            dependencies.add(
              createCellKey(cellAddress.rowIndex, cellAddress.columnIndex),
            );

            return cellAddress;
          },
        ),
    );

    return {
      type: "range",
      cells: cellsInRange,
    };
  }

  function scalarizeFormulaValue(value: FormulaValue): ScalarFormulaValue {
    if (value.type !== "range") {
      return value;
    }

    if (value.cells.length !== 1 || value.cells[0]?.length !== 1) {
      return createErrorValue("VALUE");
    }

    const referencedCell = value.cells[0][0];

    return evaluateCell(referencedCell.rowIndex, referencedCell.columnIndex)
      .value;
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

    const comparison = compareScalarValues(
      normalizedOperands.left,
      normalizedOperands.right,
    );

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
      leftScalar.type === "blank"
        ? coerceBlankForComparison(rightScalar)
        : leftScalar;
    const normalizedRight =
      rightScalar.type === "blank"
        ? coerceBlankForComparison(leftScalar)
        : rightScalar;

    return {
      left: normalizedLeft,
      right: normalizedRight,
    };
  }

  function coerceBlankForComparison(
    other: ScalarFormulaValue,
  ): ScalarFormulaValue {
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

  function compareScalarValues(
    left: ScalarFormulaValue,
    right: ScalarFormulaValue,
  ): number {
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

  function getScalarArgument(
    arg: FormulaAst,
    dependencies: Set<CellKey>,
  ): ScalarFormulaValue {
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

    const flattenedValues = flattenRangeCells(value.cells);

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

  function flattenRangeCells(
    cellsInRange: readonly (readonly CellAddress[])[],
  ): ScalarFormulaValue[] | ErrorValue {
    const flattenedValues: ScalarFormulaValue[] = [];

    for (const row of cellsInRange) {
      for (const cellAddress of row) {
        const cellValue = evaluateCell(
          cellAddress.rowIndex,
          cellAddress.columnIndex,
        ).value;

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

  function getRangeArgument(
    arg: FormulaAst,
    dependencies: Set<CellKey>,
  ): RangeValue | ErrorValue {
    const value = evaluateAst(arg, dependencies);

    if (value.type !== "range") {
      return createErrorValue("VALUE");
    }

    return value;
  }

  function getFirstRangeCell(rangeValue: RangeValue): CellAddress | ErrorValue {
    const firstCell = rangeValue.cells[0]?.[0];

    if (!firstCell) {
      return createErrorValue("REF");
    }

    return firstCell;
  }

  function getVectorAddresses(
    rangeValue: RangeValue,
  ): CellAddress[] | ErrorValue {
    if (rangeValue.cells.length === 0 || rangeValue.cells[0]?.length === 0) {
      return createErrorValue("REF");
    }

    if (rangeValue.cells.length === 1) {
      return [...rangeValue.cells[0]];
    }

    if (rangeValue.cells.every((row) => row.length === 1)) {
      return rangeValue.cells.map((row) => row[0]);
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

  function evaluateSum(
    args: FormulaAst[],
    dependencies: Set<CellKey>,
  ): FormulaValue {
    const numericValues = collectAggregateNumbers(args, dependencies);

    if (isErrorValue(numericValues)) {
      return numericValues;
    }

    return {
      type: "number",
      value: numericValues.reduce((total, value) => total + value, 0),
    };
  }

  function evaluateProduct(
    args: FormulaAst[],
    dependencies: Set<CellKey>,
  ): FormulaValue {
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

  function evaluateMin(
    args: FormulaAst[],
    dependencies: Set<CellKey>,
  ): FormulaValue {
    const numericValues = collectAggregateNumbers(args, dependencies);

    if (isErrorValue(numericValues)) {
      return numericValues;
    }

    return {
      type: "number",
      value: numericValues.length === 0 ? 0 : Math.min(...numericValues),
    };
  }

  function evaluateMax(
    args: FormulaAst[],
    dependencies: Set<CellKey>,
  ): FormulaValue {
    const numericValues = collectAggregateNumbers(args, dependencies);

    if (isErrorValue(numericValues)) {
      return numericValues;
    }

    return {
      type: "number",
      value: numericValues.length === 0 ? 0 : Math.max(...numericValues),
    };
  }

  function evaluateAverage(
    args: FormulaAst[],
    dependencies: Set<CellKey>,
  ): FormulaValue {
    const numericValues = collectAggregateNumbers(args, dependencies);

    if (isErrorValue(numericValues)) {
      return numericValues;
    }

    if (numericValues.length === 0) {
      return createErrorValue("DIV0");
    }

    return {
      type: "number",
      value:
        numericValues.reduce((total, value) => total + value, 0) /
        numericValues.length,
    };
  }

  function evaluateCount(
    args: FormulaAst[],
    dependencies: Set<CellKey>,
  ): FormulaValue {
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

  function evaluateCountA(
    args: FormulaAst[],
    dependencies: Set<CellKey>,
  ): FormulaValue {
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

  function evaluateAbs(
    args: FormulaAst[],
    dependencies: Set<CellKey>,
  ): FormulaValue {
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

  function evaluateRound(
    args: FormulaAst[],
    dependencies: Set<CellKey>,
  ): FormulaValue {
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

  function evaluateInt(
    args: FormulaAst[],
    dependencies: Set<CellKey>,
  ): FormulaValue {
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

  function evaluateMod(
    args: FormulaAst[],
    dependencies: Set<CellKey>,
  ): FormulaValue {
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

  function evaluatePower(
    args: FormulaAst[],
    dependencies: Set<CellKey>,
  ): FormulaValue {
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

  function evaluateSqrt(
    args: FormulaAst[],
    dependencies: Set<CellKey>,
  ): FormulaValue {
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

  function evaluateAnd(
    args: FormulaAst[],
    dependencies: Set<CellKey>,
  ): FormulaValue {
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

  function evaluateOr(
    args: FormulaAst[],
    dependencies: Set<CellKey>,
  ): FormulaValue {
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

  function evaluateNot(
    args: FormulaAst[],
    dependencies: Set<CellKey>,
  ): FormulaValue {
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

  function evaluateIf(
    args: FormulaAst[],
    dependencies: Set<CellKey>,
  ): FormulaValue {
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

  function evaluateIfError(
    args: FormulaAst[],
    dependencies: Set<CellKey>,
  ): FormulaValue {
    const argumentError = expectArgumentCount(args, 2);

    if (argumentError) {
      return argumentError;
    }

    const firstValue = scalarizeFormulaValue(
      evaluateAst(args[0], dependencies),
    );

    if (firstValue.type === "error") {
      return evaluateAst(args[1], dependencies);
    }

    return firstValue;
  }

  function evaluateLen(
    args: FormulaAst[],
    dependencies: Set<CellKey>,
  ): FormulaValue {
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

  function evaluateLeft(
    args: FormulaAst[],
    dependencies: Set<CellKey>,
  ): FormulaValue {
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

  function evaluateRight(
    args: FormulaAst[],
    dependencies: Set<CellKey>,
  ): FormulaValue {
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

  function evaluateMid(
    args: FormulaAst[],
    dependencies: Set<CellKey>,
  ): FormulaValue {
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

  function evaluateTrim(
    args: FormulaAst[],
    dependencies: Set<CellKey>,
  ): FormulaValue {
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

  function evaluateLower(
    args: FormulaAst[],
    dependencies: Set<CellKey>,
  ): FormulaValue {
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

  function evaluateUpper(
    args: FormulaAst[],
    dependencies: Set<CellKey>,
  ): FormulaValue {
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

  function evaluateConcat(
    args: FormulaAst[],
    dependencies: Set<CellKey>,
  ): FormulaValue {
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

  function evaluateTextJoin(
    args: FormulaAst[],
    dependencies: Set<CellKey>,
  ): FormulaValue {
    const argumentError = expectArgumentCount(
      args,
      3,
      Number.POSITIVE_INFINITY,
    );

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

  function evaluateValue(
    args: FormulaAst[],
    dependencies: Set<CellKey>,
  ): FormulaValue {
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

  function evaluateChoose(
    args: FormulaAst[],
    dependencies: Set<CellKey>,
  ): FormulaValue {
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

  function evaluateRow(
    args: FormulaAst[],
    dependencies: Set<CellKey>,
  ): FormulaValue {
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

  function evaluateColumn(
    args: FormulaAst[],
    dependencies: Set<CellKey>,
  ): FormulaValue {
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

  function evaluateIndex(
    args: FormulaAst[],
    dependencies: Set<CellKey>,
  ): FormulaValue {
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

    const height = arrayValue.cells.length;
    const width = arrayValue.cells[0]?.length ?? 0;
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

    const targetCell = arrayValue.cells[resolvedRow - 1]?.[resolvedColumn - 1];

    if (!targetCell) {
      return createErrorValue("REF");
    }

    return {
      type: "range",
      cells: [[targetCell]],
    };
  }

  function evaluateMatch(
    args: FormulaAst[],
    dependencies: Set<CellKey>,
  ): FormulaValue {
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
      const cellValue = evaluateCell(
        lookupVector[index].rowIndex,
        lookupVector[index].columnIndex,
      ).value;

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

  function evaluateXLookup(
    args: FormulaAst[],
    dependencies: Set<CellKey>,
  ): FormulaValue {
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
      const candidateValue = evaluateCell(
        lookupVector[index].rowIndex,
        lookupVector[index].columnIndex,
      ).value;

      if (candidateValue.type === "error") {
        return candidateValue;
      }

      if (compareScalarValues(candidateValue, lookupValue) === 0) {
        return evaluateCell(
          returnVector[index].rowIndex,
          returnVector[index].columnIndex,
        ).value;
      }
    }

    if (args[3]) {
      return evaluateAst(args[3], dependencies);
    }

    return createErrorValue("NA");
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
    ["CHOOSE", evaluateChoose],
    ["ROW", evaluateRow],
    ["COLUMN", evaluateColumn],
    ["INDEX", evaluateIndex],
    ["MATCH", evaluateMatch],
    ["XLOOKUP", evaluateXLookup],
  ]);

  for (let rowIndex = 0; rowIndex < rowCount; rowIndex += 1) {
    for (let columnIndex = 0; columnIndex < columnCount; columnIndex += 1) {
      evaluateCell(rowIndex, columnIndex);
    }
  }

  return {
    sheetId: sheet.id,
    workbookVersion,
    cells,
    dependents,
    precedents,
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
  return (
    typeof value === "object" &&
    value !== null &&
    "type" in value &&
    value.type === "error"
  );
}

function getDisplayForValue(value: ScalarFormulaValue): string {
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
      return value.value;
  }
}

function formatNumericDisplay(value: number): string {
  const normalizedValue = Object.is(value, -0) ? 0 : value;

  return String(normalizedValue);
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
