import {
  getSheetColumnCount,
  getSheetRowCount,
  isFormulaInput,
  parseCellReference,
  type FormulaErrorCode,
  type WorkbookSheet,
} from "./workbook-core";

export type CellKey = `${number}:${number}`;

export interface CellEvaluation {
  input: string;
  display: string;
  isFormula: boolean;
  numericValue?: number;
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
  | { type: "cell"; rowIndex: number; columnIndex: number }
  | { type: "leftParen" }
  | { type: "number"; value: number }
  | { type: "operator"; value: "+" | "-" | "*" | "/" }
  | { type: "rightParen" };

type FormulaAst =
  | {
      type: "binary";
      operator: "+" | "-" | "*" | "/";
      left: FormulaAst;
      right: FormulaAst;
    }
  | { type: "number"; value: number }
  | { type: "reference"; rowIndex: number; columnIndex: number }
  | { type: "unary"; operator: "+" | "-"; operand: FormulaAst };

type NumericEvaluationResult = {
  errorCode?: FormulaErrorCode;
  numericValue?: number;
};

const EMPTY_DEPENDENCIES: CellKey[] = [];
const NUMBER_LITERAL_PATTERN = /^[+-]?(?:\d+(?:\.\d+)?|\.\d+)$/;

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

    if (
      character === "+" ||
      character === "-" ||
      character === "*" ||
      character === "/"
    ) {
      tokens.push({
        type: "operator",
        value: character,
      });
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

    if (/[A-Za-z]/.test(character)) {
      const columnStart = index;

      while (index < expression.length && /[A-Za-z]/.test(expression[index])) {
        index += 1;
      }

      const rowStart = index;

      while (index < expression.length && /[0-9]/.test(expression[index])) {
        index += 1;
      }

      if (rowStart === index) {
        throw new Error(
          `Invalid cell reference near "${expression.slice(columnStart)}".`,
        );
      }

      const reference = expression.slice(columnStart, index);
      const parsedReference = parseCellReference(reference);

      tokens.push({
        type: "cell",
        ...parsedReference,
      });
      continue;
    }

    const numberMatch = /^(\d+(?:\.\d+)?|\.\d+)/.exec(expression.slice(index));

    if (numberMatch) {
      tokens.push({
        type: "number",
        value: Number(numberMatch[1]),
      });
      index += numberMatch[1].length;
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
    let node = parseTerm();
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
        right: parseTerm(),
      };
      token = tokens[index];
    }

    return node;
  }

  function parseTerm(): FormulaAst {
    let node = parseUnary();
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
        right: parseUnary(),
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

    return parsePrimary();
  }

  function parsePrimary(): FormulaAst {
    const token = tokens[index];

    if (!token) {
      throw new Error("Formula ended unexpectedly.");
    }

    if (token.type === "number") {
      index += 1;
      return token;
    }

    if (token.type === "cell") {
      index += 1;
      return {
        type: "reference",
        rowIndex: token.rowIndex,
        columnIndex: token.columnIndex,
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
        evaluation = {
          input,
          display: input,
          isFormula: false,
          numericValue: parseRawNumericValue(input),
          dependencies: EMPTY_DEPENDENCIES,
        };
      } else {
        evaluation = evaluateFormulaCell(input, cellKey);
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
  ): CellEvaluation {
    let ast: FormulaAst;

    try {
      ast = parseFormula(input);
    } catch {
      return createErrorEvaluation(input, "PARSE", EMPTY_DEPENDENCIES);
    }

    const dependencies = new Set<CellKey>();
    const result = evaluateAst(ast, dependencies);
    const normalizedDependencies = [...dependencies];
    const errorCode = cycleCellKeys.has(cellKey) ? "CYCLE" : result.errorCode;

    if (errorCode) {
      return createErrorEvaluation(input, errorCode, normalizedDependencies);
    }

    const numericValue = result.numericValue ?? 0;

    return {
      input,
      display: formatNumericDisplay(numericValue),
      isFormula: true,
      numericValue,
      dependencies: normalizedDependencies,
    };
  }

  function evaluateAst(
    ast: FormulaAst,
    dependencies: Set<CellKey>,
  ): NumericEvaluationResult {
    switch (ast.type) {
      case "number":
        return {
          numericValue: ast.value,
        };
      case "reference": {
        if (
          ast.rowIndex < 0 ||
          ast.rowIndex >= rowCount ||
          ast.columnIndex < 0 ||
          ast.columnIndex >= columnCount
        ) {
          return {
            errorCode: "REF",
          };
        }

        const dependencyKey = createCellKey(ast.rowIndex, ast.columnIndex);

        dependencies.add(dependencyKey);

        const referencedCell = evaluateCell(ast.rowIndex, ast.columnIndex);

        if (referencedCell.errorCode) {
          return {
            errorCode: referencedCell.errorCode,
          };
        }

        if (referencedCell.numericValue === undefined) {
          return {
            errorCode: "VALUE",
          };
        }

        return {
          numericValue: referencedCell.numericValue,
        };
      }
      case "unary": {
        const operandResult = evaluateAst(ast.operand, dependencies);

        if (operandResult.errorCode) {
          return operandResult;
        }

        const numericValue = operandResult.numericValue ?? 0;

        return {
          numericValue: ast.operator === "-" ? -numericValue : numericValue,
        };
      }
      case "binary": {
        const leftResult = evaluateAst(ast.left, dependencies);

        if (leftResult.errorCode) {
          return leftResult;
        }

        const rightResult = evaluateAst(ast.right, dependencies);

        if (rightResult.errorCode) {
          return rightResult;
        }

        const leftValue = leftResult.numericValue ?? 0;
        const rightValue = rightResult.numericValue ?? 0;

        switch (ast.operator) {
          case "+":
            return {
              numericValue: leftValue + rightValue,
            };
          case "-":
            return {
              numericValue: leftValue - rightValue,
            };
          case "*":
            return {
              numericValue: leftValue * rightValue,
            };
          case "/":
            if (rightValue === 0) {
              return {
                errorCode: "DIV0",
              };
            }

            return {
              numericValue: leftValue / rightValue,
            };
        }
      }
    }
  }

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

function createErrorEvaluation(
  input: string,
  errorCode: FormulaErrorCode,
  dependencies: readonly CellKey[],
): CellEvaluation {
  return {
    input,
    display: getErrorDisplay(errorCode),
    isFormula: isFormulaInput(input),
    errorCode,
    dependencies: [...dependencies],
  };
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
  }
}

function parseRawNumericValue(input: string): number | undefined {
  if (input.length === 0) {
    return 0;
  }

  if (!NUMBER_LITERAL_PATTERN.test(input)) {
    return undefined;
  }

  return Number(input);
}
