import type { CellDataResult } from "./workbook-core";

export function getFormulaBarPreview(
  cellData: CellDataResult | null,
): string {
  return cellData?.display ?? "idle";
}
