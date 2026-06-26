import type { SheetSummary } from "./workbook-core";

export function filterSheetQuickOpenResults(
  sheets: readonly SheetSummary[],
  query: string,
): SheetSummary[] {
  const normalizedQuery = normalizeSheetQuickOpenText(query);

  if (!normalizedQuery) {
    return [...sheets];
  }

  const scoredSheets: Array<{ index: number; score: number; sheet: SheetSummary }> = [];

  sheets.forEach((sheet, index) => {
    const score = getSheetQuickOpenScore(sheet.name, normalizedQuery);

    if (score === null) {
      return;
    }

    scoredSheets.push({ index, score, sheet });
  });

  return scoredSheets
    .sort((left, right) => left.score - right.score || left.index - right.index)
    .map((entry) => entry.sheet);
}

function getSheetQuickOpenScore(name: string, normalizedQuery: string): number | null {
  const normalizedName = normalizeSheetQuickOpenText(name);

  if (normalizedName === normalizedQuery) {
    return 0;
  }

  if (normalizedName.startsWith(normalizedQuery)) {
    return 1;
  }

  if (
    normalizedName.split(/\s+/).some((part) => part.length > 0 && part.startsWith(normalizedQuery))
  ) {
    return 2;
  }

  return normalizedName.includes(normalizedQuery) ? 3 : null;
}

function normalizeSheetQuickOpenText(value: string): string {
  return value.trim().toLocaleLowerCase();
}
