import type { FlatJsonRow } from "./leanix/types";

export async function writeTestCell(): Promise<void> {
  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    sheet.getRange("A1").values = [["LeanIX Add-in Connected"]];
    await context.sync();
  });
}

export async function writeJsonArrayToWorksheet(value: unknown): Promise<void> {
  const rows = toFlatRows(value);
  if (!rows) {
    throw new Error("This object cannot yet be converted automatically to an Excel table.");
  }

  const headers = Array.from(new Set(rows.flatMap((row) => Object.keys(row))));
  const values = [
    headers.map(toTitle),
    ...rows.map((row) => headers.map((header) => row[header] ?? ""))
  ];

  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const existingTables = sheet.tables;
    existingTables.load("items/name");
    await context.sync();

    for (const table of existingTables.items) {
      if (table.name === "LeanIXStorageData") {
        table.delete();
      }
    }

    const range = sheet.getRangeByIndexes(0, 0, values.length, headers.length);
    range.values = values;
    const table = sheet.tables.add(range, true);
    table.name = "LeanIXStorageData";
    table.getRange().format.autofitColumns();
    await context.sync();
  });
}

function toFlatRows(value: unknown): FlatJsonRow[] | null {
  if (!Array.isArray(value) || value.length === 0) {
    return null;
  }

  const rows = value.filter(isFlatObject);
  return rows.length === value.length ? rows : null;
}

function isFlatObject(value: unknown): value is FlatJsonRow {
  if (!value || typeof value !== "object" || Array.isArray(value)) {
    return false;
  }

  return Object.values(value).every(
    (item) =>
      item === null ||
      typeof item === "string" ||
      typeof item === "number" ||
      typeof item === "boolean"
  );
}

function toTitle(key: string): string {
  return key
    .replace(/([a-z0-9])([A-Z])/g, "$1 $2")
    .replace(/[_-]+/g, " ")
    .replace(/\w\S*/g, (word) => word.charAt(0).toUpperCase() + word.slice(1));
}
