import * as XLSX from "xlsx";
import { OCX_VERSION_HEADER, STAKEHOLDER_TABLE_HEADERS } from "../constants";

export function findHeaderCell(
  sheet: XLSX.WorkSheet,
  header: string
): { r: number; c: number } | null {
  const range = XLSX.utils.decode_range(sheet["!ref"] || "");
  for (let R = range.s.r; R <= range.e.r; ++R) {
    for (let C = range.s.c; C <= range.e.c; ++C) {
      const cellAddress = XLSX.utils.encode_cell({ r: R, c: C });
      const cell = sheet[cellAddress];
      if (cell && cell.v && cell.v.toString().trim().includes(header)) {
        return { r: R, c: C };
      }
    }
  }
  return null;
}

export function calculateColSpan(
  sheet: XLSX.WorkSheet,
  cell: { r: number; c: number }
): number {
  const mergeRanges = sheet["!merges"] || [];
  for (const range of mergeRanges) {
    if (range.s.r === cell.r && range.s.c === cell.c) {
      return range.e.c - range.s.c + 1;
    }
  }
  return 1; // Default to 1 if no merge is found
}

export function formatDateString(value: any): string {
  if (typeof value === "number") {
    // Assume it's an Excel date serial number
    const date = XLSX.SSF.parse_date_code(value);
    if (date) {
      return `${date.y}-${String(date.m).padStart(2, "0")}-${String(
        date.d
      ).padStart(2, "0")}`;
    }
  } else if (typeof value === "string") {
    // Assume it's already a valid date string
    const date = new Date(value);
    if (!isNaN(date.getTime())) {
      return date.toISOString().split("T")[0]; // Return YYYY-MM-DD format
    }
  }
  return value;
}

export function isDateFormat(cell: XLSX.CellObject): boolean {
  if (cell && cell.z) {
    const formatString = cell.z as string;
    return (
      formatString.toLowerCase().includes("d") ||
      formatString.toLowerCase().includes("m") ||
      formatString.toLowerCase().includes("y")
    );
  }
  return false;
}

export function getCurrencySymbol(cell: XLSX.CellObject): string | null {
  if (cell && cell.z) {
    const formatString = cell.z as string;
    const currencySymbols = ["$", "€", "£", "¥", "₹"];
    for (const symbol of currencySymbols) {
      if (formatString.includes(symbol)) {
        return symbol;
      }
    }
  }
  return null;
}

export function extractOcxVersion(sheet: XLSX.WorkSheet): string | null {
  const range = XLSX.utils.decode_range(sheet["!ref"] || "");

  for (let R = range.s.r; R <= range.e.r; ++R) {
    for (let C = range.s.c; C <= range.e.c; ++C) {
      const cellAddress = XLSX.utils.encode_cell({ r: R, c: C });
      const cell = sheet[cellAddress];
      if (
        cell &&
        cell.v &&
        typeof cell.v === "string" &&
        cell.v.includes(OCX_VERSION_HEADER)
      ) {
        const versionMatch = cell.v.match(
          new RegExp(`${OCX_VERSION_HEADER}\\s*(.*)`)
        );
        if (versionMatch) {
          return versionMatch[1].trim();
        }
      }
    }
  }

  return null;
}

export function extractEntityName(sheet: XLSX.WorkSheet): string {
  const range = XLSX.utils.decode_range(sheet["!ref"] || "");

  for (let R = range.s.r; R <= range.e.r; ++R) {
    for (let C = range.s.c; C <= range.e.c; ++C) {
      const cellAddress = XLSX.utils.encode_cell({ r: R, c: C });
      const cell = sheet[cellAddress];
      if (
        cell &&
        cell.v &&
        typeof cell.v === "string" &&
        cell.v.includes(STAKEHOLDER_TABLE_HEADERS.STAKEHOLDER_CAP)
      ) {
        const entityName = cell.v
          .replace(STAKEHOLDER_TABLE_HEADERS.STAKEHOLDER_CAP, "")
          .trim();
        return entityName;
      }
    }
  }

  return "";
}
