import * as XLSX from "xlsx";
import { calculateColSpan, findHeaderCell, getCurrencySymbol } from "../utils";
import { SUMMARY_TABLE_HEADERS, SHARE_CLASSES_IN_SUMMARY } from "../constants";
import { Securities, ShareClass, ShareClassType, Summary } from "../types";

function parseSummaryTable(sheet: XLSX.WorkSheet): any[] {
  const header = SUMMARY_TABLE_HEADERS.SUMMARY_CAP;
  const startCell = findHeaderCell(sheet, header);
  if (!startCell) {
    throw new Error(`Header "${header}" not found.`);
  }

  const colSpan = calculateColSpan(sheet, startCell);
  const startRow = startCell.r;
  const startCol = startCell.c;

  const headers: string[] = [];
  for (let C = startCol; C < startCol + colSpan; C++) {
    const headerCell = sheet[XLSX.utils.encode_cell({ r: startRow + 1, c: C })];
    if (headerCell && headerCell.v) {
      headers.push(headerCell.v.toString().trim());
    }
  }

  const shareClasses = SHARE_CLASSES_IN_SUMMARY;
  const data: any[] = [];
  let currentShareClass: string | null = null;

  for (
    let R = startRow + 2;
    R <= XLSX.utils.decode_range(sheet["!ref"] || "").e.r;
    R++
  ) {
    const firstCell = sheet[XLSX.utils.encode_cell({ r: R, c: startCol })];
    if (!firstCell || !firstCell.v) {
      if (currentShareClass) {
        // End of the current share class if the row is empty
        currentShareClass = null;
      }
      continue;
    }

    const cellValue = firstCell.v.toString().trim();
    if (cellValue.toLowerCase() === "total") {
      // End of the main table if "Total" is found in the first column
      break;
    }

    if (
      shareClasses.some((sc) =>
        cellValue.toLowerCase().includes(sc.toLowerCase())
      )
    ) {
      currentShareClass = cellValue;
      continue;
    }

    if (currentShareClass) {
      const row: any = { shareClass: currentShareClass };
      let colIndex = 0;
      for (let C = startCol; C < startCol + colSpan; C++) {
        const cell = sheet[XLSX.utils.encode_cell({ r: R, c: C })];
        if (headers[colIndex] !== "") {
          // Ignore empty columns
          row[headers[colIndex]] = cell ? cell.v : null;
        }
        colIndex++;
      }
      data.push(row);
    }
  }

  return data;
}

function parseOutstandingConvertibleSecurities(sheet: XLSX.WorkSheet): any[] {
  const header = SUMMARY_TABLE_HEADERS.SECURTIES;
  const startCell = findHeaderCell(sheet, header);
  if (!startCell) {
    throw new Error(`Header "${header}" not found.`);
  }

  const colSpan = calculateColSpan(sheet, startCell);
  const startRow = startCell.r;
  const startCol = startCell.c;

  const headers: string[] = [];
  for (let C = startCol; C < startCol + colSpan; C++) {
    const headerCell = sheet[XLSX.utils.encode_cell({ r: startRow + 1, c: C })];
    if (headerCell && headerCell.v) {
      headers.push(headerCell.v.toString().trim());
    }
  }

  const data: any[] = [];
  for (
    let R = startRow + 2;
    R <= XLSX.utils.decode_range(sheet["!ref"] || "").e.r;
    R++
  ) {
    const firstCell = sheet[XLSX.utils.encode_cell({ r: R, c: startCol })];
    if (!firstCell || !firstCell.v) {
      continue;
    }

    const cellValue = firstCell.v.toString().trim();
    if (cellValue.toLowerCase() === "total") {
      break;
    }

    const row: any = {};
    let colIndex = 0;
    for (let C = startCol; C < startCol + colSpan; C++) {
      const cell = sheet[XLSX.utils.encode_cell({ r: R, c: C })];
      if (headers[colIndex] !== "") {
        const currency = getCurrencySymbol(cell);
        if (currency) {
          row[headers[colIndex]] = {
            value: cell.v,
            currency,
          };
        } else {
          row[headers[colIndex]] = cell ? cell.v : null;
        }
      }
      colIndex++;
    }
    data.push(row);
  }

  return data;
}

function extractShareClasses(data: any): ShareClass {
  const result: ShareClass = {
    name: data["Share Class"],
    type: ShareClassType.COMMON,
    authorisedShares: data["Shares Authorized/\nReserved"] ?? null,
    outshandingShares: data["Outstanding Shares*"] ?? null,
    fullDilutedShares: data["Fully Diluted Shares**"] ?? null,
    dilutePercentage: data["% Fully Diluted**"] ?? null,
    votingMutiplier: data["Voting\nMultiplier"] ?? null,
  };

  const shareClass = data.shareClass?.toLowerCase();
  let shareClassType: ShareClassType = ShareClassType.COMMON;

  if (shareClass) {
    if (shareClass.includes("common stock")) {
      shareClassType = ShareClassType.COMMON;
    } else if (shareClass.includes("preferred stock")) {
      shareClassType = ShareClassType.PREFERRED;
    } else if (shareClass.includes("warrants & non-plan awards")) {
      shareClassType = ShareClassType.WARRANT;
    } else {
      shareClassType = ShareClassType.STOCK_PLAN;
    }
  }

  result.type = shareClassType;

  return result;
}

function extractSecurities(data: any): Securities {
  const security: Securities = {
    name: "",
    numberOfSecurities: 0,
    outstandingAmount: {
      value: 0,
      currency: "",
    },
    discount: 0,
    valuationCap: {
      value: 0,
      currency: "",
    },
  };

  Object.keys(data).forEach((key) => {
    if (key.toLowerCase().includes("security type")) {
      security.name = data[key];
    } else if (key.toLowerCase().includes("#")) {
      security.numberOfSecurities = data[key];
    } else if (key.toLowerCase().includes("outstanding")) {
      security.outstandingAmount = data[key];
    } else if (key.toLowerCase().includes("discount")) {
      security.discount = data[key];
    } else if (key.toLowerCase().includes("valuation cap")) {
      security.valuationCap = data[key];
    }
  });

  return security;
}

export function parseSummary(sheet: XLSX.WorkSheet): Summary {
  const summaryData = parseSummaryTable(sheet);
  const securitiesData = parseOutstandingConvertibleSecurities(sheet);

  const capTableSummary = summaryData.map(extractShareClasses);
  const securities = securitiesData.map(extractSecurities);

  return { capTableSummary, securities };
}
