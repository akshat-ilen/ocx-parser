import * as XLSX from "xlsx";
import { ContextData, Share, SheetStakeholder, Stakeholder } from "../types";
import { findHeaderCell } from "../utils";
import {
  AVAILABLE_FOR_GRANT_STAKEHOLDER,
  STAKEHOLDER_TABLE_HEADERS,
} from "../constants";

function extractStakeholderTable(
  sheet: XLSX.WorkSheet,
  startCell: { r: number; c: number },
  rowLength?: number
): { headers: string[]; data: any[] } {
  const headers: string[] = [];
  const data: any[] = [];
  const range = XLSX.utils.decode_range(sheet["!ref"] || "");
  const startRow = startCell.r;
  let startCol = startCell.c;

  // Extract headers
  for (let C = startCol; C <= range.e.c; C++) {
    const headerCell = sheet[XLSX.utils.encode_cell({ r: startRow + 1, c: C })];
    if (headerCell && headerCell.v) {
      headers.push(headerCell.v.toString().trim());
    } else {
      // Stop when we hit an empty column
      if (headers.length > 0) {
        break;
      }
    }
  }

  // Extract data
  for (let R = startRow + 2; R <= range.e.r; R++) {
    const firstCell = sheet[XLSX.utils.encode_cell({ r: R, c: startCol })];
    if (rowLength) {
      if (rowLength && R > startRow + rowLength + 1) {
        break;
      }
    } else {
      if (
        firstCell &&
        firstCell.v &&
        firstCell.v.toString().trim().toLowerCase() === "total"
      ) {
        break;
      }
    }

    const row: any = {};
    headers.forEach((header, index) => {
      const cell = sheet[XLSX.utils.encode_cell({ r: R, c: startCol + index })];
      row[header] = cell ? cell.v : null;
    });

    if (
      rowLength ||
      Object.values(row).some(
        (value) => value !== null && value !== undefined && value !== ""
      )
    ) {
      data.push(row);
    }
  }

  return { headers, data };
}

function extractTables(sheet: XLSX.WorkSheet): {
  capTable: any;
  additionalTable: any;
} {
  const capHeaderCell = findHeaderCell(
    sheet,
    STAKEHOLDER_TABLE_HEADERS.STAKEHOLDER_CAP
  );
  const additionalHeaderCell = findHeaderCell(
    sheet,
    STAKEHOLDER_TABLE_HEADERS.ADDITIONAL_DETAILS
  );

  if (!capHeaderCell || !additionalHeaderCell) {
    throw new Error("Stakeholder headers not found");
  }

  const capTable = extractStakeholderTable(sheet, capHeaderCell);
  const additionalTable = extractStakeholderTable(
    sheet,
    additionalHeaderCell,
    capTable.data.length
  );

  return {
    capTable,
    additionalTable,
  };
}

function parseStakeholderSheet(
  sheet: XLSX.WorkSheet,
  context: ContextData
): SheetStakeholder[] {
  const stakeholders: SheetStakeholder[] = [];

  const { capTable, additionalTable } = extractTables(sheet);

  capTable.data.forEach((capRow: { [key: string]: any }, index: number) => {
    const stakeholder: SheetStakeholder = {
      name: capRow[capTable.headers[0]],
      group: capRow[capTable.headers[1]],
      stocksByRound: {},
      stocksByStockPlanHistory: {},
      additionalDetails: {},
    };

    capTable.headers.slice(2).forEach((header: string) => {
      const round = context.capRounds.find((round) =>
        header.includes(round.name)
      )?.name;
      if (round) {
        stakeholder.stocksByRound[round] = capRow[header] || 0;
        return;
      }

      const stockPlanDetails = context.stockPlans.find((stockPlan) =>
        header.includes(stockPlan.name)
      )?.name;
      if (stockPlanDetails) {
        stakeholder.stocksByStockPlanHistory[stockPlanDetails] =
          capRow[header] || 0;

        return;
      }
    });

    additionalTable.headers.forEach((header: string) => {
      stakeholder.additionalDetails[header] =
        additionalTable.data[index][header];
    });

    stakeholders.push(stakeholder);
  });

  return stakeholders;
}

function transformToStakeholder(stakeholder: SheetStakeholder): Stakeholder {
  return {
    name: stakeholder.name,
    group: stakeholder.group,
    sharesByRound: Object.keys(stakeholder.stocksByRound).map((round) => ({
      name: round,
      shares: stakeholder.stocksByRound[round],
    })),
    sharesByStockPlan: Object.keys(stakeholder.stocksByStockPlanHistory).map(
      (stockPlan) => ({
        name: stockPlan,
        shares: stakeholder.stocksByStockPlanHistory[stockPlan],
      })
    ),
    additionaDetails: {
      primaryStakeholderType:
        stakeholder.additionalDetails["Primary Stakeholder Type"],
      secondaryStakeholderType:
        stakeholder.additionalDetails["Secondary Stakeholder Types"],
      address: {
        addressLine1: stakeholder.additionalDetails["Mailing Address Line 1"],
        addressLine2: stakeholder.additionalDetails["Mailing Address Line 2"],
        city: stakeholder.additionalDetails["City"],
        state: stakeholder.additionalDetails["State"],
        countryCode: stakeholder.additionalDetails["Country"],
        postalCode: stakeholder.additionalDetails["Zip Code"],
      },
      email: stakeholder.additionalDetails["Email Address"],
      notes: stakeholder.additionalDetails["Notes"],
    },
  };
}

export function parseStakeholders(
  sheet: XLSX.WorkSheet,
  context: ContextData
): { stakeholders: Stakeholder[]; availableForGrant: Share[] } {
  const stakeholdersFromSheets = parseStakeholderSheet(sheet, context);

  let availableForGrant: Share[] = [];
  const stakeholders: Stakeholder[] = [];

  stakeholdersFromSheets.forEach((stakeholderFromSheet) => {
    const stakeholder = transformToStakeholder(stakeholderFromSheet);
    if (stakeholder.name === AVAILABLE_FOR_GRANT_STAKEHOLDER) {
      availableForGrant = stakeholder.sharesByStockPlan;
    } else {
      stakeholders.push(stakeholder);
    }
  });

  return { stakeholders, availableForGrant };
}
