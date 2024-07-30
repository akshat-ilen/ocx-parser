import * as XLSX from "xlsx";
import {
  type FinancialHistory,
  type SheetValuation,
  type StockPlanHistory,
  type StockPlanDetails,
  type ContextData,
  type CapRound,
  type StockPlan,
  StockClass,
  Valuation,
} from "../types";

import {
  findHeaderCell,
  calculateColSpan,
  formatDateString,
  getCurrencySymbol,
  isDateFormat,
} from "../utils";

import { CONTEXT_TABLE_HEADERS } from "../constants";

function extractContextTable(
  sheet: XLSX.WorkSheet,
  startCell: { r: number; c: number },
  colSpan: number
): any[] {
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
      break;
    }

    const row: any = {};
    for (let C = startCol; C < startCol + colSpan; C++) {
      const cell = sheet[XLSX.utils.encode_cell({ r: R, c: C })];
      let cellValue = cell ? cell.v : null;

      // Check if the cell is a date format and format it
      if (isDateFormat(cell) && cellValue) {
        cellValue = formatDateString(cellValue);
      }

      const currency = getCurrencySymbol(cell);

      if (currency) {
        row[headers[C - startCol]] = {
          value: cellValue,
          currency,
        };
      } else {
        row[headers[C - startCol]] = cellValue;
      }
    }
    data.push(row);
  }

  return data;
}

function parseContextSheet(
  sheet: XLSX.WorkSheet
): Omit<ContextData, "capRounds" | "stockPlans" | "valuations"> {
  const rounds: FinancialHistory[] = [];
  const valuations: SheetValuation[] = [];
  const stockPlanHistories: StockPlanHistory[] = [];
  const stockPlanDetails: StockPlanDetails[] = [];

  const tableHeaders = Object.values(CONTEXT_TABLE_HEADERS);

  tableHeaders.forEach((header) => {
    const headerCell = findHeaderCell(sheet, header);
    if (headerCell) {
      const colSpan = calculateColSpan(sheet, headerCell);
      const tableData = extractContextTable(sheet, headerCell, colSpan);
      switch (header) {
        case CONTEXT_TABLE_HEADERS.ROUNDS:
          rounds.push(...tableData);
          break;
        case CONTEXT_TABLE_HEADERS.VALUATIONS:
          valuations.push(...tableData);
          break;
        case CONTEXT_TABLE_HEADERS.STOCK_PLAN_HISTORIES:
          stockPlanHistories.push(...tableData);
          break;
        case CONTEXT_TABLE_HEADERS.STOCK_PLAN_DETAILS:
          stockPlanDetails.push(...tableData);
          break;
      }
    }
  });

  return {
    rounds,
    sheetValuations: valuations,
    stockPlanHistories,
    stockPlanDetails,
  };
}

function financialHistoryToRound(financialHistory: FinancialHistory): CapRound {
  return {
    name: financialHistory["Round"],
    stockClass: financialHistory["Round"].toLowerCase().includes("preferred")
      ? StockClass.PREFERRED
      : StockClass.COMMOM,
    closingDate: financialHistory["Initial Closing Date"],
    issuePrice: financialHistory["Original Issue Price"],
    liquidationMultiple: financialHistory["Liquidation Multiple"],
  };
}

function getStockPlan(stockPlanHistories: StockPlanHistory): StockPlan {
  return {
    name: stockPlanHistories["Plan Name"],
    date: stockPlanHistories["Date"],
  };
}

function getValuation(valuation: SheetValuation): Valuation {
  const result: Valuation = {
    date: "",
    pricePerShare: 0,
    firm: "",
  };

  Object.keys(valuation).forEach((key) => {
    if (key.toLowerCase().includes("date")) {
      result.date = valuation[key];
    } else if (key.toLowerCase().includes("price")) {
      result.pricePerShare = valuation[key];
    } else if (key.toLowerCase().includes("firm")) {
      result.firm = valuation[key];
    }
  });

  return result;
}

export function parseContext(contextSheet: XLSX.WorkSheet): ContextData {
  const contextData = parseContextSheet(contextSheet);

  const capRounds = contextData.rounds.map(financialHistoryToRound);
  const valuations = contextData.sheetValuations.map(getValuation);
  const stockPlans = contextData.stockPlanHistories.map(getStockPlan);

  return { ...contextData, capRounds, stockPlans, valuations };
}
