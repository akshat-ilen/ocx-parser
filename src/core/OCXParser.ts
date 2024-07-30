import * as XLSX from "xlsx";
import { ParseResult, ParseStatus, CapTable } from "../types/parser";
import { OCXParserError } from "../types/common";
import { parseContext } from "../parsers/context";
import { parseStakeholders } from "../parsers/stakeholder";
import { extractEntityName, extractOcxVersion } from "../utils";
import { SUPPORTED_VERSIONS } from "../constants/versions";
import { SHEETS } from "../constants";
import { parseSummary } from "../parsers";

export class OCXParser {
  public async parseCapTable(file: File | Buffer): Promise<ParseResult> {
    try {
      let arrayBuffer: ArrayBuffer;
      if (file instanceof File) {
        arrayBuffer = await file.arrayBuffer();
      } else {
        arrayBuffer = file.buffer.slice(
          file.byteOffset,
          file.byteOffset + file.byteLength
        );
      }

      const workbook = XLSX.read(new Uint8Array(arrayBuffer), {
        type: "array",
        cellNF: true,
      });

      // Extract sheets
      const contextSheet = workbook.Sheets[SHEETS.CONTEXT];
      const stakeholderSheet = workbook.Sheets[SHEETS.STAKEHOLDERS];
      const summarySheet = workbook.Sheets[SHEETS.SUMMARY];

      const ocxVersion = extractOcxVersion(contextSheet);
      if (!ocxVersion || !SUPPORTED_VERSIONS.includes(ocxVersion)) {
        throw new OCXParserError(`Unsupported OCX version: ${ocxVersion}`);
      }

      const summary = parseSummary(summarySheet);
      const contextData = parseContext(contextSheet);
      const { stakeholders, availableForGrant } = parseStakeholders(
        stakeholderSheet,
        contextData
      );

      const entityName = extractEntityName(stakeholderSheet);

      const capTable: CapTable = {
        ocxVersion,
        entityName,
        rounds: contextData.capRounds,
        stockPlans: contextData.stockPlans,
        stakeholders,
        availableForGrant,
        summary,
        valuations: contextData.valuations,
      };

      return {
        status: ParseStatus.SUCCESS,
        data: capTable,
        error: null,
      };
    } catch (error) {
      return {
        status: ParseStatus.ERROR,
        data: null,
        error: new OCXParserError(
          error instanceof Error ? error.message : String(error)
        ),
      };
    }
  }
}

export function createOCXParser(): OCXParser {
  return new OCXParser();
}
