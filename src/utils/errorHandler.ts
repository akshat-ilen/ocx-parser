import { ParseResult, ParseStatus, OCXParserError } from "../types";

export function handleError(error: unknown): ParseResult {
  const errorMessage = error instanceof Error ? error.message : String(error);

  return {
    status: ParseStatus.ERROR,
    data: null,
    error: new OCXParserError(errorMessage),
  };
}
