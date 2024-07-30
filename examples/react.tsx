import React, { useState } from "react";
import { createOCXParser, ParseResult, CapTable } from "ocx-parser";

const OCXFileParser: React.FC = () => {
  const [parseResult, setParseResult] = useState<ParseResult | null>(null);

  const handleFileUpload = async (
    event: React.ChangeEvent<HTMLInputElement>
  ) => {
    const file = event.target.files?.[0];
    if (file) {
      const parser = createOCXParser();
      try {
        const result = await parser.parseCapTable(file);
        setParseResult(result);
      } catch (error) {
        console.error("Error parsing file:", error);
        setParseResult({
          status: "error",
          data: null,
          error: new Error("Failed to parse file"),
        });
      }
    }
  };

  const renderCapTable = (capTable: CapTable) => {
    // Render your cap table data here
    return (
      <div>
        <h2>Cap Table Data</h2>
        <pre>{JSON.stringify(capTable, null, 2)}</pre>
      </div>
    );
  };

  return (
    <div>
      <input type="file" onChange={handleFileUpload} accept=".xlsx" />
      {parseResult && (
        <div>
          {parseResult.status === "success" && parseResult.data ? (
            renderCapTable(parseResult.data)
          ) : (
            <p>Error: {parseResult.error?.message}</p>
          )}
        </div>
      )}
    </div>
  );
};

export default OCXFileParser;
