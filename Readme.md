# OCX Parser

OCX Parser is a TypeScript/JavaScript library for parsing Open Cap Table (OCX) files. It provides a simple interface to extract cap table information from OCX files in both Node.js and browser environments.

## Installation

You can install the OCX Parser using npm:

```bash
npm install ocx-parser
```

Or using yarn:

```bash
yarn add ocx-parser
```

## Usage

### Node.js

Here's an example of how to use the OCX Parser in a Node.js environment:

```javascript
import { createOCXParser } from "ocx-parser";
import fs from "fs/promises";

async function parseOCXFile(filePath) {
  const parser = createOCXParser();

  try {
    const fileBuffer = await fs.readFile(filePath);
    const result = await parser.parseCapTable(fileBuffer);

    if (result.status === "success") {
      console.log("Parsing successful:", result.data);
    } else {
      console.error("Parsing failed:", result.error);
    }
  } catch (error) {
    console.error("Error reading or parsing file:", error);
  }
}

// Usage
parseOCXFile("/path/to/your/ocx/file.xlsx");
```

### Browser

In a browser environment, you can use the OCX Parser with file input:

```html
<!DOCTYPE html>
<html>
  <body>
    <input type="file" id="fileInput" accept=".xlsx" />
    <script type="module">
      import { createOCXParser } from "ocx-parser";

      const fileInput = document.getElementById("fileInput");
      fileInput.addEventListener("change", async (event) => {
        const file = event.target.files[0];
        if (file) {
          const parser = createOCXParser();
          try {
            const result = await parser.parseCapTable(file);
            if (result.status === "success") {
              console.log("Parsing successful:", result.data);
            } else {
              console.error("Parsing failed:", result.error);
            }
          } catch (error) {
            console.error("Error parsing file:", error);
          }
        }
      });
    </script>
  </body>
</html>
```

## API

### `createOCXParser()`

Creates and returns an instance of the OCX parser.

### `parser.parseCapTable(file: File | Buffer): Promise<ParseResult>`

Parses the given OCX file and returns a Promise that resolves to a `ParseResult` object.

- `file`: Can be a `File` object (in browser environments) or a `Buffer` (in Node.js environments).
- Returns: `Promise<ParseResult>`

### `ParseResult`

The `ParseResult` object has the following structure:

```typescript
interface ParseResult {
  status: "success" | "error";
  data: CapTable | null;
  error: Error | null;
}
```

- `status`: Indicates whether the parsing was successful or encountered an error.
- `data`: Contains the parsed cap table data if successful, null otherwise.
- `error`: Contains the error object if an error occurred, null otherwise.

### `CapTable`

The `CapTable` object represents the parsed data from the OCX file. It includes information about the OCX version, context data, stakeholders, and available shares for grant.

## Error Handling

The OCX Parser uses a `ParseResult` object to handle both successful parsing and errors. Always check the `status` of the `ParseResult` before accessing the `data` or `error` properties.

## Dependencies

This package depends on `xlsx` for parsing Excel files. Make sure it's properly installed in your project.

## License

This project is licensed under the MIT License - see the LICENSE file for details.
