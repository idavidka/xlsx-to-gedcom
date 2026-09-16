# @treeviz/xlsx-to-ged

> Part of the [@treeviz](https://www.npmjs.com/org/treeviz) organization - A collection of tools for genealogy data processing and visualization.

Convert Excel `.xlsx` spreadsheets into GEDCOM 5.5.1 with configurable column mapping.

## Features

- **TreeViz simple layout** — fixed column order (surname, given name, birth, death, sex, parent row numbers, …) works without manual mapping
- **Header guessing** — TreeViz export headers and common hu/en aliases
- **Relationship linking** — parent/spouse row numbers → `FAMC` / `FAM`, optional spouse + marriage on `FAM`
- **Browser-safe API** — `ArrayBuffer` in, GEDCOM text out
- **CLI** — `xlsx-to-ged input.xlsx -o output.ged`

## Installation

```bash
npm install @treeviz/xlsx-to-ged
```

## Development

```bash
npm i
npm run build
npm test
```

Regenerate the committed example spreadsheet:

```bash
npm run generate:example
```

## CLI

```bash
# Header guess (default) — same as the app wizard “Fejléc alapján”
npx xlsx-to-ged input.xlsx -o output.ged

# Fixed TreeViz simple column layout (no header row)
npx xlsx-to-ged input.xlsx -o output.ged --simple

# Pick a worksheet by index or name
npx xlsx-to-ged book.xlsx -o output.ged --sheet 1
npx xlsx-to-ged book.xlsx -o output.ged --sheet People
```

During development (without building):

```bash
npm run dev:xlsx -- examples/simple-family.xlsx -o output.ged
```

## Library API

```ts
import {
	xlsxToGed,
	inspectXlsx,
	DEFAULT_SIMPLE_LAYOUT,
} from "@treeviz/xlsx-to-ged";

const buffer = await file.arrayBuffer();
const inspect = await inspectXlsx(buffer);
const { gedcom, warnings } = await xlsxToGed(buffer); // default simple layout
// or
const { gedcom, warnings } = await xlsxToGed(buffer, inspect.guessed);
```

See the TreeViz monorepo `docs/XLSX_TO_GED_PLAN.md` for the full column mapping spec.

## Example file

[examples/simple-family.xlsx](./examples/simple-family.xlsx) — three-person sheet (row no. = data-row index; regenerate with `npm run generate:example`).

## Requirements

- Node.js 20+
- macOS/Linux/Windows

## License

MIT
