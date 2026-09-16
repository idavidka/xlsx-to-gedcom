# Example spreadsheets

| File | Description |
| --- | --- |
| [simple-family.xlsx](./simple-family.xlsx) | TreeViz **simple layout** with English headers. Row no. = 1-based data-row index (no extra column); column J = optional GEDCOM INDI id. Three people; Péter links father/mother via row nos. `1` and `2`. |

Regenerate after changing `src/template.ts`:

```bash
npm run generate:example --workspace=@treeviz/xlsx-to-ged
```

Quick check with the CLI-less API:

```bash
npx tsx -e "import { readFileSync } from 'node:fs'; import { xlsxToGed } from './src/xlsx-to-ged.ts'; const b = readFileSync('examples/simple-family.xlsx'); const { gedcom } = await xlsxToGed(b); console.log(gedcom.slice(0, 400));"
```

(from the package directory)
