import { mkdir, writeFile } from "node:fs/promises";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";

import { createSimpleTemplateBuffer } from "../src/template";

const root = join(dirname(fileURLToPath(import.meta.url)), "..");
const outDir = join(root, "examples");
const outPath = join(outDir, "simple-family.xlsx");

const buffer = await createSimpleTemplateBuffer();
await mkdir(outDir, { recursive: true });
await writeFile(outPath, Buffer.from(buffer));

console.log(`Wrote ${outPath}`);
