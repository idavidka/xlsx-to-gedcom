#!/usr/bin/env node
import fs from "node:fs/promises";
import path from "node:path";

import yargs from "yargs";
import { hideBin } from "yargs/helpers";

import { DEFAULT_SIMPLE_LAYOUT } from "./constants";
import { inspectXlsx } from "./inspect";
import { xlsxToGed } from "./xlsx-to-ged";

const argv = await yargs(hideBin(process.argv))
	.scriptName("xlsx-to-ged")
	.usage("Usage: $0 <input.xlsx> -o <output.ged> [--simple] [--sheet <index|name>]")
	.command(
		"$0 [input]",
		"Convert XLSX spreadsheet to GEDCOM",
		(yargsBuilder) =>
			yargsBuilder
				.positional("input", {
					describe: "Input XLSX file path",
					type: "string",
				})
				.option("i", {
					alias: "in",
					type: "string",
					describe: "Input XLSX file path (alternative to positional)",
				})
				.option("o", {
					alias: "out",
					type: "string",
					demandOption: true,
					describe: "Output GED file path",
				})
				.option("simple", {
					type: "boolean",
					default: false,
					describe:
						"Use TreeViz simple fixed column layout (skip header guess)",
				})
				.option("sheet", {
					type: "string",
					describe:
						"Worksheet index (0-based) or name (default: first sheet)",
				})
	)
	.help()
	.strict()
	.parse();

const inputPath = (argv.in ?? argv.input) as string | undefined;

if (!inputPath) {
	console.error(
		"Missing input file. Example: xlsx-to-ged input.xlsx -o output.ged"
	);
	process.exit(1);
}

const parseSheetRef = (raw?: string): string | number | undefined => {
	if (raw == null || raw === "") {
		return undefined;
	}
	const asNumber = Number(raw);
	return Number.isNaN(asNumber) ? raw : asNumber;
};

try {
	const buffer = await fs.readFile(inputPath);
	const fileName = path.basename(inputPath);
	const sheet = parseSheetRef(argv.sheet as string | undefined);

	let mapping;
	if (argv.simple) {
		mapping = {
			...DEFAULT_SIMPLE_LAYOUT,
			sheet: sheet ?? DEFAULT_SIMPLE_LAYOUT.sheet,
		};
	} else {
		const inspect = await inspectXlsx(
			buffer,
			sheet != null ? { sheet } : undefined
		);
		mapping = {
			...inspect.guessed,
			sheet: sheet ?? inspect.activeSheet,
		};
	}

	const { gedcom, warnings } = await xlsxToGed(buffer, { mapping, fileName });
	await fs.writeFile(argv.out as string, gedcom, "utf8");

	for (const warning of warnings) {
		console.warn("Warning:", warning);
	}

	console.log(`Done: ${path.resolve(argv.out as string)}`);
} catch (e) {
	console.error("Conversion failed:", e instanceof Error ? e.message : e);
	process.exit(1);
}
