import { readFileSync } from "node:fs";
import { join } from "node:path";

import ExcelJS from "exceljs";
import { describe, expect, it } from "vitest";

import { DEFAULT_SIMPLE_LAYOUT } from "../constants";
import { inspectXlsx } from "../inspect";
import { xlsxToGed } from "../xlsx-to-ged";

const buildSimpleWorkbook = async () => {
	const workbook = new ExcelJS.Workbook();
	const sheet = workbook.addWorksheet("People");
	sheet.addRow(["Kovács", "János", "12 JUN 1880", "3 JAN 1945", "M", "", ""]);
	sheet.addRow(["Nagy", "Mária", "1885", "1960", "F", "", ""]);
	sheet.addRow(["Kovács", "Péter", "1910", "", "M", "1", "2"]);
	return (await workbook.xlsx.writeBuffer()) as ArrayBuffer;
};

describe("xlsx-to-ged", () => {
	it("converts default simple layout without mapping", async () => {
		const buffer = await buildSimpleWorkbook();
		const { gedcom, warnings } = await xlsxToGed(buffer);
		expect(gedcom).toContain("0 HEAD");
		expect(gedcom).toContain("1 SOUR XLSX2GED");
		expect(gedcom).toContain("Kovács");
		expect(gedcom).toContain("János");
		expect(gedcom).toContain("0 @I3@ INDI");
		expect(gedcom).toContain("1 FAMC @F1@");
		expect(gedcom).toContain("1 HUSB @I1@");
		expect(gedcom).toContain("1 WIFE @I2@");
		expect(gedcom).toContain("1 CHIL @I3@");
		expect(warnings.some((w) => w.includes("skipped"))).toBe(false);
	});

	it("inspects workbook sheets and preview", async () => {
		const buffer = await buildSimpleWorkbook();
		const result = await inspectXlsx(buffer);
		expect(result.sheets[0]?.name).toBe("People");
		expect(result.previewRows.length).toBeGreaterThan(0);
		expect(result.guessed.columns.surname).toBe(
			DEFAULT_SIMPLE_LAYOUT.columns.surname
		);
	});

	it("converts committed example spreadsheet", async () => {
		const bytes = readFileSync(
			join(import.meta.dirname, "../../examples/simple-family.xlsx")
		);
		const buffer = bytes.buffer.slice(
			bytes.byteOffset,
			bytes.byteOffset + bytes.byteLength
		);
		const inspect = await inspectXlsx(buffer);
		const { gedcom } = await xlsxToGed(buffer, inspect.guessed);
		expect(gedcom).toContain("Kovács");
		expect(gedcom).toContain("Péter");
		expect(gedcom).toContain("1 FAMC @F1@");
	});

	it("links by row no. while explicit id sets INDI xref", async () => {
		const workbook = new ExcelJS.Workbook();
		const sheet = workbook.addWorksheet("People");
		sheet.addRow([
			"Kovács",
			"János",
			"",
			"",
			"M",
			"",
			"",
			"",
			"",
			"JANOS1",
		]);
		sheet.addRow([
			"Nagy",
			"Mária",
			"",
			"",
			"F",
			"",
			"",
			"",
			"",
			"MARIA2",
		]);
		sheet.addRow([
			"Kovács",
			"Péter",
			"",
			"",
			"M",
			"1",
			"2",
			"",
			"",
			"PETER3",
		]);
		const buffer = (await workbook.xlsx.writeBuffer()) as ArrayBuffer;
		const { gedcom } = await xlsxToGed(buffer);
		expect(gedcom).toContain("0 @JANOS1@ INDI");
		expect(gedcom).toContain("0 @PETER3@ INDI");
		expect(gedcom).toContain("1 HUSB @JANOS1@");
		expect(gedcom).toContain("1 WIFE @MARIA2@");
		expect(gedcom).toContain("1 CHIL @PETER3@");
	});

	it("throws when sheet has no names", async () => {
		const workbook = new ExcelJS.Workbook();
		workbook.addWorksheet("Empty").addRow(["", ""]);
		const buffer = (await workbook.xlsx.writeBuffer()) as ArrayBuffer;
		await expect(xlsxToGed(buffer)).rejects.toThrow(/No (valid persons|data rows)/);
	});
});
