import ExcelJS from "exceljs";

import { SIMPLE_LAYOUT_HEADERS_EN } from "./constants";

/** Blank .xlsx with English headers for the TreeViz simple layout. */
export const createSimpleTemplateBuffer = async (): Promise<ArrayBuffer> => {
	const workbook = new ExcelJS.Workbook();
	const sheet = workbook.addWorksheet("People");
	sheet.addRow(SIMPLE_LAYOUT_HEADERS_EN);
	sheet.addRow([
		"Kovács",
		"János",
		"12 JUN 1880",
		"3 JAN 1945",
		"M",
		"",
		"",
		"",
		"",
		"",
		"",
		"",
		"",
		"",
	]);
	sheet.addRow([
		"Nagy",
		"Mária",
		"1885",
		"1960",
		"F",
		"",
		"",
		"",
		"",
		"",
		"",
		"",
		"",
		"",
	]);
	sheet.addRow([
		"Kovács",
		"Péter",
		"1910",
		"",
		"M",
		"1",
		"2",
		"",
		"",
		"",
		"",
		"",
		"",
		"",
	]);
	const buffer = await workbook.xlsx.writeBuffer();
	return buffer as ArrayBuffer;
};
