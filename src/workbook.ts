import ExcelJS from "exceljs";

import type { Binary } from "./types";
import { cellToString } from "./cells";

export type LoadedWorkbook = {
	workbook: ExcelJS.Workbook;
	getSheetRows: (
		sheet: string | number
	) => { name: string; index: number; rows: unknown[][] };
};

const toBuffer = (input: Binary): ArrayBuffer | Buffer => {
	if (input instanceof ArrayBuffer) {
		return input;
	}
	if (input instanceof Uint8Array) {
		return input.buffer.slice(
			input.byteOffset,
			input.byteOffset + input.byteLength
		) as ArrayBuffer;
	}
	return input;
};

export const loadWorkbook = async (input: Binary): Promise<LoadedWorkbook> => {
	const workbook = new ExcelJS.Workbook();
	await workbook.xlsx.load(toBuffer(input) as ArrayBuffer);

	const getSheetRows = (sheetRef: string | number) => {
		const worksheet =
			typeof sheetRef === "number"
				? workbook.worksheets[sheetRef]
				: workbook.getWorksheet(sheetRef);
		if (!worksheet) {
			throw new Error(`Sheet not found: ${String(sheetRef)}`);
		}
		const rows: unknown[][] = [];
		worksheet.eachRow({ includeEmpty: true }, (row) => {
			const cells: unknown[] = [];
			row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
				cells[colNumber - 1] = cellToString(cell.value);
			});
			rows.push(cells);
		});
		return {
			name: worksheet.name,
			index: worksheet.id - 1,
			rows,
		};
	};

	return { workbook, getSheetRows };
};

export const getSheetInfos = (workbook: ExcelJS.Workbook) => {
	return workbook.worksheets.map((ws, index) => ({
		name: ws.name,
		index,
		rowCount: ws.rowCount,
		columnCount: ws.columnCount,
	}));
};
