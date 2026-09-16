import { looksLikeHeaderRow, guessMappingFromHeaders } from "./mapping";
import { loadWorkbook, getSheetInfos } from "./workbook";
import type { Binary, InspectResult, XlsxMapping } from "./types";

const PREVIEW_ROW_LIMIT = 25;

export const inspectXlsx = async (
	input: Binary,
	options?: { sheet?: string | number }
): Promise<InspectResult> => {
	const { getSheetRows, workbook } = await loadWorkbook(input);
	const sheetRef =
		options?.sheet ??
		(workbook.worksheets[0]?.id != null
			? workbook.worksheets[0].id - 1
			: 0);
	const { rows, index: activeSheet } = getSheetRows(sheetRef);
	const firstRow = rows[0] ?? [];
	const hasHeader = looksLikeHeaderRow(firstRow);
	const headerCandidates = hasHeader
		? firstRow.map((c) => String(c ?? ""))
		: [];
	/** Always from sheet row 0 so the wizard can toggle hasHeader without re-fetch. */
	const previewRows = rows.slice(0, PREVIEW_ROW_LIMIT);
	const guessed = guessMappingFromHeaders(
		hasHeader ? headerCandidates : [],
		sheetRef
	);
	if (!hasHeader) {
		guessed.hasHeader = false;
	}

	return {
		sheets: getSheetInfos(workbook),
		activeSheet,
		headerCandidates,
		previewRows,
		guessed,
	};
};

export const inspectSheet = async (
	input: Binary,
	mapping: Pick<XlsxMapping, "sheet" | "hasHeader" | "headerRow">
): Promise<{ previewRows: unknown[][]; headerCandidates: string[] }> => {
	const { getSheetRows } = await loadWorkbook(input);
	const { rows } = getSheetRows(mapping.sheet);
	const headerRow = rows[mapping.headerRow] ?? [];
	const start = mapping.hasHeader ? mapping.headerRow + 1 : 0;
	return {
		headerCandidates: mapping.hasHeader
			? headerRow.map((c) => String(c ?? ""))
			: [],
		previewRows: rows.slice(start, start + PREVIEW_ROW_LIMIT),
	};
};
