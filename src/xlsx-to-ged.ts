import { resolveMapping } from "./mapping";
import { modelToGed } from "./model-to-ged";
import { rowsToModel } from "./xlsx-to-model";
import { loadWorkbook } from "./workbook";
import type {
	Binary,
	XlsxMapping,
	XlsxToGedOptions,
	XlsxToGedResult,
} from "./types";

export const xlsxToGed = async (
	input: Binary,
	options?: XlsxToGedOptions | XlsxMapping
): Promise<XlsxToGedResult> => {
	const normalized: XlsxToGedOptions =
		options && "columns" in options
			? { mapping: options }
			: (options ?? {});
	const mapping = resolveMapping(normalized);
	const { getSheetRows } = await loadWorkbook(input);
	const { rows } = getSheetRows(mapping.sheet);
	const { persons, families, warnings } = rowsToModel(rows, mapping);
	const gedcom = modelToGed(
		persons,
		families,
		normalized.fileName?.replace(/\.xlsx$/i, "")
	);
	return { gedcom, warnings };
};
