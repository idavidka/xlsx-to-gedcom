import type { XlsxMapping } from "./types";

/** Fixed column indices for the TreeViz simple layout (see docs/XLSX_TO_GED_PLAN.md). */
export const DEFAULT_SIMPLE_LAYOUT: XlsxMapping = {
	sheet: 0,
	hasHeader: false,
	headerRow: 0,
	nameOrder: "last-first",
	columns: {
		surname: 0,
		givenName: 1,
		birthDate: 2,
		deathDate: 3,
		sex: 4,
		fatherId: 5,
		motherId: 6,
		birthPlace: 7,
		deathPlace: 8,
		id: 9,
		suffix: 10,
		spouseId: 11,
		marriageDate: 12,
		marriagePlace: 13,
	},
};

export const SIMPLE_LAYOUT_COLUMN_COUNT =
	Math.max(...Object.values(DEFAULT_SIMPLE_LAYOUT.columns).map((c) => c ?? 0)) +
	1;

/** English header row for the simple layout template. Row no. = 1-based data-row index. */
export const SIMPLE_LAYOUT_HEADERS_EN = [
	"Surname",
	"Given name",
	"Birth",
	"Death",
	"Sex",
	"Father no.",
	"Mother no.",
	"Birth place",
	"Death place",
	"ID",
	"Suffix",
	"Spouse no.",
	"Marriage date",
	"Marriage place",
];

/** Hungarian headers — still recognized when guessing mapping. */
export const SIMPLE_LAYOUT_HEADERS_HU = [
	"Vezetéknév",
	"Keresztnév",
	"Születés",
	"Halál",
	"Nem",
	"Apa sorszám",
	"Anya sorszám",
	"Születési hely",
	"Halálozási hely",
	"Azonosító",
	"Utótag",
	"Házastárs sorszám",
	"Házasság dátuma",
	"Házasság helye",
];
