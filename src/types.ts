export type Binary = ArrayBuffer | Uint8Array | Buffer;

export type NameOrder = "first-last" | "last-first";

export type TargetField =
	| "serial"
	| "id"
	| "surname"
	| "givenName"
	| "suffix"
	| "prefix"
	| "fullName"
	| "sex"
	| "birthDate"
	| "birthPlace"
	| "deathDate"
	| "deathPlace"
	| "fatherId"
	| "motherId"
	| "fatherGivenName"
	| "fatherSurname"
	| "fatherSuffix"
	| "motherGivenName"
	| "motherSurname"
	| "motherSuffix"
	| "spouseId"
	| "spouseGivenName"
	| "spouseSurname"
	| "spouseSuffix"
	| "childGivenName"
	| "childSurname"
	| "childSuffix"
	| "marriageDate"
	| "marriagePlace"
	| "occupation"
	| "note"
	| "link";

export type ColumnIndex = number;

export interface XlsxMapping {
	sheet: string | number;
	hasHeader: boolean;
	headerRow: number;
	dataStartRow?: number;
	nameOrder?: NameOrder;
	columns: Partial<Record<TargetField, ColumnIndex>>;
}

export interface SheetInfo {
	name: string;
	index: number;
	rowCount: number;
	columnCount: number;
}

export interface InspectResult {
	sheets: SheetInfo[];
	activeSheet: number;
	headerCandidates: string[];
	previewRows: unknown[][];
	guessed: XlsxMapping;
}

export interface XlsxToGedOptions {
	mapping?: XlsxMapping;
	layout?: "simple" | "export";
	fileName?: string;
}

export interface XlsxToGedResult {
	gedcom: string;
	warnings: string[];
}

export interface MappedPersonPreview {
	serial: number;
	givenName?: string;
	surname?: string;
	suffix?: string;
	sex?: string;
	birthDate?: string;
	deathDate?: string;
	fatherSerial?: number;
	motherSerial?: number;
	spouseSerial?: number;
}
