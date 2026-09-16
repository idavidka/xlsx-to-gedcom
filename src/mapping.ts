import {
	DEFAULT_SIMPLE_LAYOUT,
	SIMPLE_LAYOUT_HEADERS_EN,
	SIMPLE_LAYOUT_HEADERS_HU,
} from "./constants";
import type { TargetField, XlsxMapping } from "./types";

const HEADER_ALIASES: Record<TargetField, string[]> = {
	serial: [
		"no.",
		"no",
		"number",
		"row no",
		"row no.",
		"row number",
		"#",
		"sorszám",
		"sorszam",
		"serial",
		"serial no",
		"serial number",
	],
	id: [
		"id",
		"indi id",
		"gedcom id",
		"person id",
		"azonosító",
		"azonosito",
	],
	surname: [
		"surname",
		"vezetéknév",
		"vezeteknev",
		"családnév",
		"csaladnev",
		"last name",
		"lastname",
	],
	givenName: [
		"given name",
		"givenname",
		"keresztnév",
		"keresztnev",
		"utónév",
		"utonev",
		"first name",
		"firstname",
	],
	suffix: ["suffix", "utótag", "utotag", "jr", "ifj"],
	prefix: ["prefix", "előtag", "elotag"],
	fullName: ["name", "név", "nev", "full name", "fullname"],
	sex: ["sex", "gender", "nem"],
	birthDate: ["birth", "birth date", "birthdate", "születés", "szuletes", "született"],
	birthPlace: [
		"birth place",
		"birthplace",
		"születési hely",
		"szuletesi hely",
	],
	deathDate: ["death", "death date", "deathdate", "halál", "halal", "meghalt"],
	deathPlace: ["death place", "deathplace", "halálozási hely", "halalozasi hely"],
	fatherId: [
		"father no.",
		"father no",
		"father number",
		"father sorszám",
		"father sorszam",
		"father serial",
		"apa sorszám",
		"apa sorszam",
	],
	motherId: [
		"mother no.",
		"mother no",
		"mother number",
		"mother sorszám",
		"mother sorszam",
		"mother serial",
		"anya sorszám",
		"anya sorszam",
	],
	fatherGivenName: ["father given name", "apa keresztneve", "apa keresztnev"],
	fatherSurname: ["father surname", "apa vezetékneve", "apa vezetekneve"],
	fatherSuffix: ["father suffix", "apa utótag"],
	motherGivenName: ["mother given name", "anya keresztneve", "anya keresztnev"],
	motherSurname: ["mother surname", "anya vezetékneve", "anya vezetekneve"],
	motherSuffix: ["mother suffix", "anya utótag"],
	spouseId: [
		"spouse no.",
		"spouse no",
		"spouse number",
		"spouse sorszám",
		"spouse sorszam",
		"spouse serial",
		"házastárs sorszám",
		"hazastars sorszam",
		"házastárs",
	],
	spouseGivenName: ["spouse given name", "házastárs keresztneve"],
	spouseSurname: ["spouse surname", "házastárs vezetékneve"],
	spouseSuffix: ["spouse suffix", "házastárs utótag"],
	childGivenName: ["child given name", "gyerek keresztneve"],
	childSurname: ["child surname", "gyerek vezetékneve"],
	childSuffix: ["child suffix", "gyerek utótag"],
	marriageDate: [
		"marriage",
		"marriage date",
		"házasság dátuma",
		"hazassag datuma",
	],
	marriagePlace: [
		"marriage place",
		"házasság helye",
		"hazassag helye",
	],
	occupation: ["occupation", "foglalkozás", "foglalkozas"],
	note: ["note", "notes", "megjegyzés", "megjegyzes"],
	link: ["link", "url", "hivatkozás", "hivatkozas"],
};

const normalizeHeader = (value: string) =>
	value
		.trim()
		.toLowerCase()
		.normalize("NFD")
		.replace(/\p{M}/gu, "");

export const matchHeaderToField = (header: string): TargetField | undefined => {
	const normalized = normalizeHeader(header);
	if (!normalized) {
		return undefined;
	}
	for (const [field, aliases] of Object.entries(HEADER_ALIASES) as Array<
		[TargetField, string[]]
	>) {
		if (
			aliases.some(
				(alias) =>
					normalized === normalizeHeader(alias) ||
					normalized.includes(normalizeHeader(alias))
			)
		) {
			return field;
		}
	}
	return undefined;
};

export const looksLikeHeaderRow = (cells: unknown[]): boolean => {
	const strings = cells.map((c) => String(c ?? "").trim()).filter(Boolean);
	if (strings.length < 2) {
		return false;
	}
	const matched = strings.filter((s) => matchHeaderToField(s)).length;
	return matched >= 2;
};

export const guessMappingFromHeaders = (
	headers: string[],
	sheet: string | number = 0
): XlsxMapping => {
	const columns: Partial<Record<TargetField, number>> = {};
	const used = new Set<TargetField>();

	headers.forEach((header, index) => {
		const field = matchHeaderToField(String(header ?? ""));
		if (field && !used.has(field)) {
			columns[field] = index;
			used.add(field);
		}
	});

	const countSimpleHeaderMatch = (expected: string[]) =>
		expected.filter((h, i) => {
			const header = headers[i];
			return (
				typeof header === "string" &&
				normalizeHeader(header) === normalizeHeader(h)
			);
		}).length;

	const simpleMatch = Math.max(
		countSimpleHeaderMatch(SIMPLE_LAYOUT_HEADERS_EN),
		countSimpleHeaderMatch(SIMPLE_LAYOUT_HEADERS_HU)
	);

	if (simpleMatch >= 4) {
		return {
			...DEFAULT_SIMPLE_LAYOUT,
			sheet,
			hasHeader: true,
			headerRow: 0,
		};
	}

	if (Object.keys(columns).length >= 2) {
		return {
			sheet,
			hasHeader: true,
			headerRow: 0,
			nameOrder: columns.surname != null ? "last-first" : "first-last",
			columns,
		};
	}

	return { ...DEFAULT_SIMPLE_LAYOUT, sheet };
};

export const resolveMapping = (
	options?: { mapping?: XlsxMapping; layout?: "simple" | "export" }
): XlsxMapping => {
	if (options?.mapping) {
		return options.mapping;
	}
	if (options?.layout === "export") {
		return {
			sheet: 0,
			hasHeader: true,
			headerRow: 0,
			nameOrder: "last-first",
			columns: {},
		};
	}
	return { ...DEFAULT_SIMPLE_LAYOUT };
};

export const getDataStartRow = (mapping: XlsxMapping): number => {
	if (mapping.dataStartRow != null) {
		return mapping.dataStartRow;
	}
	return mapping.hasHeader ? mapping.headerRow + 1 : 0;
};

export const getCellValue = (
	row: unknown[],
	field: TargetField,
	mapping: XlsxMapping
): string => {
	const col = mapping.columns[field];
	if (col == null) {
		return "";
	}
	return String(row[col] ?? "").trim();
};

export const mappingFromColumnTargets = (
	columnTargets: Array<TargetField | "">,
	base: XlsxMapping
): XlsxMapping => {
	const columns: Partial<Record<TargetField, number>> = {};
	columnTargets.forEach((target, index) => {
		if (target) {
			columns[target] = index;
		}
	});
	return { ...base, columns };
};
