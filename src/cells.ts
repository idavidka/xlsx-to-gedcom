import type { CellValue } from "exceljs";

const EXCEL_EPOCH_MS = Date.UTC(1899, 11, 30);

export const cellToString = (value: CellValue | undefined | null): string => {
	if (value === undefined || value === null) {
		return "";
	}
	if (typeof value === "string") {
		return value.trim();
	}
	if (typeof value === "number") {
		return formatExcelNumber(value);
	}
	if (value instanceof Date) {
		return formatDate(value);
	}
	if (typeof value === "object") {
		if ("text" in value && typeof value.text === "string") {
			return value.text.trim();
		}
		if ("result" in value && value.result != null) {
			return cellToString(value.result as CellValue);
		}
		if ("richText" in value && Array.isArray(value.richText)) {
			return value.richText
				.map((part) => part.text ?? "")
				.join("")
				.trim();
		}
	}
	return String(value).trim();
};

const formatExcelNumber = (value: number): string => {
	if (Number.isInteger(value) && value >= 1 && value <= 500000) {
		const asDate = new Date(EXCEL_EPOCH_MS + value * 86400000);
		if (!Number.isNaN(asDate.getTime())) {
			const year = asDate.getUTCFullYear();
			if (year >= 1000 && year <= 2200) {
				return formatDate(asDate);
			}
		}
	}
	if (Number.isInteger(value)) {
		return String(value);
	}
	return String(value);
};

const formatDate = (date: Date): string => {
	const day = date.getUTCDate();
	const month = date.toLocaleString("en-US", {
		month: "short",
		timeZone: "UTC",
	});
	const year = date.getUTCFullYear();
	return `${day} ${month.toUpperCase()} ${year}`;
};

/** Normalize a user-provided GEDCOM INDI xref (without @ delimiters). */
export const normalizeIndiId = (raw: string): string | undefined => {
	let value = raw.trim();
	if (!value) {
		return undefined;
	}
	value = value.replace(/^@+|@+$/g, "");
	if (!/^[A-Za-z0-9_-]+$/.test(value)) {
		return undefined;
	}
	return value;
};

export const parseSerial = (raw: string): number | undefined => {
	const trimmed = raw.trim();
	if (!trimmed) {
		return undefined;
	}
	const n = Number(trimmed);
	if (!Number.isFinite(n) || n < 1) {
		return undefined;
	}
	return Math.floor(n);
};

export const normalizeSex = (raw: string): "M" | "F" | "U" | undefined => {
	const v = raw.trim().toLowerCase();
	if (!v) {
		return undefined;
	}
	if (["m", "male", "férfi", "ferfi", "ffi"].includes(v)) {
		return "M";
	}
	if (["f", "female", "nő", "no", "női"].includes(v)) {
		return "F";
	}
	if (["u", "unknown", "ismeretlen", "?"].includes(v)) {
		return "U";
	}
	return undefined;
};

export const columnIndexToLetter = (index: number): string => {
	let n = index + 1;
	let label = "";
	while (n > 0) {
		const rem = (n - 1) % 26;
		label = String.fromCharCode(65 + rem) + label;
		n = Math.floor((n - 1) / 26);
	}
	return label;
};
