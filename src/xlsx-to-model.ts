import { getCellValue, getDataStartRow } from "./mapping";
import { normalizeIndiId, normalizeSex, parseSerial } from "./cells";
import type { Family, Person, PersonName } from "./model";
import type { TargetField, XlsxMapping } from "./types";

interface RowDraft {
	serial: number;
	explicitIndiId?: string;
	rowIndex: number;
	name: PersonName;
	sex?: "M" | "F" | "U";
	birthDate?: string;
	birthPlace?: string;
	deathDate?: string;
	deathPlace?: string;
	fatherSerial?: number;
	motherSerial?: number;
	spouseSerial?: number;
	marriageDate?: string;
	marriagePlace?: string;
	link?: string;
	note?: string;
	occupation?: string;
	relativeNames: Partial<
		Record<
			"father" | "mother" | "spouse",
			{ givenName?: string; surname?: string; suffix?: string }
		>
	>;
}

const isRowEmpty = (row: unknown[]) =>
	!row.some((cell) => String(cell ?? "").trim());

const parseNameParts = (
	given: string,
	surname: string,
	fullName: string,
	nameOrder: "first-last" | "last-first"
): PersonName => {
	if (given || surname) {
		return { givenName: given || undefined, surname: surname || undefined };
	}
	if (!fullName) {
		return {};
	}
	const parts = fullName.trim().split(/\s+/);
	if (parts.length === 1) {
		return nameOrder === "last-first"
			? { surname: parts[0] }
			: { givenName: parts[0] };
	}
	if (nameOrder === "last-first") {
		return { surname: parts[0], givenName: parts.slice(1).join(" ") };
	}
	return {
		givenName: parts.slice(0, -1).join(" "),
		surname: parts[parts.length - 1],
	};
};

const readRelativeName = (
	row: unknown[],
	prefix: "father" | "mother" | "spouse",
	mapping: XlsxMapping
) => {
	const given = getCellValue(row, `${prefix}GivenName` as TargetField, mapping);
	const surname = getCellValue(row, `${prefix}Surname` as TargetField, mapping);
	const suffix = getCellValue(row, `${prefix}Suffix` as TargetField, mapping);
	if (!given && !surname && !suffix) {
		return undefined;
	}
	return { givenName: given || undefined, surname: surname || undefined, suffix: suffix || undefined };
};

export const rowsToModel = (
	rows: unknown[][],
	mapping: XlsxMapping
): { persons: Person[]; families: Family[]; warnings: string[] } => {
	const warnings: string[] = [];
	const start = getDataStartRow(mapping);
	const dataRows = rows.slice(start).filter((row) => !isRowEmpty(row));

	if (!dataRows.length) {
		throw new Error("No data rows found in spreadsheet");
	}

	const drafts: RowDraft[] = [];

	dataRows.forEach((row, offset) => {
		const dataRowNumber = offset + 1;
		const explicitSerial = parseSerial(getCellValue(row, "serial", mapping));
		const serial = explicitSerial ?? dataRowNumber;
		const rawIndiId = getCellValue(row, "id", mapping);
		const explicitIndiId = normalizeIndiId(rawIndiId);
		if (rawIndiId && !explicitIndiId) {
			warnings.push(
				`Row ${start + offset + 1}: invalid INDI id "${rawIndiId}", using auto id`
			);
		}

		const name = parseNameParts(
			getCellValue(row, "givenName", mapping),
			getCellValue(row, "surname", mapping),
			getCellValue(row, "fullName", mapping),
			mapping.nameOrder ?? "last-first"
		);
		name.suffix = getCellValue(row, "suffix", mapping) || undefined;
		name.prefix = getCellValue(row, "prefix", mapping) || undefined;

		if (!name.givenName && !name.surname) {
			warnings.push(`Row ${start + offset + 1}: skipped (no name)`);
			return;
		}

		const sexRaw = getCellValue(row, "sex", mapping);
		const sex = normalizeSex(sexRaw);
		if (sexRaw && !sex) {
			warnings.push(`Row ${start + offset + 1}: unrecognized sex "${sexRaw}"`);
		}

		drafts.push({
			serial,
			explicitIndiId,
			rowIndex: start + offset,
			name,
			sex,
			birthDate: getCellValue(row, "birthDate", mapping) || undefined,
			birthPlace: getCellValue(row, "birthPlace", mapping) || undefined,
			deathDate: getCellValue(row, "deathDate", mapping) || undefined,
			deathPlace: getCellValue(row, "deathPlace", mapping) || undefined,
			fatherSerial: parseSerial(getCellValue(row, "fatherId", mapping)),
			motherSerial: parseSerial(getCellValue(row, "motherId", mapping)),
			spouseSerial: parseSerial(getCellValue(row, "spouseId", mapping)),
			marriageDate:
				getCellValue(row, "marriageDate", mapping) || undefined,
			marriagePlace:
				getCellValue(row, "marriagePlace", mapping) || undefined,
			link: getCellValue(row, "link", mapping) || undefined,
			note: getCellValue(row, "note", mapping) || undefined,
			occupation: getCellValue(row, "occupation", mapping) || undefined,
			relativeNames: {
				father: readRelativeName(row, "father", mapping),
				mother: readRelativeName(row, "mother", mapping),
				spouse: readRelativeName(row, "spouse", mapping),
			},
		});
	});

	if (!drafts.length) {
		throw new Error("No valid persons found (every row missing a name)");
	}

	const serialToGedId = new Map<number, string>();
	const usedIndiIds = new Set<string>();
	const persons: Person[] = [];
	let nextIndi = 1;

	const personIdBySerial = (serial: number): string | undefined =>
		serialToGedId.get(serial);

	const allocateAutoIndiId = (): string => {
		while (usedIndiIds.has(`I${nextIndi}`)) {
			nextIndi++;
		}
		const id = `I${nextIndi++}`;
		usedIndiIds.add(id);
		return id;
	};

	const createStub = (
		rel: { givenName?: string; surname?: string; suffix?: string },
		sex?: "M" | "F" | "U"
	): string => {
		const id = allocateAutoIndiId();
		persons.push({
			id,
			name: rel,
			sex,
			events: [],
			famc: [],
			fams: [],
		});
		warnings.push(
			`Created stub person ${id} (${[rel.surname, rel.givenName].filter(Boolean).join(" ")})`
		);
		return id;
	};

	for (const draft of drafts) {
		let id: string;
		if (draft.explicitIndiId) {
			if (usedIndiIds.has(draft.explicitIndiId)) {
				warnings.push(
					`Row ${draft.rowIndex + 1}: duplicate INDI id "${draft.explicitIndiId}", using auto id`
				);
				id = allocateAutoIndiId();
			} else {
				id = draft.explicitIndiId;
				usedIndiIds.add(id);
			}
		} else {
			id = allocateAutoIndiId();
		}
		serialToGedId.set(draft.serial, id);
		const events = [];
		if (draft.birthDate || draft.birthPlace) {
			events.push({
				type: "BIRT" as const,
				date: draft.birthDate,
				place: draft.birthPlace,
			});
		}
		if (draft.deathDate || draft.deathPlace) {
			events.push({
				type: "DEAT" as const,
				date: draft.deathDate,
				place: draft.deathPlace,
			});
		}
		persons.push({
			id,
			name: draft.name,
			sex: draft.sex,
			events,
			link: draft.link,
			note: draft.note,
			occupation: draft.occupation,
			famc: [],
			fams: [],
		});
	}

	const personById = new Map(persons.map((p) => [p.id, p]));
	const families: Family[] = [];
	let nextFam = 1;

	const famKey = (husb?: string, wife?: string) =>
		`${husb ?? ""}|${wife ?? ""}`;

	const famByKey = new Map<string, Family>();

	const getOrCreateFamily = (
		husb?: string,
		wife?: string,
		marriage?: { date?: string; place?: string }
	): Family => {
		const key = famKey(husb, wife);
		let fam = famByKey.get(key);
		if (!fam) {
			fam = {
				id: `F${nextFam++}`,
				husb,
				wife,
				chil: [],
				events: [],
			};
			if (marriage?.date || marriage?.place) {
				fam.events!.push({
					type: "MARR",
					date: marriage.date,
					place: marriage.place,
				});
			}
			families.push(fam);
			famByKey.set(key, fam);
			if (husb) {
				personById.get(husb)?.fams?.push(fam.id);
			}
			if (wife) {
				personById.get(wife)?.fams?.push(fam.id);
			}
		}
		return fam;
	};

	const resolveParentId = (
		serial: number | undefined,
		relName:
			| { givenName?: string; surname?: string; suffix?: string }
			| undefined,
		defaultSex: "M" | "F"
	): string | undefined => {
		if (serial != null) {
			const existing = personIdBySerial(serial);
			if (existing) {
				return existing;
			}
			warnings.push(`Parent serial ${serial} not found in sheet`);
		}
		if (relName && (relName.givenName || relName.surname)) {
			return createStub(relName, defaultSex);
		}
		return undefined;
	};

	for (const draft of drafts) {
		const childId = serialToGedId.get(draft.serial)!;
		const fatherId = resolveParentId(
			draft.fatherSerial,
			draft.relativeNames.father,
			"M"
		);
		const motherId = resolveParentId(
			draft.motherSerial,
			draft.relativeNames.mother,
			"F"
		);

		if (fatherId || motherId) {
			const fam = getOrCreateFamily(fatherId, motherId);
			if (!fam.chil!.includes(childId)) {
				fam.chil!.push(childId);
			}
			const child = personById.get(childId);
			if (child && !child.famc!.includes(fam.id)) {
				child.famc!.push(fam.id);
			}
		}

		if (draft.spouseSerial != null || draft.relativeNames.spouse) {
			let spouseId = draft.spouseSerial
				? personIdBySerial(draft.spouseSerial)
				: undefined;
			if (draft.spouseSerial && !spouseId) {
				warnings.push(
					`Spouse serial ${draft.spouseSerial} not found for row ${draft.rowIndex + 1}`
				);
			}
			if (!spouseId && draft.relativeNames.spouse) {
				spouseId = createStub(draft.relativeNames.spouse, "U");
			}
			if (spouseId) {
				const self = personById.get(childId)!;
				const spouse = personById.get(spouseId)!;
				let husb = childId;
				let wife = spouseId;
				if (self.sex === "F") {
					husb = spouseId;
					wife = childId;
				} else if (self.sex === "M") {
					husb = childId;
					wife = spouseId;
				} else if (spouse.sex === "M") {
					husb = spouseId;
					wife = childId;
				} else if (spouse.sex === "F") {
					husb = childId;
					wife = spouseId;
				}
				getOrCreateFamily(husb, wife, {
					date: draft.marriageDate,
					place: draft.marriagePlace,
				});
			}
		}
	}

	return { persons, families, warnings };
};
