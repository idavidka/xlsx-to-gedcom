import type { Family, Person } from "./model";

const escapeGedcom = (value: string) => value.replace(/\n/g, " ").trim();

export const modelToGed = (
	persons: Person[],
	families: Family[],
	treeName?: string
): string => {
	const lines: string[] = [];
	lines.push("0 HEAD");
	lines.push("1 SOUR XLSX2GED");
	lines.push("1 GEDC");
	lines.push("2 VERS 5.5.1");
	lines.push("2 FORM LINEAGE-LINKED");
	lines.push("1 CHAR UTF-8");

	if (treeName) {
		lines.push(`1 FILE ${treeName}.xlsx`);
		lines.push(`1 _TREE ${treeName}`);
	}

	for (const person of persons) {
		lines.push(`0 @${person.id}@ INDI`);
		const { givenName, surname, suffix, prefix } = person.name;
		const displayGiven = givenName ?? "";
		const displaySurname = surname ?? "";
		if (displayGiven || displaySurname) {
			if (displaySurname) {
				lines.push(
					`1 NAME ${escapeGedcom(displayGiven)} /${escapeGedcom(displaySurname)}/`
				);
			} else {
				lines.push(`1 NAME ${escapeGedcom(displayGiven)}`);
			}
			if (displayGiven) {
				lines.push(`2 GIVN ${escapeGedcom(displayGiven)}`);
			}
			if (displaySurname) {
				lines.push(`2 SURN ${escapeGedcom(displaySurname)}`);
			}
			if (suffix) {
				lines.push(`2 NSFX ${escapeGedcom(suffix)}`);
			}
			if (prefix) {
				lines.push(`2 NPFX ${escapeGedcom(prefix)}`);
			}
		}
		if (person.sex && person.sex !== "U") {
			lines.push(`1 SEX ${person.sex}`);
		}
		for (const event of person.events ?? []) {
			lines.push(`1 ${event.type}`);
			if (event.date) {
				lines.push(`2 DATE ${escapeGedcom(event.date)}`);
			}
			if (event.place) {
				lines.push(`2 PLAC ${escapeGedcom(event.place)}`);
			}
		}
		if (person.occupation) {
			lines.push(`1 OCCU ${escapeGedcom(person.occupation)}`);
		}
		if (person.note) {
			lines.push(`1 NOTE ${escapeGedcom(person.note)}`);
		}
		if (person.link) {
			lines.push(`1 WWW ${escapeGedcom(person.link)}`);
		}
		for (const famc of person.famc ?? []) {
			lines.push(`1 FAMC @${famc}@`);
		}
		for (const fams of person.fams ?? []) {
			lines.push(`1 FAMS @${fams}@`);
		}
	}

	for (const family of families) {
		lines.push(`0 @${family.id}@ FAM`);
		if (family.husb) {
			lines.push(`1 HUSB @${family.husb}@`);
		}
		if (family.wife) {
			lines.push(`1 WIFE @${family.wife}@`);
		}
		for (const child of family.chil ?? []) {
			lines.push(`1 CHIL @${child}@`);
		}
		for (const event of family.events ?? []) {
			lines.push(`1 ${event.type}`);
			if (event.date) {
				lines.push(`2 DATE ${escapeGedcom(event.date)}`);
			}
			if (event.place) {
				lines.push(`2 PLAC ${escapeGedcom(event.place)}`);
			}
		}
	}

	lines.push("0 TRLR");
	return lines.join("\n") + "\n";
};
