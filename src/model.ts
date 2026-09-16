export type Sex = "M" | "F" | "U";

export interface PersonName {
	givenName?: string;
	surname?: string;
	suffix?: string;
	prefix?: string;
}

export interface Event {
	type: "BIRT" | "DEAT" | "MARR" | string;
	date?: string;
	place?: string;
}

export interface Person {
	id: string;
	name: PersonName;
	sex?: Sex;
	events?: Event[];
	famc?: string[];
	fams?: string[];
	link?: string;
	note?: string;
	occupation?: string;
}

export interface Family {
	id: string;
	husb?: string;
	wife?: string;
	chil?: string[];
	events?: Event[];
}
