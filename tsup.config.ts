import { defineConfig } from "tsup";

export default defineConfig({
	entry: {
		index: "src/index.ts",
		"cli-xlsx-to-ged": "src/cli-xlsx-to-ged.ts",
	},
	format: ["esm"],
	dts: true,
	clean: true,
	sourcemap: false,
	splitting: false,
	treeshake: true,
	minify: false,
});
