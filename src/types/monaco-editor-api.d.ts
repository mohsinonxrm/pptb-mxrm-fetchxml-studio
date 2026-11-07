// @monaco-editor/react 4.7 and @monaco-editor/loader 1.6 type their `monaco` parameters with
// `import * as monaco from "monaco-editor/esm/vs/editor/editor.api"`. Since 0.56, monaco-editor
// has an `exports` map that doesn't expose that deep path, so the import doesn't resolve (#145).
//
// The same API is now published as "monaco-editor/editor", which is also what FXS hands the
// loader (`loader.config({ monaco })` in main.tsx). Point the old path at it. This file must
// stay a global script (no top-level import/export) so the `declare module` below is an ambient
// declaration, not an augmentation. Remove it once a stable @monaco-editor/react release
// imports a path that monaco-editor exports.
declare module "monaco-editor/esm/vs/editor/editor.api" {
	export * from "monaco-editor/editor";
}
