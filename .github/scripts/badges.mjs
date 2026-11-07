// README badges as SVG files in assets/badges/, generated from package.json (`npm run badges`).
// PPTB shows the README in its main window, whose CSP only allows images from GitHub and
// *.githubusercontent.com, so img.shields.io badges don't show there (RELEASING.md → README
// images). The look follows shields.io's flat style, as Command Bar Studio's badges do.
// check-release.mjs checks the files and the README links are current.

/** Approximate Verdana 11px advance widths, in px (close to what shields.io measures) */
const widthOf = (text) =>
	[...text].reduce((w, c) => {
		if (/[0-9]/.test(c)) return w + 7;
		if (/[.\-/]/.test(c)) return w + 4.2;
		if (c === " ") return w + 3.9;
		if (c === "@") return w + 11;
		if (/[a-z]/.test(c)) return w + 6.5;
		return w + 7.5;
	}, 0);

const escapeXml = (text) => text.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");

/** One flat badge: grey label, coloured message (dark text on light colours) */
export function badgeSvg(label, message, color, { darkText = false } = {}) {
	const pad = 10;
	const lw = Math.round(widthOf(label) + pad);
	const mw = Math.round(widthOf(message) + pad);
	const w = lw + mw;
	const textLength = (s) => Math.round(widthOf(s) * 10);
	const text = (x, s, fill, shadow) =>
		`<text aria-hidden="true" x="${x}" y="150" fill="${shadow}" fill-opacity=".3" transform="scale(.1)" textLength="${textLength(s)}">${escapeXml(s)}</text>` +
		`<text x="${x}" y="140" transform="scale(.1)" fill="${fill}" textLength="${textLength(s)}">${escapeXml(s)}</text>`;
	const title = escapeXml(`${label}: ${message}`);
	return (
		`<svg xmlns="http://www.w3.org/2000/svg" width="${w}" height="20" role="img" aria-label="${title}">` +
		`<title>${title}</title>` +
		`<linearGradient id="s" x2="0" y2="100%"><stop offset="0" stop-color="#bbb" stop-opacity=".1"/><stop offset="1" stop-opacity=".1"/></linearGradient>` +
		`<clipPath id="r"><rect width="${w}" height="20" rx="3" fill="#fff"/></clipPath>` +
		`<g clip-path="url(#r)"><rect width="${lw}" height="20" fill="#555"/><rect x="${lw}" width="${mw}" height="20" fill="${color}"/><rect width="${w}" height="20" fill="url(#s)"/></g>` +
		`<g fill="#fff" text-anchor="middle" font-family="Verdana,Geneva,DejaVu Sans,sans-serif" text-rendering="geometricPrecision" font-size="110">` +
		text(lw * 5, label, "#fff", "#010101") +
		(darkText ? text(lw * 10 + mw * 5, message, "#333", "#ccc") : text(lw * 10 + mw * 5, message, "#fff", "#010101")) +
		`</g></svg>\n`
	);
}

/** "^18.3.1" → "18.3.1" */
const plainVersion = (range) => String(range ?? "").replace(/^[^\d]*/, "");

/** owner/repo from package.json's repository.url */
export function repositoryOf(pkg) {
	const url = typeof pkg.repository === "string" ? pkg.repository : pkg.repository?.url;
	const match = /github\.com[/:]([^/]+\/[^/.]+)/.exec(url ?? "");
	if (!match) throw new Error("package.json repository.url isn't a GitHub repository.");
	return match[1];
}

/** Where the README links a badge: the file on main, as raw.githubusercontent.com serves it */
export const badgeUrl = (pkg, file) => `https://raw.githubusercontent.com/${repositoryOf(pkg)}/refs/heads/main/assets/badges/${file}`;

/**
 * Every README badge, from package.json: the file name, its alt text and the SVG. The order is the
 * README's.
 */
export function badges(pkg) {
	const deps = { ...pkg.dependencies, ...pkg.devDependencies };
	const version = pkg.version;
	const react = plainVersion(deps.react).split(".")[0];
	const typescript = plainVersion(deps.typescript).split(".").slice(0, 2).join(".");
	const fluent = plainVersion(deps["@fluentui/react-components"]).split(".")[0];
	const pptbTypes = plainVersion(deps["@pptb/types"]);
	const minApi = pkg.features?.minAPI;
	const list = [
		// Pre-releases yellow, releases blue, as Command Bar Studio's version badge
		{ file: "version.svg", alt: "Version", link: `https://www.npmjs.com/package/${pkg.name}`, svg: badgeSvg("version", version, version.includes("-") ? "#dfb317" : "#007ec6") },
	];
	if (minApi) list.push({ file: "pptb.svg", alt: `Power Platform ToolBox ${minApi} or later`, svg: badgeSvg("Power Platform ToolBox", `≥ ${minApi}`, "#742774") });
	if (pptbTypes) list.push({ file: "pptb-types.svg", alt: `@pptb/types ${pptbTypes}`, svg: badgeSvg("@pptb/types", pptbTypes, "#007ec6") });
	if (react) list.push({ file: "react.svg", alt: `React ${react}`, svg: badgeSvg("React", react, "#61DAFB", { darkText: true }) });
	if (typescript) list.push({ file: "typescript.svg", alt: `TypeScript ${typescript}`, svg: badgeSvg("TypeScript", typescript, "#3178C6") });
	if (fluent) list.push({ file: "fluent-ui.svg", alt: `Fluent UI v${fluent}`, svg: badgeSvg("Fluent UI", `v${fluent}`, "#0078D4") });
	if (pkg.license) list.push({ file: "license.svg", alt: `License ${pkg.license}`, svg: badgeSvg("license", pkg.license, "#97ca00") });
	return list;
}

/**
 * The README's badge lines, one image per line. The version badge's alt text has no version and it
 * links to npm, so a release changes only the SVG (`npm run badges`), not the README.
 */
export const badgeMarkdown = (pkg) =>
	badges(pkg).map((b) => {
		const image = `![${b.alt}](${badgeUrl(pkg, b.file)})`;
		return b.link ? `[${image}](${b.link})` : image;
	});
