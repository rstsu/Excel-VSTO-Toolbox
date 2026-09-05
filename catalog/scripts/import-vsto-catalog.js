#!/usr/bin/env node

const fs = require("node:fs");
const path = require("node:path");

const sourceRoot = process.argv[2];
const outputFile = process.argv[3] || path.resolve(__dirname, "../dist/demo-data.js");
const version = process.argv[4] || "unbekannt";

if (!sourceRoot) {
  console.error("Aufruf: node scripts/import-vsto-catalog.js <Repository-Ordner> [Ausgabedatei] [Version]");
  process.exit(1);
}

const sourceSpecs = [
  { file: "RegExPQ/DemoCatalog.PowerQuery.vb", category: "Power Query", prefix: "PQ" },
  { file: "RegExPQ/DemoCatalog.Regex.vb", category: "Regex", prefix: "RX" },
  { file: "RegExPQ/DemoCatalog.Vba.vb", category: "VBA", prefix: "VB" },
  { file: "RegExPQ/DemoCatalog.Formula.vb", category: "Formeln", prefix: "FX" }
];

function fromVbString(value) {
  return value.replace(/""/g, "\"");
}

function extractVbString(block, field) {
  const match = block.match(new RegExp(`\\.${field}\\s*=\\s*\"((?:\"\"|[^\"])*)\"`));
  if (!match) throw new Error(`${field} fehlt in einem Demo-Block`);
  return fromVbString(match[1]);
}

function extractTags(block) {
  const match = block.match(/\.Tags\s*=\s*\{([^}]*)\}/);
  if (!match) return [];
  return Array.from(match[1].matchAll(/"((?:""|[^"])*)"/g), item => fromVbString(item[1]));
}

function extractCData(block, field) {
  const start = block.indexOf(`.${field} = TextBlock(`);
  if (start < 0) return "";
  const fieldBlock = block.slice(start);
  const match = fieldBlock.match(/<!\[CDATA\[\s*([\s\S]*?)\s*\]\]>/);
  return match ? match[1].replace(/\r\n/g, "\n").trim() : "";
}

function makeSummary(description) {
  const paragraphs = description
    .split(/\n\s*\n/)
    .map(text => text.replace(/\s*\n\s*/g, " ").trim())
    .filter(Boolean)
    .filter(text => !/^!{3,}/.test(text));
  let summary = paragraphs[0] || "Beschreibung im Demo-Katalog der Excel-VSTO-Toolbox.";
  if (summary.length > 190) {
    const shortened = summary.slice(0, 187);
    summary = `${shortened.slice(0, shortened.lastIndexOf(" "))} …`;
  }
  return summary;
}

function getPackageMap() {
  const runnerPath = path.join(sourceRoot, "RegExPQ/DemoRunner.vb");
  const runner = fs.readFileSync(runnerPath, "utf8").replace(/^\uFEFF/, "");
  const methodPackages = new Map();

  for (const match of runner.matchAll(/Private Sub\s+(\w+)\(\)([\s\S]*?)(?=\n\s*Private (?:Sub|Function)|\nEnd Class)/g)) {
    const packageMatch = match[2].match(/ExtractAndOpen\(\s*App,\s*"([^"]+\.zip)",\s*"([^"]+)"/);
    if (packageMatch) {
      methodPackages.set(match[1], { fileName: packageMatch[1], folderName: packageMatch[2] });
    }
  }

  const packageMap = new Map();
  for (const match of runner.matchAll(/Case\s+"([^"]+)"([\s\S]*?)(?=\n\s*Case\s+"|\n\s*End Select)/g)) {
    const method = match[2].match(/\b(Create\w+)\(\)/)?.[1];
    const packageInfo = methodPackages.get(method);
    if (!packageInfo) continue;

    const packagePath = path.join(sourceRoot, "RegExPQ/Demos", packageInfo.fileName);
    if (!fs.existsSync(packagePath)) throw new Error(`Demo-Archiv fehlt: ${packageInfo.fileName}`);

    packageMap.set(match[1], {
      ...packageInfo,
      sizeBytes: fs.statSync(packagePath).size,
      downloadUrl: `https://github.com/rstsu/Excel-VSTO-Toolbox/raw/refs/heads/main/RegExPQ/Demos/${encodeURIComponent(packageInfo.fileName)}`
    });
  }

  return packageMap;
}

const packageMap = getPackageMap();

function parseFile(spec) {
  const absolute = path.join(sourceRoot, spec.file);
  const source = fs.readFileSync(absolute, "utf8").replace(/^\uFEFF/, "");
  const starts = Array.from(source.matchAll(/\.Id\s*=\s*"/g), match => match.index);

  return starts.map((start, index) => {
    const end = starts[index + 1] ?? source.length;
    const block = source.slice(start, end);
    const sourceId = extractVbString(block, "Id");
    const numericId = Number(sourceId.match(/(\d+)$/)?.[1]);
    const description = extractCData(block, "Description");
    return {
      id: `${spec.prefix}-${String(numericId).padStart(3, "0")}`,
      sourceId,
      title: extractVbString(block, "Title"),
      category: spec.category,
      summary: makeSummary(description),
      description,
      tags: extractTags(block),
      codeText: extractCData(block, "CodeText"),
      isNew: false,
      package: packageMap.get(sourceId) || null,
      sourceUrl: `https://github.com/rstsu/Excel-VSTO-Toolbox/blob/main/${spec.file}#L${source.slice(0, start).split("\n").length}`
    };
  });
}

const demos = sourceSpecs.flatMap(parseFile);
for (const category of sourceSpecs.map(spec => spec.category)) {
  const inCategory = demos.filter(demo => demo.category === category);
  const newestNumber = Math.max(...inCategory.map(demo => Number(demo.sourceId.match(/(\d+)$/)[1])));
  for (const demo of inCategory) {
    demo.isNew = Number(demo.sourceId.match(/(\d+)$/)[1]) === newestNumber;
  }
}

const counts = Object.fromEntries(sourceSpecs.map(spec => [spec.category, demos.filter(demo => demo.category === spec.category).length]));
if (demos.length !== 64 || Object.values(counts).join(",") !== "21,21,11,11") {
  throw new Error(`Unerwartete Kataloggröße: ${demos.length} (${JSON.stringify(counts)})`);
}

const sourceCommit = require("node:child_process")
  .execFileSync("git", ["-C", sourceRoot, "rev-parse", "HEAD"], { encoding: "utf8" })
  .trim();

const header = `/* Automatisch aus dem VSTO-Katalog erzeugt.\n * Quelle: https://github.com/rstsu/Excel-VSTO-Toolbox\n * Erneut erzeugen: node scripts/import-vsto-catalog.js <Repository-Ordner> dist/demo-data.js <Version>\n */\n`;
const meta = {
  version,
  sourceCommit,
  repositoryUrl: "https://github.com/rstsu/Excel-VSTO-Toolbox",
  counts,
  packageDemoCount: demos.filter(demo => demo.package).length,
  uniquePackageCount: new Set(demos.filter(demo => demo.package).map(demo => demo.package.fileName)).size
};
const output = `${header}window.CATALOG_META = ${JSON.stringify(meta, null, 2)};\n\nwindow.DEMO_CATALOG = ${JSON.stringify(demos, null, 2)};\n`;
fs.writeFileSync(outputFile, output, "utf8");

console.log(`✓ ${demos.length} Demos importiert: ${JSON.stringify(counts)}`);
