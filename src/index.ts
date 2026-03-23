import * as fs from "fs";
import * as path from "path";
import { buildReport } from "./builder";
import { AnnualReport } from "./types";

async function main() {
  const args = process.argv.slice(2);

  if (args.length < 1) {
    console.error("Usage:   ts-node src/index.ts <report.json> [output.docx]");
    console.error("Example: ts-node src/index.ts data/sample.json output/report.docx");
    process.exit(1);
  }

  const jsonPath   = path.resolve(args[0]);
  const outputPath = args[1]
    ? path.resolve(args[1])
    : path.join(path.dirname(jsonPath), path.basename(jsonPath, ".json") + "_report.docx");

  if (!fs.existsSync(jsonPath)) {
    console.error(`❌ JSON file not found: ${jsonPath}`);
    process.exit(1);
  }

  console.log(`📄 Reading: ${jsonPath}`);

  let report: AnnualReport;
  try {
    const raw = fs.readFileSync(jsonPath, "utf-8");
    report = JSON.parse(raw);
  } catch (err) {
    console.error(`❌ Failed to parse JSON: ${err}`);
    process.exit(1);
  }

  if (!report.meta?.companyName || !report.meta?.reportYear || !report.chapters) {
    console.error("❌ JSON must contain: meta.companyName, meta.reportYear, chapters[]");
    process.exit(1);
  }

  console.log(`🏢 Company:  ${report.meta.companyName}`);
  console.log(`📅 Year:     ${report.meta.reportYear}`);
  console.log(`📚 Chapters: ${report.chapters.length}`);

  const totalSections = report.chapters.reduce((n, c) => n + c.sections.length, 0);
  console.log(`📝 Sections: ${totalSections}`);

  const jsonDir = path.dirname(jsonPath);

  try {
    console.log("\n⚙️  Building report...");
    const buffer = await buildReport(report, jsonDir);

    const outputDir = path.dirname(outputPath);
    if (!fs.existsSync(outputDir)) fs.mkdirSync(outputDir, { recursive: true });

    fs.writeFileSync(outputPath, buffer);
    const sizeKb = Math.round(buffer.length / 1024);
    console.log(`✅ Report saved: ${outputPath} (${sizeKb} KB)`);
  } catch (err) {
    console.error(`❌ Build failed: ${err}`);
    process.exit(1);
  }
}

main();
