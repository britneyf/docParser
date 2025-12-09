import * as fs from "fs";
import * as path from "path";
import { parseDocument } from "./parser";

async function main() {
  const docxPath = process.argv[2] || path.join(__dirname, "Fire_Hazard_Audit_Report_Enhanced.docx");
  
  if (!fs.existsSync(docxPath)) {
    console.error(`File not found: ${docxPath}`);
    process.exit(1);
  }

  console.log(`Loading DOCX file: ${docxPath}...`);
  const buffer = fs.readFileSync(docxPath);
  
  console.log("Parsing document...");
  const result = await parseDocument(buffer);
  
  // Export to JSON
  const output = {
    metadata: {
      totalBlocks: result.blocks.length,
      totalSemanticBlocks: result.semantic.length,
      document: path.basename(docxPath),
      parsedAt: new Date().toISOString()
    },
    classifications: result.semantic.map((block, idx) => ({
      index: idx + 1,
      type: block.type,
      text: block.text,
      headingLevel: block.headingLevel,
      listLevel: block.listLevel,
      pageNumber: block.pageNumber,
      lineNumber: block.lineNumber,
      metadata: {
        styleName: block.raw.type === "paragraph" ? block.raw.styleName : undefined,
        alignment: block.raw.type === "paragraph" ? block.raw.alignment : undefined,
        isInTable: block.raw.type === "paragraph" ? block.raw.isInTable : undefined,
        hasPageBreak: block.raw.type === "paragraph" ? block.raw.hasPageBreak : undefined,
      }
    })),
    summary: (() => {
      const counts: Record<string, number> = {};
      result.semantic.forEach(block => {
        counts[block.type] = (counts[block.type] || 0) + 1;
      });
      return counts;
    })()
  };
  
  const jsonPath = path.join(__dirname, "classifications.json");
  fs.writeFileSync(jsonPath, JSON.stringify(output, null, 2));
  
  console.log(`\n✅ Classifications exported to: ${jsonPath}`);
  console.log(`\nSummary:`);
  Object.entries(output.summary)
    .sort((a, b) => b[1] - a[1])
    .forEach(([type, count]) => {
      console.log(`  ${type}: ${count}`);
    });
}

main().catch(console.error);

