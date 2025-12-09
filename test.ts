import * as fs from "fs";
import * as path from "path";
import { parseDocument } from "./parser";

async function main() {
  const docxPath = path.join(__dirname, "Healthcare_Audit_Report.docx");
  
  if (!fs.existsSync(docxPath)) {
    console.error(`File not found: ${docxPath}`);
    process.exit(1);
  }

  console.log("Loading DOCX file...");
  const buffer = fs.readFileSync(docxPath);
  
  console.log("Parsing document...");
  const result = await parseDocument(buffer, docxPath);
  
  console.log(`\n=== PARSING RESULTS ===\n`);
  console.log(`Total blocks extracted: ${result.blocks.length}`);
  console.log(`Total semantic blocks: ${result.semantic.length}\n`);
  
  // Show classification summary
  const classificationCounts: Record<string, number> = {};
  result.semantic.forEach(block => {
    classificationCounts[block.type] = (classificationCounts[block.type] || 0) + 1;
  });
  
  console.log("=== CLASSIFICATION SUMMARY ===");
  Object.entries(classificationCounts)
    .sort((a, b) => b[1] - a[1])
    .forEach(([type, count]) => {
      console.log(`${type}: ${count}`);
    });
  
  // Show ALL classifications in detail
  console.log("\n=== ALL CLASSIFICATIONS ===\n");
  result.semantic.forEach((block, idx) => {
    const textPreview = block.text.length > 100 
      ? block.text.substring(0, 100) + "..." 
      : block.text || "(empty)";
    
    let details = "";
    if (block.headingLevel !== undefined) {
      details = ` (Level ${block.headingLevel})`;
    }
    if (block.listLevel !== undefined) {
      details = ` (List Level ${block.listLevel})`;
    }
    if (block.pageNumber !== undefined) {
      details += ` [Page ${block.pageNumber}]`;
    }
    
    let output = `${String(idx + 1).padStart(3, " ")}. [${block.type.padEnd(12, " ")}]${details} ${textPreview}`;
    
    console.log(output);
  });
  
  // Group by type for easier review
  console.log("\n\n=== CLASSIFICATIONS BY TYPE ===\n");
  const blocksByType: Record<string, typeof result.semantic> = {};
  result.semantic.forEach(block => {
    if (!blocksByType[block.type]) {
      blocksByType[block.type] = [];
    }
    blocksByType[block.type].push(block);
  });
  
  Object.entries(blocksByType)
    .sort((a, b) => b[1].length - a[1].length)
    .forEach(([type, blocks]) => {
      console.log(`\n--- ${type} (${blocks.length} blocks) ---`);
      blocks.forEach((block, idx) => {
        const textPreview = block.text.length > 120 
          ? block.text.substring(0, 120) + "..." 
          : block.text || "(empty)";
        let extra = "";
        if (block.headingLevel !== undefined) extra += ` [H${block.headingLevel}]`;
        if (block.listLevel !== undefined) extra += ` [L${block.listLevel}]`;
        console.log(`  ${idx + 1}.${extra} ${textPreview}`);
      });
    });
}

main().catch(console.error);

