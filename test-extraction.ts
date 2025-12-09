import * as fs from "fs";
import * as path from "path";
import { parseDocument } from "./parser";

async function testExtraction() {
  const docxPath = path.join(__dirname, "Fire_Hazard_Audit_Report_Enhanced.docx");
  
  if (!fs.existsSync(docxPath)) {
    console.error(`File not found: ${docxPath}`);
    process.exit(1);
  }

  console.log("Loading DOCX file...");
  const buffer = fs.readFileSync(docxPath);
  
  console.log("Parsing document...");
  const result = await parseDocument(buffer, docxPath);
  
  console.log(`\nTotal blocks: ${result.blocks.length}`);
  console.log(`Total semantic blocks: ${result.semantic.length}\n`);
  
  // Find the blocks that should contain "Fire Hazard..." and "ABC Bank"
  console.log("=== Checking blocks for 'Fire Hazard' and 'ABC Bank' ===\n");
  
  for (let i = 0; i < result.blocks.length; i++) {
    const block = result.blocks[i];
    if (block.type === "paragraph") {
      const text = block.text;
      if (text.includes("Fire Hazard") || text.includes("ABC Bank")) {
        console.log(`Block ${i}:`);
        console.log(`  Text: "${text}"`);
        console.log(`  Text length: ${text.length}`);
        console.log(`  Number of runs: ${block.runs.length}`);
        console.log(`  Runs text: ${block.runs.map(r => `"${r.text}"`).join(" | ")}`);
        console.log(`  Has newline: ${text.includes('\n')}`);
        console.log(`  Raw text (with escapes): ${JSON.stringify(text)}`);
        console.log();
      }
    }
  }
  
  // Also check semantic blocks
  console.log("=== Checking semantic blocks ===\n");
  for (let i = 0; i < result.semantic.length; i++) {
    const semantic = result.semantic[i];
    if (semantic.text.includes("Fire Hazard") || semantic.text.includes("ABC Bank")) {
      console.log(`Semantic block ${i}:`);
      console.log(`  Type: ${semantic.type}`);
      console.log(`  Text: "${semantic.text}"`);
      console.log(`  Text length: ${semantic.text.length}`);
      console.log(`  Raw block index: ${result.blocks.indexOf(semantic.raw)}`);
      console.log();
    }
  }
}

testExtraction().catch(console.error);

