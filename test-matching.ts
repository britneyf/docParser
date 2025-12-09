import * as fs from "fs";
import * as path from "path";
import { parseDocument } from "./parser";

async function testMatching() {
  const docxPath = path.join(__dirname, "Fire_Hazard_Audit_Report_Enhanced.docx");
  const buffer = fs.readFileSync(docxPath);
  
  const result = await parseDocument(buffer, docxPath);
  
  // Test the blockTextMap logic
  const blockTextMap = new Map<string, number>();
  result.blocks.forEach((block, index) => {
    if (block.type === "paragraph") {
      const normalizedText = block.text.trim().replace(/[ \t]+/g, ' ').replace(/\n[ \t]*/g, '\n').trim();
      blockTextMap.set(normalizedText, index);
      
      // Also map each line separately
      if (normalizedText.includes('\n')) {
        const lines = normalizedText.split('\n').map(line => line.trim()).filter(line => line.length > 0);
        for (const line of lines) {
          if (!blockTextMap.has(line)) {
            blockTextMap.set(line, index);
          }
        }
      }
    }
  });
  
  console.log("=== Testing blockTextMap ===");
  console.log(`Total entries in map: ${blockTextMap.size}`);
  
  const searchText1 = "Fire Hazard & Life Safety Controls – Branch Audit";
  const searchText2 = "ABC Bank";
  
  console.log(`\nSearching for: "${searchText1}"`);
  for (const [text, idx] of blockTextMap.entries()) {
    if (text.includes(searchText1) || searchText1.includes(text)) {
      console.log(`  Found in map: "${text.substring(0, 60)}..." -> block ${idx}`);
    }
  }
  
  console.log(`\nSearching for: "${searchText2}"`);
  for (const [text, idx] of blockTextMap.entries()) {
    if (text === searchText2 || text.includes(searchText2) || searchText2.includes(text)) {
      console.log(`  Found in map: "${text}" -> block ${idx}`);
    }
  }
  
  // Check block 3 specifically
  console.log(`\n=== Block 3 details ===`);
  const block3 = result.blocks[3];
  if (block3.type === "paragraph") {
    console.log(`Block 3 text: ${JSON.stringify(block3.text)}`);
    console.log(`Block 3 normalized: ${JSON.stringify(block3.text.trim().replace(/[ \t]+/g, ' ').replace(/\n[ \t]*/g, '\n').trim())}`);
    const lines = block3.text.split('\n').map(l => l.trim()).filter(l => l.length > 0);
    console.log(`Block 3 lines: ${JSON.stringify(lines)}`);
  }
}

testMatching().catch(console.error);

