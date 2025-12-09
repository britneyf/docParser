import * as fs from "fs";
import * as path from "path";
import JSZip from "jszip";
import { XMLParser } from "fast-xml-parser";

async function testXMLStructure() {
  const docxPath = path.join(__dirname, "Fire_Hazard_Audit_Report_Enhanced.docx");
  const buffer = fs.readFileSync(docxPath);
  
  const zip = await JSZip.loadAsync(buffer);
  const xml = await zip.file("word/document.xml")!.async("string");
  
  // Find paragraphs containing "Fire Hazard" or "ABC Bank"
  const parser = new XMLParser({ 
    ignoreAttributes: false,
    preserveOrder: true,
    parseAttributeValue: true
  });
  
  const doc = parser.parse(xml);
  
  // Find w:document -> w:body
  let document: any;
  if (Array.isArray(doc)) {
    const docItem = doc.find((item: any) => item["w:document"]);
    document = docItem ? docItem["w:document"] : null;
  } else {
    document = doc["w:document"];
  }
  
  let body: any;
  if (Array.isArray(document)) {
    const bodyItem = document.find((item: any) => item["w:body"]);
    body = bodyItem ? bodyItem["w:body"] : null;
  } else {
    body = document["w:body"];
  }
  
  if (!body || !Array.isArray(body)) {
    console.log("Body is not an array or not found");
    return;
  }
  
  console.log(`Total body items: ${body.length}\n`);
  
  // Check first 10 paragraphs
  let paraCount = 0;
  for (const item of body) {
    const keys = Object.keys(item);
    for (const key of keys) {
      if (key === "w:p") {
        const pValue = item[key];
        const paragraphs = Array.isArray(pValue) ? pValue : [pValue];
        
        for (const p of paragraphs) {
          paraCount++;
          
          // Extract text from this paragraph
          const runs = p["w:r"] ? (Array.isArray(p["w:r"]) ? p["w:r"] : [p["w:r"]]) : [];
          let paraText = "";
          
          for (const r of runs) {
            const tNode = r["w:t"];
            if (typeof tNode === "string") {
              paraText += tNode;
            } else if (Array.isArray(tNode)) {
              paraText += tNode.map((t: any) => typeof t === "string" ? t : (t?.["#text"] || "")).join("");
            } else if (tNode && typeof tNode === "object") {
              paraText += tNode["#text"] || "";
            }
          }
          
          if (paraText.includes("Fire Hazard") || paraText.includes("ABC Bank")) {
            console.log(`Paragraph ${paraCount}:`);
            console.log(`  Text: "${paraText}"`);
            console.log(`  Number of runs: ${runs.length}`);
            console.log(`  Raw XML (first 200 chars): ${JSON.stringify(JSON.stringify(p)).substring(0, 200)}`);
            console.log();
          }
          
          if (paraCount >= 10) break;
        }
      }
      if (paraCount >= 10) break;
    }
    if (paraCount >= 10) break;
  }
}

testXMLStructure().catch(console.error);

