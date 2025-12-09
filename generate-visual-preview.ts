import * as fs from "fs";
import * as path from "path";
import { parseDocument } from "./parser";

async function main() {
  // Allow command line argument for file path, or use default
  const fileName = process.argv[2] || "Healthcare_Audit_Report.docx";
  const docxPath = path.join(__dirname, fileName);
  
  if (!fs.existsSync(docxPath)) {
    console.error(`File not found: ${docxPath}`);
    console.error(`Please provide a valid DOCX file path as an argument, or place the file in the project root.`);
    process.exit(1);
  }

  console.log("Loading DOCX file...");
  const buffer = fs.readFileSync(docxPath);
  
  console.log("Parsing document...");
  const result = await parseDocument(buffer, docxPath);
  
  // Generate HTML preview
  let html = `<!DOCTYPE html>
<html>
<head>
  <meta charset="UTF-8">
  <title>Document Parser Preview</title>
  <style>
    body {
      font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
      margin: 20px;
      background: #f5f5f5;
    }
    .container {
      max-width: 1200px;
      margin: 0 auto;
      background: white;
      padding: 20px;
      box-shadow: 0 2px 10px rgba(0,0,0,0.1);
    }
    h1 {
      color: #333;
      border-bottom: 3px solid #4CAF50;
      padding-bottom: 10px;
    }
    .block {
      margin: 10px 0;
      padding: 10px;
      border-left: 4px solid #ddd;
      background: #fafafa;
      transition: all 0.2s;
    }
    .block:hover {
      background: #f0f0f0;
      border-left-color: #4CAF50;
    }
    .page-break {
      border-top: 3px dashed #ff6b6b;
      margin: 30px 0;
      padding: 10px;
      background: #fff3cd;
      text-align: center;
      font-weight: bold;
      color: #856404;
    }
    .heading {
      border-left-color: #2196F3;
      font-weight: bold;
      font-size: 1.1em;
    }
    .heading.level-1 {
      border-left-color: #2196F3;
      font-size: 1.3em;
    }
    .heading.level-2 {
      border-left-color: #03A9F4;
      font-size: 1.15em;
    }
    .list-item {
      border-left-color: #9C27B0;
      padding-left: 20px;
    }
    .paragraph {
      border-left-color: #757575;
    }
    .table {
      border-left-color: #FF9800;
      background: #fff8e1;
    }
    .table-text {
      border-left-color: #FF6F00;
      background: #fff3e0;
      font-family: 'Courier New', monospace;
      font-size: 0.9em;
    }
    .metadata {
      font-size: 0.85em;
      color: #666;
      margin-top: 5px;
      padding: 5px;
      background: white;
      border-radius: 3px;
    }
    .metadata span {
      display: inline-block;
      margin-right: 15px;
      padding: 2px 8px;
      background: #e3f2fd;
      border-radius: 3px;
    }
    .page-number {
      background: #ffebee !important;
      color: #c62828;
      font-weight: bold;
    }
    .line-number {
      background: #e8f5e9 !important;
      color: #2e7d32;
    }
    .summary {
      background: #e3f2fd;
      padding: 15px;
      border-radius: 5px;
      margin-bottom: 20px;
    }
    .summary h2 {
      margin-top: 0;
      color: #1976d2;
    }
    .summary-stats {
      display: grid;
      grid-template-columns: repeat(auto-fit, minmax(150px, 1fr));
      gap: 10px;
      margin-top: 10px;
    }
    .stat {
      background: white;
      padding: 10px;
      border-radius: 3px;
      text-align: center;
    }
    .stat-value {
      font-size: 1.5em;
      font-weight: bold;
      color: #1976d2;
    }
    .stat-label {
      font-size: 0.9em;
      color: #666;
    }
  </style>
</head>
<body>
  <div class="container">
    <h1>📄 Document Parser Preview</h1>
    
    <div class="summary">
      <h2>Summary</h2>
      <div class="summary-stats">
        <div class="stat">
          <div class="stat-value">${result.blocks.length}</div>
          <div class="stat-label">Total Blocks</div>
        </div>
        <div class="stat">
          <div class="stat-value">${result.semantic.length}</div>
          <div class="stat-label">Semantic Blocks</div>
        </div>
        <div class="stat">
          <div class="stat-value">${result.semantic.filter(b => b.type === "HEADING").length}</div>
          <div class="stat-label">Headings</div>
        </div>
        <div class="stat">
          <div class="stat-value">${result.semantic.filter(b => b.pageNumber).length > 0 ? Math.max(...result.semantic.filter(b => b.pageNumber).map(b => b.pageNumber!)) : 'N/A'}</div>
          <div class="stat-label">Total Pages</div>
        </div>
      </div>
    </div>
`;

  result.semantic.forEach((block, idx) => {

    // Determine CSS class based on type
    let blockClass = block.type.toLowerCase().replace("_", "-");
    if (block.type === "HEADING" || block.type === "SUBHEADING") {
      blockClass = `heading level-${block.headingLevel || 1}`;
    } else if (block.type === "TABLE_TEXT") {
      blockClass = "table-text";
    }

    const text = block.text || "(empty)";
    const textPreview = text.length > 200 ? text.substring(0, 200) + "..." : text;

    html += `
    <div class="block ${blockClass}">
      <div style="font-weight: bold; color: #333; margin-bottom: 5px;">
        ${idx + 1}. [${block.type}] ${textPreview.split('\n')[0]}
      </div>
      <div class="metadata">
        ${block.pageNumber ? `<span class="page-number">Page ${block.pageNumber}</span>` : ''}
        ${block.headingLevel ? `<span>Level ${block.headingLevel}</span>` : ''}
        ${block.listLevel !== undefined ? `<span>List Level ${block.listLevel}</span>` : ''}
        ${block.type === "TABLE_TEXT" ? `<span style="background: #ffebee;">applyGrammarRules: ${block.applyGrammarRules !== undefined ? block.applyGrammarRules : 'N/A'}</span>` : ''}
        ${block.type === "TABLE_TEXT" ? `<span style="background: #e8f5e9;">applySpellingRules: ${block.applySpellingRules !== undefined ? block.applySpellingRules : 'N/A'}</span>` : ''}
        ${block.type === "TABLE_TEXT" ? `<span style="background: #fff3e0;">applyCapitalizationRules: ${block.applyCapitalizationRules !== undefined ? block.applyCapitalizationRules : 'N/A'}</span>` : ''}
      </div>
    </div>`;
  });

  html += `
  </div>
</body>
</html>`;

  const outputPath = path.join(__dirname, "document-preview.html");
  fs.writeFileSync(outputPath, html);
  
  console.log(`\n✅ Visual preview generated: ${outputPath}`);
  console.log(`\nOpen this file in your browser to see the document structure.`);
}

main().catch(console.error);

