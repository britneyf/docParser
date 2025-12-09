import JSZip from "jszip";
import { XMLParser } from "fast-xml-parser";
import { DocxBlock, DocxParagraph, DocxTable, DocxRun } from "./types";

export async function extractDocxStructure(buffer: Buffer): Promise<DocxBlock[]> {
  const zip = await JSZip.loadAsync(buffer);
  const xml = await zip.file("word/document.xml")!.async("string");
  const parser = new XMLParser({ 
    ignoreAttributes: false,
    preserveOrder: true, // Preserve order of elements
    parseAttributeValue: true
  });
  const doc = parser.parse(xml);
  
  // With preserveOrder, the structure is an array of objects
  // Find w:document in the parsed structure
  let document: any;
  if (Array.isArray(doc)) {
    const docItem = doc.find((item: any) => item["w:document"]);
    document = docItem ? docItem["w:document"] : null;
  } else {
    document = doc["w:document"];
  }
  
  if (!document) {
    throw new Error("Could not find document element");
  }
  
  // Find w:body in the document
  let body: any;
  if (Array.isArray(document)) {
    const bodyItem = document.find((item: any) => item["w:body"]);
    body = bodyItem ? bodyItem["w:body"] : null;
  } else {
    body = document["w:body"];
  }
  
  if (!body) {
    throw new Error("Could not find document body");
  }
  
  const blocks: DocxBlock[] = [];

  function extractParagraph(pNode: any): DocxParagraph {
    const runs: DocxRun[] = [];
    
    // With preserveOrder, pNode might be an array or have a different structure
    // Try to extract w:r runs from the paragraph node
    let rawRuns: any[] = [];
    
    if (Array.isArray(pNode)) {
      // If pNode is an array, find elements with w:r
      for (const item of pNode) {
        if (item && item["w:r"]) {
          const rItems = Array.isArray(item["w:r"]) ? item["w:r"] : [item["w:r"]];
          rawRuns.push(...rItems);
        }
      }
    } else if (pNode["w:r"]) {
      // Standard structure
      rawRuns = [].concat(pNode["w:r"]);
    }
    
    for (let runIndex = 0; runIndex < rawRuns.length; runIndex++) {
      const r = rawRuns[runIndex];
      
      // Check for line breaks (w:br) - this can be a separate run or within a run
      let hasLineBreak = false;
      let isLineBreakOnly = false;
      
      // Check if this run is ONLY a line break (no text)
      if (r["w:br"] !== undefined && r["w:br"] !== null && !r["w:t"]) {
        isLineBreakOnly = true;
        const br = Array.isArray(r["w:br"]) ? r["w:br"] : [r["w:br"]];
        for (const breakNode of br) {
          if (typeof breakNode === "string" || (breakNode && !breakNode["@_w:type"])) {
            hasLineBreak = true;
            isLineBreakOnly = true;
          } else {
            const breakType = breakNode["@_w:type"] || breakNode["@w:type"];
            if (breakType !== "page") {
              hasLineBreak = true;
              isLineBreakOnly = true;
            }
          }
        }
      } else if (r["w:br"] !== undefined && r["w:br"] !== null) {
        // Line break within a run that also has text
        const br = Array.isArray(r["w:br"]) ? r["w:br"] : [r["w:br"]];
        for (const breakNode of br) {
          if (typeof breakNode === "string") {
            hasLineBreak = true;
          } else {
            const breakType = breakNode["@_w:type"] || breakNode["@w:type"];
            if (breakType !== "page") {
              hasLineBreak = true;
            }
          }
        }
      }
      
      // If this run is only a line break, add it and continue
      if (isLineBreakOnly) {
        runs.push({ text: "\n", isBold: false, isItalic: false, fontSize: undefined });
        continue;
      }
      
      const tNode: any = r["w:t"];
      let text = "";
      
      if (typeof tNode === "string") {
        text = tNode;
      } else if (Array.isArray(tNode)) {
        // When w:t is an array, join elements - if there's a line break, join with newline
        const textParts: string[] = [];
        for (let i = 0; i < tNode.length; i++) {
          const part = typeof tNode[i] === "string" ? tNode[i] : tNode[i]?.["#text"] || "";
          if (part) {
            textParts.push(part);
            // If there's a line break and this isn't the last part, add newline
            if (hasLineBreak && i < tNode.length - 1) {
              textParts.push("\n");
            }
          }
        }
        text = textParts.join("");
      } else if (tNode) {
        text = tNode["#text"] || "";
      }
      
      const rPr = r["w:rPr"];
      const isBold = !!rPr?.["w:b"];
      const isItalic = !!rPr?.["w:i"];
      let fontSize: number | undefined;
      if (rPr?.["w:sz"]) {
        const szVal = rPr["w:sz"]["@_w:val"] || rPr["w:sz"]["@val"] || rPr["w:sz"];
        if (typeof szVal === "number") {
          fontSize = szVal / 2;
        } else if (typeof szVal === "string") {
          fontSize = parseInt(szVal) / 2;
        }
      }

      // Handle text and line breaks
      if (text) {
        runs.push({ text, isBold, isItalic, fontSize });
        // If there's a line break after text (and it's not already in the text), add it
        if (hasLineBreak && !text.endsWith("\n")) {
          runs.push({ text: "\n", isBold: false, isItalic: false, fontSize: undefined });
        }
      } else if (hasLineBreak && !isLineBreakOnly) {
        // Run with only a line break, no text (but we already handled isLineBreakOnly above)
        runs.push({ text: "\n", isBold: false, isItalic: false, fontSize: undefined });
      }
    }

    const text = runs.map(r => r.text).join("");
    const styleName = pNode["w:pPr"]?.["w:pStyle"]?.["@_w:val"];
    const alignment = pNode["w:pPr"]?.["w:jc"]?.["@_w:val"];
    const numberingPr = pNode["w:pPr"]?.["w:numPr"];
    const numbering = numberingPr
      ? {
          numId: parseInt(numberingPr["w:numId"]["@_w:val"]),
          level: parseInt(numberingPr["w:ilvl"]["@_w:val"]),
        }
      : null;

    return {
      type: "paragraph",
      runs,
      text,
      styleName,
      alignment,
      numbering
    };
  }

  function extractTable(tblNode: any): DocxTable {
    const rows = tblNode["w:tr"] || [];
    const parsedRows: DocxParagraph[][] = [];
    
    let tableRowCount = 0;
    let totalCellsExtracted = 0;
    let cellsWithText = 0;
    let cellsWithoutText = 0;

    for (const trRaw of [].concat(rows)) {
      const tr: any = trRaw;
      tableRowCount++;
      // With preserveOrder, tr might be an array of cell objects { "w:tc": {...} }
      let cells: any[] = [];
      if (Array.isArray(tr)) {
        // Array of cell objects - extract w:tc from each
        const cellItems = (tr as any[]).filter((item: any) => item && item["w:tc"]);
        cells = cellItems.map((item: any) => item["w:tc"]);
      } else if (tr && tr["w:tc"]) {
        // Standard structure - w:tc is an array or single object
        cells = Array.isArray(tr["w:tc"]) ? tr["w:tc"] : [tr["w:tc"]];
      }
      
      const parsedCells: DocxParagraph[] = [];

      for (const tc of cells) {
        if (!tc) continue;
        totalCellsExtracted++;
        
        // With preserveOrder, tc might be an array of paragraph objects { "w:p": {...} }
        let cellParas: any[] = [];
        if (Array.isArray(tc)) {
          // Array of paragraph objects - extract w:p from each
          const paraItems = (tc as any[]).filter((item: any) => item && item["w:p"]);
          cellParas = paraItems.map((item: any) => item["w:p"]);
        } else if (tc["w:p"]) {
          // Standard structure
          cellParas = Array.isArray(tc["w:p"]) ? tc["w:p"] : [tc["w:p"]];
        }
        
        // Each cell can have multiple paragraphs - preserve them with newlines
        if (cellParas.length > 0) {
          // Extract all paragraphs and preserve newlines between them
          const allRuns: DocxRun[] = [];
          let combinedText = "";
          
          for (let i = 0; i < cellParas.length; i++) {
            // Debug: inspect the paragraph structure for first few cells
            if (totalCellsExtracted <= 5 && i === 0) {
              console.log(`[docxExtractor] Cell ${totalCellsExtracted} paragraph structure:`, JSON.stringify(Object.keys(cellParas[i] || {})).substring(0, 200));
              console.log(`[docxExtractor] Cell ${totalCellsExtracted} is array?`, Array.isArray(cellParas[i]));
              console.log(`[docxExtractor] Cell ${totalCellsExtracted} has w:r?`, !!(cellParas[i] && cellParas[i]["w:r"]));
              if (Array.isArray(cellParas[i])) {
                console.log(`[docxExtractor] Cell ${totalCellsExtracted} array length:`, cellParas[i].length);
                for (let j = 0; j < Math.min(3, cellParas[i].length); j++) {
                  console.log(`[docxExtractor] Cell ${totalCellsExtracted} array[${j}] keys:`, JSON.stringify(Object.keys(cellParas[i][j] || {})).substring(0, 200));
                }
              }
              if (cellParas[i] && cellParas[i]["w:r"]) {
                const runs = [].concat(cellParas[i]["w:r"]);
                console.log(`[docxExtractor] Cell ${totalCellsExtracted} has ${runs.length} runs`);
                if (runs.length > 0) {
                  console.log(`[docxExtractor] Cell ${totalCellsExtracted} first run keys:`, JSON.stringify(Object.keys(runs[0] || {})).substring(0, 200));
                  console.log(`[docxExtractor] Cell ${totalCellsExtracted} first run has w:t?`, !!runs[0]["w:t"]);
                }
              }
            }
            
            const para = extractParagraph(cellParas[i]);
            allRuns.push(...para.runs);
            if (para.text) {
              if (combinedText) combinedText += "\n"; // Use newline, not space
              combinedText += para.text;
            }
            // Add newline run between paragraphs (except after last)
            if (i < cellParas.length - 1 && para.text) {
              allRuns.push({ text: "\n", isBold: false, isItalic: false, fontSize: undefined });
            }
          }
          
          // Debug logging for first few cells
          if (totalCellsExtracted <= 5) {
            console.log(`[docxExtractor] Extracted cell ${totalCellsExtracted} (row ${tableRowCount}): ${cellParas.length} paragraphs, text="${combinedText.substring(0, 50) || '(empty)'}${combinedText.length > 50 ? '...' : ''}", runs=${allRuns.length}`);
          }
          
          // Create a single cell with combined content (preserving newlines)
          parsedCells.push({
            type: "paragraph",
            runs: allRuns,
            text: combinedText, // Don't trim - preserve newlines
            isInTable: true
          });
          
          if (combinedText.trim()) {
            cellsWithText++;
          } else {
            cellsWithoutText++;
          }
        } else {
          // If no paragraphs in cell, create an empty one
          if (totalCellsExtracted <= 5) {
            console.log(`[docxExtractor] Extracted cell ${totalCellsExtracted} (row ${tableRowCount}): NO PARAGRAPHS FOUND`);
          }
          parsedCells.push({
            type: "paragraph",
            runs: [],
            text: "",
            isInTable: true
          });
          cellsWithoutText++;
        }
      }

      parsedRows.push(parsedCells);
    }
    
    console.log(`[docxExtractor] Extracted table: ${tableRowCount} rows, ${totalCellsExtracted} cells, ${cellsWithText} with text, ${cellsWithoutText} empty`);

    return { type: "table", rows: parsedRows };
  }

  // Process body elements in document order
  // With preserveOrder: true, body is an array where each item has a single key like "w:p" or "w:tbl"
  if (Array.isArray(body)) {
    // Body is an array - process in order
    for (const item of body) {
      const keys = Object.keys(item);
      for (const key of keys) {
        if (key === "w:p") {
          // Each item contains one paragraph (or array of paragraphs if nested)
          const pValue = item[key];
          if (Array.isArray(pValue)) {
            for (const p of pValue) {
              const para = extractParagraph(p);
              blocks.push(para);
            }
          } else {
            const para = extractParagraph(pValue);
            blocks.push(para);
          }
        } else if (key === "w:tbl") {
          // Each item contains one table
          const tblValue = item[key];
          // With preserveOrder, w:tbl is an array where:
          // - First item: w:tblPr (table properties)
          // - Subsequent items: w:tr (table rows) - each item has { "w:tr": {...} }
          if (Array.isArray(tblValue)) {
            // Filter out table properties (w:tblPr) and extract rows (w:tr)
            const rowItems = tblValue.filter((item: any) => item["w:tr"]);
            // Extract the w:tr objects from each item
            const rows = rowItems.map((item: any) => item["w:tr"]);
            // Combine into one table structure
            const table = extractTable({ "w:tr": rows });
            blocks.push(table);
          } else {
            const table = extractTable(tblValue);
            blocks.push(table);
          }
        }
        // Ignore other elements like w:sectPr
      }
    }
  } else {
    // Fallback: body is an object (old behavior)
    const paragraphs = body["w:p"] ? (Array.isArray(body["w:p"]) ? body["w:p"] : [body["w:p"]]) : [];
    const tables = body["w:tbl"] ? (Array.isArray(body["w:tbl"]) ? body["w:tbl"] : [body["w:tbl"]]) : [];
    
    // Process paragraphs
    for (const p of paragraphs) {
      const para = extractParagraph(p);
      blocks.push(para);
    }
    
    // Process tables
    for (const tbl of tables) {
      const table = extractTable(tbl);
      blocks.push(table);
    }
  }

  return blocks;
}

