import { DocxBlock, DocxParagraph, DocxTable, DocxRun } from "./types";

/**
 * Escapes XML special characters
 */
function escapeXml(text: string): string {
  return text
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;");
}

/**
 * Serializes DocxBlock[] back to DOCX XML format
 * Builds XML manually to avoid issues with XMLBuilder and preserveOrder
 */
export function serializeDocxBlocks(blocks: DocxBlock[]): string {
  let xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
  xml += '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">\n';
  xml += '  <w:body>\n';

  for (const block of blocks) {
    if (block.type === "paragraph") {
      xml += serializeParagraph(block, 4);
    } else if (block.type === "table") {
      xml += serializeTable(block, 4);
    }
  }

  xml += '  </w:body>\n';
  xml += '</w:document>';
  return xml;
}

function serializeParagraph(para: DocxParagraph, indent: number): string {
  const indentStr = " ".repeat(indent);
  let xml = `${indentStr}<w:p>\n`;

  // Add paragraph properties if they exist
  if (para.styleName || para.alignment || para.numbering) {
    xml += `${indentStr}  <w:pPr>\n`;
    if (para.styleName) {
      xml += `${indentStr}    <w:pStyle w:val="${escapeXml(para.styleName)}"/>\n`;
    }
    if (para.alignment) {
      xml += `${indentStr}    <w:jc w:val="${escapeXml(para.alignment)}"/>\n`;
    }
    if (para.numbering) {
      xml += `${indentStr}    <w:numPr>\n`;
      xml += `${indentStr}      <w:numId w:val="${para.numbering.numId}"/>\n`;
      xml += `${indentStr}      <w:ilvl w:val="${para.numbering.level}"/>\n`;
      xml += `${indentStr}    </w:numPr>\n`;
    }
    xml += `${indentStr}  </w:pPr>\n`;
  }

  // Serialize runs
  for (const run of para.runs) {
    xml += `${indentStr}  <w:r>\n`;

    // Add run properties if they exist
    const hasFormatting = run.isBold || run.isItalic || run.fontSize;
    if (hasFormatting) {
      xml += `${indentStr}    <w:rPr>\n`;
      if (run.isBold) {
        xml += `${indentStr}      <w:b/>\n`;
      }
      if (run.isItalic) {
        xml += `${indentStr}      <w:i/>\n`;
      }
      if (run.fontSize) {
        xml += `${indentStr}      <w:sz w:val="${Math.round(run.fontSize * 2)}"/>\n`;
      }
      xml += `${indentStr}    </w:rPr>\n`;
    }

    // Handle text
    if (run.text === "\n") {
      // Line break only
      xml += `${indentStr}    <w:br/>\n`;
    } else if (run.text.includes("\n")) {
      // Text with line breaks
      const parts = run.text.split("\n");
      for (let i = 0; i < parts.length; i++) {
        if (parts[i]) {
          xml += `${indentStr}    <w:t>${escapeXml(parts[i])}</w:t>\n`;
        }
        if (i < parts.length - 1) {
          xml += `${indentStr}    <w:br/>\n`;
        }
      }
    } else if (run.text) {
      // Regular text
      xml += `${indentStr}    <w:t>${escapeXml(run.text)}</w:t>\n`;
    }

    xml += `${indentStr}  </w:r>\n`;
  }

  xml += `${indentStr}</w:p>\n`;
  return xml;
}

function serializeTable(table: DocxTable, indent: number): string {
  const indentStr = " ".repeat(indent);
  let xml = `${indentStr}<w:tbl>\n`;
  
  // Add table properties (minimal)
  xml += `${indentStr}  <w:tblPr/>\n`;

  // Serialize rows
  for (const row of table.rows) {
    xml += `${indentStr}  <w:tr>\n`;

    for (const cell of row) {
      xml += `${indentStr}    <w:tc>\n`;
      
      // Add cell properties (minimal)
      xml += `${indentStr}      <w:tcPr/>\n`;

      // Serialize cell content as paragraph
      if (cell.type === "paragraph") {
        xml += serializeParagraph(cell, indent + 6);
      }

      xml += `${indentStr}    </w:tc>\n`;
    }

    xml += `${indentStr}  </w:tr>\n`;
  }

  xml += `${indentStr}</w:tbl>\n`;
  return xml;
}

