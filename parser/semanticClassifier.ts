import { DocxBlock, DocxParagraph, SemanticBlock } from "./types";

export function classifyBlocks(blocks: DocxBlock[]): SemanticBlock[] {
  const result: SemanticBlock[] = [];
  let isInTableOfContents = false; // Track if we're in a Table of Contents section

  for (const block of blocks) {
    if (block.type === "table") {
      // Process each cell in the table as TABLE_TEXT
      // For quality audit checking, we need all cells accessible
      let rowIndex = 0;
      let totalCells = 0;
      
      for (const row of block.rows) {
        const isHeaderRow = rowIndex === 0; // First row is typically headers
        
        // A row is an array of cells (DocxParagraph objects)
        for (const cell of row) {
          totalCells++;
          // Each cell is a DocxParagraph with isInTable = true
          // Extract text from cell - check both text property and runs
          let cellText = cell.text ? cell.text.trim() : "";
          
          // If text is empty but cell has runs, try to extract text from runs
          if (!cellText && cell.runs && cell.runs.length > 0) {
            cellText = cell.runs.map(r => r.text || "").join("").trim();
          }
          
          // For quality checking, extract all non-empty cells
          // This helps identify content that needs grammar/spelling checks
          if (cellText) {
            result.push({
              type: "TABLE_TEXT",
              text: cellText,
              raw: cell as DocxBlock,
              applyGrammarRules: false,
              applySpellingRules: true,
              applyCapitalizationRules: true,
              // For quality checking: headers should be checked for capitalization
              // Data cells should be checked for grammar/spelling
            });
          }
        }
        rowIndex++;
      }
      
      // Also add a TABLE marker for the structure itself
      // Include metadata about table size for quality checking
      result.push({ 
        type: "TABLE", 
        text: `[Table with ${rowIndex} rows, ${totalCells} cells]`, 
        raw: block 
      });
      continue;
    }

    const p = block as DocxParagraph;

    // Check if in table
    if (p.isInTable) {
      const tableText = p.text.trim();
      // Only add non-empty table cells
      if (tableText) {
        result.push({
          type: "TABLE_TEXT",
          text: tableText,
          raw: block,
          applyGrammarRules: false,
          applySpellingRules: true,
          applyCapitalizationRules: true,
          // section will be set by section builder/chunker if needed
        });
      }
      continue;
    }

    const text = p.text.trim();
    if (!text) {
      // Empty blocks typically indicate section breaks - reset TOC flag
      // But we don't need to include them in the semantic classification results
      if (isInTableOfContents) {
        isInTableOfContents = false;
      }
      continue; // Skip empty blocks - don't add them to result
    }

    // Check for caption
    // Rules:
    // 1. Must NOT be in a table (table cell text is not a caption)
    // 2. Must NOT be "Table of Contents"
    // 3. Must meet AT LEAST ONE of:
    //    - styleName includes "Caption"
    //    - text starts with Figure|Table|Chart|Image AND contains a number
    if (!p.isInTable && text.toLowerCase() !== "table of contents") {
      const hasCaptionStyle = p.styleName?.toLowerCase().includes("caption") || false;
      
      // Pattern: starts with Figure|Table|Chart|Image and contains a number
      const captionPattern = /^(Figure|Table|Chart|Image|Exhibit|Diagram)\s+[\dA-Z]/i;
      const hasCaptionPattern = captionPattern.test(text);
      
      // Must have at least one condition
      if (hasCaptionStyle || hasCaptionPattern) {
        result.push({
          type: "CAPTION",
          text,
          raw: block,
        });
        continue;
      }
    }

    // Special case: "Table of Contents" should always be a HEADING
    if (text.toLowerCase() === "table of contents") {
      isInTableOfContents = true; // Mark that we're entering TOC section
      result.push({
        type: "HEADING",
        text,
        raw: block,
        headingLevel: 1,
      });
      continue;
    }

    // Calculate heading score FIRST (before list item detection)
    // This ensures numbered section headers like "1. Executive Summary" are classified as headings
    const isBold = p.runs.some(r => r.isBold);
    const maxFontSize = Math.max(...p.runs.map(r => r.fontSize || 0), 0);
    const isAllCaps = text === text.toUpperCase() && text.length < 80 && text.length > 0;
    const hasHeadingStyle = p.styleName?.toLowerCase().includes("heading") || false;
    const noEndingPunctuation = !/[.!?]$/.test(text);
    const isShort = text.length < 80;
    const isVeryShort = text.length < 30;
    const isTitleCase = /^[A-Z][a-z]+(\s+[A-Z][a-z]+)*$/.test(text);
    
    // Check if this looks like a document title (all caps, short, no punctuation)
    const isDocumentTitle = isAllCaps && isShort && noEndingPunctuation && text.length > 5;
    
    // Check if this is a numbered section header (e.g., "1. Executive Summary")
    // Numbered items that are short, title case, and have heading characteristics should be headings
    // BUT: If we're in Table of Contents, numbered items should be LIST_ITEM, not HEADING
    const isNumberedSectionHeader = /^\d+\.\s+[A-Z]/.test(text) && isShort && noEndingPunctuation;
    
    // For numbered section headers, check if the text after the number is title case
    // Allow common words like "of", "the", "and", "in", "on", "at", "to", "for", "with"
    const textAfterNumber = text.replace(/^\d+\.\s+/, "");
    const titleCasePattern = /^[A-Z][a-z]+(\s+(of|the|and|in|on|at|to|for|with|a|an)\s+)*[A-Z][a-z]+(\s+[A-Z][a-z]+)*$/i;
    const isTitleCaseAfterNumber = titleCasePattern.test(textAfterNumber) || /^[A-Z][a-z]+(\s+[A-Z][a-z]+)*$/.test(textAfterNumber);
    
    // Check if this is an observation-style subheading (e.g., "Observation 1: Fire Extinguishers...")
    const isObservationSubheading = /^Observation\s+\d+:\s+[A-Z]/.test(text) && isShort && noEndingPunctuation;

    // If we're in Table of Contents and this is a numbered item, skip heading classification
    // It will be classified as LIST_ITEM below
    if (isInTableOfContents && isNumberedSectionHeader) {
      // Skip heading classification, fall through to list item detection
    } else {
      const headingScore =
        (hasHeadingStyle ? 5 : 0) +
        (isBold ? 2 : 0) +
        (maxFontSize >= 16 ? 2 : 0) +
        (isAllCaps && isShort ? 3 : 0) + // Increased weight for all-caps
        (isTitleCase && isVeryShort ? 2 : 0) +
        (isTitleCaseAfterNumber && isNumberedSectionHeader ? 2 : 0) + // Title case after number
        (noEndingPunctuation && isShort ? 1 : 0) +
        (isDocumentTitle ? 2 : 0) + // Bonus for document title pattern
        (isNumberedSectionHeader ? 3 : 0) + // Bonus for numbered section headers
        (isObservationSubheading ? 4 : 0); // Bonus for observation-style subheadings

      // Lower threshold for document titles, numbered section headers, observation subheadings, or if score is high enough
      // For numbered section headers with title case text, lower threshold to 5
      // For observation subheadings, lower threshold to 5
      if (headingScore >= 6 || 
          (isDocumentTitle && headingScore >= 4) || 
          (isNumberedSectionHeader && isTitleCaseAfterNumber && headingScore >= 5) ||
          (isObservationSubheading && headingScore >= 5)) {
        // Determine heading level from style name or pattern
        let headingLevel = 1;
        if (p.styleName) {
          const levelMatch = p.styleName.match(/(\d+)/);
          if (levelMatch) {
            headingLevel = parseInt(levelMatch[1]);
          } else if (p.styleName.toLowerCase().includes("heading")) {
            headingLevel = 2;
          }
        }
        
        // Observation-style items should be subheadings (level 2)
        if (isObservationSubheading) {
          headingLevel = 2;
        }

        // If we encounter a heading that's not "Table of Contents", we've left the TOC section
        if (isInTableOfContents && text.toLowerCase() !== "table of contents") {
          isInTableOfContents = false;
        }
        
        result.push({
          type: headingLevel === 1 ? "HEADING" : "SUBHEADING",
          text,
          raw: block,
          headingLevel,
        });
        continue;
      }
    }

    // Check for list item (only if not already classified as heading)
    if (p.numbering) {
      result.push({
        type: "LIST_ITEM",
        text,
        raw: block,
        listLevel: p.numbering.level,
      });
      continue;
    }

    // Enhanced list detection heuristics
    // Pattern 1: Traditional list items (bullet, dash, number at start)
    const traditionalListPattern = /^[\s]*([•\-\*]|\d+[\.\)]|\([a-z]\)|[a-z][\.\)])\s+/;
    
    // Pattern 2: Label-value pairs (text ending with colon, followed by value)
    // Examples: "Reviewed period:", "On-site review:", "Areas covered:"
    const labelValuePattern = /^[\s]*\*\*?[A-Z][^:]{0,50}:\*\*?\s+/i; // Bold label ending with colon
    const labelPattern = /^[\s]*[A-Z][^:]{0,50}:\s+/; // Regular label ending with colon
    
    // Pattern 3: Short lines that look like labels (ends with colon, under 60 chars)
    const shortLabelPattern = /^[\s]*[A-Za-z][^:]{0,50}:\s+.{1,100}$/;
    
    // If text contains newlines, split into separate paragraphs (one per line)
    // Each line that was created with Enter key should be a separate paragraph
    if (text.includes("\n")) {
      const lines = text.split("\n").filter(line => line.trim().length > 0);
      
      // Check if this looks like a structured list
      let looksLikeList = false;
      let matchingLines = 0;
      
      for (const line of lines) {
        const trimmed = line.trim();
        if (traditionalListPattern.test(trimmed) || 
            labelValuePattern.test(trimmed) || 
            labelPattern.test(trimmed) ||
            (shortLabelPattern.test(trimmed) && trimmed.length < 100)) {
          matchingLines++;
        }
      }
      
      // If at least 2 lines match list patterns, treat as list
      if (matchingLines >= 2) {
        looksLikeList = true;
      } else if (lines.length >= 2) {
        // Also check if lines are short and follow a pattern (likely list items)
        const shortLines = lines.filter(l => l.trim().length < 100);
        if (shortLines.length >= 2 && shortLines.length === lines.length) {
          // All lines are short - might be a list
          looksLikeList = true;
        }
      }
      
      if (looksLikeList) {
        // Split into individual list items
        for (const line of lines) {
          const trimmed = line.trim();
          if (trimmed.length > 0) {
            result.push({
              type: "LIST_ITEM",
              text: trimmed,
              raw: block,
              listLevel: 0,
            });
          }
        }
        continue;
      } else {
        // Not a list pattern - split into separate paragraphs (one per line)
        for (const line of lines) {
          const trimmed = line.trim();
          if (trimmed.length > 0) {
            // Check if this single line matches a list pattern
            if (traditionalListPattern.test(trimmed) || 
                labelValuePattern.test(trimmed) || 
                labelPattern.test(trimmed)) {
              result.push({
                type: "LIST_ITEM",
                text: trimmed,
                raw: block,
                listLevel: 0,
              });
            } else {
              // Default to paragraph for each line
              result.push({
                type: "PARAGRAPH",
                text: trimmed,
                raw: block,
              });
            }
          }
        }
        continue;
      }
    }
    
    // Check for concatenated list items (no newlines, but multiple list markers)
    // Only split dash-separated items (not bullet points), as bullets on same line should stay together
    // Pattern: "- Item1- Item2- Item3" (dash items that should be split)
    // Don't split bullet points (•) - they're often intentionally on the same line
    const dashListPattern = /([-\*]\s+[A-Z][^-\*]*?)(?=[-\*]\s+[A-Z]|$)/g;
    const dashMatches = Array.from(text.matchAll(dashListPattern));
    
    // Only split dash-separated items if they have substantial content
    // This handles cases like "- Assess fire safety- Verify compliance" which should be split
    if (dashMatches.length >= 2) {
      const hasSubstantialContent = dashMatches.some(match => {
        const content = match[1].replace(/^[-\*]\s+/, "");
        return content.length > 15; // At least 15 chars of content
      });
      
      // Only split if there's substantial content, indicating they should be separate
      if (hasSubstantialContent) {
        const parts: string[] = [];
        
        for (let i = 0; i < dashMatches.length; i++) {
          const match = dashMatches[i];
          const nextMatch = dashMatches[i + 1];
          
          if (match.index !== undefined) {
            const start = match.index;
            const end = nextMatch ? nextMatch.index : text.length;
            const part = text.substring(start, end).trim();
            
            if (part) {
              parts.push(part);
            }
          }
        }
        
        // If we successfully split into 2+ parts, add each as a list item
        if (parts.length >= 2) {
          for (const part of parts) {
            const trimmed = part.trim();
            if (trimmed) {
              result.push({
                type: "LIST_ITEM",
                text: trimmed,
                raw: block,
                listLevel: 0,
              });
            }
          }
          continue;
        }
      }
    }
    
    // Check for concatenated label-value pairs (no newlines, but multiple labels)
    // Pattern: "Label1: value1Label2: value2Label3: value3"
    // Find all label patterns: word(s) starting with capital, may be hyphenated, ending with colon
    // Match: [A-Z] followed by lowercase letters, spaces, hyphens, ending with colon
    // This should match: "Reviewed period:", "On-site review:", "Areas covered:"
    const labelStartPattern = /([A-Z][a-z]+(?:\s+[a-z]+)*(?:-[a-z]+)*\s*[a-z]*:)/g;
    const labelMatches = Array.from(text.matchAll(labelStartPattern));
    
    // If we find 2+ labels, split the text at each label boundary
    if (labelMatches.length >= 2) {
      const parts: string[] = [];
      
      for (let i = 0; i < labelMatches.length; i++) {
        const match = labelMatches[i];
        const nextMatch = labelMatches[i + 1];
        
        if (match.index !== undefined) {
          const start = match.index;
          const end = nextMatch ? nextMatch.index : text.length;
          const part = text.substring(start, end).trim();
          
          if (part) {
            parts.push(part);
          }
        }
      }
      
      // If we successfully split into 2+ parts, add each as a list item
      if (parts.length >= 2) {
        for (const part of parts) {
          const trimmed = part.trim();
          if (trimmed) {
            result.push({
              type: "LIST_ITEM",
              text: trimmed,
              raw: block,
              listLevel: 0,
            });
          }
        }
        continue;
      }
    }
    
    // Check for concatenated label-description pairs (no newlines, but multiple items)
    // Pattern: "High – DescriptionMedium – DescriptionLow – Description"
    // Find all positions where a capital word is followed by dash (em dash – or regular -)
    // Split the text at these boundaries
    const labelDashBoundaryPattern = /([A-Z][a-z]+\s*[–-])/g;
    const dashBoundaryMatches = Array.from(text.matchAll(labelDashBoundaryPattern));
    
    // If we find 2+ label-dash patterns, split the text at these boundaries
    if (dashBoundaryMatches.length >= 2) {
      const parts: string[] = [];
      
      for (let i = 0; i < dashBoundaryMatches.length; i++) {
        const match = dashBoundaryMatches[i];
        const nextMatch = dashBoundaryMatches[i + 1];
        
        if (match.index !== undefined) {
          const start = match.index;
          const end = nextMatch ? nextMatch.index : text.length;
          const part = text.substring(start, end).trim();
          
          if (part) {
            parts.push(part);
          }
        }
      }
      
      // If we successfully split into 2+ parts, add each as a list item
      if (parts.length >= 2) {
        for (const part of parts) {
          const trimmed = part.trim();
          if (trimmed) {
            result.push({
              type: "LIST_ITEM",
              text: trimmed,
              raw: block,
              listLevel: 0,
            });
          }
        }
        continue;
      }
    }
    
    // Single list item detection (no newlines or only one item)
    if (traditionalListPattern.test(text) || 
        labelValuePattern.test(text) || 
        labelPattern.test(text)) {
      result.push({
        type: "LIST_ITEM",
        text,
        raw: block,
        listLevel: 0,
      });
      continue;
    }

    // Check for missing line breaks in text (capital letter immediately followed by capital)
    // Pattern: "AuditABC Bank" should be split into "Audit" and "ABC Bank"
    // This handles cases where line breaks weren't properly extracted from DOCX
    const missingLineBreakPattern = /([a-z])([A-Z])/g;
    const lineBreakMatches = Array.from(text.matchAll(missingLineBreakPattern));
    
    // If we find potential missing line breaks, split the text
    if (lineBreakMatches.length > 0) {
      // Split at positions where lowercase is followed by uppercase
      const parts: string[] = [];
      let lastIndex = 0;
      
      for (const match of lineBreakMatches) {
        if (match.index !== undefined) {
          // Split before the capital letter (after the lowercase)
          const splitPos = match.index + 1;
          if (splitPos > lastIndex) {
            const part = text.substring(lastIndex, splitPos).trim();
            if (part) {
              parts.push(part);
            }
            lastIndex = splitPos;
          }
        }
      }
      
      // Add remaining text
      if (lastIndex < text.length) {
        const remaining = text.substring(lastIndex).trim();
        if (remaining) {
          parts.push(remaining);
        }
      }
      
      // If we successfully split into 2+ parts, create separate paragraphs
      if (parts.length >= 2) {
        for (const part of parts) {
          const trimmed = part.trim();
          if (trimmed) {
            result.push({
              type: "PARAGRAPH",
              text: trimmed,
              raw: block,
            });
          }
        }
        // Reset TOC flag if needed
        if (isInTableOfContents) {
          isInTableOfContents = false;
        }
        continue;
      }
    }
    
    // Default to paragraph
    // If we encounter a paragraph (not a list item), we've left the TOC section
    if (isInTableOfContents) {
      isInTableOfContents = false;
    }
    
    result.push({
      type: "PARAGRAPH",
      text,
      raw: block,
    });
  }

  return result;
}

