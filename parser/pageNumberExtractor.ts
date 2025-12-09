import * as fs from "fs";
import * as path from "path";
import * as os from "os";
import { exec } from "child_process";
import { promisify } from "util";
import * as pdfjs from "pdfjs-dist";
import { DocxBlock } from "./types";

const execAsync = promisify(exec);

interface PDFTextBlock {
  text: string;
  page: number;
}

/**
 * Convert DOCX to PDF using LibreOffice
 */
async function convertDocxToPDF(docxPath: string, outputDir: string): Promise<string> {
  const pdfPath = path.join(outputDir, "temp.pdf");
  
  // Try different paths for LibreOffice on macOS
  const possiblePaths = [
    "/Applications/LibreOffice.app/Contents/MacOS/soffice",
    "/usr/local/bin/soffice",
    "/opt/homebrew/bin/soffice"
  ];
  
  let sofficePath: string | null = null;
  
  // First try to find in known locations
  for (const p of possiblePaths) {
    if (fs.existsSync(p)) {
      sofficePath = p;
      break;
    }
  }
  
  // If not found, try which
  if (!sofficePath) {
    try {
      const { stdout } = await execAsync("which soffice 2>/dev/null");
      const trimmed = stdout.trim();
      if (trimmed && fs.existsSync(trimmed)) {
        sofficePath = trimmed;
      }
    } catch {
      // Ignore
    }
  }
  
  if (!sofficePath) {
    throw new Error("LibreOffice not found. Please install LibreOffice (brew install --cask libreoffice) or ensure 'soffice' is in PATH.");
  }
  
  // Convert DOCX to PDF
  const command = `"${sofficePath}" --headless --convert-to pdf --outdir "${outputDir}" "${docxPath}" 2>&1`;
  
  try {
    const { stdout, stderr } = await execAsync(command, { timeout: 30000 });
    
    // Check if PDF was created
    if (!fs.existsSync(pdfPath)) {
      // Sometimes LibreOffice creates PDF with different name
      const files = fs.readdirSync(outputDir).filter(f => f.endsWith('.pdf'));
      if (files.length > 0) {
        const actualPdfPath = path.join(outputDir, files[0]);
        // Rename to expected name
        fs.renameSync(actualPdfPath, pdfPath);
      } else {
        throw new Error(`PDF conversion failed. Output: ${stdout} ${stderr}`);
      }
    }
    
    return pdfPath;
  } catch (error: any) {
    throw new Error(`Failed to convert DOCX to PDF: ${error.message || error}`);
  }
}

/**
 * Extract text from PDF with page numbers
 */
async function extractTextFromPDF(pdfPath: string): Promise<PDFTextBlock[]> {
  const data = fs.readFileSync(pdfPath);
  const uint8Array = new Uint8Array(data);
  
  const loadingTask = pdfjs.getDocument({ data: uint8Array });
  const pdf = await loadingTask.promise;
  
  const blocks: PDFTextBlock[] = [];
  
  for (let pageNum = 1; pageNum <= pdf.numPages; pageNum++) {
    const page = await pdf.getPage(pageNum);
    const textContent = await page.getTextContent();
    
    const pageText = textContent.items
      .map((item: any) => item.str)
      .join(" ")
      .trim();
    
    if (pageText) {
      blocks.push({
        text: pageText,
        page: pageNum
      });
    }
  }
  
  return blocks;
}

/**
 * Match DOCX blocks to PDF pages by text similarity
 * Ensures pages progress sequentially (never go backwards)
 */
function matchBlocksToPDFPages(
  docxBlocks: DocxBlock[],
  pdfBlocks: PDFTextBlock[]
): Map<number, number> {
  const pageMap = new Map<number, number>();
  
  // Check if PDF has an extra cover page by looking for document title
  // Look for "INTERNAL AUDIT REPORT" or similar first heading on PDF pages
  let pageOffset = 0;
  if (pdfBlocks.length > 1 && docxBlocks.length > 0) {
    // Find the first significant heading in DOCX (like "INTERNAL AUDIT REPORT")
    let firstHeading = "";
    for (let i = 0; i < Math.min(10, docxBlocks.length); i++) {
      const block = docxBlocks[i];
      if (block.type === "paragraph") {
        const text = block.text?.trim() || "";
        // Look for all-caps headings or document titles
        if (text.length > 5 && text.length < 50 && text === text.toUpperCase()) {
          firstHeading = text;
          break;
        }
      }
    }
    
    if (firstHeading) {
      // Check which PDF page contains this heading
      const firstPageHasTitle = pdfBlocks[0].text.toUpperCase().includes(firstHeading.toUpperCase());
      const secondPageHasTitle = pdfBlocks.length > 1 && pdfBlocks[1].text.toUpperCase().includes(firstHeading.toUpperCase());
      
      // If title is on second page but not first, we have a 1-page offset
      if (!firstPageHasTitle && secondPageHasTitle) {
        pageOffset = 1;
      }
    } else {
      // Fallback: check if first PDF page is very short (likely blank/cover)
      const firstPdfLength = (pdfBlocks[0].text || "").trim().length;
      if (firstPdfLength < 50 && pdfBlocks.length > 1) {
        // First page is very short, likely a cover - apply offset
        pageOffset = 1;
      }
    }
  }
  
  let currentPDFPage = 1 + pageOffset;
  let cumulativeText = ""; // Track cumulative text to detect page transitions
  let blocksOnCurrentPage = 0; // Track how many blocks we've been on current page
  let foundTableOfContents = false; // Track when we find TOC
  let tocPageNumber = 0; // Track which page TOC is on
  let inTOCSection = false; // Track if we're still in TOC section (between TOC heading and first real section)
  
  for (let i = 0; i < docxBlocks.length; i++) {
    const block = docxBlocks[i];
    
    if (block.type === "table") {
      // For tables, assign page based on previous block
      // Also check if table spans to next page by looking at next block
      if (i > 0) {
        const prevPage = pageMap.get(i - 1);
        if (prevPage) {
          let tablePage = prevPage;
          
          // Check if table might span to next page
          // If next block is on a different page and is close to this table, table might span
          if (i < docxBlocks.length - 1) {
            const nextBlock = docxBlocks[i + 1];
            if (nextBlock.type === "paragraph") {
              const nextText = nextBlock.text?.trim() || "";
              // Try to find next block's page (will be assigned later, but we can estimate)
              // For now, assign table to previous block's page
              // Tables that span pages will be on the starting page
            }
          }
          
          pageMap.set(i, tablePage);
          currentPDFPage = tablePage + pageOffset; // Convert back to PDF page index
        }
      }
      continue;
    }
    
    const blockText = block.text.trim();
    if (!blockText) {
      // Empty blocks get the current page
      pageMap.set(i, currentPDFPage - pageOffset);
      continue;
    }
    
    // Detect Table of Contents - will set page number after matching
    if (blockText.toLowerCase() === "table of contents") {
      foundTableOfContents = true;
      inTOCSection = true; // We're entering TOC section
    }
    
    // If we're in TOC section, keep all blocks on the same page as TOC (Page 2)
    if (inTOCSection && tocPageNumber > 0) {
      // Check if this is a numbered item that could be a TOC entry
      const isNumberedItem = /^\d+\.\s+[A-Z]/.test(blockText) && blockText.length < 50;
      
      // Count how many numbered TOC entries we've seen since TOC heading
      let tocEntryCount = 0;
      for (let j = i - 1; j >= 0; j--) {
        const prevBlock = docxBlocks[j];
        if (prevBlock.type === "paragraph") {
          const prevText = prevBlock.text?.trim() || "";
          if (prevText.toLowerCase() === "table of contents") {
            break; // Found TOC heading, stop counting
          }
          if (/^\d+\.\s+[A-Z]/.test(prevText) && prevText.length < 50) {
            tocEntryCount++;
          }
        }
      }
      
      // If it's a numbered item, check if it's a TOC entry or real section
      if (isNumberedItem) {
        // Check if next block is a substantial paragraph (indicates real section, not TOC entry)
        let isLikelyRealSection = false;
        if (i < docxBlocks.length - 1) {
          const nextBlock = docxBlocks[i + 1];
          if (nextBlock.type === "paragraph") {
            const nextText = nextBlock.text?.trim() || "";
            // If next block is a substantial paragraph (not another numbered item), it's likely a real section
            if (nextText.length > 50 && !/^\d+\.\s+[A-Z]/.test(nextText)) {
              isLikelyRealSection = true;
            }
          }
        }
        
        if (isLikelyRealSection) {
          // This is the first real section heading - force to Page 3 and exit TOC section
          inTOCSection = false;
          currentPDFPage = 3 + pageOffset;
          blocksOnCurrentPage = 0;
          cumulativeText = blockText;
          pageMap.set(i, 3);
          continue;
        } else {
          // It's a TOC entry - keep it on TOC page
          pageMap.set(i, tocPageNumber);
          currentPDFPage = tocPageNumber + pageOffset;
          blocksOnCurrentPage++;
          cumulativeText = blockText;
          continue;
        }
      }
      
      // If we encounter a non-numbered block that's not a TOC entry, we've left TOC
      if (blockText.toLowerCase() !== "table of contents" && blockText.length > 30) {
        inTOCSection = false;
      }
    }
    
    // Add to cumulative text (for better matching)
    cumulativeText += " " + blockText;
    cumulativeText = cumulativeText.trim();
    
    // Check if block text appears on current page
    let currentPageScore = 0;
    let bestPage = currentPDFPage;
    let bestScore = 0;
    
    // Always check current page and next 2 pages (account for offset)
    const searchStart = Math.max(1, currentPDFPage);
    const searchEnd = Math.min(pdfBlocks.length, currentPDFPage + 2);
    
    // Check if this looks like a heading (short, starts with number or capital)
    const isHeading = blockText.length < 60 && (/^\d+\./.test(blockText) || /^[A-Z][a-z]+/.test(blockText));
    
    // For headings that might appear in TOC, use surrounding context for better matching
    let useContext = false;
    let contextText = blockText;
    if (isHeading && i > 0 && i < docxBlocks.length - 1) {
      // Get text from previous and next blocks for context
      const prevBlock = docxBlocks[i - 1];
      const nextBlock = docxBlocks[i + 1];
      if (prevBlock.type === "paragraph" && nextBlock.type === "paragraph") {
        const prevText = prevBlock.text?.trim().substring(0, 50) || "";
        const nextText = nextBlock.text?.trim().substring(0, 50) || "";
        if (prevText && nextText) {
          contextText = `${prevText} ${blockText} ${nextText}`;
          useContext = true;
        }
      }
    }
    
    for (let p = searchStart; p <= searchEnd; p++) {
      const pageText = pdfBlocks[p - 1].text;
      let score = calculateTextSimilarity(blockText, pageText);
      
      // For headings with context, also check context match
      if (useContext && isHeading) {
        const contextScore = calculateTextSimilarity(contextText, pageText);
        // Use the better of the two scores, but weight context slightly higher
        score = Math.max(score, contextScore * 1.1);
      }
      
      // For headings, require more exact matching (they might appear in TOC)
      if (isHeading && !useContext) {
        const normalizedBlock = blockText.toLowerCase().trim();
        const normalizedPage = pageText.toLowerCase();
        
        // Check for exact heading match (with word boundaries)
        const exactMatch = new RegExp(`\\b${normalizedBlock.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}\\b`, 'i');
        if (exactMatch.test(normalizedPage)) {
          score = Math.max(score, 0.7); // Boost for exact match
        } else if (normalizedPage.includes(normalizedBlock)) {
          score = score * 0.7; // More reduction for substring match (likely in TOC)
        }
      }
      
      if (p === currentPDFPage) {
        currentPageScore = score;
      }
      
      if (score > bestScore) {
        bestScore = score;
        bestPage = p;
      }
    }
    
    // Progressive page advancement with better distribution:
    // 1. Stay on current page if match is good (>= 0.4)
    // 2. Advance if current match is weak (< 0.35) AND next page is better (>= 0.3)
    // 3. After many blocks on same page (15+), advance if match is weak (< 0.4)
    // 4. Never jump more than 1 page ahead at a time
    
    const nextPageScore = (bestPage === currentPDFPage + 1) ? bestScore : 0;
    
    // For headings, be more willing to advance if match isn't very strong
    const headingAdvanceThreshold = isHeading ? 0.5 : 0.4;
    const shouldAdvance = 
      (currentPageScore < 0.3 && nextPageScore >= 0.4) ||
      (isHeading && currentPageScore < 0.4 && nextPageScore >= 0.35) ||
      (blocksOnCurrentPage >= 25 && currentPageScore < 0.3 && currentPDFPage < pdfBlocks.length);
    
      // Calculate assigned page first
    let assignedPage = currentPDFPage - pageOffset;
    
    if (shouldAdvance && currentPDFPage < pdfBlocks.length) {
      currentPDFPage++;
      // Apply page offset when assigning (subtract offset to get actual page number)
      assignedPage = currentPDFPage - pageOffset;
      pageMap.set(i, assignedPage);
      cumulativeText = blockText;
      blocksOnCurrentPage = 0;
    } else if (currentPageScore >= headingAdvanceThreshold) {
      // Good match on current page - stay
      assignedPage = currentPDFPage - pageOffset;
      pageMap.set(i, assignedPage);
      blocksOnCurrentPage++;
    } else {
      // Default: stay on current page
      assignedPage = currentPDFPage - pageOffset;
      pageMap.set(i, assignedPage);
      blocksOnCurrentPage++;
    }
    
    // Track TOC page number after assignment
    if (blockText.toLowerCase() === "table of contents" && tocPageNumber === 0) {
      tocPageNumber = assignedPage;
    }
    
    // AFTER assignment: If TOC is on Page 2 and this is a numbered heading that got assigned to Page 4+,
    // it's likely the first section heading and should be on Page 3
    if (foundTableOfContents && tocPageNumber === 2 && assignedPage >= 4) {
      const isNumberedHeading = /^\d+\.\s+[A-Z]/.test(blockText) && blockText.length < 50;
      if (isNumberedHeading) {
        // Check if Page 3 already has a numbered heading (excluding TOC entries)
        // TOC entries are typically right after "Table of Contents" heading
        let page3HasHeading = false;
        for (let j = 0; j < i; j++) {
          const prevAssignedPage = pageMap.get(j);
          if (prevAssignedPage === 3) {
            const prevBlock = docxBlocks[j];
            if (prevBlock.type === "paragraph") {
              const prevText = prevBlock.text?.trim() || "";
              // Check if this is a numbered heading AND not a TOC entry
              // TOC entries are typically within 10 blocks of "Table of Contents"
              const isNumbered = /^\d+\.\s+[A-Z]/.test(prevText) && prevText.length < 50;
              if (isNumbered) {
                // Check if it's likely a TOC entry (close to TOC heading)
                let isTOCEntry = false;
                for (let k = Math.max(0, j - 10); k < j; k++) {
                  const checkBlock = docxBlocks[k];
                  if (checkBlock.type === "paragraph") {
                    const checkText = checkBlock.text?.trim() || "";
                    if (checkText.toLowerCase() === "table of contents") {
                      isTOCEntry = true;
                      break;
                    }
                  }
                }
                if (!isTOCEntry) {
                  page3HasHeading = true;
                  break;
                }
              }
            }
          }
        }
        
        // If Page 3 doesn't have a numbered heading (excluding TOC), correct this one to Page 3
        if (!page3HasHeading) {
          pageMap.set(i, 3);
          currentPDFPage = 3 + pageOffset;
          blocksOnCurrentPage = 0;
        }
      }
    }
    
    // Limit cumulative text size to avoid memory issues
    if (cumulativeText.length > 200) {
      cumulativeText = cumulativeText.slice(-100);
    }
  }
  
  return pageMap;
}

/**
 * Calculate text similarity using multiple methods
 */
function calculateTextSimilarity(text1: string, text2: string): number {
  const normalize = (s: string) => s.toLowerCase().replace(/\s+/g, " ").trim();
  const norm1 = normalize(text1);
  const norm2 = normalize(text2);
  
  if (!norm1 || !norm2) return 0;
  
  // Exact match or containment (highest score)
  if (norm2.includes(norm1) && norm1.length > 5) {
    return 0.9;
  }
  
  // Check if significant portion of text1 is in text2
  const words1 = norm1.split(/\s+/).filter(w => w.length > 2);
  const words2 = norm2.split(/\s+/).filter(w => w.length > 2);
  
  if (words1.length === 0) return 0;
  
  // Count word matches
  let matches = 0;
  const wordSet2 = new Set(words2);
  for (const word of words1) {
    if (wordSet2.has(word)) {
      matches++;
    }
  }
  
  const wordMatchRatio = matches / words1.length;
  
  // Also check for substring matches (for short phrases)
  if (norm1.length < 50) {
    // For short text, check if it appears as a substring
    if (norm2.includes(norm1)) {
      return Math.max(0.8, wordMatchRatio);
    }
  }
  
  return wordMatchRatio;
}

/**
 * Extract page numbers for DOCX blocks by converting to PDF and matching
 */
export async function extractPageNumbers(
  docxPath: string,
  docxBlocks: DocxBlock[]
): Promise<Map<number, number>> {
  const tempDir = os.tmpdir();
  
  try {
    // Convert DOCX to PDF
    const pdfPath = await convertDocxToPDF(docxPath, tempDir);
    
    // Extract text from PDF with page numbers
    const pdfBlocks = await extractTextFromPDF(pdfPath);
    
    // Match DOCX blocks to PDF pages
    const pageMap = matchBlocksToPDFPages(docxBlocks, pdfBlocks);
    
    // Clean up temp PDF
    try {
      fs.unlinkSync(pdfPath);
    } catch {
      // Ignore cleanup errors
    }
    
    return pageMap;
  } catch (error) {
    console.warn(`Warning: Could not extract page numbers: ${error}`);
    return new Map();
  }
}

