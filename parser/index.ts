import { extractDocxStructure } from "./docxExtractor";
import { classifyBlocks } from "./semanticClassifier";
import { extractPageNumbers } from "./pageNumberExtractor";
import { DocxBlock, SemanticBlock } from "./types";

export async function parseDocument(buffer: Buffer, docxPath?: string): Promise<{
  blocks: DocxBlock[];
  semantic: SemanticBlock[];
}> {
  const blocks = await extractDocxStructure(buffer);
  const semantic = classifyBlocks(blocks);
  
  // Add page numbers if docxPath is provided
  if (docxPath) {
    try {
      const pageMap = await extractPageNumbers(docxPath, blocks);
      
      // Assign page numbers to semantic blocks (without changing classification)
      // Track current page for blocks that were split
      let currentPage = 1;
      
      semantic.forEach((block) => {
        const blockIndex = blocks.findIndex(b => b === block.raw);
        if (blockIndex >= 0) {
          const pageNum = pageMap.get(blockIndex);
          if (pageNum) {
            block.pageNumber = pageNum;
            currentPage = pageNum; // Update current page
          } else {
            // If no match found, use current page (for split blocks)
            block.pageNumber = currentPage;
          }
        } else {
          // For blocks that don't match (shouldn't happen), use current page
          block.pageNumber = currentPage;
        }
      });
    } catch (error) {
      console.warn(`Warning: Could not extract page numbers: ${error}`);
    }
  }
  
  return { blocks, semantic };
}

export * from "./types";
export * from "./docxExtractor";
export * from "./semanticClassifier";
