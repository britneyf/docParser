import { extractDocxStructure } from "./docxExtractor";
import { classifyBlocks } from "./semanticClassifier";
import { DocxBlock, SemanticBlock } from "./types";

export async function parseDocument(buffer: Buffer): Promise<{
  blocks: DocxBlock[];
  semantic: SemanticBlock[];
}> {
  const blocks = await extractDocxStructure(buffer);
  const semantic = classifyBlocks(blocks);
  
  return { blocks, semantic };
}

export * from "./types";
export * from "./docxExtractor";
export * from "./semanticClassifier";
