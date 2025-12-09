export interface DocxRun {
  text: string;
  isBold?: boolean;
  isItalic?: boolean;
  fontSize?: number;
}

export interface DocxParagraph {
  type: "paragraph";
  runs: DocxRun[];
  text: string;
  styleName?: string;
  alignment?: string;
  numbering?: { level: number; numId: number } | null;
  isInTable?: boolean;
}

export interface DocxTable {
  type: "table";
  rows: DocxParagraph[][]; 
}

export type DocxBlock = DocxParagraph | DocxTable;

export interface SemanticBlock {
  type: "HEADING" | "SUBHEADING" | "PARAGRAPH" | "LIST_ITEM" | "TABLE" | "TABLE_TEXT" | "CAPTION" | "FOOTNOTE" | "HEADER" | "FOOTER" | "UNKNOWN";
  text: string;
  raw: DocxBlock;
  headingLevel?: number;
  listLevel?: number;
  // Table-specific properties
  applyGrammarRules?: boolean;
  applySpellingRules?: boolean;
  applyCapitalizationRules?: boolean;
  section?: string; // Current section name (inherited from parent heading)
  pageNumber?: number; // Page number where this block appears
}

