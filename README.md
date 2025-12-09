# DOCX Parser

A modern, properly layered TypeScript document parser for Microsoft Word (.docx) files.

## Architecture

The parser follows a clean 4-stage architecture:

1. **Stage 1: Load DOCX** - Extracts the document from ZIP archive
2. **Stage 2: Extract structural blocks** - Parses XML to extract paragraphs and tables
3. **Stage 3: Semantic classification** - Classifies blocks using heuristics
4. **Stage 4: Chunk assembly** - (Future) Assembles sections, lists, and tables

## Features

- ✅ No Mammoth dependency
- ✅ Stable paragraph extraction
- ✅ Proper XML parsing with fast-xml-parser
- ✅ Semantic classification (HEADING, SUBHEADING, PARAGRAPH, LIST_ITEM, TABLE, etc.)
- ✅ Preserves formatting information (bold, italic, font size)
- ✅ Handles tables and nested structures
- ✅ TypeScript with full type safety

## Installation

```bash
npm install
```

## Usage

```typescript
import { parseDocument } from "./parser";
import * as fs from "fs";

const buffer = fs.readFileSync("document.docx");
const result = await parseDocument(buffer);

console.log(`Extracted ${result.blocks.length} blocks`);
console.log(`Classified ${result.semantic.length} semantic blocks`);
```

## Running Tests

```bash
npm run test
```

This will parse `Fire_Hazard_Audit_Report_Enhanced.docx` and display all classifications in the console, including:
- Summary statistics
- All blocks with their classifications
- Blocks grouped by type

## Exporting Classifications

To export classifications to a JSON file:

```bash
npm run export
```

Or specify a different document:

```bash
npm run export path/to/your/document.docx
```

This creates a `classifications.json` file with all classification data, including metadata and summary statistics.

## Project Structure

```
parser/
  ├── types.ts              # Type definitions
  ├── docxExtractor.ts      # Stage 2: XML structure extraction
  ├── semanticClassifier.ts # Stage 3: Heuristic classification
  └── index.ts              # Main entry point
```

## Classification Types

### Heuristic Classifier (Default)
- `HEADING` - Main section headings
- `SUBHEADING` - Subsection headings
- `PARAGRAPH` - Regular text paragraphs
- `LIST_ITEM` - Numbered or bulleted list items
- `TABLE` - Table structures
- `TABLE_TEXT` - Individual table cells (text content from tables)
- `CAPTION` - Figure/Table captions
- `FOOTNOTE` - Footnotes
- `HEADER` - Document headers
- `FOOTER` - Document footers
- `UNKNOWN` - Unclassified blocks

### LLM Classifier (Minimal Categories)
- `HEADING` - Section headers or titles
- `PARAGRAPH` - Normal prose blocks
- `LIST_ITEM` - Bulleted or numbered items
- `TABLE_TEXT` - Text from table cells

## Using LLM Classification

For more accurate classification, you can use an LLM (like Cursor's AI) to classify blocks:

### Step 1: Generate Prompts

```bash
npm run llm-prompts
```

This generates prompts for each block and saves them to `llm-prompts.txt`.

### Step 2: Classify with Cursor AI

1. Open `llm-prompts.txt`
2. Copy each prompt
3. Paste it into Cursor's AI chat
4. Copy the JSON response
5. Save all responses to a file (one JSON per line)

### Step 3: Process Responses

```bash
npm run process-llm responses.txt
```

This will parse the LLM responses and generate a classification report.

### Programmatic Usage

```typescript
import { formatBlockForClassification, parseLLMResponse } from "./parser";

const block = result.blocks[0];
const { prompt } = formatBlockForClassification(block);

// Use prompt with your LLM
const llmResponse = await callYourLLM(prompt);

// Parse the response
const classification = parseLLMResponse(llmResponse);
```

