# DOCX Parser & Quality Check System

A modern TypeScript-based document parser with semantic classification and quality checking capabilities for audit reports.

## Features

- **Document Parsing**: Extracts structured data from DOCX files
- **Semantic Classification**: Automatically classifies document blocks (headings, paragraphs, list items, etc.)
- **Page Number Extraction**: Accurate page number assignment using PDF conversion
- **Quality Checking**: Identifies grammar, spelling, formatting, and consistency issues
- **Interactive UI**: Web-based interface for document upload and quality check results

## Installation

```bash
npm install
```

## Usage

### Run the Web Server

```bash
npm run dev
```

Then open http://localhost:3000 in your browser.

### Run Tests

```bash
npm test
```

### Generate Preview

```bash
npm run preview
```

## Project Structure

```
docParser/
├── parser/              # Core parsing logic
│   ├── docxExtractor.ts    # DOCX structure extraction
│   ├── semanticClassifier.ts  # Semantic block classification
│   ├── pageNumberExtractor.ts # Page number extraction
│   └── types.ts            # Type definitions
├── public/              # Web UI
│   ├── index.html         # Main HTML
│   ├── style.css          # Styles
│   └── app.js             # Frontend logic
├── server.ts            # Express server
└── test.ts              # Test script
```

## API Endpoints

- `POST /api/quality-check` - Upload and process a document
- `POST /api/apply-changes` - Apply selected recommendations to document
- `GET /api/download/:filename` - Download updated document

## Quality Check Rules

Quality check rules will be implemented based on your requirements. The system currently supports:

- Grammar checking
- Spelling checking
- Formatting consistency
- Content consistency

## License

MIT
