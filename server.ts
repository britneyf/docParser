import express, { Request, Response } from "express";
import multer from "multer";
import * as path from "path";
import * as fs from "fs";
import { parseDocument, DocxBlock, DocxParagraph } from "./parser";
import JSZip from "jszip";

// Jargon to simple language mapping
export const JARGON_MAP: Record<string, string> = {
    "on account of the fact that": "because",
    "in possession of": "have",
    "a large number of": "many",
    "made a statement saying": "stated",
    "in the vicinity of": "near",
    "admin": "administration",
    "vs": "compared to",
    "in order to": "to"
};

// Bad/vague phrase patterns with replacements for observations
// Order matters: simpler patterns (single words) should come before complex patterns (phrases with capture groups)
export const BAD_PHRASE_PATTERNS: { pattern: RegExp; replacement: string | ((substring: string, ...args: any[]) => string) }[] = [
    // Garbage fillers - remove entirely (process these first before hedges)
    { pattern: /\bkind of\s+/gi, replacement: "" },
    { pattern: /\bsort of\s+/gi, replacement: "" },
    { pattern: /\bjust\s+/gi, replacement: "" },
    { pattern: /\bonly\s+/gi, replacement: "" },
    { pattern: /\bbasically\s+/gi, replacement: "" },
    { pattern: /\bliterally\s+/gi, replacement: "" },
    { pattern: /\btruly\s+/gi, replacement: "" },
    { pattern: /\bsurely\s+/gi, replacement: "" },
    
    // Hedging adverbs modifying adjectives - remove the hedge, keep the adjective
    // The captured group should preserve original case, but ensure it's lowercase (adjectives in middle of sentence)
    { pattern: /\bsomewhat\s+([a-zA-Z]+)/gi, replacement: (match, word) => {
        return word.toLowerCase();
    }},
    { pattern: /\bslightly\s+([a-zA-Z]+)/gi, replacement: (match, word) => {
        return word.toLowerCase();
    }},
    { pattern: /\bquite\s+([a-zA-Z]+)/gi, replacement: (match, word) => {
        return word.toLowerCase();
    }},
    { pattern: /\breally\s+([a-zA-Z]+)/gi, replacement: (match, word) => {
        return word.toLowerCase();
    }},
    { pattern: /\bfairly\s+([a-zA-Z]+)/gi, replacement: (match, word) => {
        return word.toLowerCase();
    }},
    { pattern: /\bpractically\s+([a-zA-Z]+)/gi, replacement: (match, word) => {
        return word.toLowerCase();
    }},
    { pattern: /\ba bit\s+([a-zA-Z]+)/gi, replacement: (match, word) => {
        return word.toLowerCase();
    }},
    { pattern: /\ba little\s+([a-zA-Z]+)/gi, replacement: (match, word) => {
        return word.toLowerCase();
    }},
    
    // Quantifiers - replace with more precise terms
    // Use function replacement to preserve capitalization
    { pattern: /\ba handful of\s+([a-zA-Z]+)/gi, replacement: (match, word, offset, string) => {
        // Check if this is at the start of a sentence (after number or at beginning)
        const beforeMatch = string.substring(Math.max(0, offset - 10), offset);
        const isSentenceStart = /^[^a-zA-Z]*$/.test(beforeMatch) || /^\d+\.\s*$/.test(beforeMatch.trim());
        const quantifier = isSentenceStart ? "Several" : "several";
        // Keep the noun lowercase (it's a noun, not an adjective)
        const noun = word.toLowerCase();
        return quantifier + " " + noun;
    }},
    { pattern: /\ba lot of\s+([a-zA-Z]+)/gi, replacement: (match, word, offset, string) => {
        const beforeMatch = string.substring(Math.max(0, offset - 10), offset);
        const isSentenceStart = /^[^a-zA-Z]*$/.test(beforeMatch) || /^\d+\.\s*$/.test(beforeMatch.trim());
        const quantifier = isSentenceStart ? "Many" : "many";
        const noun = word.toLowerCase();
        return quantifier + " " + noun;
    }},
    { pattern: /\ba number of\s+([a-zA-Z]+)/gi, replacement: (match, word, offset, string) => {
        const beforeMatch = string.substring(Math.max(0, offset - 10), offset);
        const isSentenceStart = /^[^a-zA-Z]*$/.test(beforeMatch) || /^\d+\.\s*$/.test(beforeMatch.trim());
        const quantifier = isSentenceStart ? "Multiple" : "multiple";
        const noun = word.toLowerCase();
        return quantifier + " " + noun;
    }},
    
    // Approximations
    { pattern: /\bclose to\s+([a-zA-Z]+)/gi, replacement: (match, word) => {
        // Keep the word lowercase (it's in the middle of a sentence)
        return "near " + word.toLowerCase();
    }},
    { pattern: /\balmost\s+([a-zA-Z]+)/gi, replacement: (match, word) => {
        // Preserve original capitalization
        return "nearly " + word;
    }},
    { pattern: /\bnearly\s+([a-zA-Z]+)/gi, replacement: (match, word) => {
        // Preserve original capitalization
        return word;
    }},
    
    // Weak verbs
    { pattern: /\bideate\b/gi, replacement: "think" },
    { pattern: /\bponder\b/gi, replacement: "consider" },
    { pattern: /\bthink about\b/gi, replacement: "consider" },
    { pattern: /\bthink through\b/gi, replacement: "analyze" },
    { pattern: /\bstudy\b/gi, replacement: "review" },
];

// Month abbreviations mapping
const MONTH_ABBREVIATIONS: Record<string, string> = {
    "Jan": "January",
    "Feb": "February",
    "Mar": "March",
    "Apr": "April",
    "Aug": "August",
    "Sept": "September",
    "Oct": "October",
    "Nov": "November",
    "Dec": "December"
};

const ALLOWED_ABBREVIATIONS = new Set(["May", "June", "July"]);

const app = express();
const PORT = process.env.PORT || 3000;

// Configure multer for file uploads
const projectRoot = path.resolve(__dirname, '..');
const uploadsDir = path.join(projectRoot, 'uploads');
const upload = multer({ 
    dest: uploadsDir,
    limits: { fileSize: 50 * 1024 * 1024 } // 50MB limit
});

// Serve static files
app.use(express.static('public'));

// Middleware
app.use(express.json());

// Request logging middleware - logs ALL requests
app.use((req: Request, res: Response, next) => {
    console.log(`\n[${new Date().toISOString()}] ${req.method} ${req.path}`);
    if (req.method === 'POST') {
        console.log('  Content-Type:', req.get('content-type'));
        console.log('  Has file:', !!req.file);
    }
    next();
});

// Routes
app.get('/', (req: Request, res: Response) => {
    res.sendFile(path.join(__dirname, '../public/index.html'));
});

// Test endpoint to verify server is working
app.get('/api/test', (req: Request, res: Response) => {
    console.log('TEST ENDPOINT CALLED');
    res.json({ message: 'Server is working', timestamp: new Date().toISOString() });
});

// Upload and process document
app.post('/api/quality-check', upload.single('document'), async (req: Request, res: Response) => {
    try {
        if (!req.file) {
            return res.status(400).json({ error: 'No file uploaded' });
        }

        const filePath = req.file.path;
        const buffer = fs.readFileSync(filePath);
        const reportType = req.body.reportType || 'draft';

        // Parse the document
        const result = await parseDocument(buffer, filePath);

        // TODO: Run quality checks based on rules
        // For now, return mock results
        const qualityResults = await runQualityChecks(result, reportType);

        // Clean up uploaded file
        fs.unlinkSync(filePath);

        res.json({
            success: true,
            results: qualityResults,
            summary: {
                totalBlocks: result.blocks.length,
                semanticBlocks: result.semantic.length
            }
        });
    } catch (error) {
        console.error('Error processing document:', error);
        res.status(500).json({ error: 'Failed to process document' });
    }
});

// Apply changes to document
app.post('/api/apply-changes', (req: Request, res: Response, next) => {
    // Log BEFORE multer processes the request
    console.log('\n\n========================================');
    console.log('=== APPLY CHANGES API CALLED (BEFORE MULTER) ===');
    console.log('========================================\n');
    next();
}, upload.single('document'), async (req: Request, res: Response) => {
    // Force immediate output - multiple methods to ensure it shows
    process.stdout.write('\n\n');
    process.stdout.write('========================================\n');
    process.stdout.write('=== APPLY CHANGES API CALLED (AFTER MULTER) ===\n');
    process.stdout.write('========================================\n');
    process.stdout.write(`Time: ${new Date().toISOString()}\n`);
    process.stdout.write('\n');
    
    console.error('\n\n========================================');
    console.error('=== APPLY CHANGES API CALLED ===');
    console.error('========================================\n');
    
    console.log('\n\n========================================');
    console.log('=== APPLY CHANGES API CALLED ===');
    console.log('========================================\n');
    
    try {
        console.log('Request body keys:', Object.keys(req.body));
        console.log('Request file:', req.file ? `${req.file.originalname} (${req.file.size} bytes)` : 'NONE');
        
        if (!req.file) {
            console.error('ERROR: No file uploaded');
            return res.status(400).json({ error: 'No file uploaded' });
        }

        if (!req.body.changes) {
            console.error('ERROR: No changes data in request body');
            return res.status(400).json({ error: 'No changes data provided' });
        }

        const filePath = req.file.path;
        console.log(`Reading file from: ${filePath}`);
        const buffer = fs.readFileSync(filePath);
        console.log(`File read: ${buffer.length} bytes`);
        
        let changes;
        try {
            changes = JSON.parse(req.body.changes);
            console.log(`Parsed changes: ${Array.isArray(changes) ? changes.length : 'NOT AN ARRAY'}`);
        } catch (parseError) {
            console.error('ERROR parsing changes JSON:', parseError);
            console.error('Raw changes data:', req.body.changes);
            return res.status(400).json({ error: 'Invalid changes data format' });
        }
        
        if (!Array.isArray(changes)) {
            console.error('ERROR: changes is not an array:', typeof changes);
            return res.status(400).json({ error: 'Changes must be an array' });
        }
        
        console.log(`\nReceived ${changes.length} changes to apply:`);
        changes.forEach((change: any, index: number) => {
            console.log(`  ${index + 1}. "${change.originalText?.substring(0, 40)}..." -> "${change.recommendedText?.substring(0, 40)}..."`);
        });

        // Apply changes to the document
        console.log('\nCalling applyChangesToDocument...');
        let updatedBuffer: Buffer;
        try {
            updatedBuffer = await applyChangesToDocument(buffer, changes);
            console.log('applyChangesToDocument returned buffer of size:', updatedBuffer.length);
        } catch (docError) {
            console.error('ERROR in applyChangesToDocument:', docError);
            if (docError instanceof Error) {
                console.error('Error message:', docError.message);
                console.error('Error stack:', docError.stack);
            }
            throw docError; // Re-throw to be caught by outer catch
        }

        // Save updated document
        const outputDir = path.join(projectRoot, 'output');
        console.log(`Output directory: ${outputDir}`);
        if (!fs.existsSync(outputDir)) {
            console.log(`Creating output directory: ${outputDir}`);
            fs.mkdirSync(outputDir, { recursive: true });
        }
        
        const outputPath = path.join(outputDir, `updated_${req.file.originalname}`);
        console.log(`Writing updated document to: ${outputPath}`);
        fs.writeFileSync(outputPath, updatedBuffer);
        
        const stats = fs.statSync(outputPath);
        console.log(`File saved successfully: ${stats.size} bytes`);
        console.log(`File exists: ${fs.existsSync(outputPath)}`);

        // Clean up uploaded file
        fs.unlinkSync(filePath);
        console.log('Cleaned up uploaded file');

        const downloadUrl = `/api/download/${path.basename(outputPath)}`;
        console.log(`\nSending response with downloadUrl: ${downloadUrl}`);
        
        res.json({
            success: true,
            downloadUrl: downloadUrl
        });
        
        console.log('\n=== APPLY CHANGES COMPLETE ===\n');
    } catch (error) {
        // Force error output to stderr
        process.stderr.write('\n!!! ERROR APPLYING CHANGES !!!\n');
        console.error('\n!!! ERROR APPLYING CHANGES !!!');
        console.error('Error type:', error instanceof Error ? error.constructor.name : typeof error);
        console.error('Error message:', error instanceof Error ? error.message : String(error));
        if (error instanceof Error) {
            console.error('Error stack:', error.stack);
            process.stderr.write(`Error: ${error.message}\n`);
            process.stderr.write(`Stack: ${error.stack}\n`);
        } else {
            console.error('Error (not Error instance):', error);
            process.stderr.write(`Error: ${String(error)}\n`);
        }
        res.status(500).json({ 
            error: 'Failed to apply changes', 
            details: error instanceof Error ? error.message : String(error),
            stack: error instanceof Error ? error.stack : undefined
        });
    }
});

// Download updated document
app.get('/api/download/:filename', (req: Request, res: Response) => {
    const filename = req.params.filename;
    console.log(`\n=== Download Request ===`);
    console.log(`Requested filename: ${filename}`);
    
    const filePath = path.join(projectRoot, 'output', filename);
    console.log(`Looking for file at: ${filePath}`);
    console.log(`File exists: ${fs.existsSync(filePath)}`);
    
    if (!fs.existsSync(filePath)) {
        console.error(`File not found: ${filePath}`);
        // List files in output directory for debugging
        const outputDir = path.join(projectRoot, 'output');
        if (fs.existsSync(outputDir)) {
            const files = fs.readdirSync(outputDir);
            console.log(`Files in output directory: ${files.join(', ')}`);
        } else {
            console.log(`Output directory does not exist: ${outputDir}`);
        }
        return res.status(404).json({ error: 'File not found', requested: filename, path: filePath });
    }

    const stats = fs.statSync(filePath);
    console.log(`File found: ${stats.size} bytes`);
    console.log(`Sending file...`);
    
    res.download(filePath, (err: Error | null) => {
        if (err) {
            console.error('Error downloading file:', err);
            res.status(500).json({ error: 'Failed to download file' });
        } else {
            console.log('File sent successfully');
        }
    });
});

// Quality check functions
async function runQualityChecks(parseResult: any, reportType: string): Promise<any[]> {
    const results: any[] = [];
    
    // Rule 1: Check if report title on first page is fully uppercase
    // Debug: Log all blocks to see page numbers
    console.log('=== Quality Check Debug ===');
    console.log('Total semantic blocks:', parseResult.semantic.length);
    const blocksWithPages = parseResult.semantic.filter((b: any) => b.pageNumber !== undefined);
    console.log('Blocks with page numbers:', blocksWithPages.length);
    
    // Find the first page number that appears (in case page 1 is actually page 0 or 2)
    const firstPageNumber = parseResult.semantic.find((b: any) => b.pageNumber !== undefined)?.pageNumber || 1;
    console.log('First page number found:', firstPageNumber);
    
    // Find all blocks on the first page
    // If no page numbers are assigned, check the first 10 blocks (likely on page 1)
    interface BlockWithIndex {
        block: any;
        index: number;
    }
    
    const firstPageBlocks: BlockWithIndex[] = parseResult.semantic
        .map((block: any, index: number): BlockWithIndex => ({ block, index }))
        .filter((item: BlockWithIndex) => {
            const { block, index } = item;
            const isPage1 = block.pageNumber === firstPageNumber || 
                           (block.pageNumber === undefined && index < 10) ||
                           block.pageNumber === 1;
            // Include all text block types (HEADING, PARAGRAPH, LIST_ITEM, SUBHEADING, etc.)
            // Exclude only non-text types like TABLE, FOOTNOTE, etc.
            const isValidType = block.type !== 'TABLE' && 
                               block.type !== 'FOOTNOTE' && 
                               block.type !== 'HEADER' && 
                               block.type !== 'FOOTER';
            const hasText = block.text && block.text.trim().length > 0;
            
            if (isPage1 && isValidType && hasText) {
                console.log(`Found page 1 block [${block.type}] (index ${index}): "${block.text.substring(0, 50)}..." (pageNumber: ${block.pageNumber})`);
            }
            
            return isPage1 && isValidType && hasText;
        });
    
    console.log(`Total blocks on page 1: ${firstPageBlocks.length}`);
    
    // Sort to check headings first (they're typically the report title)
    firstPageBlocks.sort((a: BlockWithIndex, b: BlockWithIndex) => {
        if (a.block.type === 'HEADING' && b.block.type !== 'HEADING') return -1;
        if (a.block.type !== 'HEADING' && b.block.type === 'HEADING') return 1;
        return a.index - b.index;
    });
    
    // Check each block on page 1
    firstPageBlocks.forEach((item: BlockWithIndex) => {
        const { block, index: blockIndex } = item;
        const text = block.text.trim();
        
        // Check if text contains letters (not just numbers/symbols)
        const hasLetters = /[a-zA-Z]/.test(text);
        if (!hasLetters) {
            return; // Skip blocks with no letters
        }
        
        // Check if the text is NOT fully uppercase
        // Allow for spaces, numbers, and punctuation, but all letters must be uppercase
        const lettersOnly = text.replace(/[^a-zA-Z]/g, '');
        const isUppercase = lettersOnly === lettersOnly.toUpperCase();
        
        console.log(`Checking: "${text.substring(0, 50)}..." - Letters only: "${lettersOnly}", Is uppercase: ${isUppercase}`);
        
        if (lettersOnly.length > 0 && !isUppercase) {
            console.log(`  -> ISSUE FOUND: Text is not fully uppercase`);
            results.push({
                id: `uppercase-title-${blockIndex}`,
                page: block.pageNumber || 1,
                section: getSectionName(parseResult.semantic, blockIndex),
                confidence: 1.0,
                severity: 'high',
                category: 'formatting',
                originalText: text,
                recommendedText: text.toUpperCase(),
                rationale: 'All text on the first page must be written in uppercase letters. The report title should be fully capitalized.'
            });
        }
    });
    
    console.log(`Total issues found from Rule 1: ${results.length}`);
    console.log('=== End Rule 1 Debug ===');

    // Rule 2: Jargon / heavy language detection
    // Only check paragraphs, list items, headings, and table text (not tables themselves)
    console.log('\n=== Rule 2: Jargon Detection ===');
    console.log(`Checking ${parseResult.semantic.length} semantic blocks...`);
    let jargonChecks = 0;
    const jargonIssuesBefore = results.length;
    for (let i = 0; i < parseResult.semantic.length; i++) {
        const block = parseResult.semantic[i];
        if (block.type !== "PARAGRAPH" && block.type !== "HEADING" && block.type !== "LIST_ITEM" && block.type !== "TABLE_TEXT") {
            continue;
        }

        const text = block.text;
        if (!text) continue;

        jargonChecks++;
        const lower = text.toLowerCase();

        // Collect all jargon phrases found in this block
        const foundJargon: Array<{ jargon: string; simple: string }> = [];
        let recommendedText = text;

        // Check each jargon phrase in the map and collect all matches
        for (const jargon in JARGON_MAP) {
            // Use word boundaries to avoid partial matches (e.g., "admin" shouldn't match "administration")
            const escapedJargon = jargon.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
            // For multi-word phrases, use word boundaries; for single words, require word boundaries
            const regexPattern = jargon.includes(' ') 
                ? `\\b${escapedJargon}\\b`  // Multi-word: word boundaries on both sides
                : `\\b${escapedJargon}\\b`;  // Single word: word boundaries on both sides
            const regex = new RegExp(regexPattern, 'gi');
            
            if (regex.test(text)) {
                const simple = JARGON_MAP[jargon];
                foundJargon.push({ jargon, simple });
                // Apply replacement to build up the recommended text
                recommendedText = recommendedText.replace(regex, simple);
            }
        }

        // If we found any jargon, create a single result with all replacements applied
        if (foundJargon.length > 0) {
            const jargonList = foundJargon.map(f => `"${f.jargon}"`).join(", ");
            const simpleList = foundJargon.map(f => `"${f.simple}"`).join(", ");
            
            console.log(`  Found ${foundJargon.length} jargon phrase(s) in block ${i} (${block.type}, page ${block.pageNumber || 'unknown'}): ${jargonList}`);

            results.push({
                id: `jargon-${i}-${foundJargon.map(f => f.jargon.replace(/\s+/g, '-')).join('-')}`,
                page: block.pageNumber || 1,
                section: getSectionName(parseResult.semantic, i),
                category: "content",
                severity: "medium",
                confidence: 0.95,
                originalText: text,
                recommendedText: recommendedText,
                rationale: `Replace jargon phrase(s) ${jargonList} with simpler term(s) ${simpleList}.`
            });
        }
    }
    const jargonIssuesFound = results.length - jargonIssuesBefore;
    console.log(`Rule 2: Checked ${jargonChecks} blocks, found ${jargonIssuesFound} jargon issues`);

    // Rule 3: Bad/vague phrase patterns detection (all content blocks, not headers/subheaders)
    // Check paragraphs, list items, and table text in all sections
    console.log('\n=== Rule 3: Bad Words Detection ===');
    let badWordChecks = 0;
    const badWordIssuesBefore = results.length;
    
    for (let i = 0; i < parseResult.semantic.length; i++) {
        const block = parseResult.semantic[i];
        
        // Only check content block types (not headers, subheaders, or tables themselves)
        if (block.type !== "PARAGRAPH" && block.type !== "LIST_ITEM" && block.type !== "TABLE_TEXT") {
            continue;
        }

        const section = getSectionName(parseResult.semantic, i);
        
        // Skip if this block is itself a header or subheader section
        // (We only want content blocks, not section titles)
        if (block.type === "HEADING" || block.type === "SUBHEADING") {
            continue;
        }

        const text = block.text;
        if (!text) continue;

        badWordChecks++;

        // Apply each pattern and collect matches
        let recommendedText = text;
        const foundPatterns: Array<{ pattern: string; match: string }> = [];
        
        for (const { pattern, replacement } of BAD_PHRASE_PATTERNS) {
            // First, collect all matches to track what we found
            const matches = Array.from(text.matchAll(pattern));
            if (matches.length > 0) {
                for (const match of matches) {
                    const matchArray = match as RegExpMatchArray;
                    if (matchArray[0]) {
                        foundPatterns.push({
                            pattern: pattern.toString(),
                            match: matchArray[0]
                        });
                    }
                }
                
                // Apply replacement using the regex pattern directly
                // This ensures capture groups like $1 work correctly
                if (typeof replacement === 'string') {
                    recommendedText = recommendedText.replace(pattern, replacement);
                } else if (typeof replacement === 'function') {
                    recommendedText = recommendedText.replace(pattern, replacement);
                }
            }
        }
        
        // Clean up extra spaces that might result from replacements
        recommendedText = recommendedText.replace(/\s+/g, ' ').trim();
        
        // Only create a result if we found patterns and the text changed
        if (foundPatterns.length > 0 && recommendedText !== text) {
            const matchedPhrases = foundPatterns.map(p => p.match).filter((v, i, a) => a.indexOf(v) === i);
            console.log(`  Found bad words in block ${i} (${block.type}, page ${block.pageNumber || 'unknown'}): ${matchedPhrases.join(', ')}`);
            
            results.push({
                id: `badphrase-${i}-${Date.now()}`,
                category: "Content",
                severity: "medium",
                page: block.pageNumber || 1,
                section: section,
                confidence: 0.9,
                originalText: text,
                recommendedText: recommendedText,
                rationale: `Weak or vague terms found: "${matchedPhrases.join('", "')}". Observations should use precise, measurable language.`
            });
        }
    }
    const badWordIssuesFound = results.length - badWordIssuesBefore;
    console.log(`Rule 3: Checked ${badWordChecks} content blocks (excluding headers/subheaders), found ${badWordIssuesFound} issues`);

    // Rule 4: Month abbreviations (all sections)
    console.log('\n=== Rule 4: Month Abbreviations Detection ===');
    
    // First, count how many TABLE_TEXT blocks we have
    const tableTextBlocks = parseResult.semantic.filter((b: any) => b.type === "TABLE_TEXT");
    console.log(`Found ${tableTextBlocks.length} TABLE_TEXT blocks in document`);
    if (tableTextBlocks.length > 0) {
        console.log(`Sample TABLE_TEXT blocks (first 5):`);
        tableTextBlocks.slice(0, 5).forEach((block: any, idx: number) => {
            console.log(`  [${idx}] "${block.text?.substring(0, 60) || '(empty)'}${block.text && block.text.length > 60 ? '...' : ''}" (page ${block.pageNumber || 'unknown'})`);
        });
    }
    
    let monthChecks = 0;
    let monthIssuesFound = 0;
    let tableTextChecks = 0;
    
    for (let i = 0; i < parseResult.semantic.length; i++) {
        const block = parseResult.semantic[i];

        // Only apply to content blocks
        if (!["PARAGRAPH", "LIST_ITEM", "TABLE_TEXT"].includes(block.type)) {
            continue;
        }

        const text = block.text;
        if (!text) continue;

        monthChecks++;
        if (block.type === "TABLE_TEXT") {
            tableTextChecks++;
        }
        
        const section = getSectionName(parseResult.semantic, i);

        // Collect all month abbreviations found in this block
        const foundAbbreviations: Array<{ abbr: string; full: string; isInDate: boolean }> = [];
        let recommendedText = text;
        
        // Pattern to detect if abbreviation is part of a date (Month Day Year format)
        const datePattern = /\b(\w+)\s+(\d{1,2})\s+(\d{4})\b/g;

        // Check each forbidden abbreviation and collect all matches
        for (const abbr in MONTH_ABBREVIATIONS) {
            const full = MONTH_ABBREVIATIONS[abbr];

            // Word boundary, match exact abbreviation
            const regex = new RegExp(`\\b${abbr}\\b`, "g");
            const matches = Array.from(text.matchAll(regex)) as RegExpMatchArray[];
            
            if (matches.length > 0) {
                // Check if this abbreviation is part of a date
                const isInDate = matches.some((match: RegExpMatchArray) => {
                    const beforeMatch = text.substring(Math.max(0, (match.index || 0) - 20), match.index || 0);
                    const afterMatch = text.substring((match.index || 0) + match[0].length, Math.min(text.length, (match.index || 0) + match[0].length + 20));
                    // Check if followed by " Day Year" pattern
                    return /\d{1,2}\s+\d{4}/.test(afterMatch.trim());
                });
                
                foundAbbreviations.push({ abbr, full, isInDate });
                
                // Apply replacement to build up the recommended text
                recommendedText = recommendedText.replace(regex, full);
            }
        }

        // If we found any abbreviations, create a single result with all replacements applied
        if (foundAbbreviations.length > 0) {
            // For dates, also add commas after the day
            const datesInText = foundAbbreviations.filter(f => f.isInDate);
            if (datesInText.length > 0) {
                // Add commas to dates: "Month Day Year" -> "Month Day, Year"
                const fullMonthNames = "January|February|March|April|May|June|July|August|September|October|November|December";
                const dateCommaPattern = new RegExp(`\\b(${fullMonthNames})\\s+(\\d{1,2})\\s+(\\d{4})\\b`, "g");
                recommendedText = recommendedText.replace(dateCommaPattern, '$1 $2, $3');
            }
            
            const abbrList = foundAbbreviations.map(f => `"${f.abbr}"`).join(", ");
            const fullList = foundAbbreviations.map(f => `"${f.full}"`).join(", ");
            
            const dateNote = datesInText.length > 0 ? " (dates also formatted with commas)" : "";
            
            console.log(`  Found ${foundAbbreviations.length} month abbreviation(s) in block ${i} (${block.type}, page ${block.pageNumber || 'unknown'}): ${abbrList}${dateNote}`);

            results.push({
                id: `month-abbrev-${i}-${foundAbbreviations.map(f => f.abbr).join('-')}`,
                page: block.pageNumber || 1,
                section: section,
                category: "formatting",
                severity: "medium",
                confidence: 0.95,
                originalText: text,
                recommendedText: recommendedText,
                rationale: `The abbreviation(s) ${abbrList} ${foundAbbreviations.length === 1 ? 'is' : 'are'} not permitted. Use ${fullList} instead.${datesInText.length > 0 ? ' Dates formatted with commas.' : ''}`
            });
            monthIssuesFound++;
        }

        // Ignore allowed abbreviations (May, June, July) - no action needed
    }
    console.log(`Rule 4: Checked ${monthChecks} blocks (${tableTextChecks} from tables), found ${monthIssuesFound} month abbreviation issues`);

    // Rule 5: Date format - add commas to dates (Month Day Year -> Month Day, Year)
    console.log('\n=== Rule 5: Date Format Standardization ===');
    const dateIssuesBefore = results.length;
    let dateChecks = 0;
    let dateIssuesFound = 0;
    
    // Pattern to match dates in format "Month Day Year" (without comma)
    // Matches both full month names and abbreviations: January 15 2026, Jan 15 2026, etc.
    const fullMonthNames = "January|February|March|April|May|June|July|August|September|October|November|December";
    const monthAbbreviations = Object.keys(MONTH_ABBREVIATIONS).join("|");
    const datePattern = new RegExp(`\\b(${fullMonthNames}|${monthAbbreviations})\\s+(\\d{1,2})\\s+(\\d{4})\\b`, "g");
    
    for (let i = 0; i < parseResult.semantic.length; i++) {
        const block = parseResult.semantic[i];

        // Only apply to content blocks
        if (!["PARAGRAPH", "LIST_ITEM", "TABLE_TEXT"].includes(block.type)) {
            continue;
        }

        const text = block.text;
        if (!text) continue;

        dateChecks++;
        
        // Check if text contains dates without commas
        // Only check for dates with FULL month names (not abbreviations - those are handled by Rule 4)
        const fullMonthNames = "January|February|March|April|May|June|July|August|September|October|November|December";
        const fullMonthDatePattern = new RegExp(`\\b(${fullMonthNames})\\s+(\\d{1,2})\\s+(\\d{4})\\b`, "g");
        const matches = Array.from(text.matchAll(fullMonthDatePattern)) as RegExpMatchArray[];
        
        if (matches.length > 0) {
            // Add commas: "Month Day Year" -> "Month Day, Year"
            const recommendedText = text.replace(fullMonthDatePattern, '$1 $2, $3');
            
            // Only create result if the text actually changed
            if (recommendedText !== text) {
                const dateList = matches.map(m => m[0]).filter((v, idx, arr) => arr.indexOf(v) === idx);
                
                console.log(`  Found ${matches.length} date(s) without comma in block ${i} (${block.type}, page ${block.pageNumber || 'unknown'}): ${dateList.join(', ')}`);
                
                const section = getSectionName(parseResult.semantic, i);
                results.push({
                    id: `date-format-${i}-${Date.now()}`,
                    page: block.pageNumber || 1,
                    section: section,
                    category: "formatting",
                    severity: "low",
                    confidence: 0.95,
                    originalText: text,
                    recommendedText: recommendedText,
                    rationale: `Add comma(s) to date(s) for standard formatting: ${dateList.map(d => `"${d}" -> "${d.replace(/(\w+)\s+(\d+)\s+(\d+)/, '$1 $2, $3')}"`).join(', ')}.`
                });
                dateIssuesFound++;
            }
        }
    }
    const dateIssuesFoundCount = results.length - dateIssuesBefore;
    console.log(`Rule 5: Checked ${dateChecks} blocks, found ${dateIssuesFoundCount} date format issues`);

    console.log(`\n=== Total Quality Check Results: ${results.length} issues found ===`);
    
    // Print all issues for verification
    console.log('\n=== All Quality Check Issues ===');
    results.forEach((result, index) => {
        console.log(`\nIssue ${index + 1}:`);
        console.log(`  ID: ${result.id}`);
        console.log(`  Category: ${result.category}`);
        console.log(`  Severity: ${result.severity}`);
        console.log(`  Page: ${result.page}`);
        console.log(`  Section: ${result.section}`);
        console.log(`  Confidence: ${result.confidence}`);
        console.log(`  Original: "${result.originalText}"`);
        console.log(`  Recommended: "${result.recommendedText}"`);
        console.log(`  Rationale: ${result.rationale}`);
    });
    console.log('\n=== End All Issues ===\n');

    return results;
}

function getSectionName(semanticBlocks: any[], currentIndex: number): string {
    // First, try to find an explicit HEADING or SUBHEADING
    for (let i = currentIndex - 1; i >= 0; i--) {
        const block = semanticBlocks[i];
        if (block.type === 'HEADING' || block.type === 'SUBHEADING') {
            return block.text;
        }
    }
    
    // Fallback: Look for paragraphs that look like headings
    // This handles cases where headings are misclassified as paragraphs
    for (let i = currentIndex - 1; i >= 0; i--) {
        const block = semanticBlocks[i];
        if (block.type === 'PARAGRAPH') {
            const text = block.text?.trim() || '';
            
            // Check if this looks like a heading:
            // 1. Numbered section header (e.g., "1. Executive Summary")
            // 2. Short text (likely a heading)
            // 3. Title case or all caps
            // 4. No ending punctuation
            const isNumberedHeader = /^\d+\.\s+[A-Z]/.test(text);
            const isShort = text.length < 80 && text.length > 0;
            const isAllCaps = text === text.toUpperCase() && text.length < 80;
            const isTitleCase = /^[A-Z][a-z]+(\s+[A-Z][a-z]+)*$/.test(text);
            const noEndingPunctuation = !/[.!?]$/.test(text);
            
            // Check if it's an observation-style heading
            const isObservationHeading = /^Observation\s+\d+:\s+[A-Z]/.test(text);
            
            // If it matches heading characteristics, treat it as a heading
            if ((isNumberedHeader || isObservationHeading || (isShort && (isAllCaps || isTitleCase) && noEndingPunctuation)) && 
                text.length > 3) {
                return text;
            }
        }
    }
    
    return 'Introduction';
}

async function applyChangesToDocument(buffer: Buffer, changes: any[]): Promise<Buffer> {
    console.log(`\n=== Applying ${changes.length} changes to document ===`);
    
    // Step 1: Load the DOCX file and get the original XML
    let zip: JSZip;
    try {
        console.log('Loading DOCX as ZIP archive...');
        zip = await JSZip.loadAsync(buffer);
        console.log('ZIP loaded successfully');
    } catch (zipError) {
        console.error('ERROR loading ZIP:', zipError);
        throw new Error(`Failed to load DOCX file: ${zipError instanceof Error ? zipError.message : String(zipError)}`);
    }
    
    // Get the main document XML
    const docFile = zip.file("word/document.xml");
    if (!docFile) {
        throw new Error('Could not find word/document.xml in DOCX file');
    }
    
    let documentXml: string;
    try {
        documentXml = await docFile.async("string");
        console.log(`Document XML length: ${documentXml.length} characters`);
    } catch (xmlError) {
        console.error('ERROR reading document.xml:', xmlError);
        throw new Error(`Failed to read document XML: ${xmlError instanceof Error ? xmlError.message : String(xmlError)}`);
    }
    
    // Step 2: Parse the document to get structured blocks for matching
    console.log('Parsing document structure for matching...');
    const { blocks } = await parseDocument(buffer);
    console.log(`Parsed ${blocks.length} blocks`);
    
    // Step 3: Create a map of block text to block index for quick lookup
    // Also map individual lines (split on newlines) to handle cases where a block contains multiple paragraphs
    const blockTextMap = new Map<string, number>();
    blocks.forEach((block, index) => {
        if (block.type === "paragraph") {
            // Normalize whitespace but preserve newlines
            const normalizedText = block.text.trim().replace(/[ \t]+/g, ' ').replace(/\n[ \t]*/g, '\n').trim();
            blockTextMap.set(normalizedText, index);
            
            // Also map each line separately (for blocks that contain multiple lines separated by newlines)
            if (normalizedText.includes('\n')) {
                const lines = normalizedText.split('\n').map(line => line.trim()).filter(line => line.length > 0);
                for (const line of lines) {
                    // Only add if not already in map (to avoid overwriting)
                    if (!blockTextMap.has(line)) {
                        blockTextMap.set(line, index);
                    }
                }
            }
        } else if (block.type === "table") {
            // For tables, map each cell
            block.rows.forEach((row, rowIndex) => {
                row.forEach((cell, cellIndex) => {
                    const normalizedText = cell.text.trim().replace(/[ \t]+/g, ' ').replace(/\n[ \t]*/g, '\n').trim();
                    blockTextMap.set(normalizedText, index);
                    
                    // Also map individual lines in table cells
                    if (normalizedText.includes('\n')) {
                        const lines = normalizedText.split('\n').map(line => line.trim()).filter(line => line.length > 0);
                        for (const line of lines) {
                            if (!blockTextMap.has(line)) {
                                blockTextMap.set(line, index);
                            }
                        }
                    }
                });
            });
        }
    });
    
    // Step 4: Apply changes using careful XML string manipulation
    // Only replace text within <w:t> tags to preserve XML structure
    let modifiedXml = documentXml;
    let replacementCount = 0;
    
    for (let i = 0; i < changes.length; i++) {
        const change = changes[i];
        const originalText = change.originalText.trim();
        const recommendedText = change.recommendedText.trim();
        // Normalize whitespace but preserve newlines (they indicate paragraph structure)
        const normalizedOriginal = originalText.replace(/[ \t]+/g, ' ').replace(/\n[ \t]*/g, '\n').trim();
        
        console.log(`\n--- Change ${i + 1}/${changes.length} ---`);
        console.log(`Original: "${originalText}"`);
        console.log(`Recommended: "${recommendedText}"`);
        
        // Find the block that matches
        let blockIndex = -1;
        for (const [text, idx] of blockTextMap.entries()) {
            if (text === normalizedOriginal || text.includes(normalizedOriginal)) {
                blockIndex = idx;
                break;
            }
        }
        
        if (blockIndex === -1) {
            console.log(`  ✗ WARNING: Could not find text to replace`);
            continue;
        }
        
        const block = blocks[blockIndex];
        console.log(`  Found matching block at index ${blockIndex}`);
        if (block.type === 'paragraph') {
            console.log(`  Block text: "${block.text}"`);
            console.log(`  Block text length: ${block.text.length}`);
        }
        console.log(`  Searching for normalized text: "${normalizedOriginal}"`);
        
        // Method 1: Try finding the exact paragraph that matches
        // Search through all paragraphs and find the best match
        console.log(`  Attempting paragraph-level search...`);
        const paraMatch = findExactParagraphMatch(modifiedXml, normalizedOriginal, originalText);
        if (paraMatch) {
            console.log(`  ✓ Found matching paragraph`);
            console.log(`  Paragraph text: "${paraMatch.paragraphText}"`);
            console.log(`  Paragraph XML length: ${paraMatch.paragraphXml.length}`);
            
            // Rebuild the paragraph with the new text
            const newPara = rebuildParagraphWithText(paraMatch.paragraphXml, normalizedOriginal, recommendedText);
            if (newPara !== paraMatch.paragraphXml) {
                modifiedXml = modifiedXml.replace(paraMatch.fullMatch, newPara);
                console.log(`  ✓ Replaced text in paragraph`);
                replacementCount++;
                continue;
            } else {
                console.log(`  ✗ Paragraph found but text replacement returned unchanged paragraph`);
            }
        } else {
            console.log(`  ✗ No matching paragraph found`);
        }
        
        // Method 2: Try replacing text only within <w:t> tags (exact match)
        // Pattern: <w:t>text</w:t> or <w:t xml:space="preserve">text</w:t>
        const escapedOriginal = escapeRegex(originalText);
        const textTagPattern = new RegExp(
            `(<w:t[^>]*>)(${escapedOriginal})(</w:t>)`,
            'gi'
        );
        
        const textTagMatches = modifiedXml.match(textTagPattern);
        if (textTagMatches && textTagMatches.length > 0) {
            modifiedXml = modifiedXml.replace(textTagPattern, (match, openTag, text, closeTag) => {
                return openTag + escapeXml(recommendedText) + closeTag;
            });
            console.log(`  ✓ Replaced ${textTagMatches.length} occurrence(s) within <w:t> tags`);
            replacementCount++;
            continue;
        }
        
        // Method 3: Try case-insensitive match within w:t tags
        const caseInsensitivePattern = new RegExp(
            `(<w:t[^>]*>)([^<]*?${escapeRegex(originalText)}[^<]*?)(</w:t>)`,
            'gi'
        );
        
        const caseMatches = modifiedXml.match(caseInsensitivePattern);
        if (caseMatches && caseMatches.length > 0) {
            modifiedXml = modifiedXml.replace(caseInsensitivePattern, (match, openTag, text, closeTag) => {
                const newText = text.replace(new RegExp(escapeRegex(originalText), 'gi'), recommendedText);
                return openTag + escapeXml(newText) + closeTag;
            });
            console.log(`  ✓ Replaced ${caseMatches.length} occurrence(s) using case-insensitive match`);
            replacementCount++;
            continue;
        }
        
        // Debug: Try to find the text in XML to see what's different
        console.log(`  Debug: Searching for first 20 chars: "${originalText.substring(0, 20)}"`);
        const first20 = originalText.substring(0, 20);
        if (modifiedXml.includes(first20)) {
            console.log(`  Found first 20 chars in XML`);
            // Try to find surrounding context
            const contextMatch = modifiedXml.match(new RegExp(`.{0,50}${escapeRegex(first20)}.{0,50}`, 'i'));
            if (contextMatch) {
                console.log(`  Context: "${contextMatch[0]}"`);
            }
        } else {
            console.log(`  First 20 chars not found in XML`);
        }
        
        console.log(`  ✗ WARNING: Could not find text to replace in XML`);
    }
    
    console.log(`\nTotal replacements made: ${replacementCount}/${changes.length}`);
    
    if (replacementCount === 0) {
        console.log(`\n⚠️  WARNING: No replacements were made! The document will be unchanged.`);
    }
    
    // Validate XML is not empty
    if (!modifiedXml || modifiedXml.trim().length === 0) {
        throw new Error('Modified XML is empty');
    }
    
    // Basic XML validation - check for well-formed tags
    const openTags = (modifiedXml.match(/<w:[^>]+>/g) || []).length;
    const closeTags = (modifiedXml.match(/<\/w:[^>]+>/g) || []).length;
    if (Math.abs(openTags - closeTags) > 10) { // Allow some difference for self-closing tags
        console.warn(`⚠️  WARNING: XML tag mismatch detected. Open tags: ${openTags}, Close tags: ${closeTags}`);
    }
    
    // Check for unclosed tags (basic check)
    if (modifiedXml.includes('<w:t>') && !modifiedXml.includes('</w:t>')) {
        throw new Error('XML structure appears corrupted: unclosed w:t tags detected');
    }
    
    console.log(`Modified XML length: ${modifiedXml.length} characters`);
    
    // Step 7: Update the DOCX file
    try {
        zip.file("word/document.xml", modifiedXml);
        console.log('Updated document.xml in ZIP');
    } catch (updateError) {
        console.error('ERROR updating ZIP file:', updateError);
        throw new Error(`Failed to update document in ZIP: ${updateError instanceof Error ? updateError.message : String(updateError)}`);
    }
    
    // Generate the new DOCX buffer
    let newBuffer: Buffer;
    try {
        console.log('Generating new DOCX buffer...');
        newBuffer = await zip.generateAsync({ 
            type: "nodebuffer",
            compression: "DEFLATE",
            compressionOptions: { level: 6 }
        });
        console.log(`Document modification complete. Original size: ${buffer.length} bytes, New size: ${newBuffer.length} bytes`);
    } catch (generateError) {
        console.error('ERROR generating DOCX buffer:', generateError);
        throw new Error(`Failed to generate updated DOCX: ${generateError instanceof Error ? generateError.message : String(generateError)}`);
    }
    
    console.log(`=== End Apply Changes ===\n`);
    
    return newBuffer;
}

/**
 * Escapes special regex characters in a string
 */
function escapeRegex(str: string): string {
    return str.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

/**
 * Finds the exact paragraph that matches the search text
 * Prioritizes exact matches and paragraphs where the text is the primary content
 */
function findExactParagraphMatch(xml: string, normalizedSearchText: string, originalSearchText: string): { fullMatch: string; paragraphXml: string; paragraphText: string } | null {
    // Match all paragraphs
    const paraPattern = /<w:p[^>]*>[\s\S]*?<\/w:p>/gi;
    const matches = Array.from(xml.matchAll(paraPattern));
    
    let bestMatch: { fullMatch: string; paragraphXml: string; paragraphText: string; score: number } | null = null;
    
    for (const match of matches) {
        const paraXml = match[0];
        const textPattern = /<w:t[^>]*>([^<]*)<\/w:t>/gi;
        const textMatches = Array.from(paraXml.matchAll(textPattern));
        
        if (textMatches.length === 0) {
            continue;
        }
        
        // Extract and decode text
        const textParts = textMatches.map(m => {
            let text = m[1];
            text = text.replace(/&#(\d+);/g, (match, num) => String.fromCharCode(parseInt(num, 10)));
            text = text.replace(/&#x([0-9a-fA-F]+);/g, (match, hex) => String.fromCharCode(parseInt(hex, 16)));
            text = text.replace(/&amp;/g, '&').replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&apos;/g, "'");
            return text;
        });
        
        // Preserve spacing between elements, including line breaks
        // Need to check for <w:r> runs that contain only <w:br/> between text runs
        let paraText = '';
        for (let i = 0; i < textParts.length; i++) {
            paraText += textParts[i];
            if (i < textMatches.length - 1) {
                const currentEnd = paraXml.indexOf(textMatches[i][0]) + textMatches[i][0].length;
                const nextStart = paraXml.indexOf(textMatches[i + 1][0], currentEnd);
                if (nextStart > currentEnd) {
                    const betweenText = paraXml.substring(currentEnd, nextStart);
                    // Check for line breaks - could be <w:br/> or <w:r><w:br/></w:r>
                    // Pattern: <w:r> that contains <w:br/> and no <w:t>
                    const lineBreakRunPattern = /<w:r[^>]*>[\s\S]*?<w:br[^>]*\/?>[\s\S]*?<\/w:r>/i;
                    if (betweenText.includes('<w:br') || betweenText.includes('<w:br/') || lineBreakRunPattern.test(betweenText)) {
                        paraText += '\n';
                    } else {
                        // Otherwise preserve whitespace
                        const whitespaceMatch = betweenText.match(/^(\s+)/);
                        if (whitespaceMatch) {
                            paraText += whitespaceMatch[1];
                        }
                    }
                }
            }
        }
        
        // Normalize only spaces/tabs, preserve newlines for paragraph structure
        const normalizedParaText = paraText.replace(/[ \t]+/g, ' ').replace(/\n[ \t]*/g, '\n').trim();
        
        // Score the match (higher is better)
        let score = 0;
        
        // Exact match gets highest score (this is what we want for "ABC Bank")
        if (normalizedParaText === normalizedSearchText) {
            score = 1000;
        }
        // Paragraph text exactly equals search text (case-insensitive but exact length match)
        else if (normalizedParaText.toLowerCase() === normalizedSearchText.toLowerCase() && 
                 Math.abs(normalizedParaText.length - normalizedSearchText.length) <= 2) {
            score = 950; // Very close to exact match
        }
        // Starts with search text (and paragraph is not much longer - likely the paragraph IS the search text)
        else if ((normalizedParaText.startsWith(normalizedSearchText + ' ') || 
                 normalizedParaText.startsWith(normalizedSearchText)) &&
                 normalizedParaText.length <= normalizedSearchText.length * 1.2) {
            score = 500;
        }
        // Ends with search text (and paragraph is not much longer)
        // Also check if it ends with newline + search text (last line in paragraph)
        else if ((normalizedParaText.endsWith(' ' + normalizedSearchText) || 
                 normalizedParaText.endsWith(normalizedSearchText) ||
                 normalizedParaText.endsWith('\n' + normalizedSearchText)) &&
                 normalizedParaText.length <= normalizedSearchText.length * 1.2) {
            score = 400;
        }
        // Check if search text is the last line in a multi-line paragraph
        else if (normalizedParaText.includes('\n') && 
                 normalizedParaText.split('\n').pop()?.trim() === normalizedSearchText) {
            score = 450; // Higher score for last line match
        }
        // Contains search text but paragraph is significantly longer (lower priority)
        else if (normalizedParaText.includes(normalizedSearchText)) {
            const lengthRatio = normalizedParaText.length / normalizedSearchText.length;
            if (lengthRatio <= 1.2) {
                score = 300 - Math.floor(lengthRatio * 10); // Prefer shorter paragraphs
            } else if (lengthRatio <= 1.5) {
                score = 200;
            } else {
                score = 50; // Much lower priority for paragraphs that are much longer
            }
        }
        
        // Also check original text (case-sensitive) - bonus points
        if (paraText === originalSearchText) {
            score += 100; // Big bonus for exact case-sensitive match
        } else if (paraText.includes(originalSearchText)) {
            score += 25; // Smaller bonus for case-sensitive contains
        }
        
        if (score > 0 && (!bestMatch || score > bestMatch.score)) {
            bestMatch = {
                fullMatch: match[0],
                paragraphXml: paraXml,
                paragraphText: normalizedParaText,
                score: score
            };
        }
    }
    
    if (bestMatch) {
        return {
            fullMatch: bestMatch.fullMatch,
            paragraphXml: bestMatch.paragraphXml,
            paragraphText: bestMatch.paragraphText
        };
    }
    
    return null;
}

/**
 * Finds a paragraph containing the specified text (legacy function, kept for compatibility)
 */
function findParagraphWithText(xml: string, searchText: string): { fullMatch: string; paragraphXml: string } | null {
    // Match a complete paragraph: <w:p ...>...</w:p>
    // Use non-greedy matching but ensure we get complete paragraphs
    const paraPattern = /<w:p[^>]*>[\s\S]*?<\/w:p>/gi;
    const matches = Array.from(xml.matchAll(paraPattern));
    
    for (const match of matches) {
        const paraXml = match[0];
        // Extract all text from w:t elements in this paragraph
        // Handle both simple <w:t>text</w:t> and <w:t xml:space="preserve">text</w:t>
        const textPattern = /<w:t[^>]*>([^<]*)<\/w:t>/gi;
        const textMatches = Array.from(paraXml.matchAll(textPattern));
        
        if (textMatches.length === 0) {
            continue; // No text in this paragraph
        }
        
        // Combine all text, handling XML entities
        const allTextParts = textMatches.map(m => {
            let text = m[1];
            // Decode numeric character entities (e.g., &#8211; for em dash, &#38; for &)
            text = text.replace(/&#(\d+);/g, (match, num) => String.fromCharCode(parseInt(num, 10)));
            text = text.replace(/&#x([0-9a-fA-F]+);/g, (match, hex) => String.fromCharCode(parseInt(hex, 16)));
            // Decode XML entities
            text = text
                .replace(/&amp;/g, '&')
                .replace(/&lt;/g, '<')
                .replace(/&gt;/g, '>')
                .replace(/&quot;/g, '"')
                .replace(/&apos;/g, "'");
            return text;
        });
        
        // Join text parts without adding spaces (preserve original structure)
        // Check for whitespace between elements in the original XML
        let allText = '';
        for (let i = 0; i < allTextParts.length; i++) {
            allText += allTextParts[i];
            // Check if there's whitespace between this element and the next in the original XML
            if (i < textMatches.length - 1) {
                const currentMatch = textMatches[i];
                const nextMatch = textMatches[i + 1];
                // Find the position after current match and before next match
                const currentEnd = paraXml.indexOf(currentMatch[0]) + currentMatch[0].length;
                const nextStart = paraXml.indexOf(nextMatch[0], currentEnd);
                if (nextStart > currentEnd) {
                    const betweenText = paraXml.substring(currentEnd, nextStart);
                    const whitespaceMatch = betweenText.match(/^(\s+)/);
                    if (whitespaceMatch) {
                        allText += whitespaceMatch[1];
                    }
                }
            }
        }
        
        // Normalize for comparison
        const normalizedAllText = allText.replace(/\s+/g, ' ').trim();
        const normalizedSearch = searchText.replace(/\s+/g, ' ').trim();
        
        // Debug: log first few paragraphs to see what we're comparing
        if (matches.indexOf(match) < 3) {
            console.log(`    Paragraph ${matches.indexOf(match) + 1} text: "${normalizedAllText.substring(0, 100)}"`);
            if (matches.indexOf(match) === 0) {
                console.log(`    Searching for: "${normalizedSearch}"`);
            }
        }
        
        // Check for exact match first (most precise)
        if (normalizedAllText === normalizedSearch) {
            return {
                fullMatch: match[0],
                paragraphXml: paraXml
            };
        }
        
        // Check if paragraph text starts with search text (handles cases where paragraph has more content)
        if (normalizedAllText.startsWith(normalizedSearch + ' ') || 
            normalizedAllText.startsWith(normalizedSearch)) {
            return {
                fullMatch: match[0],
                paragraphXml: paraXml
            };
        }
        
        // Check if paragraph text ends with search text
        if (normalizedAllText.endsWith(' ' + normalizedSearch) || 
            normalizedAllText.endsWith(normalizedSearch)) {
            return {
                fullMatch: match[0],
                paragraphXml: paraXml
            };
        }
        
        // Last resort: check if it contains the text (but be more careful)
        // Only match if the paragraph is not much longer than the search text
        if (normalizedAllText.includes(normalizedSearch) && 
            normalizedAllText.length <= normalizedSearch.length * 1.5) {
            return {
                fullMatch: match[0],
                paragraphXml: paraXml
            };
        }
    }
    
    return null;
}

/**
 * Escapes XML special characters
 */
function escapeXml(text: string): string {
    return text
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&apos;');
}

/**
 * Rebuilds a paragraph XML with new text, preserving structure
 */
function rebuildParagraphWithText(paraXml: string, oldText: string, newText: string): string {
    // Find all w:t elements in the paragraph
    const textPattern = /<w:t([^>]*)>([^<]*)<\/w:t>/gi;
    const textElements: Array<{ full: string; attrs: string; text: string; index: number }> = [];
    
    let match;
    while ((match = textPattern.exec(paraXml)) !== null) {
        textElements.push({
            full: match[0],
            attrs: match[1],
            text: match[2],
            index: match.index
        });
    }
    
    if (textElements.length === 0) {
        return paraXml; // No text elements found, return as-is
    }
    
    // Combine all text to verify it matches (handle XML entities)
    const textParts = textElements.map(e => {
        // Decode numeric character entities first
        let text = e.text;
        text = text.replace(/&#(\d+);/g, (m, num) => String.fromCharCode(parseInt(num, 10)));
        text = text.replace(/&#x([0-9a-fA-F]+);/g, (m, hex) => String.fromCharCode(parseInt(hex, 16)));
        // Decode XML entities
        text = text
            .replace(/&amp;/g, '&')
            .replace(/&lt;/g, '<')
            .replace(/&gt;/g, '>')
            .replace(/&quot;/g, '"')
            .replace(/&apos;/g, "'");
        return text;
    });
    
    // Join text parts - preserve original text exactly as it appears
    // Check for whitespace and line breaks between elements in the original XML
    let combinedText = '';
    for (let i = 0; i < textParts.length; i++) {
        combinedText += textParts[i];
        
        // Check if there's whitespace or line breaks between this element and the next in the original XML
        if (i < textElements.length - 1) {
            const currentEnd = textElements[i].index + textElements[i].full.length;
            const nextStart = textElements[i + 1].index;
            const betweenText = paraXml.substring(currentEnd, nextStart);
            
            // Check for line breaks first - could be <w:br/> or <w:r><w:br/></w:r>
            const lineBreakRunPattern = /<w:r[^>]*>[\s\S]*?<w:br[^>]*\/?>[\s\S]*?<\/w:r>/i;
            if (betweenText.includes('<w:br') || betweenText.includes('<w:br/') || lineBreakRunPattern.test(betweenText)) {
                combinedText += '\n';
            } else {
                // Otherwise preserve whitespace
                const whitespaceMatch = betweenText.match(/^(\s+)/);
                if (whitespaceMatch) {
                    combinedText += whitespaceMatch[1];
                }
            }
        }
    }
    
    // Normalize whitespace for comparison, but preserve newlines
    // Replace multiple spaces/tabs with single space, but keep newlines
    const normalizedCombined = combinedText.replace(/[ \t]+/g, ' ').replace(/\n[ \t]*/g, '\n').trim();
    const normalizedOld = oldText.replace(/[ \t]+/g, ' ').replace(/\n[ \t]*/g, '\n').trim();
    
    // Special case: if paragraph text exactly matches (after normalization), replace entire paragraph
    if (normalizedCombined === normalizedOld) {
        // If newText contains newlines, we need to preserve the line break structure
        // Check if the original paragraph had line breaks between text elements
        const hasLineBreaks = combinedText.includes('\n');
        
        if (hasLineBreaks && newText.includes('\n')) {
            // Split newText by newlines and distribute to text elements, preserving line break structure
            const newTextParts = newText.split('\n');
            let result = paraXml;
            
            // Find all <w:r> runs to preserve line break structure
            const runPattern = /<w:r[^>]*>[\s\S]*?<\/w:r>/gi;
            const runs = Array.from(paraXml.matchAll(runPattern));
            const textRunIndices: number[] = [];
            
            // Find which runs contain text elements
            for (let i = 0; i < runs.length; i++) {
                if (runs[i][0].includes('<w:t')) {
                    textRunIndices.push(i);
                }
            }
            
            // Replace text in each text element, preserving line breaks between them
            for (let i = 0; i < Math.min(newTextParts.length, textElements.length); i++) {
                const elem = textElements[i];
                const escapedPart = escapeXml(newTextParts[i]);
                const newElement = `<w:t${elem.attrs}>${escapedPart}</w:t>`;
                result = result.replace(elem.full, newElement);
            }
            
            // Clear remaining text elements if newText has fewer parts
            for (let i = newTextParts.length; i < textElements.length; i++) {
                const elem = textElements[i];
                const lastIndex = result.lastIndexOf(elem.full);
                if (lastIndex !== -1) {
                    const emptyElement = `<w:t${elem.attrs}></w:t>`;
                    result = result.substring(0, lastIndex) + emptyElement + result.substring(lastIndex + elem.full.length);
                }
            }
            
            return result;
        } else {
            // No line breaks or newText doesn't have newlines - simple replacement
            const escapedNewText = escapeXml(newText);
            let result = paraXml;
            
            // Replace the first element with the new text
            const firstElement = textElements[0];
            const newFirstElement = `<w:t${firstElement.attrs}>${escapedNewText}</w:t>`;
            result = result.replace(firstElement.full, newFirstElement);
            
            // Clear text from other elements (keep structure but empty)
            for (let i = 1; i < textElements.length; i++) {
                const elem = textElements[i];
                const lastIndex = result.lastIndexOf(elem.full);
                if (lastIndex !== -1) {
                    const emptyElement = `<w:t${elem.attrs}></w:t>`;
                    result = result.substring(0, lastIndex) + emptyElement + result.substring(lastIndex + elem.full.length);
                }
            }
            
            return result;
        }
    }
    
    // Check if the combined text contains the old text (substring replacement)
    if (normalizedCombined.includes(normalizedOld)) {
            // Find where oldText appears in the original (non-normalized) text
            // Try to match with flexible whitespace but preserve newlines
            const oldTextPattern = escapeRegex(oldText).replace(/[ \t]+/g, '[ \\t]+').replace(/\n/g, '\\n');
            const oldTextRegex = new RegExp(oldTextPattern, 'i');
            const match = combinedText.match(oldTextRegex);
            
            if (match && match.index !== undefined) {
                // Replace the matched text with newText, preserving surrounding text and structure
                const beforeMatch = combinedText.substring(0, match.index);
                const afterMatch = combinedText.substring(match.index + match[0].length);
                const newCombinedText = beforeMatch + newText + afterMatch;
                
                // Check if the replacement crosses a line break boundary
                const matchEndIndex = match.index + match[0].length;
                const hasNewlineBeforeMatch = beforeMatch.includes('\n');
                const hasNewlineInMatch = match[0].includes('\n');
                const hasNewlineAfterMatch = afterMatch.includes('\n');
                const newTextHasNewline = newText.includes('\n');
                
                // If the original had line breaks and we're preserving them, we need to handle them specially
                if (combinedText.includes('\n') && (hasNewlineInMatch || newTextHasNewline)) {
                    // Split the new combined text by newlines and distribute to text elements
                    const newTextParts = newCombinedText.split('\n');
                    let result = paraXml;
                    
                    // Replace text in each element, preserving line break structure
                    for (let i = 0; i < Math.min(newTextParts.length, textElements.length); i++) {
                        const elem = textElements[i];
                        const escapedPart = escapeXml(newTextParts[i]);
                        const newElement = `<w:t${elem.attrs}>${escapedPart}</w:t>`;
                        result = result.replace(elem.full, newElement);
                    }
                    
                    // Clear remaining text elements if newText has fewer parts
                    for (let i = newTextParts.length; i < textElements.length; i++) {
                        const elem = textElements[i];
                        const lastIndex = result.lastIndexOf(elem.full);
                        if (lastIndex !== -1) {
                            const emptyElement = `<w:t${elem.attrs}></w:t>`;
                            result = result.substring(0, lastIndex) + emptyElement + result.substring(lastIndex + elem.full.length);
                        }
                    }
                    
                    return result;
                } else {
                    // No line breaks involved - simple replacement
                    const escapedNewText = escapeXml(newCombinedText);
                    let result = paraXml;
                    
                    // Replace the first element with the new combined text
                    const firstElement = textElements[0];
                    const newFirstElement = `<w:t${firstElement.attrs}>${escapedNewText}</w:t>`;
                    result = result.replace(firstElement.full, newFirstElement);
                    
                    // Clear text from other elements (keep structure but empty)
                    for (let i = 1; i < textElements.length; i++) {
                        const elem = textElements[i];
                        const lastIndex = result.lastIndexOf(elem.full);
                        if (lastIndex !== -1) {
                            const emptyElement = `<w:t${elem.attrs}></w:t>`;
                            result = result.substring(0, lastIndex) + emptyElement + result.substring(lastIndex + elem.full.length);
                        }
                    }
                    
                    return result;
                }
            } else {
                // Fallback: try simple replace on normalized text
                const newCombinedText = normalizedCombined.replace(normalizedOld, newText);
                const escapedNewText = escapeXml(newCombinedText);
                
                let result = paraXml;
                const firstElement = textElements[0];
                const newFirstElement = `<w:t${firstElement.attrs}>${escapedNewText}</w:t>`;
                result = result.replace(firstElement.full, newFirstElement);
                
                for (let i = 1; i < textElements.length; i++) {
                    const elem = textElements[i];
                    const lastIndex = result.lastIndexOf(elem.full);
                    if (lastIndex !== -1) {
                        const emptyElement = `<w:t${elem.attrs}></w:t>`;
                        result = result.substring(0, lastIndex) + emptyElement + result.substring(lastIndex + elem.full.length);
                    }
                }
                
                return result;
            }
        }
    
    return paraXml; // No match, return as-is
}

// Create necessary directories
const dirs = ['uploads', 'output'];
dirs.forEach(dir => {
    const dirPath = path.join(projectRoot, dir);
    console.log(`Ensuring directory exists: ${dirPath}`);
    if (!fs.existsSync(dirPath)) {
        fs.mkdirSync(dirPath, { recursive: true });
        console.log(`Created directory: ${dirPath}`);
    } else {
        console.log(`Directory already exists: ${dirPath}`);
    }
});

// Global error handler for unhandled errors
process.on('unhandledRejection', (reason, promise) => {
    console.error('UNHANDLED REJECTION:', reason);
    process.stderr.write(`UNHANDLED REJECTION: ${reason}\n`);
});

process.on('uncaughtException', (error) => {
    console.error('UNCAUGHT EXCEPTION:', error);
    process.stderr.write(`UNCAUGHT EXCEPTION: ${error.message}\n`);
    process.stderr.write(`Stack: ${error.stack}\n`);
});

app.listen(PORT, () => {
    console.log(`Server running on http://localhost:${PORT}`);
    process.stdout.write(`Server running on http://localhost:${PORT}\n`);
});

