import { JARGON_MAP } from "./server";

// Test the jargon replacement logic
function testJargonReplacement() {
    console.log("=== Testing Jargon Replacement ===\n");
    
    // Test cases
    const testCases = [
        "We need to do this on account of the fact that it's important.",
        "The team is in possession of the documents.",
        "There are a large number of issues to address.",
        "He made a statement saying that it was correct.",
        "The building is in the vicinity of the park.",
        "Please contact the admin for more information.",
        "We compared option A vs option B.",
        "In order to complete the task, we need more time."
    ];
    
    console.log("Test cases:\n");
    testCases.forEach((text, index) => {
        console.log(`${index + 1}. Original: "${text}"`);
        
        let modifiedText = text;
        const foundJargon: string[] = [];
        
        // Apply all jargon replacements
        for (const jargon in JARGON_MAP) {
            const lower = text.toLowerCase();
            if (lower.includes(jargon)) {
                const simple = JARGON_MAP[jargon];
                const regex = new RegExp(jargon.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), 'gi');
                modifiedText = modifiedText.replace(regex, simple);
                foundJargon.push(`${jargon} → ${simple}`);
            }
        }
        
        if (foundJargon.length > 0) {
            console.log(`   Found jargon: ${foundJargon.join(', ')}`);
            console.log(`   Recommended: "${modifiedText}"`);
        } else {
            console.log(`   No jargon found`);
        }
        console.log();
    });
    
    // Test edge cases
    console.log("\n=== Edge Cases ===\n");
    const edgeCases = [
        "On Account Of The Fact That (capitalized)",
        "in order to do something in order to do something else (multiple instances)",
        "admin vs admin (multiple jargon phrases)"
    ];
    
    edgeCases.forEach((text, index) => {
        console.log(`${index + 1}. Original: "${text}"`);
        
        let modifiedText = text;
        const foundJargon: string[] = [];
        
        for (const jargon in JARGON_MAP) {
            const lower = text.toLowerCase();
            if (lower.includes(jargon)) {
                const simple = JARGON_MAP[jargon];
                const regex = new RegExp(jargon.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), 'gi');
                modifiedText = modifiedText.replace(regex, simple);
                foundJargon.push(`${jargon} → ${simple}`);
            }
        }
        
        if (foundJargon.length > 0) {
            console.log(`   Found jargon: ${foundJargon.join(', ')}`);
            console.log(`   Recommended: "${modifiedText}"`);
        } else {
            console.log(`   No jargon found`);
        }
        console.log();
    });
}

testJargonReplacement();

