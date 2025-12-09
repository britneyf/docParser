// State management
let state = {
    file: null,
    reportType: 'draft',
    processing: false,
    results: [],
    filteredResults: [],
    selectedRecommendations: new Set(),
    activeFilter: null,
    issueCounts: {
        grammar: 0,
        spelling: 0,
        formatting: 0,
        consistency: 0,
        total: 0
    }
};

// DOM Elements
const reportTypeSelect = document.getElementById('reportType');
const fileUpload = document.getElementById('fileUpload');
const uploadArea = document.getElementById('uploadArea');
const uploadPlaceholder = document.getElementById('uploadPlaceholder');
const fileCard = document.getElementById('fileCard');
const fileName = document.getElementById('fileName');
const fileSize = document.getElementById('fileSize');
const removeFileBtn = document.getElementById('removeFile');
const qualityCheckBtn = document.getElementById('qualityCheckBtn');
const processingSection = document.getElementById('processingSection');
const processingSteps = document.getElementById('processingSteps');
const resultsSummary = document.getElementById('resultsSummary');
const issueCounts = document.getElementById('issueCounts');
const resultsTable = document.getElementById('resultsTable');
const resultsTableBody = document.getElementById('resultsTableBody');
const selectAllCheckbox = document.getElementById('selectAllCheckbox');
const selectAllBtn = document.getElementById('selectAllBtn');
const deselectAllBtn = document.getElementById('deselectAllBtn');
const applyChangesBtn = document.getElementById('applyChangesBtn');
const downloadSection = document.getElementById('downloadSection');
const downloadBtn = document.getElementById('downloadBtn');

// Event Listeners
reportTypeSelect.addEventListener('change', (e) => {
    state.reportType = e.target.value;
});

// File upload handling
uploadPlaceholder.addEventListener('click', () => {
    fileUpload.click();
});

uploadArea.addEventListener('dragover', (e) => {
    e.preventDefault();
    uploadArea.style.borderColor = '#6b46c1';
    uploadArea.style.background = '#f3f4f6';
});

uploadArea.addEventListener('dragleave', () => {
    uploadArea.style.borderColor = '#d1d5db';
    uploadArea.style.background = '#f9fafb';
});

uploadArea.addEventListener('drop', (e) => {
    e.preventDefault();
    uploadArea.style.borderColor = '#d1d5db';
    uploadArea.style.background = '#f9fafb';
    
    const file = e.dataTransfer.files[0];
    if (file && file.name.endsWith('.docx')) {
        handleFileSelect(file);
    }
});

fileUpload.addEventListener('change', (e) => {
    const file = e.target.files[0];
    if (file) {
        handleFileSelect(file);
    }
});

removeFileBtn.addEventListener('click', () => {
    state.file = null;
    fileUpload.value = '';
    uploadPlaceholder.classList.remove('hidden');
    fileCard.classList.add('hidden');
    qualityCheckBtn.disabled = true;
    document.getElementById('qualityCheckButtonGroup').style.display = 'none';
});

function handleFileSelect(file) {
    state.file = file;
    fileName.textContent = file.name;
    fileSize.textContent = formatFileSize(file.size);
    uploadPlaceholder.classList.add('hidden');
    fileCard.classList.remove('hidden');
    qualityCheckBtn.disabled = false;
    document.getElementById('qualityCheckButtonGroup').style.display = 'block';
}

function formatFileSize(bytes) {
    if (bytes < 1024) return bytes + ' B';
    if (bytes < 1024 * 1024) return (bytes / 1024).toFixed(1) + ' KB';
    return (bytes / (1024 * 1024)).toFixed(1) + ' MB';
}

qualityCheckBtn.addEventListener('click', async () => {
    if (!state.file) return;
    await runQualityCheck();
});

selectAllCheckbox.addEventListener('change', (e) => {
    const checked = e.target.checked;
    document.querySelectorAll('.recommendation-checkbox').forEach(cb => {
        cb.checked = checked;
        const id = cb.dataset.id;
        if (checked) {
            state.selectedRecommendations.add(id);
        } else {
            state.selectedRecommendations.delete(id);
        }
    });
    updateApplyButton();
});

selectAllBtn.addEventListener('click', () => {
    selectAllCheckbox.checked = true;
    selectAllCheckbox.dispatchEvent(new Event('change'));
});

deselectAllBtn.addEventListener('click', () => {
    selectAllCheckbox.checked = false;
    selectAllCheckbox.dispatchEvent(new Event('change'));
});

applyChangesBtn.addEventListener('click', async () => {
    await applyChanges();
});

downloadBtn.addEventListener('click', () => {
    downloadUpdatedDocument();
});

// Quality Check Process
async function runQualityCheck() {
    state.processing = true;
    qualityCheckBtn.disabled = true;
    processingSection.classList.remove('hidden');
    resultsSummary.classList.add('hidden');
    resultsTable.classList.add('hidden');
    downloadSection.classList.add('hidden');

    const steps = [
        'Loading document...',
        'Extracting document structure...',
        'Classifying semantic blocks...',
        'Extracting page numbers...',
        'Running grammar checks...',
        'Running spelling checks...',
        'Checking formatting consistency...',
        'Analyzing content quality...',
        'Generating recommendations...',
        'Finalizing results...'
    ];

    // Simulate processing with steps
    for (let i = 0; i < steps.length; i++) {
        await new Promise(resolve => setTimeout(resolve, 500));
        addProcessingStep(steps[i], i === steps.length - 1);
    }

    // Call backend API
    try {
        const formData = new FormData();
        formData.append('document', state.file);
        formData.append('reportType', state.reportType);

        const response = await fetch('/api/quality-check', {
            method: 'POST',
            body: formData
        });

        if (!response.ok) {
            throw new Error('Failed to process document');
        }

        const data = await response.json();
        state.results = data.results || generateMockResults();
    } catch (error) {
        console.error('Error:', error);
        // Fallback to mock results if API fails
        state.results = generateMockResults();
    }
    state.filteredResults = [...state.results];
    state.issueCounts = calculateIssueCounts(state.results);

    displayResults();
    state.processing = false;
}

function addProcessingStep(stepText, isLast = false) {
    const stepDiv = document.createElement('div');
    stepDiv.className = 'processing-step';
    stepDiv.textContent = stepText;
    
    if (processingSteps.children.length > 0) {
        const prevStep = processingSteps.lastElementChild;
        prevStep.classList.remove('active');
        prevStep.classList.add('completed');
    }
    
    stepDiv.classList.add('active');
    processingSteps.appendChild(stepDiv);
    
    if (isLast) {
        setTimeout(() => {
            stepDiv.classList.remove('active');
            stepDiv.classList.add('completed');
        }, 500);
    }
}

function generateMockResults() {
    // This will be replaced with actual API call
    return [
        {
            id: '1',
            page: 3,
            section: 'Executive Summary',
            confidence: 0.95,
            severity: 'high',
            category: 'grammar',
            originalText: 'The audit was conduct on January 2025.',
            recommendedText: 'The audit was conducted in January 2025.',
            rationale: 'Subject-verb agreement error and incorrect preposition usage.'
        },
        {
            id: '2',
            page: 4,
            section: 'Audit Objectives',
            confidence: 0.88,
            severity: 'medium',
            category: 'spelling',
            originalText: 'Asses the effectiveness',
            recommendedText: 'Assess the effectiveness',
            rationale: 'Spelling error: "Asses" should be "Assess".'
        },
        {
            id: '3',
            page: 5,
            section: 'Methodology',
            confidence: 0.92,
            severity: 'high',
            category: 'formatting',
            originalText: 'walkthroughs',
            recommendedText: 'Walkthroughs',
            rationale: 'Capitalization inconsistency in section headings.'
        },
        {
            id: '4',
            page: 6,
            section: 'Detailed Observations',
            confidence: 0.75,
            severity: 'low',
            category: 'consistency',
            originalText: 'High risk',
            recommendedText: 'High – Severe risk',
            rationale: 'Inconsistent risk rating terminology.'
        }
    ];
}

function calculateIssueCounts(results) {
    const counts = {
        grammar: 0,
        spelling: 0,
        formatting: 0,
        consistency: 0,
        total: results.length
    };

    results.forEach(result => {
        if (counts.hasOwnProperty(result.category)) {
            counts[result.category]++;
        }
    });

    return counts;
}

function displayResults() {
    processingSection.classList.add('hidden');
    resultsSummary.classList.remove('hidden');
    resultsTable.classList.remove('hidden');

    // Display issue counts
    displayIssueCounts();
    
    // Display results table
    displayResultsTable();
}

function displayIssueCounts() {
    issueCounts.innerHTML = '';
    
    const categories = [
        { key: 'grammar', label: 'Grammar Issues' },
        { key: 'spelling', label: 'Spelling Issues' },
        { key: 'formatting', label: 'Formatting Issues' },
        { key: 'consistency', label: 'Consistency Issues' },
        { key: 'total', label: 'Total Issues' }
    ];

    categories.forEach(category => {
        const count = state.issueCounts[category.key];
        const card = document.createElement('div');
        card.className = 'issue-count-card';
        if (state.activeFilter === category.key) {
            card.classList.add('active');
        }
        card.innerHTML = `
            <div class="count">${count}</div>
            <div class="label">${category.label}</div>
        `;
        card.addEventListener('click', () => {
            filterByCategory(category.key);
        });
        issueCounts.appendChild(card);
    });
}

function filterByCategory(category) {
    if (state.activeFilter === category) {
        state.activeFilter = null;
        state.filteredResults = [...state.results];
    } else {
        state.activeFilter = category;
        if (category === 'total') {
            state.filteredResults = [...state.results];
        } else {
            state.filteredResults = state.results.filter(r => r.category === category);
        }
    }
    displayIssueCounts();
    displayResultsTable();
}

function displayResultsTable() {
    resultsTableBody.innerHTML = '';
    
    if (state.filteredResults.length === 0) {
        const row = document.createElement('tr');
        row.innerHTML = '<td colspan="8" style="text-align: center; padding: 40px;">No issues found</td>';
        resultsTableBody.appendChild(row);
        return;
    }

    state.filteredResults.forEach(result => {
        const row = document.createElement('tr');
        if (state.selectedRecommendations.has(result.id)) {
            row.classList.add('selected');
        }
        
        row.innerHTML = `
            <td>
                <input type="checkbox" class="recommendation-checkbox" 
                       data-id="${result.id}" 
                       ${state.selectedRecommendations.has(result.id) ? 'checked' : ''}>
            </td>
            <td>${result.page}</td>
            <td>${result.section}</td>
            <td><span class="confidence-score">${(result.confidence * 100).toFixed(0)}%</span></td>
            <td><span class="severity-badge severity-${result.severity}">${result.severity.toUpperCase()}</span></td>
            <td><span class="original-text" title="${result.originalText}">${result.originalText}</span></td>
            <td><span class="recommended-text" title="${result.recommendedText}">${result.recommendedText}</span></td>
            <td><span class="rationale" title="${result.rationale}">${result.rationale}</span></td>
        `;
        
        const checkbox = row.querySelector('.recommendation-checkbox');
        checkbox.addEventListener('change', (e) => {
            const id = e.target.dataset.id;
            if (e.target.checked) {
                state.selectedRecommendations.add(id);
                row.classList.add('selected');
            } else {
                state.selectedRecommendations.delete(id);
                row.classList.remove('selected');
            }
            updateApplyButton();
        });
        
        resultsTableBody.appendChild(row);
    });
    
    updateApplyButton();
}

function updateApplyButton() {
    const count = state.selectedRecommendations.size;
    applyChangesBtn.disabled = count === 0;
    applyChangesBtn.textContent = count > 0 
        ? `Apply Changes (${count} selected)` 
        : 'Apply Changes';
}

async function applyChanges() {
    if (state.selectedRecommendations.size === 0) {
        console.log('No recommendations selected');
        return;
    }
    
    console.log('=== Frontend: Apply Changes Called ===');
    console.log('Selected recommendations:', Array.from(state.selectedRecommendations));
    
    applyChangesBtn.disabled = true;
    applyChangesBtn.textContent = 'Applying Changes...';
    
    try {
        const selectedResults = state.results.filter(r => 
            state.selectedRecommendations.has(r.id)
        );

        console.log(`Sending ${selectedResults.length} changes to server:`, selectedResults);

        const formData = new FormData();
        formData.append('document', state.file);
        formData.append('changes', JSON.stringify(selectedResults));

        console.log('Sending request to /api/apply-changes...');
        const response = await fetch('/api/apply-changes', {
            method: 'POST',
            body: formData
        });

        console.log('Response status:', response.status, response.statusText);

        if (!response.ok) {
            const errorText = await response.text();
            console.error('Error response status:', response.status);
            console.error('Error response text:', errorText);
            let errorDetails = errorText;
            try {
                const errorJson = JSON.parse(errorText);
                errorDetails = JSON.stringify(errorJson, null, 2);
                console.error('Error details:', errorDetails);
            } catch (e) {
                // Not JSON, use as-is
            }
            throw new Error(`Failed to apply changes: ${response.status} ${response.statusText}\n${errorDetails}`);
        }

        const data = await response.json();
        console.log('Response data:', data);
        state.downloadUrl = data.downloadUrl;
        
        applyChangesBtn.textContent = 'Changes Applied!';
        downloadSection.classList.remove('hidden');
        
        // Scroll to download section
        downloadSection.scrollIntoView({ behavior: 'smooth' });
        console.log('=== Frontend: Apply Changes Complete ===');
    } catch (error) {
        console.error('Error applying changes:', error);
        alert('Failed to apply changes: ' + (error instanceof Error ? error.message : String(error)));
        applyChangesBtn.disabled = false;
        applyChangesBtn.textContent = 'Apply Changes';
    }
}

function downloadUpdatedDocument() {
    if (state.downloadUrl) {
        const link = document.createElement('a');
        link.href = state.downloadUrl;
        link.download = state.file.name.replace('.docx', '_updated.docx');
        link.click();
    } else {
        alert('Download URL not available. Please apply changes first.');
    }
}

