// State
let state = {
    baseWorkbook: null,
    baseFileName: '',
    sheetNames: [],
    tiers: [],
    pricingGrids: {} // Store detected pricing grid locations per sheet
};

// DOM Elements
const baseUploadZone = document.getElementById('base-upload-zone');
const baseFileInput = document.getElementById('base-file-input');
const baseUploadSuccess = document.getElementById('base-upload-success');
const baseFilename = document.getElementById('base-filename');
const baseDetails = document.getElementById('base-details');
const baseRemove = document.getElementById('base-remove');

const step2 = document.getElementById('step-2');
const step3 = document.getElementById('step-3');

const addTierBtn = document.getElementById('add-tier-btn');
const importDeltasBtn = document.getElementById('import-deltas-btn');
const deltasFileInput = document.getElementById('deltas-file-input');
const tiersContainer = document.getElementById('tiers-container');

const generateBtn = document.getElementById('generate-btn');
const separateFilesCheckbox = document.getElementById('separate-files');
const includeBaseCheckbox = document.getElementById('include-base');
const outputPreview = document.getElementById('output-preview');
const fileList = document.getElementById('file-list');

// Initialize
document.addEventListener('DOMContentLoaded', init);

function init() {
    setupUploadZone();
    setupTierControls();
    setupGenerateControls();
}

// ============================================
// FILE UPLOAD HANDLING
// ============================================

function setupUploadZone() {
    // Click to upload
    baseUploadZone.addEventListener('click', (e) => {
        if (!baseUploadZone.classList.contains('has-file')) {
            baseFileInput.click();
        }
    });

    // File input change
    baseFileInput.addEventListener('change', (e) => {
        if (e.target.files.length > 0) {
            handleBaseFile(e.target.files[0]);
        }
    });

    // Drag and drop
    baseUploadZone.addEventListener('dragover', (e) => {
        e.preventDefault();
        baseUploadZone.classList.add('dragover');
    });

    baseUploadZone.addEventListener('dragleave', () => {
        baseUploadZone.classList.remove('dragover');
    });

    baseUploadZone.addEventListener('drop', (e) => {
        e.preventDefault();
        baseUploadZone.classList.remove('dragover');
        if (e.dataTransfer.files.length > 0) {
            handleBaseFile(e.dataTransfer.files[0]);
        }
    });

    // Remove file
    baseRemove.addEventListener('click', (e) => {
        e.stopPropagation();
        resetBaseFile();
    });
}

function handleBaseFile(file) {
    if (!file.name.match(/\.xlsx?$/i)) {
        alert('Please upload an Excel file (.xlsx or .xls)');
        return;
    }

    const reader = new FileReader();
    reader.onload = (e) => {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            
            state.baseWorkbook = workbook;
            state.baseFileName = file.name;
            state.sheetNames = workbook.SheetNames;
            
            // Detect pricing grids in each sheet
            detectPricingGrids(workbook);
            
            // Update UI
            showBaseFileSuccess(file.name, state.sheetNames.length);
            enableStep2();
            
        } catch (err) {
            console.error('Error reading file:', err);
            alert('Error reading file. Please make sure it\'s a valid Excel file.');
        }
    };
    reader.readAsArrayBuffer(file);
}

function detectPricingGrids(workbook) {
    state.pricingGrids = {};
    
    workbook.SheetNames.forEach(sheetName => {
        const sheet = workbook.Sheets[sheetName];
        const gridInfo = findPricingGrid(sheet);
        if (gridInfo) {
            state.pricingGrids[sheetName] = gridInfo;
        }
    });
}

function findPricingGrid(sheet) {
    // Look for the pricing grid pattern:
    // Row with "Rate" and lock period headers (7 Day, 15 Day, etc.)
    // Followed by rows with rate values and pricing
    
    const range = XLSX.utils.decode_range(sheet['!ref'] || 'A1');
    let headerRow = -1;
    let rateCol = -1;
    let pricingStartCol = -1;
    let pricingEndCol = -1;
    
    // Find header row with "Rate" and lock periods
    for (let r = 0; r <= Math.min(range.e.r, 20); r++) {
        for (let c = 0; c <= range.e.c; c++) {
            const cellRef = XLSX.utils.encode_cell({ r, c });
            const cell = sheet[cellRef];
            if (cell && String(cell.v).toLowerCase() === 'rate') {
                // Check if next cells have lock period headers
                const nextCell = sheet[XLSX.utils.encode_cell({ r, c: c + 1 })];
                if (nextCell && String(nextCell.v).toLowerCase().includes('day')) {
                    headerRow = r;
                    rateCol = c;
                    
                    // Find pricing columns (those with "Day" in header)
                    for (let pc = c + 1; pc <= range.e.c; pc++) {
                        const headerCell = sheet[XLSX.utils.encode_cell({ r, c: pc })];
                        if (headerCell && String(headerCell.v).toLowerCase().includes('day')) {
                            if (pricingStartCol === -1) pricingStartCol = pc;
                            pricingEndCol = pc;
                        }
                    }
                    break;
                }
            }
        }
        if (headerRow !== -1) break;
    }
    
    if (headerRow === -1) return null;
    
    // Find data rows (rows with numeric rate values)
    let dataStartRow = headerRow + 1;
    let dataEndRow = dataStartRow;
    
    for (let r = headerRow + 1; r <= range.e.r; r++) {
        const rateCell = sheet[XLSX.utils.encode_cell({ r, c: rateCol })];
        if (rateCell && !isNaN(parseFloat(rateCell.v))) {
            dataEndRow = r;
        } else if (dataEndRow > dataStartRow) {
            // Hit a non-numeric row after finding data
            break;
        }
    }
    
    return {
        headerRow,
        rateCol,
        pricingStartCol,
        pricingEndCol,
        dataStartRow,
        dataEndRow
    };
}

function showBaseFileSuccess(filename, sheetCount) {
    baseUploadZone.classList.add('has-file');
    baseUploadZone.querySelector('.upload-content').hidden = true;
    baseUploadSuccess.hidden = false;
    baseFilename.textContent = filename;
    baseDetails.textContent = `${sheetCount} product tabs detected`;
}

function resetBaseFile() {
    state.baseWorkbook = null;
    state.baseFileName = '';
    state.sheetNames = [];
    state.pricingGrids = {};
    
    baseFileInput.value = '';
    baseUploadZone.classList.remove('has-file');
    baseUploadZone.querySelector('.upload-content').hidden = false;
    baseUploadSuccess.hidden = true;
    
    disableStep2();
    disableStep3();
}

// ============================================
// TIER MANAGEMENT
// ============================================

function setupTierControls() {
    addTierBtn.addEventListener('click', () => {
        addTier(`Tier ${state.tiers.length + 2}`);
    });

    importDeltasBtn.addEventListener('click', () => {
        deltasFileInput.click();
    });

    deltasFileInput.addEventListener('change', (e) => {
        if (e.target.files.length > 0) {
            importDeltasFile(e.target.files[0]);
        }
    });
}

function enableStep2() {
    step2.classList.remove('disabled');
    
    // Add initial tier if none exist
    if (state.tiers.length === 0) {
        addTier('Tier 2');
    }
}

function disableStep2() {
    step2.classList.add('disabled');
    state.tiers = [];
    tiersContainer.innerHTML = '';
}

function addTier(name = 'New Tier') {
    const tier = {
        id: Date.now(),
        name: name,
        deltas: {}
    };
    
    // Initialize all deltas to 0
    state.sheetNames.forEach(sheet => {
        tier.deltas[sheet] = 0;
    });
    
    state.tiers.push(tier);
    renderTier(tier);
    updateGenerateButton();
}

function renderTier(tier) {
    const tierEl = document.createElement('div');
    tierEl.className = 'tier-card';
    tierEl.dataset.tierId = tier.id;
    
    tierEl.innerHTML = `
        <div class="tier-header">
            <input type="text" class="tier-name-input" value="${tier.name}" 
                   onchange="updateTierName(${tier.id}, this.value)">
            <div class="tier-actions">
                <button class="btn btn-small btn-outline" onclick="copyTier(${tier.id})">
                    Duplicate
                </button>
                <button class="btn btn-small btn-danger" onclick="removeTier(${tier.id})">
                    Remove
                </button>
            </div>
        </div>
        <div class="tier-body">
            <div class="delta-grid">
                ${state.sheetNames.map(sheet => `
                    <div class="delta-item">
                        <label title="${sheet}">${sheet}</label>
                        <input type="number" 
                               class="delta-input" 
                               step="0.125" 
                               value="${tier.deltas[sheet] || 0}"
                               onchange="updateDelta(${tier.id}, '${sheet}', this.value)"
                               oninput="styleDeltaInput(this)">
                    </div>
                `).join('')}
            </div>
        </div>
    `;
    
    tiersContainer.appendChild(tierEl);
    
    // Apply initial styling to inputs
    tierEl.querySelectorAll('.delta-input').forEach(styleDeltaInput);
}

function updateTierName(tierId, name) {
    const tier = state.tiers.find(t => t.id === tierId);
    if (tier) {
        tier.name = name;
    }
}

function updateDelta(tierId, sheetName, value) {
    const tier = state.tiers.find(t => t.id === tierId);
    if (tier) {
        tier.deltas[sheetName] = parseFloat(value) || 0;
    }
    updateGenerateButton();
}

function styleDeltaInput(input) {
    const value = parseFloat(input.value) || 0;
    input.classList.remove('positive', 'negative');
    if (value > 0) {
        input.classList.add('positive');
    } else if (value < 0) {
        input.classList.add('negative');
    }
}

function copyTier(tierId) {
    const sourceTier = state.tiers.find(t => t.id === tierId);
    if (!sourceTier) return;
    
    const newTier = {
        id: Date.now(),
        name: sourceTier.name + ' (Copy)',
        deltas: { ...sourceTier.deltas }
    };
    
    state.tiers.push(newTier);
    renderTier(newTier);
    updateGenerateButton();
}

function removeTier(tierId) {
    const index = state.tiers.findIndex(t => t.id === tierId);
    if (index === -1) return;
    
    state.tiers.splice(index, 1);
    
    const tierEl = document.querySelector(`.tier-card[data-tier-id="${tierId}"]`);
    if (tierEl) {
        tierEl.remove();
    }
    
    updateGenerateButton();
}

function importDeltasFile(file) {
    const reader = new FileReader();
    reader.onload = (e) => {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            const sheet = workbook.Sheets[workbook.SheetNames[0]];
            const rows = XLSX.utils.sheet_to_json(sheet, { header: 1 });
            
            // Expected format: Tier Name | Tab Name | Delta
            // Skip header row if present
            const startRow = (rows[0] && typeof rows[0][0] === 'string' && 
                             rows[0][0].toLowerCase().includes('tier')) ? 1 : 0;
            
            // Group by tier
            const tierMap = new Map();
            
            for (let i = startRow; i < rows.length; i++) {
                const [tierName, tabName, delta] = rows[i];
                if (!tierName || !tabName || delta === undefined) continue;
                
                const normalizedTier = String(tierName).trim();
                const normalizedTab = String(tabName).trim();
                const deltaValue = parseFloat(delta) || 0;
                
                if (!tierMap.has(normalizedTier)) {
                    tierMap.set(normalizedTier, {});
                }
                tierMap.get(normalizedTier)[normalizedTab] = deltaValue;
            }
            
            // Clear existing tiers and add imported ones
            state.tiers = [];
            tiersContainer.innerHTML = '';
            
            tierMap.forEach((deltas, tierName) => {
                const tier = {
                    id: Date.now() + Math.random(),
                    name: tierName,
                    deltas: {}
                };
                
                // Initialize all sheets to 0, then apply imported deltas
                state.sheetNames.forEach(sheet => {
                    tier.deltas[sheet] = deltas[sheet] || 0;
                });
                
                state.tiers.push(tier);
                renderTier(tier);
            });
            
            updateGenerateButton();
            alert(`Imported ${tierMap.size} tier(s) from file`);
            
        } catch (err) {
            console.error('Error importing deltas:', err);
            alert('Error importing deltas file. Please check the format.');
        }
    };
    reader.readAsArrayBuffer(file);
}

// ============================================
// GENERATION
// ============================================

function setupGenerateControls() {
    generateBtn.addEventListener('click', generateRatesheets);
}

function updateGenerateButton() {
    const hasDeltas = state.tiers.some(tier => 
        Object.values(tier.deltas).some(d => d !== 0)
    );
    
    generateBtn.disabled = state.tiers.length === 0;
    
    if (state.tiers.length > 0) {
        step3.classList.remove('disabled');
    } else {
        step3.classList.add('disabled');
    }
}

function disableStep3() {
    step3.classList.add('disabled');
    outputPreview.hidden = true;
}

function generateRatesheets() {
    if (!state.baseWorkbook || state.tiers.length === 0) return;
    
    const separateFiles = separateFilesCheckbox.checked;
    const includeBase = includeBaseCheckbox.checked;
    
    const generatedFiles = [];
    
    if (separateFiles) {
        // Generate separate file for each tier
        state.tiers.forEach(tier => {
            const newWorkbook = applyDeltasToWorkbook(state.baseWorkbook, tier);
            const filename = generateFilename(state.baseFileName, tier.name);
            generatedFiles.push({ workbook: newWorkbook, filename, tierName: tier.name });
        });
        
        if (includeBase) {
            generatedFiles.unshift({
                workbook: state.baseWorkbook,
                filename: state.baseFileName,
                tierName: 'Base'
            });
        }
    } else {
        // All tiers in one file (different approach - add tier name to sheet names)
        // For simplicity, we'll still do separate files
        state.tiers.forEach(tier => {
            const newWorkbook = applyDeltasToWorkbook(state.baseWorkbook, tier);
            const filename = generateFilename(state.baseFileName, tier.name);
            generatedFiles.push({ workbook: newWorkbook, filename, tierName: tier.name });
        });
    }
    
    // Show preview and trigger downloads
    showOutputPreview(generatedFiles);
    
    // Auto-download all files
    generatedFiles.forEach((file, index) => {
        setTimeout(() => {
            downloadWorkbook(file.workbook, file.filename);
        }, index * 500); // Stagger downloads
    });
}

function applyDeltasToWorkbook(sourceWorkbook, tier) {
    // Create a deep copy of the workbook
    const newWorkbook = XLSX.utils.book_new();
    
    sourceWorkbook.SheetNames.forEach(sheetName => {
        const sourceSheet = sourceWorkbook.Sheets[sheetName];
        const delta = tier.deltas[sheetName] || 0;
        const gridInfo = state.pricingGrids[sheetName];
        
        // Clone the sheet
        const newSheet = JSON.parse(JSON.stringify(sourceSheet));
        
        // Apply delta to pricing cells if grid was detected and delta is non-zero
        if (gridInfo && delta !== 0) {
            for (let r = gridInfo.dataStartRow; r <= gridInfo.dataEndRow; r++) {
                for (let c = gridInfo.pricingStartCol; c <= gridInfo.pricingEndCol; c++) {
                    const cellRef = XLSX.utils.encode_cell({ r, c });
                    const cell = newSheet[cellRef];
                    
                    if (cell && typeof cell.v === 'number') {
                        cell.v = Math.round((cell.v + delta) * 10000) / 10000; // Round to 4 decimals
                        // Update formatted value if present
                        if (cell.w) {
                            delete cell.w; // Let XLSX regenerate formatted value
                        }
                    }
                }
            }
        }
        
        XLSX.utils.book_append_sheet(newWorkbook, newSheet, sheetName);
    });
    
    return newWorkbook;
}

function generateFilename(baseFilename, tierName) {
    const baseName = baseFilename.replace(/\.xlsx?$/i, '');
    const sanitizedTier = tierName.replace(/[^a-zA-Z0-9_-]/g, '_');
    return `${baseName}_${sanitizedTier}.xlsx`;
}

function downloadWorkbook(workbook, filename) {
    const wbout = XLSX.write(workbook, { bookType: 'xlsx', type: 'array' });
    const blob = new Blob([wbout], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
}

function showOutputPreview(files) {
    outputPreview.hidden = false;
    
    fileList.innerHTML = files.map(file => `
        <div class="file-item">
            <div class="file-item-info">
                <div class="file-item-icon">
                    <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                        <path d="M9 12l2 2 4-4m6 2a9 9 0 11-18 0 9 9 0 0118 0z"/>
                    </svg>
                </div>
                <span class="file-item-name">${file.filename}</span>
            </div>
            <button class="btn btn-small btn-outline" onclick="redownloadFile('${file.tierName}')">
                Download Again
            </button>
        </div>
    `).join('');
}

// Make redownload available globally
window.redownloadFile = function(tierName) {
    let workbook, filename;
    
    if (tierName === 'Base') {
        workbook = state.baseWorkbook;
        filename = state.baseFileName;
    } else {
        const tier = state.tiers.find(t => t.name === tierName);
        if (!tier) return;
        workbook = applyDeltasToWorkbook(state.baseWorkbook, tier);
        filename = generateFilename(state.baseFileName, tier.name);
    }
    
    downloadWorkbook(workbook, filename);
};

// Make tier functions globally available
window.updateTierName = updateTierName;
window.updateDelta = updateDelta;
window.styleDeltaInput = styleDeltaInput;
window.copyTier = copyTier;
window.removeTier = removeTier;
