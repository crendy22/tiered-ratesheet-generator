// State
let state = {
    baseFileBuffer: null,  // Store raw file buffer for ExcelJS
    baseFileName: '',
    sheetNames: [],
    tiers: [],
    pricingGrids: {}
};

// Navigation elements
const navStep1 = document.getElementById('nav-step-1');
const navStep2 = document.getElementById('nav-step-2');
const navStep3 = document.getElementById('nav-step-3');
const statusDot = document.getElementById('status-dot');
const statusText = document.getElementById('status-text');

function updateNavigation(activeStep) {
    [navStep1, navStep2, navStep3].forEach((nav, i) => {
        nav.classList.remove('active', 'complete');
        if (i + 1 < activeStep) {
            nav.classList.add('complete');
        } else if (i + 1 === activeStep) {
            nav.classList.add('active');
        }
    });
    
    document.getElementById('step-1-number').classList.toggle('complete', activeStep > 1);
    document.getElementById('step-2-number').classList.toggle('complete', activeStep > 2);
}

function updateStatus(text, active = true) {
    statusText.textContent = text;
    statusDot.classList.toggle('inactive', !active);
}

// DOM Elements
const baseUploadZone = document.getElementById('base-upload-zone');
const baseFileInput = document.getElementById('base-file-input');
const baseUploadSuccess = document.getElementById('base-upload-success');
const baseFilename = document.getElementById('base-filename');
const baseDetails = document.getElementById('base-details');
const baseRemove = document.getElementById('base-remove');

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
    baseUploadZone.addEventListener('click', (e) => {
        if (!baseUploadZone.classList.contains('has-file')) {
            baseFileInput.click();
        }
    });

    baseFileInput.addEventListener('change', (e) => {
        if (e.target.files.length > 0) {
            handleBaseFile(e.target.files[0]);
        }
    });

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

    baseRemove.addEventListener('click', (e) => {
        e.stopPropagation();
        resetBaseFile();
    });
}

async function handleBaseFile(file) {
    if (!file.name.match(/\.xlsx?$/i)) {
        alert('Please upload an Excel file (.xlsx or .xls)');
        return;
    }

    try {
        updateStatus('Loading file...', true);
        
        const buffer = await file.arrayBuffer();
        state.baseFileBuffer = buffer;
        state.baseFileName = file.name;
        
        // Load with ExcelJS to get sheet names and detect grids
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.load(buffer);
        
        state.sheetNames = workbook.worksheets.map(ws => ws.name);
        
        // Detect pricing grids
        detectPricingGrids(workbook);
        
        showBaseFileSuccess(file.name, state.sheetNames.length);
        enableStep2();
        
    } catch (err) {
        console.error('Error reading file:', err);
        alert('Error reading file. Please make sure it\'s a valid Excel file.');
        updateStatus('Error', false);
    }
}

function detectPricingGrids(workbook) {
    state.pricingGrids = {};
    
    workbook.worksheets.forEach(worksheet => {
        const gridInfo = findPricingGrid(worksheet);
        if (gridInfo) {
            state.pricingGrids[worksheet.name] = gridInfo;
        }
    });
}

function findPricingGrid(worksheet) {
    let headerRow = -1;
    let rateCol = -1;
    let pricingStartCol = -1;
    let pricingEndCol = -1;
    
    // Search first 25 rows for header
    for (let r = 1; r <= Math.min(25, worksheet.rowCount); r++) {
        const row = worksheet.getRow(r);
        for (let c = 1; c <= Math.min(15, worksheet.columnCount); c++) {
            const cell = row.getCell(c);
            const value = cell.value;
            if (value && String(value).toLowerCase() === 'rate') {
                // Check next cell for "day"
                const nextCell = row.getCell(c + 1);
                if (nextCell.value && String(nextCell.value).toLowerCase().includes('day')) {
                    headerRow = r;
                    rateCol = c;
                    
                    // Find all pricing columns
                    for (let pc = c + 1; pc <= Math.min(c + 10, worksheet.columnCount); pc++) {
                        const headerCell = row.getCell(pc);
                        if (headerCell.value && String(headerCell.value).toLowerCase().includes('day')) {
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
    
    // Find data rows
    let dataStartRow = headerRow + 1;
    let dataEndRow = dataStartRow;
    
    for (let r = headerRow + 1; r <= worksheet.rowCount; r++) {
        const row = worksheet.getRow(r);
        const rateCell = row.getCell(rateCol);
        const rateValue = rateCell.value;
        
        if (rateValue !== null && rateValue !== undefined && !isNaN(parseFloat(rateValue))) {
            dataEndRow = r;
        } else if (dataEndRow > dataStartRow) {
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
    state.baseFileBuffer = null;
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
    document.getElementById('step-2').classList.remove('disabled');
    updateNavigation(2);
    updateStatus('Ratesheet loaded');
    
    if (state.tiers.length === 0) {
        addTier('Tier 2');
    }
}

function disableStep2() {
    document.getElementById('step-2').classList.add('disabled');
    state.tiers = [];
    tiersContainer.innerHTML = '';
    updateNavigation(1);
    updateStatus('Ready', false);
}

function addTier(name = 'New Tier') {
    const tier = {
        id: Date.now(),
        name: name,
        deltas: {}
    };
    
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
                <button class="btn btn-ghost btn-small" onclick="copyTier(${tier.id})">
                    <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" width="14" height="14">
                        <rect x="9" y="9" width="13" height="13" rx="2" ry="2"></rect>
                        <path d="M5 15H4a2 2 0 0 1-2-2V4a2 2 0 0 1 2-2h9a2 2 0 0 1 2 2v1"></path>
                    </svg>
                </button>
                <button class="btn btn-ghost btn-small" onclick="removeTier(${tier.id})">
                    <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" width="14" height="14">
                        <path d="M18 6L6 18M6 6l12 12"/>
                    </svg>
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
    tierEl.querySelectorAll('.delta-input').forEach(styleDeltaInput);
}

function updateTierName(tierId, name) {
    const tier = state.tiers.find(t => t.id === tierId);
    if (tier) tier.name = name;
}

function updateDelta(tierId, sheetName, value) {
    const tier = state.tiers.find(t => t.id === tierId);
    if (tier) tier.deltas[sheetName] = parseFloat(value) || 0;
    updateGenerateButton();
}

function styleDeltaInput(input) {
    const value = parseFloat(input.value) || 0;
    input.classList.remove('positive', 'negative');
    if (value > 0) input.classList.add('positive');
    else if (value < 0) input.classList.add('negative');
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
    if (tierEl) tierEl.remove();
    
    updateGenerateButton();
}

async function importDeltasFile(file) {
    try {
        const buffer = await file.arrayBuffer();
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.load(buffer);
        
        const worksheet = workbook.worksheets[0];
        const rows = [];
        
        worksheet.eachRow((row, rowNum) => {
            rows.push(row.values.slice(1)); // slice(1) because ExcelJS is 1-indexed
        });
        
        // Detect header row and column mapping
        const headerRow = rows[0] || [];
        const headers = headerRow.map(h => String(h || '').toLowerCase().trim());
        
        let tabCol = -1, tierCol = -1, deltaCol = -1;
        
        headers.forEach((h, i) => {
            if (h === 'tab' || h === 'tab name' || h === 'sheet' || h === 'product') tabCol = i;
            if (h === 'tier' || h === 'tier name' || h === 'tier number') tierCol = i;
            if (h === 'adjustment' || h === 'delta' || h === 'adj' || h === 'change') deltaCol = i;
        });
        
        // Fallback detection
        if (tabCol === -1 || tierCol === -1 || deltaCol === -1) {
            if (rows[1] && state.sheetNames.includes(String(rows[1][0]).trim())) {
                tabCol = 0; tierCol = 1; deltaCol = 2;
            } else if (rows[1] && state.sheetNames.includes(String(rows[1][1]).trim())) {
                tierCol = 0; tabCol = 1; deltaCol = 2;
            }
        }
        
        if (tabCol === -1 || tierCol === -1 || deltaCol === -1) {
            alert('Could not detect column format. Please use columns: Tab, Tier, Adjustment');
            return;
        }
        
        // Group by tier
        const tierMap = new Map();
        
        for (let i = 1; i < rows.length; i++) {
            const row = rows[i];
            if (!row || row.length === 0) continue;
            
            const tabName = String(row[tabCol] || '').trim();
            const tierValue = row[tierCol];
            const delta = parseFloat(row[deltaCol]);
            
            if (!tabName || tierValue === undefined || isNaN(delta)) continue;
            
            let tierName = (typeof tierValue === 'number' || !isNaN(Number(tierValue))) 
                ? `Tier ${tierValue}` 
                : String(tierValue).trim();
            
            if (!tierMap.has(tierName)) tierMap.set(tierName, {});
            tierMap.get(tierName)[tabName] = delta;
        }
        
        if (tierMap.size === 0) {
            alert('No valid tier data found in file.');
            return;
        }
        
        // Clear and rebuild tiers
        state.tiers = [];
        tiersContainer.innerHTML = '';
        
        const sortedTiers = Array.from(tierMap.entries()).sort((a, b) => {
            const numA = parseInt(a[0].replace(/\D/g, '')) || 0;
            const numB = parseInt(b[0].replace(/\D/g, '')) || 0;
            return numA - numB;
        });
        
        sortedTiers.forEach(([tierName, deltas]) => {
            const tier = {
                id: Date.now() + Math.random(),
                name: tierName,
                deltas: {}
            };
            
            state.sheetNames.forEach(sheetName => {
                tier.deltas[sheetName] = deltas[sheetName] || 0;
            });
            
            state.tiers.push(tier);
            renderTier(tier);
        });
        
        updateGenerateButton();
        
    } catch (err) {
        console.error('Error importing deltas:', err);
        alert('Error importing deltas file. Please check the format.');
    }
}

// ============================================
// GENERATION
// ============================================

function setupGenerateControls() {
    generateBtn.addEventListener('click', generateRatesheets);
}

function updateGenerateButton() {
    generateBtn.disabled = state.tiers.length === 0;
    
    if (state.tiers.length > 0) {
        document.getElementById('step-3').classList.remove('disabled');
        updateNavigation(3);
        updateStatus(`${state.tiers.length} tier(s) configured`);
    } else {
        document.getElementById('step-3').classList.add('disabled');
    }
}

function disableStep3() {
    document.getElementById('step-3').classList.add('disabled');
    outputPreview.hidden = true;
}

async function generateRatesheets() {
    if (!state.baseFileBuffer || state.tiers.length === 0) return;
    
    const includeBase = includeBaseCheckbox.checked;
    
    updateStatus('Generating files...', true);
    generateBtn.disabled = true;
    
    try {
        const generatedFiles = [];
        
        // Generate file for each tier
        for (const tier of state.tiers) {
            const buffer = await applyDeltasToWorkbook(tier);
            const filename = generateFilename(state.baseFileName, tier.name);
            generatedFiles.push({ buffer, filename, tierName: tier.name });
        }
        
        // Include base if requested
        if (includeBase) {
            generatedFiles.unshift({
                buffer: state.baseFileBuffer,
                filename: state.baseFileName,
                tierName: 'Base'
            });
        }
        
        showOutputPreview(generatedFiles);
        updateStatus(`${generatedFiles.length} file(s) generated`);
        
        // Store generated files for re-download
        state.generatedFiles = generatedFiles;
        
        // Auto-download all files
        for (let i = 0; i < generatedFiles.length; i++) {
            await new Promise(resolve => setTimeout(resolve, 300));
            downloadBuffer(generatedFiles[i].buffer, generatedFiles[i].filename);
        }
        
    } catch (err) {
        console.error('Error generating files:', err);
        alert('Error generating files: ' + err.message);
        updateStatus('Error', false);
    } finally {
        generateBtn.disabled = false;
    }
}

async function applyDeltasToWorkbook(tier) {
    // Load a fresh copy of the workbook from the original buffer
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(state.baseFileBuffer);
    
    // Apply deltas to each worksheet
    workbook.worksheets.forEach(worksheet => {
        const delta = tier.deltas[worksheet.name] || 0;
        const gridInfo = state.pricingGrids[worksheet.name];
        
        if (gridInfo && delta !== 0) {
            // Apply delta to pricing cells
            for (let r = gridInfo.dataStartRow; r <= gridInfo.dataEndRow; r++) {
                const row = worksheet.getRow(r);
                for (let c = gridInfo.pricingStartCol; c <= gridInfo.pricingEndCol; c++) {
                    const cell = row.getCell(c);
                    const value = cell.value;
                    
                    // Only modify numeric values (not "na" or empty)
                    if (typeof value === 'number') {
                        cell.value = Math.round((value + delta) * 10000) / 10000;
                    }
                }
            }
        }
    });
    
    // Write to buffer (preserves all formatting)
    const buffer = await workbook.xlsx.writeBuffer();
    return buffer;
}

function generateFilename(baseFilename, tierName) {
    const baseName = baseFilename.replace(/\.xlsx?$/i, '');
    const sanitizedTier = tierName.replace(/[^a-zA-Z0-9_-]/g, '_');
    return `${baseName}_${sanitizedTier}.xlsx`;
}

function downloadBuffer(buffer, filename) {
    const blob = new Blob([buffer], { 
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' 
    });
    
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
            <button class="btn btn-small btn-secondary" onclick="redownloadFile('${file.tierName}')">
                Download
            </button>
        </div>
    `).join('');
}

// Global functions
window.redownloadFile = function(tierName) {
    const file = state.generatedFiles?.find(f => f.tierName === tierName);
    if (file) {
        downloadBuffer(file.buffer, file.filename);
    }
};

window.updateTierName = updateTierName;
window.updateDelta = updateDelta;
window.styleDeltaInput = styleDeltaInput;
window.copyTier = copyTier;
window.removeTier = removeTier;
