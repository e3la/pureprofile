// --- CONFIGURATION ---
const CONSTANTS = {
    ContractType: 'openended',
    Primary: 'no',
    StaffType: 'academic',
    EndDate: '',
    WebEn: '',
    Phone: '',
    Mobile: '',
    Fax: '',
    DefaultPhoto: 'UMSLlouie2.png',
    Gender: 'UNKNOWN',
    FTE: ''
};

// --- GLOBAL STATE ---
const APP_STATE = {
    workbook: null,
    zipObject: null,
    hasExcel: false,
    hasZip: false,

    rawPersons: [],
    rawStaff: [],
    sheetHeaders: {
        person: [],
        staff: []
    },

    // Org Data
    allOrgs: [], // Stores full object with visibility

    // Job Title Data
    topJobTitles: [],  
    allJobTitles: [], 

    employmentOptions: ['faculty', 'staff', 'emeritus', 'other'],

    // User Settings
    settings: {
        showUUIDs: false,       // Default hidden
        includeRestricted: false // Default Public only
    },

    stagedEdits: new Map(), 
    stagedNewPersons: [],   
    stagedPhotos: new Map(), 
    
    currentMode: 'add', 
    currentEditBlob: null, 
    currentOriginalPhotoName: null
};

// --- INITIALIZATION ---
document.addEventListener('DOMContentLoaded', () => {
    setupDragAndDrop();
});

// --- HELPER: ROBUST DOWNLOADER ---
function triggerDownload(blob, filename) {
    if (typeof saveAs !== 'undefined') {
        saveAs(blob, filename); 
    } else if (window.saveAs) {
        window.saveAs(blob, filename); 
    } else {
        const url = window.URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.style.display = 'none';
        a.href = url;
        a.download = filename;
        document.body.appendChild(a);
        a.click();
        window.URL.revokeObjectURL(url);
        document.body.removeChild(a);
    }
}

// --- UTILITIES (Shared) ---
function formatDate(incoming) {
    if(!incoming) return '';
    if(typeof incoming === 'string') return incoming;
    let d = null;
    if(typeof incoming === 'number') {
        d = new Date(Math.round((incoming - 25569) * 864e5));
    } else {
        d = new Date(incoming);
    }
    if(!d || isNaN(d.getTime())) return incoming; 
    const dd = String(d.getDate()).padStart(2, '0');
    const mm = String(d.getMonth() + 1).padStart(2, '0');
    const yyyy = d.getFullYear();
    return `${dd}-${mm}-${yyyy}`;
}

function getTodayStr() {
    const d = new Date();
    const dd = String(d.getDate()).padStart(2, '0');
    const mm = String(d.getMonth() + 1).padStart(2, '0');
    const yyyy = d.getFullYear();
    return `${dd}-${mm}-${yyyy}`;
}

function processInputDate(input) {
    if (!input) return getTodayStr();
    const clean = input.trim();
    const match = clean.match(/^(\d{1,2})[-/](\d{1,2})[-/](\d{4})$/);
    if (match) {
        let p1 = parseInt(match[1]); 
        let p2 = parseInt(match[2]); 
        const year = match[3];
        if (p2 > 12 && p1 <= 12) { const temp = p1; p1 = p2; p2 = temp; }
        return `${String(p1).padStart(2, '0')}-${String(p2).padStart(2, '0')}-${year}`;
    }
    return clean; 
}

// --- NAVIGATION ---

function goHome() {
    const hasChanges = (APP_STATE.stagedNewPersons.length > 0 || APP_STATE.stagedEdits.size > 0);
    if (hasChanges) {
        const discard = confirm("You have unsaved changes! \n\nClick OK to DISCARD data and Reset.\nClick Cancel to stay here and save.");
        if (!discard) return; 
    }
    window.location.reload();
}

function startPersonManager() {
    document.getElementById('landingPage').classList.add('hidden');
    document.getElementById('personManagerPage').classList.remove('hidden');
    if(typeof setEditorMode === 'function') setEditorMode('add');
}

function clearStaging() {
    APP_STATE.stagedNewPersons = [];
    APP_STATE.stagedEdits.clear();
    APP_STATE.stagedPhotos.clear();
    
    if(typeof renderUnifiedStaging === 'function') renderUnifiedStaging();
}

// --- FILE LOADING ---
function setupDragAndDrop() {
    const dropExcel = document.getElementById('dropExcel');
    const inputExcel = document.getElementById('globalExcelInput');
    if(dropExcel && inputExcel) {
        dropExcel.addEventListener('click', () => inputExcel.click());
        dropExcel.addEventListener('dragover', (e) => { e.preventDefault(); dropExcel.classList.add('dragover'); });
        dropExcel.addEventListener('dragleave', () => dropExcel.classList.remove('dragover'));
        dropExcel.addEventListener('drop', (e) => {
            e.preventDefault(); dropExcel.classList.remove('dragover');
            if(e.dataTransfer.files.length) handleExcelLoad(e.dataTransfer.files[0]);
        });
        inputExcel.onchange = (e) => { if(e.target.files.length) handleExcelLoad(e.target.files[0]); };
    }

    const dropZip = document.getElementById('dropZip');
    const inputZip = document.getElementById('globalZipInput');
    if(dropZip && inputZip) {
        dropZip.addEventListener('click', () => inputZip.click());
        dropZip.addEventListener('dragover', (e) => { e.preventDefault(); dropZip.classList.add('dragover'); });
        dropZip.addEventListener('dragleave', () => dropZip.classList.remove('dragover'));
        dropZip.addEventListener('drop', (e) => {
            e.preventDefault(); dropZip.classList.remove('dragover');
            if(e.dataTransfer.files.length) handleZipLoad(e.dataTransfer.files[0]);
        });
        inputZip.onchange = (e) => { if(e.target.files.length) handleZipLoad(e.target.files[0]); };
    }
}

async function handleExcelLoad(file) {
    const el = document.getElementById('statusExcel');
    if(el) el.innerHTML = 'â³ Reading...';
    try {
        const buf = await file.arrayBuffer();
        APP_STATE.workbook = XLSX.read(buf, {type: 'array'});
        processWorkbookData();
        APP_STATE.hasExcel = true;
        if(el) { el.innerHTML = 'âœ… Loaded'; el.className = 'badge bg-success'; }
        
        const dropZone = document.getElementById('dropExcel');
        const content = document.getElementById('excelContent');
        if(dropZone) dropZone.classList.add('drop-success');
        if(content) content.innerHTML = `<div style="font-size: 2rem;">âœ…</div><p><strong>Excel Loaded</strong><br><span class="small">${file.name}</span></p>`;
        
        checkUnlockStatus();
    } catch(err) {
        if(el) { el.innerHTML = 'âŒ Error'; el.className = 'badge bg-danger'; }
        alert("Excel Error: " + err.message);
    }
}

async function handleZipLoad(file) {
    const el = document.getElementById('statusZip');
    if(el) el.innerHTML = 'â³ Reading...';
    try {
        APP_STATE.zipObject = await JSZip.loadAsync(file);
        APP_STATE.hasZip = true;
        if(el) { el.innerHTML = 'âœ… Loaded'; el.className = 'badge bg-success'; }

        const dropZone = document.getElementById('dropZip');
        const content = document.getElementById('zipContent');
        if(dropZone) dropZone.classList.add('drop-success');
        if(content) content.innerHTML = `<div style="font-size: 2rem;">âœ…</div><p><strong>Zip Loaded</strong><br><span class="small">${file.name}</span></p>`;

        checkUnlockStatus();
    } catch(err) {
        if(el) { el.innerHTML = 'âŒ Error'; el.className = 'badge bg-danger'; }
        alert("Zip Error: " + err.message);
    }
}

function processWorkbookData() {
    const wb = APP_STATE.workbook;
    // 1. Orgs
    const orgSheetName = wb.SheetNames.find(n => n.toLowerCase().includes('organis') && !n.toLowerCase().includes('hierarch'));
    if(orgSheetName) {
        const rawRows = XLSX.utils.sheet_to_json(wb.Sheets[orgSheetName]);
        // Filter out bad rows, sort by name
        APP_STATE.allOrgs = rawRows.filter(r => r.OrganisationID && !r.OrganisationID.toString().includes(' '))
            .map(o => ({
                id: o.OrganisationID,
                name: o.Name_en || o.Name || "Unknown",
                visibility: o.Visibility || 'Public' // Capture visibility for filtering
            })).sort((a,b) => a.name.localeCompare(b.name));
    }
    // 2. Persons
    const persSheetName = wb.SheetNames.find(n => n.toLowerCase().includes('person'));
    if(persSheetName) {
        const worksheet = wb.Sheets[persSheetName];
        APP_STATE.sheetHeaders.person = XLSX.utils.sheet_to_json(worksheet, {header: 1})[0];
        APP_STATE.rawPersons = XLSX.utils.sheet_to_json(worksheet);
    }
    // 3. Staff & Jobs
    const staffSheetName = wb.SheetNames.find(n => n.toLowerCase().includes('staff'));
    if(staffSheetName) {
        const worksheet = wb.Sheets[staffSheetName];
        APP_STATE.sheetHeaders.staff = XLSX.utils.sheet_to_json(worksheet, {header: 1})[0];
        APP_STATE.rawStaff = XLSX.utils.sheet_to_json(worksheet);
        
        const counts = {};
        APP_STATE.rawStaff.forEach(s => {
            const j = s.JobDescription ? s.JobDescription.trim() : "";
            if(j) counts[j] = (counts[j] || 0) + 1;
        });
        const sortedJobs = Object.keys(counts).sort((a,b) => counts[b] - counts[a]);
        APP_STATE.topJobTitles = sortedJobs.slice(0, 10);
        APP_STATE.allJobTitles = Object.keys(counts).sort();
    }
    // 4. Options
    const classSheetName = wb.SheetNames.find(n => n.toLowerCase().includes('classif') || n.toLowerCase().includes('dictionar'));
    if (classSheetName) {
        try {
            const rows = XLSX.utils.sheet_to_json(wb.Sheets[classSheetName], {header: 1});
            let empRowIndex = -1;
            for(let i=0; i<rows.length; i++) {
                if(rows[i] && rows[i].some(cell => String(cell).includes('/dk/atira/pure/person/employmenttypes'))) {
                    empRowIndex = i; break;
                }
            }
            if(empRowIndex > -1 && rows[empRowIndex + 1]) {
                const cleanOptions = rows[empRowIndex + 1].filter(opt => opt && String(opt).toLowerCase() !== 'uri');
                if(cleanOptions.length > 0) APP_STATE.employmentOptions = cleanOptions;
            }
        } catch (e) { console.warn("Failed to parse classification sheet", e); }
    }
}

function checkUnlockStatus() {
    if(APP_STATE.hasExcel && APP_STATE.hasZip) {
        const personCount = APP_STATE.rawPersons.length;
        const photoCount = Object.keys(APP_STATE.zipObject.files).filter(f => !APP_STATE.zipObject.files[f].dir).length;
        const statsEl = document.getElementById('fileStats');
        if(statsEl) {
            statsEl.classList.remove('hidden');
            statsEl.innerHTML = `âœ… Ready! Found <strong>${personCount}</strong> Profiles and <strong>${photoCount}</strong> Photos.`;
        }
        setTimeout(() => { startPersonManager(); }, 1000);
    }
}

function generateChangeLog() {
    let log = "PURE PERSON MANAGER - CHANGE LOG\n";
    log += "Generated: " + new Date().toLocaleString() + "\n";
    log += "========================================\n\n";

    if (APP_STATE.stagedNewPersons.length > 0) {
        log += "--- NEW PERSONS CREATED (" + APP_STATE.stagedNewPersons.length + ") ---\n";
        APP_STATE.stagedNewPersons.forEach(p => {
            log += `ID: ${p.person.PersonID} | Name: ${p.person.Firstname} ${p.person.Lastname}\n`;
            log += p.photoBlob ? `   - Photo: ${p.photoName}\n` : `   - Photo: (Default)\n`;
            log += `   - Affiliations (${p.staff.length}):\n`;
            p.staff.forEach(s => log += `      -> [${s.OrganisationID}] : ${s.JobDescription}\n`);
            log += "\n";
        });
    }

    if (APP_STATE.stagedEdits.size > 0) {
        log += "--- EDITED PERSONS (" + APP_STATE.stagedEdits.size + ") ---\n";
        APP_STATE.stagedEdits.forEach((val, key) => {
            log += `ID: ${key} | Name: ${val.person.Firstname} ${val.person.Lastname}\n`;
            if (APP_STATE.stagedPhotos.has(key)) {
                const pData = APP_STATE.stagedPhotos.get(key);
                log += `   - Photo CHANGED: ${pData.originalFilename || '(None)'} --> ${pData.filename}\n`;
            }
            log += `   - Final Affiliations (${val.staff.length}):\n`;
            val.staff.forEach(s => log += `      -> [${s.OrganisationID}] : ${s.JobDescription}\n`);
            log += "\n";
        });
    }

    if (APP_STATE.stagedNewPersons.length === 0 && APP_STATE.stagedEdits.size === 0) {
        log += "No changes detected.\n";
    }
    return log;
}

async function downloadAllData() {
    // 1. Clone the original data to avoid modifying the live state
    let outPersons = JSON.parse(JSON.stringify(APP_STATE.rawPersons));
    let outStaff = JSON.parse(JSON.stringify(APP_STATE.rawStaff));

    // Arrays to hold the data that will be appended to the bottom
    const newPersonRows = [];
    const newStaffRows = [];

    // --- HELPER: Create a blank object based on headers to ensure an empty row is rendered ---
    const createBlankRow = (headers) => {
        const blank = {};
        if (headers && Array.isArray(headers)) {
            headers.forEach(h => blank[h] = "");
        }
        return blank;
    };

    // 2. Handle NEW Persons (These always go to the bottom)
    APP_STATE.stagedNewPersons.forEach(item => {
        newPersonRows.push(item.person);
        if(Array.isArray(item.staff)) newStaffRows.push(...item.staff);
        else newStaffRows.push(item.staff);
    });

    // 3. Handle EDITS (Blank old rows, move new data to bottom)
    APP_STATE.stagedEdits.forEach((data, pid) => {
        // --- A. PROCESS PERSON SHEET ---
        const pIdx = outPersons.findIndex(p => p.PersonID === pid);
        if(pIdx > -1) {
            // Replace the original data with a blank row
            outPersons[pIdx] = createBlankRow(APP_STATE.sheetHeaders.person);
            
            // Add the updated data to the "New Rows" queue
            newPersonRows.push(data.person);
        }

        // --- B. PROCESS STAFF SHEET ---
        // Staff is tricky because one person might have multiple rows.
        // We find ALL rows belonging to this ID, blank them, and add the new arrangement to the end.
        for(let i = 0; i < outStaff.length; i++) {
            if(outStaff[i].PersonID === pid) {
                outStaff[i] = createBlankRow(APP_STATE.sheetHeaders.staff);
            }
        }
        // Add the updated staff rows to the "New Rows" queue
        newStaffRows.push(...data.staff);
    });

    // 4. Combine: (Originals with Blanks) + (New/Moved Rows)
    const finalPersons = [...outPersons, ...newPersonRows];
    const finalStaff = [...outStaff, ...newStaffRows];

    // 5. Write Excel
    const wb = APP_STATE.workbook;
    const writeSafeSheet = (keyword, dataArray, headerArray) => {
        const sheetName = wb.SheetNames.find(n => n.toLowerCase().includes(keyword));
        if(sheetName && headerArray && headerArray.length > 0) {
            // json_to_sheet with these options ensures blank objects become empty rows
            const newSheet = XLSX.utils.json_to_sheet(dataArray, { header: headerArray, skipHeader: false });
            wb.Sheets[sheetName] = newSheet;
        }
    };
    
    writeSafeSheet('person', finalPersons, APP_STATE.sheetHeaders.person);
    writeSafeSheet('staff', finalStaff, APP_STATE.sheetHeaders.staff);

    XLSX.writeFile(wb, "Pure_Updated_Masterlist.xlsx");

    // 6. Log (With small delay)
    setTimeout(() => {
        const logContent = generateChangeLog();
        const logBlob = new Blob([logContent], {type: "text/plain;charset=utf-8"});
        triggerDownload(logBlob, "change_log.txt");
    }, 500);

    // 7. Zip (With larger delay to prevent blocking)
    const hasNewPhotos = APP_STATE.stagedNewPersons.some(x => x.photoBlob);
    const hasPhotoEdits = APP_STATE.stagedPhotos.size > 0;
    
    if(hasNewPhotos || hasPhotoEdits) {
        setTimeout(async () => {
            try {
                const newZip = new JSZip();
                const filesToRemove = new Set();
                APP_STATE.stagedPhotos.forEach(val => { if(val.originalFilename) filesToRemove.add(val.originalFilename); });

                // Copy existing photos (if zip loaded)
                if(APP_STATE.zipObject && APP_STATE.zipObject.files) {
                    for (const [filename, fileData] of Object.entries(APP_STATE.zipObject.files)) {
                        if(!fileData.dir && !filesToRemove.has(filename)) {
                            const content = await fileData.async('blob');
                            newZip.file(filename, content);
                        }
                    }
                }

                // Add New
                APP_STATE.stagedNewPersons.forEach(item => { 
                    if(item.photoBlob) newZip.file(item.photoName, item.photoBlob); 
                });
                APP_STATE.stagedPhotos.forEach(val => {
                    newZip.file(val.filename, val.file);
                });

                const blob = await newZip.generateAsync({type:"blob"});
                triggerDownload(blob, "Updated_Photos.zip");
                
                // Cleanup
                clearStaging();
            } catch(e) {
                console.error("Zip Error:", e);
                alert("Error generating photo zip: " + e.message);
            }
        }, 1500);
    } else {
        setTimeout(() => clearStaging(), 1000);
    }
}
