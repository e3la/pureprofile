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
    if(el) el.innerHTML = '⏳ Reading...';
    try {
        const buf = await file.arrayBuffer();
        APP_STATE.workbook = XLSX.read(buf, {type: 'array'});
        processWorkbookData();
        APP_STATE.hasExcel = true;
        if(el) { el.innerHTML = '✅ Loaded'; el.className = 'badge bg-success'; }
        
        const dropZone = document.getElementById('dropExcel');
        const content = document.getElementById('excelContent');
        if(dropZone) dropZone.classList.add('drop-success');
        if(content) content.innerHTML = `<div style="font-size: 2rem;">✅</div><p><strong>Excel Loaded</strong><br><span class="small">${file.name}</span></p>`;
        
        checkUnlockStatus();
    } catch(err) {
        if(el) { el.innerHTML = '❌ Error'; el.className = 'badge bg-danger'; }
        alert("Excel Error: " + err.message);
    }
}

async function handleZipLoad(file) {
    const el = document.getElementById('statusZip');
    if(el) el.innerHTML = '⏳ Reading...';
    try {
        APP_STATE.zipObject = await JSZip.loadAsync(file);
        APP_STATE.hasZip = true;
        if(el) { el.innerHTML = '✅ Loaded'; el.className = 'badge bg-success'; }

        const dropZone = document.getElementById('dropZip');
        const content = document.getElementById('zipContent');
        if(dropZone) dropZone.classList.add('drop-success');
        if(content) content.innerHTML = `<div style="font-size: 2rem;">✅</div><p><strong>Zip Loaded</strong><br><span class="small">${file.name}</span></p>`;

        checkUnlockStatus();
    } catch(err) {
        if(el) { el.innerHTML = '❌ Error'; el.className = 'badge bg-danger'; }
        alert("Zip Error: " + err.message);
    }
}

function processWorkbookData() {
    const wb = APP_STATE.workbook;
    // Orgs
    const orgSheetName = wb.SheetNames.find(n => n.toLowerCase().includes('organis') && !n.toLowerCase().includes('hierarch'));
    if(orgSheetName) {
        const rawRows = XLSX.utils.sheet_to_json(wb.Sheets[orgSheetName]);
        APP_STATE.allOrgs = rawRows.filter(r => r.OrganisationID && !r.OrganisationID.toString().includes(' '))
            .map(o => ({
                id: o.OrganisationID,
                name: o.Name_en || o.Name || "Unknown",
                visibility: o.Visibility || 'Public'
            })).sort((a,b) => a.name.localeCompare(b.name));
    }
    // Persons
    const persSheetName = wb.SheetNames.find(n => n.toLowerCase().includes('person'));
    if(persSheetName) {
        const worksheet = wb.Sheets[persSheetName];
        APP_STATE.sheetHeaders.person = XLSX.utils.sheet_to_json(worksheet, {header: 1})[0];
        APP_STATE.rawPersons = XLSX.utils.sheet_to_json(worksheet);
    }
    // Staff & Jobs
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
    // Options
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