document.addEventListener('DOMContentLoaded', () => {
    setupDragAndDrop();
    
    // Setup inputs
    const elFirst = document.getElementById('p_Firstname');
    const elLast = document.getElementById('p_Lastname');
    if(elFirst) elFirst.addEventListener('input', generateID);
    if(elLast) elLast.addEventListener('input', generateID);
    
    const pPhotoInput = document.getElementById('p_PhotoInput');
    if(pPhotoInput) {
        pPhotoInput.addEventListener('change', function(e) {
            if(!e.target.files.length) return;
            const file = e.target.files[0];
            
            if(file.size > 1024 * 1024) { 
                alert("File too large. Max 1MB.");
                this.value = '';
                return;
            }
            const imgEl = document.getElementById('p_PhotoPreview');
            const phEl = document.getElementById('photoPlaceholder');
            const status = document.getElementById('photoStatus');
            
            imgEl.src = URL.createObjectURL(file);
            imgEl.classList.remove('hidden');
            phEl.classList.add('hidden');
            status.innerText = "New Upload Selected";
        });
    }

    // Search Logic
    const searchBox = document.getElementById('searchBox');
    if(searchBox) {
        searchBox.addEventListener('keyup', function() {
            if(APP_STATE.currentMode !== 'edit') return;
            const q = this.value.toLowerCase();
            const res = document.getElementById('searchResults');
            res.innerHTML = '';
            
            if(q.length < 2) { res.classList.add('hidden'); return; }

            const matches = APP_STATE.rawPersons.filter(p => {
                const matchesText = (p.Firstname + " " + p.Lastname + " " + p.PersonID).toLowerCase().includes(q);
                const matchesVis = APP_STATE.settings.includeRestricted || (p.Visibility === 'Public' || p.Visibility === 'public');
                return matchesText && matchesVis;
            }).slice(0, 8);

            matches.forEach(p => {
                const div = document.createElement('div');
                div.className = 'search-item';
                div.innerHTML = `<strong>${p.Firstname} ${p.Lastname}</strong> <small class="text-muted">(${p.PersonID})</small>`;
                div.onclick = () => loadPersonForEdit(p);
                res.appendChild(div);
            });
            if(matches.length) res.classList.remove('hidden');
        });
    }
});

function updateSettings() {
    if(APP_STATE.settings) {
        APP_STATE.settings.showUUIDs = document.getElementById('toggleUUIDs').checked;
        APP_STATE.settings.includeRestricted = document.getElementById('toggleRestricted').checked;
    }
}

function setEditorMode(mode) {
    APP_STATE.currentMode = mode;
    const searchWrapper = document.getElementById('searchContainerWrapper');
    const personDetailsCard = document.getElementById('personDetailsCard');
    const formTitle = document.getElementById('formTitle');
    const regenBtn = document.getElementById('regenIdBtn');
    const idField = document.getElementById('p_PersonID');
    const stageBtn = document.getElementById('stageBtn');

    resetEditor(false); 

    if(mode === 'add') {
        searchWrapper.classList.add('hidden');
        formTitle.innerText = "New Person Details";
        formTitle.className = "text-primary m-0";
        personDetailsCard.classList.add('border-primary');
        personDetailsCard.classList.remove('border-warning');
        
        regenBtn.disabled = false;
        idField.readOnly = true; 
        idField.classList.remove('bg-secondary', 'text-white');
        idField.classList.add('bg-light');
        
        stageBtn.innerText = "Stage New Person";
        stageBtn.className = "btn btn-primary w-100 btn-lg";
        toggleFormInputs(true);
        addAffiliationRow();
    } else {
        searchWrapper.classList.remove('hidden');
        formTitle.innerText = "Edit Existing Person";
        formTitle.className = "text-warning m-0";
        personDetailsCard.classList.remove('border-primary');
        personDetailsCard.classList.add('border-warning');
        regenBtn.disabled = true;
        idField.readOnly = true; 
        stageBtn.innerText = "Stage Edits";
        stageBtn.className = "btn btn-warning w-100 btn-lg";
        toggleFormInputs(false);
        document.getElementById('affRows').innerHTML = '<div class="text-muted text-center py-4">Search a person to edit affiliations</div>';
    }
}

function toggleFormInputs(enable) {
    const selector = '#personDetailsCard input, #personDetailsCard select, #personDetailsCard button:not(#regenIdBtn)';
    document.querySelectorAll(selector).forEach(el => el.disabled = !enable);
    document.getElementById('p_PhotoInput').disabled = !enable;
}

function resetEditor(fullReset = true) {
    document.querySelectorAll('#personDetailsCard input').forEach(i => i.value = '');
    document.getElementById('p_Visibility').value = 'public';
    const preview = document.getElementById('p_PhotoPreview');
    const placeholder = document.getElementById('photoPlaceholder');
    const status = document.getElementById('photoStatus');
    
    if(preview) {
        preview.src = '';
        preview.classList.add('hidden');
    }
    if(placeholder) placeholder.classList.remove('hidden');
    if(status) status.innerText = "None";
    
    document.getElementById('editStagedIndex').value = "-1";
    document.getElementById('editStagedType').value = "";
    APP_STATE.currentEditBlob = null;
    APP_STATE.currentOriginalPhotoName = null;
    
    const stageBtn = document.getElementById('stageBtn');
    if(APP_STATE.currentMode === 'add') stageBtn.innerText = "Stage New Person";
    else stageBtn.innerText = "Stage Edits";
    document.getElementById('affRows').innerHTML = '';
    
    if(fullReset && APP_STATE.currentMode === 'add') addAffiliationRow();
    if(fullReset && APP_STATE.currentMode === 'edit') {
        document.getElementById('affRows').innerHTML = '<div class="text-muted text-center py-4">Search a person to edit affiliations</div>';
        toggleFormInputs(false);
    }
    document.getElementById('idWarning').classList.add('hidden');
}

function renderUnifiedStaging() {
    const tbody = document.getElementById('stagingBody');
    if(!tbody) return; 
    tbody.innerHTML = '';
    
    let count = 0;
    APP_STATE.stagedNewPersons.forEach((item, idx) => {
        count++;
        const pImg = item.photoBlob ? '📷' : '🦁';
        tbody.innerHTML += `
            <tr class="table-success">
                <td><span class="badge bg-success">NEW</span></td>
                <td>${item.person.PersonID}</td>
                <td>${item.person.Firstname} ${item.person.Lastname}</td>
                <td title="${item.person.ProfilePhoto}">${pImg}</td>
                <td class="small">${item.staff.length} affiliation(s)</td>
                <td class="text-end">
                    <button class="btn btn-sm btn-info text-white me-1" onclick="loadStagedForEdit('new', ${idx})">Edit</button>
                    <button class="btn btn-sm btn-outline-danger" onclick="removeStagedNew(${idx})">Remove</button>
                </td>
            </tr>`;
    });

    APP_STATE.stagedEdits.forEach((val, key) => {
        count++;
        const hasPhotoChange = APP_STATE.stagedPhotos.has(key);
        const pImg = hasPhotoChange ? '📷 (Updated)' : '-';
        tbody.innerHTML += `
            <tr class="table-warning">
                <td><span class="badge bg-warning text-dark">EDIT</span></td>
                <td>${key}</td>
                <td>${val.person.Firstname} ${val.person.Lastname}</td>
                <td>${pImg}</td>
                <td class="small">Data update + ${val.staff.length} affiliations</td>
                <td class="text-end">
                    <button class="btn btn-sm btn-info text-white me-1" onclick="loadStagedForEdit('edit', '${key}')">Edit</button>
                    <button class="btn btn-sm btn-outline-danger" onclick="removeStagedEdit('${key}')">Remove</button>
                </td>
            </tr>`;
    });
    const counter = document.getElementById('totalStageCount');
    if(counter) counter.innerText = count;
    if(count > 0) document.getElementById('stagingCard').classList.remove('hidden');
    else document.getElementById('stagingCard').classList.add('hidden');
}

function removeStagedNew(idx) {
    APP_STATE.stagedNewPersons.splice(idx, 1);
    renderUnifiedStaging();
}
function removeStagedEdit(key) {
    APP_STATE.stagedEdits.delete(key);
    APP_STATE.stagedPhotos.delete(key); 
    renderUnifiedStaging();
}

function checkUnlockStatus() {
    if(APP_STATE.hasExcel && APP_STATE.hasZip) {
        const personCount = APP_STATE.rawPersons.length;
        const photoCount = Object.keys(APP_STATE.zipObject.files).filter(f => !APP_STATE.zipObject.files[f].dir).length;
        const statsEl = document.getElementById('fileStats');
        if(statsEl) {
            statsEl.classList.remove('hidden');
            statsEl.innerHTML = `✅ Ready! Found <strong>${personCount}</strong> Profiles and <strong>${photoCount}</strong> Photos.`;
        }
        setTimeout(() => { startPersonManager(); }, 1000);
    }
}

function goHome() {
    const hasChanges = (APP_STATE.stagedNewPersons.length > 0 || APP_STATE.stagedEdits.size > 0);
    if (hasChanges) {
        const discard = confirm("You have unsaved changes! \n\nClick OK to DISCARD data and leave.\nClick Cancel to stay here and save.");
        if (!discard) return; 
    }
    window.location.href = 'index.html';
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