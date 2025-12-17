/**
 * PURE ORGANISATION MANAGER
 * Updated: With Staging Areas for Hierarchy and Names
 */

const ORG_APP = {
    wb: null,
    
    // Data storage (Originals)
    hierData: [],       
    hierDescRow: null,  
    orgData: [],        
    orgDescRow: null,   
    orgHeader: [],      

    // Staging State
    stagedHierAdd: [],        // Array of new objects { ParentOrganisationID, ChildOrganisationID, ... }
    stagedHierDel: new Set(), // Set of _originalIndex integers to be removed
    stagedOrgEdits: new Map(), // Map<OrgID, { Name_en, Visibility }>

    // State
    lookupMap: {},      
    sortState: {
        hier: { col: null, asc: true },
        names: { col: 'name', asc: true }
    },

    init: function() {
        this.setupListeners();
    },

    setupListeners: function() {
        const drop = document.getElementById('dropZone');
        const inp = document.getElementById('fileInput');

        if(drop && inp) {
            drop.addEventListener('click', () => inp.click());
            drop.addEventListener('dragover', (e) => { e.preventDefault(); drop.style.borderColor = '#0d6efd'; drop.style.backgroundColor = '#eef5ff'; });
            drop.addEventListener('dragleave', () => { drop.style.borderColor = '#0d6efd'; drop.style.backgroundColor = '#fff'; });
            drop.addEventListener('drop', (e) => {
                e.preventDefault();
                drop.style.borderColor = '#0d6efd';
                drop.style.backgroundColor = '#fff';
                if(e.dataTransfer.files.length) this.loadFile(e.dataTransfer.files[0]);
            });
            inp.addEventListener('change', (e) => {
                if(e.target.files.length) this.loadFile(e.target.files[0]);
            });
        }
    },

    loadFile: async function(file) {
        const msg = document.getElementById('statusMsg');
        msg.innerText = "Reading Excel...";
        msg.className = "badge bg-info text-dark fs-6 fw-normal px-3 py-2";
        
        try {
            const buf = await file.arrayBuffer();
            this.wb = XLSX.read(buf, { type: 'array' });
            this.processWorkbook();
            
            document.getElementById('loadScreen').classList.add('d-none');
            document.getElementById('mainTool').classList.remove('d-none');
            this.render();
        } catch(e) {
            console.error(e);
            alert("Error loading file: " + e.message);
            msg.innerText = "Error Loading File";
            msg.className = "badge bg-danger fs-6 fw-normal px-3 py-2";
        }
    },

    processWorkbook: function() {
        // --- 1. PROCESS ORGANISATIONS SHEET ---
        const orgSheetName = this.wb.SheetNames.find(n => n.toLowerCase().includes('organis') && !n.toLowerCase().includes('hierarch'));
        if(!orgSheetName) throw new Error("Could not find 'Organisations' sheet.");

        const orgWs = this.wb.Sheets[orgSheetName];
        let orgRaw = XLSX.utils.sheet_to_json(orgWs); 
        this.orgHeader = XLSX.utils.sheet_to_json(orgWs, { header: 1 })[0];

        if (orgRaw.length > 0) {
            const firstRow = orgRaw[0];
            const checkVal = (firstRow.Visibility || '') + (firstRow.Name_en || '');
            if (checkVal.length > 50 || (firstRow.OrganisationID && firstRow.OrganisationID.includes('unique identifier'))) {
                this.orgDescRow = orgRaw.shift();
            }
        }
        this.orgData = orgRaw;

        // Build Lookup Map
        this.lookupMap = {};
        this.orgData.forEach(r => {
            if(r.OrganisationID) {
                this.lookupMap[r.OrganisationID] = {
                    name: r.Name_en || r.Name || "(No Name)",
                    vis: (r.Visibility || 'Public').toLowerCase()
                };
            }
        });

        // --- 2. PROCESS HIERARCHY SHEET ---
        const hierSheetName = this.wb.SheetNames.find(n => n.toLowerCase().includes('hierarch'));
        if(!hierSheetName) throw new Error("Could not find 'OrganisationalHierarchy' sheet.");

        const hierWs = this.wb.Sheets[hierSheetName];
        let hierRaw = XLSX.utils.sheet_to_json(hierWs);

        if (hierRaw.length > 0) {
            const firstRow = hierRaw[0];
            if (firstRow.ParentOrganisationID && firstRow.ParentOrganisationID.toString().toLowerCase().includes('id entered')) {
                this.hierDescRow = firstRow;
                hierRaw.shift();
            }
        }
        // Add original index for safe deletion tracking
        this.hierData = hierRaw.map((row, index) => ({ ...row, _originalIndex: index }));

        this.updateDatalist();
    },

    updateDatalist: function() {
        const list = document.getElementById('orgList');
        if(!list) return;

        const toggle = document.getElementById('toggleRestricted');
        const includeRestricted = toggle ? toggle.checked : false;
        
        list.innerHTML = '';
        
        const sorted = [...this.orgData]
            .filter(o => {
                 if(!includeRestricted && (o.Visibility || 'public').toLowerCase() === 'restricted') return false;
                 return true;
            })
            .sort((a,b) => (a.Name_en || '').localeCompare(b.Name_en || ''));
        
        sorted.forEach(o => {
            if(!o.OrganisationID) return;
            const opt = document.createElement('option');
            opt.value = `${o.Name_en || o.Name} [${o.OrganisationID}]`;
            list.appendChild(opt);
        });
    },

    render: function() {
        this.updateDatalist();
        
        if(document.getElementById('hierBody')) {
            this.renderHierarchy();
            this.renderStagedHier();
        }
        if(document.getElementById('orgBody')) {
            this.renderNames();
            this.renderStagedOrgs();
        }
    },

    getOrgMeta: function(id) {
        // Check staged edits first
        if (this.stagedOrgEdits.has(id)) {
            const staged = this.stagedOrgEdits.get(id);
            return { name: staged.Name_en, vis: staged.Visibility.toLowerCase() };
        }
        return this.lookupMap[id] || { name: 'Unknown', vis: 'public' };
    },

    /* ================= HIERARCHY LOGIC ================= */
    
    sortHier: function(col, thElement) {
        if(this.sortState.hier.col === col) {
            this.sortState.hier.asc = !this.sortState.hier.asc;
        } else {
            this.sortState.hier.col = col;
            this.sortState.hier.asc = true;
        }
        this.updateHeaderIcons('hierBody', thElement, this.sortState.hier.asc);
        this.renderHierarchy();
    },

    renderHierarchy: function() {
        const tbody = document.getElementById('hierBody');
        if(!tbody) return;

        tbody.innerHTML = '';
        const toggleRest = document.getElementById('toggleRestricted');
        const includeRestricted = toggleRest ? toggleRest.checked : false;
        
        const toggleUUID = document.getElementById('toggleUUIDs');
        const showUUID = toggleUUID ? toggleUUID.checked : false;

        // 1. Filter & Prepare
        let displayData = this.hierData.map(row => {
            const pMeta = this.getOrgMeta(row.ParentOrganisationID);
            const cMeta = this.getOrgMeta(row.ChildOrganisationID);
            return {
                raw: row,
                pName: pMeta.name,
                cName: cMeta.name,
                pVis: pMeta.vis,
                cVis: cMeta.vis,
                isDeleted: this.stagedHierDel.has(row._originalIndex)
            };
        }).filter(item => {
            if(!item.raw.ParentOrganisationID && !item.raw.ChildOrganisationID) return false;
            if(!includeRestricted && (item.pVis === 'restricted' || item.cVis === 'restricted')) return false;
            return true;
        });

        // 2. Sort
        const { col, asc } = this.sortState.hier;
        if(col) {
            displayData.sort((a, b) => {
                let valA = (col === 'parent' ? a.pName : a.cName).toLowerCase();
                let valB = (col === 'parent' ? b.pName : b.cName).toLowerCase();
                if(valA < valB) return asc ? -1 : 1;
                if(valA > valB) return asc ? 1 : -1;
                return 0;
            });
        }

        // 3. Render
        displayData.forEach(item => {
            const tr = document.createElement('tr');
            if(item.isDeleted) {
                tr.classList.add('table-danger');
                tr.style.textDecoration = 'line-through';
                tr.style.opacity = '0.7';
            }
            
            let visBadge = '';
            if(item.cVis === 'restricted') {
                visBadge = ` <span class="badge bg-warning text-dark border ms-2" style="font-size: 0.65rem;">Restricted</span>`;
            }

            const pDisplay = showUUID 
                ? `<strong>${item.pName}</strong><br><span class='text-muted small'>${item.raw.ParentOrganisationID}</span>` 
                : `<span class="fw-bold">${item.pName}</span>`;

            const cDisplay = showUUID 
                ? `<strong>${item.cName}</strong>${visBadge}<br><span class='text-muted small'>${item.raw.ChildOrganisationID}</span>` 
                : `<span class="fw-bold">${item.cName}</span>${visBadge}`;

            const actionBtn = item.isDeleted
                ? `<button class="btn btn-sm btn-outline-secondary" onclick="ORG_APP.toggleHierDelete(${item.raw._originalIndex})">Undo</button>`
                : `<span class="action-icon delete-row" onclick="ORG_APP.toggleHierDelete(${item.raw._originalIndex})">❌</span>`;

            tr.innerHTML = `
                <td>${pDisplay}</td>
                <td>${cDisplay}</td>
                <td class="text-end">${actionBtn}</td>
            `;
            tbody.appendChild(tr);
        });

        const countEl = document.getElementById('hierCount');
        if(countEl) countEl.innerText = `${displayData.length} existing relationships`;
    },

    renderStagedHier: function() {
        const div = document.getElementById('stagedHierArea');
        const tbody = document.getElementById('stagedHierBody');
        const countSpan = document.getElementById('stagedHierCount');
        if(!div || !tbody) return;

        tbody.innerHTML = '';
        let count = 0;

        // Show deletions
        if(this.stagedHierDel.size > 0) {
             const row = document.createElement('tr');
             row.className = "table-light text-center";
             row.innerHTML = `<td colspan="3" class="text-danger fw-bold">Pending Deletions: ${this.stagedHierDel.size} rows (marked in table above)</td>`;
             tbody.appendChild(row);
             count += this.stagedHierDel.size;
        }

        // Show Additions
        this.stagedHierAdd.forEach((item, idx) => {
            count++;
            const tr = document.createElement('tr');
            tr.className = "table-success";
            tr.innerHTML = `
                <td><span class="badge bg-success">NEW</span> ${item.pName} <br><small class="text-muted">${item.ParentOrganisationID}</small></td>
                <td>${item.cName} <br><small class="text-muted">${item.ChildOrganisationID}</small></td>
                <td class="text-end"><button class="btn btn-sm btn-outline-danger" onclick="ORG_APP.removeStagedHierAdd(${idx})">Remove</button></td>
            `;
            tbody.appendChild(tr);
        });

        if(countSpan) countSpan.innerText = count;
        if(count > 0) div.classList.remove('d-none');
        else div.classList.add('d-none');
    },

    addHierarchy: function() {
        const pVal = document.getElementById('addParent').value;
        const cVal = document.getElementById('addChild').value;
        const err = document.getElementById('addError');
        
        err.classList.add('d-none');
        const extractID = (val) => { const m = val.match(/\[(.*?)\]$/); return m ? m[1] : val.trim(); };

        const pID = extractID(pVal);
        const cID = extractID(cVal);

        if(!pID || !cID) return this.showError("Both fields required");
        if(pID === cID) return this.showError("Parent and Child cannot be the same");
        if(!this.lookupMap[pID]) return this.showError(`Parent ID not found: ${pID}`);
        if(!this.lookupMap[cID]) return this.showError(`Child ID not found: ${cID}`);

        const includeRestricted = document.getElementById('toggleRestricted').checked;
        if(!includeRestricted) {
            if(this.lookupMap[pID].vis === 'restricted' || this.lookupMap[cID].vis === 'restricted') {
                 return this.showError("Cannot add Restricted organisations while they are hidden.");
            }
        }

        // Check Existing + Staged
        const alreadyExists = this.hierData.some(r => r.ParentOrganisationID === pID && r.ChildOrganisationID === cID && !this.stagedHierDel.has(r._originalIndex));
        const alreadyStaged = this.stagedHierAdd.some(r => r.ParentOrganisationID === pID && r.ChildOrganisationID === cID);

        if(alreadyExists || alreadyStaged) return this.showError("Relationship already exists");

        // STAGE IT
        this.stagedHierAdd.push({
            ParentOrganisationID: pID,
            ChildOrganisationID: cID,
            pName: this.lookupMap[pID].name,
            cName: this.lookupMap[cID].name
        });

        document.getElementById('addParent').value = '';
        document.getElementById('addChild').value = '';
        this.renderStagedHier();
    },

    toggleHierDelete: function(originalIndex) {
        if(this.stagedHierDel.has(originalIndex)) {
            this.stagedHierDel.delete(originalIndex); // Undo delete
        } else {
            this.stagedHierDel.add(originalIndex); // Mark for delete
        }
        this.renderHierarchy();
        this.renderStagedHier();
    },

    removeStagedHierAdd: function(idx) {
        this.stagedHierAdd.splice(idx, 1);
        this.renderStagedHier();
    },

    /* ================= NAMES LOGIC ================= */

    sortNames: function(col, thElement) {
        if(this.sortState.names.col === col) {
            this.sortState.names.asc = !this.sortState.names.asc;
        } else {
            this.sortState.names.col = col;
            this.sortState.names.asc = true;
        }
        this.updateHeaderIcons('orgBody', thElement, this.sortState.names.asc);
        this.renderNames();
    },

    renderNames: function() {
        const tbody = document.getElementById('orgBody');
        if(!tbody) return;

        tbody.innerHTML = '';
        const search = document.getElementById('searchOrgs').value.toLowerCase();
        const toggleRest = document.getElementById('toggleRestricted');
        const includeRestricted = toggleRest ? toggleRest.checked : false;

        // 1. Filter
        let matches = this.orgData.filter(o => {
            if(!o.OrganisationID) return false;
            
            // Apply Staged Edit values for filtering logic
            let effectiveName = o.Name_en || o.Name || '';
            let effectiveVis = o.Visibility || 'Public';

            if(this.stagedOrgEdits.has(o.OrganisationID)) {
                const s = this.stagedOrgEdits.get(o.OrganisationID);
                effectiveName = s.Name_en;
                effectiveVis = s.Visibility;
            }

            const vis = effectiveVis.toLowerCase();
            if(!includeRestricted && vis === 'restricted') return false;

            const txt = `${effectiveName} ${o.OrganisationID}`.toLowerCase();
            return txt.includes(search);
        });

        // 2. Sort
        const { col, asc } = this.sortState.names;
        if(col) {
            matches.sort((a, b) => {
                let valA, valB;
                
                // Helper to get effective value
                const getVal = (row) => {
                    if(this.stagedOrgEdits.has(row.OrganisationID)) {
                        return this.stagedOrgEdits.get(row.OrganisationID);
                    }
                    return { Name_en: row.Name_en || row.Name, Visibility: row.Visibility };
                };

                const objA = getVal(a);
                const objB = getVal(b);

                if(col === 'id') { valA = a.OrganisationID; valB = b.OrganisationID; }
                else if(col === 'vis') { valA = objA.Visibility || ''; valB = objB.Visibility || ''; }
                else { valA = objA.Name_en || ''; valB = objB.Name_en || ''; }
                
                valA = String(valA).toLowerCase();
                valB = String(valB).toLowerCase();
                
                if(valA < valB) return asc ? -1 : 1;
                if(valA > valB) return asc ? 1 : -1;
                return 0;
            });
        }

        // 3. Render (Limit 100)
        matches.slice(0, 100).forEach(o => {
            const tr = document.createElement('tr');
            
            let name = o.Name_en || o.Name || '';
            let visibility = o.Visibility || 'Public';
            let isEdited = false;

            if(this.stagedOrgEdits.has(o.OrganisationID)) {
                const s = this.stagedOrgEdits.get(o.OrganisationID);
                name = s.Name_en;
                visibility = s.Visibility;
                isEdited = true;
                tr.classList.add('table-warning');
            }

            const visClass = visibility.toLowerCase() === 'restricted' ? 'bg-warning text-dark' : 'bg-light text-dark border';
            const editBadge = isEdited ? '<span class="badge bg-warning text-dark me-1">Modified</span>' : '';

            tr.innerHTML = `
                <td class="small text-muted font-monospace" style="font-size:0.8rem">${o.OrganisationID}</td>
                <td class="fw-bold text-primary">${editBadge}${name}</td>
                <td><span class="badge ${visClass}">${visibility}</span></td>
                <td class="text-end">
                    <span class="action-icon edit-row" onclick="ORG_APP.openEditModal('${o.OrganisationID}')">✏️</span>
                </td>
            `;
            tbody.appendChild(tr);
        });
        
        const countEl = document.getElementById('namesCount');
        if(countEl) countEl.innerText = `${matches.length} organisations shown`;
    },

    renderStagedOrgs: function() {
        const div = document.getElementById('stagedOrgArea');
        const tbody = document.getElementById('stagedOrgBody');
        const countSpan = document.getElementById('stagedOrgCount');
        if(!div || !tbody) return;

        tbody.innerHTML = '';
        if(this.stagedOrgEdits.size === 0) {
            div.classList.add('d-none');
            return;
        }

        this.stagedOrgEdits.forEach((val, id) => {
            const tr = document.createElement('tr');
            tr.innerHTML = `
                <td>${id}</td>
                <td>${val.Name_en}</td>
                <td>${val.Visibility}</td>
                <td class="text-end"><button class="btn btn-sm btn-outline-danger" onclick="ORG_APP.removeStagedOrgEdit('${id}')">Undo</button></td>
            `;
            tbody.appendChild(tr);
        });

        if(countSpan) countSpan.innerText = this.stagedOrgEdits.size;
        div.classList.remove('d-none');
    },

    removeStagedOrgEdit: function(id) {
        this.stagedOrgEdits.delete(id);
        this.renderNames();
        this.renderStagedOrgs();
    },

    /* ================= UTILS & EXPORT ================= */

    updateHeaderIcons: function(tbodyId, activeTh, isAsc) {
        const table = document.getElementById(tbodyId).closest('table');
        const ths = table.querySelectorAll('th.sortable');
        ths.forEach(th => {
            th.classList.remove('asc', 'desc');
            if(th === activeTh) {
                th.classList.add(isAsc ? 'asc' : 'desc');
            }
        });
    },

    showError: function(msg) {
        const el = document.getElementById('addError');
        el.innerText = msg;
        el.classList.remove('d-none');
        setTimeout(() => el.classList.add('d-none'), 3000);
    },

    openEditModal: function(id) {
        let name = "";
        let vis = "Public";

        // Check staged first, then raw
        if(this.stagedOrgEdits.has(id)) {
            const s = this.stagedOrgEdits.get(id);
            name = s.Name_en;
            vis = s.Visibility;
        } else {
            const org = this.orgData.find(o => o.OrganisationID === id);
            if(org) {
                name = org.Name_en || org.Name || "";
                vis = org.Visibility || 'Public';
            }
        }
        
        // Normalize Vis
        if(vis.toLowerCase() === 'restricted') vis = 'Restricted';
        else vis = 'Public';

        document.getElementById('editId').value = id;
        document.getElementById('dispEditId').value = id;
        document.getElementById('editNameEn').value = name;
        document.getElementById('editVisibility').value = vis;

        const modal = new bootstrap.Modal(document.getElementById('editModal'));
        modal.show();
    },

    saveOrgChanges: function() {
        const id = document.getElementById('editId').value;
        const newName = document.getElementById('editNameEn').value;
        const newVis = document.getElementById('editVisibility').value;

        // STAGE IT instead of writing to raw data
        this.stagedOrgEdits.set(id, { Name_en: newName, Visibility: newVis });

        this.renderNames(); 
        this.renderStagedOrgs();
        
        const el = document.getElementById('editModal');
        const modal = bootstrap.Modal.getInstance(el);
        modal.hide();
    },

    download: function() {
        if(!this.wb) return;

        // 1. Prepare Hierarchy
        // Filter out deleted indices, then remove the helper key
        let finalHierData = this.hierData
            .filter(r => !this.stagedHierDel.has(r._originalIndex))
            .map(row => {
                const clean = { ...row };
                delete clean._originalIndex;
                return clean;
            });
        
        // Append Adds
        this.stagedHierAdd.forEach(add => {
            finalHierData.push({
                ParentOrganisationID: add.ParentOrganisationID,
                ChildOrganisationID: add.ChildOrganisationID
            });
        });

        if(this.hierDescRow) finalHierData.unshift(this.hierDescRow);
        
        // Write Hier Sheet
        const newHierSheet = XLSX.utils.json_to_sheet(finalHierData);
        const hierName = this.wb.SheetNames.find(n => n.toLowerCase().includes('hierarch'));
        this.wb.Sheets[hierName] = newHierSheet;


        // 2. Prepare Organisations
        let finalOrgData = this.orgData.map(o => {
            // Clone
            let newObj = { ...o };
            // Apply Staged Edit
            if(this.stagedOrgEdits.has(o.OrganisationID)) {
                const edits = this.stagedOrgEdits.get(o.OrganisationID);
                newObj.Name_en = edits.Name_en;
                newObj.Visibility = edits.Visibility;
            }
            return newObj;
        });

        if(this.orgDescRow) finalOrgData.unshift(this.orgDescRow);

        // Write Org Sheet
        const newOrgSheet = XLSX.utils.json_to_sheet(finalOrgData, { header: this.orgHeader });
        const orgName = this.wb.SheetNames.find(n => n.toLowerCase().includes('organis') && !n.toLowerCase().includes('hierarch'));
        this.wb.Sheets[orgName] = newOrgSheet;

        // Save
        XLSX.writeFile(this.wb, "Updated_Organisations_Masterlist.xlsx");
        
        // Optional: Clear staging? 
        // this.stagedHierAdd = []; this.stagedHierDel.clear(); this.stagedOrgEdits.clear();
        // this.render();
    }
};

document.addEventListener('DOMContentLoaded', () => ORG_APP.init());