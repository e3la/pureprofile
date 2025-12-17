/**
 * ORGANISATION VIEWER
 * Logic to map Persons -> Staff -> Organisations
 */

const ORG_VIEWER = {
    wb: null,
    
    // Raw Data
    rawPersons: [],
    rawStaff: [],
    rawOrgs: [],

    // Processed Data
    deptMap: new Map(), // Map<OrgID, { name: string, vis: string, members: Array }>
    pMap: new Map(),    // Map<PersonID, { first, last, name, email }>
    sortedDeptIds: [],

    currentDeptId: null,

    init: function() {
        this.setupListeners();
    },

    setupListeners: function() {
        const drop = document.getElementById('dropZone');
        const inp = document.getElementById('fileInput');

        if(drop && inp) {
            drop.addEventListener('click', () => {
                inp.value = ''; // Reset input so same file can be selected again if needed
                inp.click();
            });
            
            drop.addEventListener('dragover', (e) => { 
                e.preventDefault(); 
                drop.style.borderColor = '#0d6efd'; 
                drop.style.backgroundColor = '#eef5ff'; 
            });
            
            drop.addEventListener('dragleave', () => { 
                drop.style.borderColor = '#0d6efd'; 
                drop.style.backgroundColor = '#fff'; 
            });
            
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
        msg.innerText = "Processing Data...";
        msg.className = "badge bg-info text-dark fs-6 fw-normal px-3 py-2";
        
        try {
            const buf = await file.arrayBuffer();
            this.wb = XLSX.read(buf, { type: 'array' });
            this.processData();
            
            document.getElementById('loadScreen').classList.add('d-none');
            document.getElementById('mainTool').classList.remove('d-none');
            this.renderSidebar();
        } catch(e) {
            console.error(e);
            alert("Error loading file: " + e.message);
            msg.innerText = "Error Loading File";
            msg.className = "badge bg-danger";
        }
    },

    processData: function() {
        // 1. Get Sheets
        const getJson = (keyword) => {
            const name = this.wb.SheetNames.find(n => n.toLowerCase().includes(keyword) && !n.toLowerCase().includes('hierarch'));
            return name ? XLSX.utils.sheet_to_json(this.wb.Sheets[name]) : [];
        };

        this.rawPersons = getJson('person');
        this.rawStaff = getJson('staff');
        this.rawOrgs = getJson('organis');

        // 2. Build Person Lookup (ID -> Name details)
        this.pMap.clear();
        this.rawPersons.forEach(p => {
            if(p.PersonID) {
                this.pMap.set(p.PersonID, {
                    first: p.Firstname || '',
                    last: p.Lastname || '',
                    name: `${p.Firstname} ${p.Lastname}`,
                    email: p.Email || '',
                    sortName: (p.Lastname || '').toLowerCase()
                });
            }
        });

        // 3. Build Org Lookup (ID -> Name) & Initialize DeptMap
        this.deptMap.clear();
        this.rawOrgs.forEach(o => {
            if(o.OrganisationID) {
                this.deptMap.set(o.OrganisationID, {
                    id: o.OrganisationID,
                    name: o.Name_en || o.Name || o.OrganisationID,
                    vis: (o.Visibility || 'Public').toLowerCase(),
                    members: []
                });
            }
        });

        // 4. Process Staff Rows (The Link)
        this.rawStaff.forEach(s => {
            if(!s.PersonID || !s.OrganisationID) return;

            // Retrieve Objects
            const person = this.pMap.get(s.PersonID);
            const dept = this.deptMap.get(s.OrganisationID);

            // If we have a valid person and the org exists in our map
            if(person && dept) {
                dept.members.push({
                    personId: s.PersonID,
                    name: person.name,
                    sortName: person.sortName,
                    email: person.email,
                    job: s.JobDescription || 'Unknown',
                    contract: s.ContractType || '',
                    fte: s.FTE || ''
                });
            }
        });

        // 5. Clean up Map
        this.sortedDeptIds = Array.from(this.deptMap.keys()).sort((a,b) => {
            const nameA = this.deptMap.get(a).name.toLowerCase();
            const nameB = this.deptMap.get(b).name.toLowerCase();
            return nameA.localeCompare(nameB);
        });
    },

    renderSidebar: function() {
        const container = document.getElementById('deptList');
        const filter = document.getElementById('searchDepts').value.toLowerCase();
        
        // Get Settings
        const toggleUUID = document.getElementById('toggleUUIDs');
        const showUUIDs = toggleUUID ? toggleUUID.checked : false;
        
        const toggleRest = document.getElementById('toggleRestricted');
        const includeRestricted = toggleRest ? toggleRest.checked : false;

        container.innerHTML = '';

        this.sortedDeptIds.forEach(id => {
            const dept = this.deptMap.get(id);
            
            // Filter by Visibility
            if(!includeRestricted && dept.vis === 'restricted') return;

            // Filter by Search Text
            const match = dept.name.toLowerCase().includes(filter) || id.toLowerCase().includes(filter);
            
            if(match) {
                const count = dept.members.length;
                let displayName = dept.name;
                if(showUUIDs) displayName = `${dept.name} [${dept.id}]`;

                const div = document.createElement('div');
                div.className = `dept-item p-2 d-flex justify-content-between align-items-center ${this.currentDeptId === id ? 'active' : ''}`;
                div.onclick = () => this.selectDept(id);
                div.innerHTML = `
                    <div class="text-truncate me-2" style="font-size:0.9rem; font-weight:500;" title="${displayName}">${displayName}</div>
                    <span class="badge bg-secondary rounded-pill" style="font-size:0.7rem;">${count}</span>
                `;
                container.appendChild(div);
            }
        });
    },

    selectDept: function(id) {
        this.currentDeptId = id;
        
        // Update Sidebar UI
        document.querySelectorAll('.dept-item').forEach(el => el.classList.remove('active'));
        this.renderSidebar(); 
        
        // Show Main View
        document.getElementById('emptyState').classList.add('d-none');
        const view = document.getElementById('deptView');
        view.classList.remove('d-none');

        const dept = this.deptMap.get(id);

        // Header
        document.getElementById('viewDeptName').innerText = dept.name;
        document.getElementById('viewDeptId').innerText = dept.id;
        document.getElementById('memberCount').innerText = dept.members.length;

        // Table
        const tbody = document.getElementById('memberTableBody');
        tbody.innerHTML = '';

        // Sort members by Last Name
        const sortedMembers = [...dept.members].sort((a,b) => a.sortName.localeCompare(b.sortName));

        if(sortedMembers.length === 0) {
            tbody.innerHTML = `<tr><td colspan="4" class="text-center text-muted py-4">No staff members found in this organisation.</td></tr>`;
        } else {
            sortedMembers.forEach(m => {
                const tr = document.createElement('tr');
                tr.innerHTML = `
                    <td>
                        <div class="fw-bold text-primary">${m.name}</div>
                        <small class="text-muted">${m.personId}</small>
                    </td>
                    <td>${m.job}</td>
                    <td>
                        <div class="small">${m.contract}</div>
                        <div class="small text-muted">FTE: ${m.fte}</div>
                    </td>
                    <td><a href="mailto:${m.email}" class="text-decoration-none">${m.email}</a></td>
                `;
                tbody.appendChild(tr);
            });
        }
    },

    exportCurrentList: function() {
        if(!this.currentDeptId) return;
        const dept = this.deptMap.get(this.currentDeptId);
        if(!dept || dept.members.length === 0) return alert("Nothing to export");

        const data = dept.members.map(m => ({
            "Person ID": m.personId,
            "Name": m.name,
            "Job Description": m.job,
            "Contract Type": m.contract,
            "FTE": m.fte,
            "Email": m.email,
            "Organisation": dept.name,
            "Organisation ID": dept.id
        }));

        const ws = XLSX.utils.json_to_sheet(data);
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, "Members");
        
        const safeName = dept.name.replace(/[^a-z0-9]/gi, '_').substring(0, 30);
        XLSX.writeFile(wb, `${safeName}_Members.xlsx`);
    },

    openBlobModal: function() {
        if(!this.rawStaff.length) return alert("Please load an Excel file first.");
        
        const lines = [];
        
        // Loop through all Departments
        this.deptMap.forEach(dept => {
            if(dept.members.length > 0) {
                dept.members.forEach(m => {
                    const p = this.pMap.get(m.personId);
                    if(p) {
                        // Format: Lastname, Firstname | Organisation
                        lines.push(`${p.last}, ${p.first} | ${dept.name}`);
                    }
                });
            }
        });

        // Sort alphabetically by the whole line
        lines.sort((a,b) => a.localeCompare(b));
        
        // Show in Modal
        document.getElementById('blobOutput').value = lines.join('\n');
        const modal = new bootstrap.Modal(document.getElementById('blobModal'));
        modal.show();
    },

    copyBlob: function() {
        const el = document.getElementById('blobOutput');
        el.select();
        document.execCommand('copy'); // Fallback for older browsers
        
        // Modern API attempt
        if(navigator.clipboard) {
            navigator.clipboard.writeText(el.value).then(() => {
                alert("Copied to clipboard!");
            });
        } else {
            alert("Copied!");
        }
    }
};

document.addEventListener('DOMContentLoaded', () => ORG_VIEWER.init());