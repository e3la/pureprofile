function generateID() {
    if(APP_STATE.currentMode !== 'add') return;
    const first = document.getElementById('p_Firstname').value.trim().toLowerCase();
    const last = document.getElementById('p_Lastname').value.trim().toLowerCase();
    if(!first && !last) {
        document.getElementById('p_PersonID').value = '';
        document.getElementById('idWarning').classList.add('hidden');
        return;
    }
    let rawID = `${first}-${last}`;
    const id = rawID.replace(/\s+/g, '-').replace(/[^a-z0-9-]/g, '');
    document.getElementById('p_PersonID').value = id;
    const existsInExcel = APP_STATE.rawPersons.some(p => p.PersonID === id);
    const existsInStaging = APP_STATE.stagedNewPersons.some(item => item.person.PersonID === id);
    const warn = document.getElementById('idWarning');
    if(existsInExcel || existsInStaging) {
        warn.classList.remove('hidden');
        warn.innerText = "âš ï¸ ID already exists! Please modify name or ID manually.";
    } else {
        warn.classList.add('hidden');
    }
}

async function loadPersonForEdit(p) {
    document.getElementById('searchResults').classList.add('hidden');
    toggleFormInputs(true); 

    document.getElementById('p_PersonID').value = p.PersonID;
    document.getElementById('p_PersonID').readOnly = true; 
    document.getElementById('p_Firstname').value = p.Firstname || '';
    document.getElementById('p_Lastname').value = p.Lastname || '';
    document.getElementById('p_Email').value = p.Email || '';
    document.getElementById('p_KnownFirst').value = p.FirstNameKnownAs || '';
    document.getElementById('p_KnownLast').value = p.LastNameKnownAs || '';
    document.getElementById('p_PostNominals').value = p.PostNominals || '';
    document.getElementById('p_Visibility').value = (p.Visibility || 'public').toLowerCase();

    const imgEl = document.getElementById('p_PhotoPreview');
    const phEl = document.getElementById('photoPlaceholder');
    const status = document.getElementById('photoStatus');
    
    imgEl.classList.add('hidden');
    phEl.classList.remove('hidden');
    status.innerText = "Checking Zip...";
    
    const photoFilename = p.ProfilePhoto; 
    APP_STATE.currentOriginalPhotoName = photoFilename;

    if(APP_STATE.zipObject && photoFilename) {
        let zipFile = APP_STATE.zipObject.file(photoFilename);
        if (!zipFile) {
            const lowerTarget = photoFilename.toLowerCase();
            for (let fileName in APP_STATE.zipObject.files) {
                if (fileName.toLowerCase() === lowerTarget && !APP_STATE.zipObject.files[fileName].dir) {
                    zipFile = APP_STATE.zipObject.files[fileName];
                    break;
                }
            }
        }
        if(zipFile) {
            const blob = await zipFile.async('blob');
            imgEl.src = URL.createObjectURL(blob);
            imgEl.classList.remove('hidden');
            phEl.classList.add('hidden');
            status.innerText = photoFilename;
            APP_STATE.currentEditBlob = blob; 
        } else {
            status.innerText = "File listed in Excel but not found in Zip";
            APP_STATE.currentEditBlob = null;
        }
    } else {
        status.innerText = "No Photo Listed in Excel";
        APP_STATE.currentEditBlob = null;
    }

    const staffRecs = APP_STATE.rawStaff.filter(s => s.PersonID === p.PersonID);
    document.getElementById('affRows').innerHTML = '';
    if(staffRecs.length === 0) {
        addAffiliationRow();
    } else {
        staffRecs.forEach(s => {
            const org = APP_STATE.allOrgs.find(o => o.id === s.OrganisationID);
            const name = org ? (APP_STATE.settings.showUUIDs ? `${org.name} [${s.OrganisationID}]` : org.name) : s.OrganisationID;
            addAffiliationRow({
                org: name,
                job: s.JobDescription,
                emp: s.EmployedAs,
                start: formatDate(s.StartDate)
            });
        });
    }
}

function loadStagedForEdit(type, key) {
    if (type === 'new') {
        document.getElementById('modeAdd').checked = true;
        setEditorMode('add'); 
    } else {
        document.getElementById('modeEdit').checked = true;
        setEditorMode('edit');
    }
    document.getElementById('personDetailsCard').scrollIntoView({ behavior: 'smooth' });
    document.getElementById('editStagedType').value = type;
    document.getElementById('editStagedIndex').value = key; 
    document.getElementById('stageBtn').innerText = "Update Staged Entry";

    let data, person, staff;
    if (type === 'new') {
        data = APP_STATE.stagedNewPersons[key];
        person = data.person;
        staff = data.staff;
        if (data.photoBlob) {
            APP_STATE.currentEditBlob = data.photoBlob;
            const preview = document.getElementById('p_PhotoPreview');
            preview.src = URL.createObjectURL(data.photoBlob);
            preview.classList.remove('hidden');
            document.getElementById('photoPlaceholder').classList.add('hidden');
            document.getElementById('photoStatus').innerText = "Staged Upload";
        }
    } else {
        data = APP_STATE.stagedEdits.get(key);
        person = data.person;
        staff = data.staff;
        if (APP_STATE.stagedPhotos.has(key)) {
            const photoData = APP_STATE.stagedPhotos.get(key);
            APP_STATE.currentEditBlob = photoData.file;
            const preview = document.getElementById('p_PhotoPreview');
            preview.src = URL.createObjectURL(photoData.file);
            preview.classList.remove('hidden');
            document.getElementById('photoPlaceholder').classList.add('hidden');
            document.getElementById('photoStatus').innerText = "Staged Upload";
        } else {
            const orig = APP_STATE.rawPersons.find(p => p.PersonID === key);
            if(orig) {
                 APP_STATE.currentOriginalPhotoName = orig.ProfilePhoto; 
                 document.getElementById('photoStatus').innerText = orig.ProfilePhoto + " (Unchanged)";
                 document.getElementById('photoPlaceholder').classList.remove('hidden');
                 document.getElementById('p_PhotoPreview').classList.add('hidden');
            }
        }
    }

    document.getElementById('p_PersonID').value = person.PersonID;
    document.getElementById('p_PersonID').readOnly = true;
    document.getElementById('p_Firstname').value = person.Firstname;
    document.getElementById('p_Lastname').value = person.Lastname;
    document.getElementById('p_Email').value = person.Email;
    document.getElementById('p_KnownFirst').value = person.FirstNameKnownAs;
    document.getElementById('p_KnownLast').value = person.LastNameKnownAs;
    document.getElementById('p_PostNominals').value = person.PostNominals;
    document.getElementById('p_Visibility').value = person.Visibility;

    document.getElementById('affRows').innerHTML = '';
    if(!staff || staff.length === 0) {
        addAffiliationRow();
    } else {
        staff.forEach(s => {
            const org = APP_STATE.allOrgs.find(o => o.id === s.OrganisationID);
            const name = org ? (APP_STATE.settings.showUUIDs ? `${org.name} [${s.OrganisationID}]` : org.name) : s.OrganisationID;
            addAffiliationRow({
                org: name,
                job: s.JobDescription,
                emp: s.EmployedAs,
                start: s.StartDate
            });
        });
    }
    toggleFormInputs(true);
}

function addAffiliationRow(data = null) {
    const defOrgName = "University of Missouri-St. Louis";
    const defOrgId = "e08c2356-15e4-4158-9cfe-34aecbd227e6";
    let showUUIDs = false;
    if (typeof APP_STATE !== 'undefined' && APP_STATE.settings) {
        showUUIDs = APP_STATE.settings.showUUIDs;
    }
    let defaultOrgStr = defOrgName;
    if(showUUIDs) defaultOrgStr = `${defOrgName} [${defOrgId}]`;
    const DEFAULT_JOB = ""; 
    const DEFAULT_EMP = "faculty";

    const div = document.createElement('div');
    div.className = 'aff-row';

    const opts = APP_STATE.employmentOptions.map(opt => 
        `<option value="${opt}" ${(data ? (data.emp == opt) : (opt == DEFAULT_EMP)) ? 'selected' : ''}>${opt}</option>`
    ).join('');

    const todayStr = getTodayStr();
    const orgVal = data ? data.org : defaultOrgStr;
    const jobVal = data ? data.job : DEFAULT_JOB;
    const startVal = (data && data.start) ? data.start : todayStr;

    div.innerHTML = `
        <div class="custom-dropdown-wrapper">
            <div class="input-group input-group-sm">
                <input type="text" class="form-control inp-org" value="${orgVal}" placeholder="Search Org..." autocomplete="off">
                <button class="btn btn-outline-secondary btn-org-arrow" type="button" tabindex="-1" style="border-left:0;">â–¼</button>
            </div>
            <div class="custom-dropdown-list org-list"></div>
        </div>
        <div class="custom-dropdown-wrapper">
            <div class="input-group input-group-sm">
                <input type="text" class="form-control inp-job" value="${jobVal}" placeholder="Job Description" autocomplete="off">
                <button class="btn btn-outline-secondary btn-job-arrow" type="button" tabindex="-1" style="border-left:0;">â–¼</button>
            </div>
            <div class="custom-dropdown-list job-list"></div>
        </div>
        <div><select class="form-select form-select-sm inp-emp">${opts}</select></div>
        <div><input type="text" class="form-control form-control-sm inp-start" value="${startVal}" placeholder="dd-mm-yyyy"></div>
        <div class="text-center"><button class="btn btn-outline-danger btn-sm rounded-circle" onclick="this.closest('.aff-row').remove()">Ã—</button></div>
    `;
    
    document.getElementById('affRows').appendChild(div);

    // ORG DROPDOWN
    const orgInput = div.querySelector('.inp-org');
    const orgArrow = div.querySelector('.btn-org-arrow');
    const orgList = div.querySelector('.org-list');

    const populateOrgList = (filterText = '') => {
        orgList.innerHTML = '';
        const lower = filterText.toLowerCase();
        let matches = APP_STATE.allOrgs.filter(o => {
            const matchesText = !filterText || `${o.name} ${o.id}`.toLowerCase().includes(lower);
            const matchesVis = APP_STATE.settings.includeRestricted || (o.visibility === 'Public' || o.visibility === 'public');
            return matchesText && matchesVis;
        });
        const displayLimit = 50; 
        const sliced = matches.slice(0, displayLimit);

        if(sliced.length === 0) {
            orgList.innerHTML = '<div style="padding:10px; color:#999;">No matches found</div>';
        } else {
            sliced.forEach(org => {
                const item = document.createElement('div');
                item.className = 'custom-dropdown-item';
                if (APP_STATE.settings.showUUIDs) {
                    item.textContent = `${org.name} [${org.id}]`;
                } else {
                    item.textContent = org.name;
                }
                item.onclick = () => {
                    orgInput.value = item.textContent;
                    orgList.style.display = 'none';
                };
                orgList.appendChild(item);
            });
            if(matches.length > displayLimit) {
                 const more = document.createElement('div');
                 more.style.padding = '5px 10px';
                 more.style.fontStyle = 'italic';
                 more.style.color = '#777';
                 more.innerText = `...and ${matches.length - displayLimit} more. Type to refine.`;
                 orgList.appendChild(more);
            }
        }
        orgList.style.display = 'block';
    };

    orgInput.addEventListener('input', () => {
        if(orgInput.value.length < 2) orgList.style.display = 'none';
        else populateOrgList(orgInput.value);
    });

    orgArrow.addEventListener('click', (e) => {
        e.stopPropagation();
        if(orgList.style.display === 'block') orgList.style.display = 'none';
        else { populateOrgList(''); orgInput.focus(); }
    });

    // JOB DROPDOWN
    const jobInput = div.querySelector('.inp-job');
    const jobArrow = div.querySelector('.btn-job-arrow');
    const jobList = div.querySelector('.job-list');

    const populateJobList = (filterText = '') => {
        jobList.innerHTML = '';
        const lower = filterText.toLowerCase();
        let sourceArray = [];
        let isTop10 = false;

        if (!filterText) {
            sourceArray = APP_STATE.topJobTitles;
            isTop10 = true;
        } else {
            sourceArray = APP_STATE.allJobTitles.filter(j => j.toLowerCase().includes(lower));
        }

        if (sourceArray.length === 0) {
            jobList.innerHTML = '<div style="padding:10px; color:#999;">No existing jobs match (New job will be created)</div>';
        } else {
            if (isTop10) {
                const head = document.createElement('div');
                head.className = 'custom-dropdown-header';
                head.innerText = 'Top 10 Most Common';
                jobList.appendChild(head);
            }
            const displayLimit = 50; 
            const sliced = sourceArray.slice(0, displayLimit);
            sliced.forEach(job => {
                const item = document.createElement('div');
                item.className = 'custom-dropdown-item';
                item.textContent = job;
                item.onclick = () => {
                    jobInput.value = job;
                    jobList.style.display = 'none';
                    checkHighlight(); 
                };
                jobList.appendChild(item);
            });
            if(!isTop10 && sourceArray.length > displayLimit) {
                 const more = document.createElement('div');
                 more.style.padding = '5px 10px';
                 more.style.fontStyle = 'italic';
                 more.style.color = '#777';
                 more.innerText = `...and ${sourceArray.length - displayLimit} more.`;
                 jobList.appendChild(more);
            }
        }
        jobList.style.display = 'block';
    };

    jobInput.addEventListener('input', () => {
        populateJobList(jobInput.value);
        checkHighlight();
    });
    
    jobInput.addEventListener('focus', () => {
        if(!jobInput.value) populateJobList('');
    });

    jobArrow.addEventListener('click', (e) => {
        e.stopPropagation();
        if(jobList.style.display === 'block') jobList.style.display = 'none';
        else { populateJobList(jobInput.value); jobInput.focus(); }
    });

    document.addEventListener('click', (e) => {
        if(!div.contains(e.target)) {
            orgList.style.display = 'none';
            jobList.style.display = 'none';
        }
    });

    const checkHighlight = () => {
         if(!jobInput.value.trim() || jobInput.value.trim() === 'Job Description') {
             div.classList.add('warning-bg');
         } else {
             div.classList.remove('warning-bg');
         }
    };
    checkHighlight(); 
}

function stagePersonChange() {
    const mode = APP_STATE.currentMode;
    const pid = document.getElementById('p_PersonID').value;
    const first = document.getElementById('p_Firstname').value.trim();
    const last = document.getElementById('p_Lastname').value.trim();

    if(!pid) return alert("Person ID is required.");
    const email = document.getElementById('p_Email').value;
    if(!email) return alert("Email is required.");

    const photoInput = document.getElementById('p_PhotoInput');
    let photoName = (mode === 'add') ? CONSTANTS.DefaultPhoto : APP_STATE.currentOriginalPhotoName;
    let photoBlob = null;
    const cleanName = `${first}-${last}`.replace(/[^a-z0-9-]/gi, '_');

    if(photoInput.files.length > 0) {
        const file = photoInput.files[0];
        const ext = file.name.split('.').pop();
        photoName = `${cleanName}.${ext}`; 
        photoBlob = file;
    } else if (APP_STATE.currentEditBlob) {
        photoBlob = APP_STATE.currentEditBlob;
        if(mode === 'add') {
             const ext = photoBlob.type === 'image/png' ? 'png' : 'jpg';
             photoName = `${cleanName}.${ext}`;
        }
    }

    const staffArr = [];
    document.querySelectorAll('.aff-row').forEach((row, index) => {
        const rawOrg = row.querySelector('.inp-org').value;
        let orgID = null;
        const match = rawOrg.match(/\[(.*?)\]$/);
        if (match) {
            orgID = match[1];
        } else {
            const found = APP_STATE.allOrgs.find(o => o.name === rawOrg.trim());
            if (found) orgID = found.id;
            else orgID = rawOrg; 
        }
        let jobDesc = row.querySelector('.inp-job').value;
        if(orgID && orgID.toLowerCase() === 'experts') {
            if(!jobDesc.startsWith('Media Contact - ')) {
                jobDesc = 'Media Contact - ' + jobDesc;
            }
        }
        const rawDate = row.querySelector('.inp-start').value;
        const processedDate = processInputDate(rawDate); 

        if(orgID) {
            staffArr.push({
                PersonID: pid,
                OrganisationID: orgID,
                ContractType: CONSTANTS.ContractType,
                JobTitle: "", 
                JobDescription: jobDesc, 
                JobDescription_translated: "",
                EmployedAs: row.querySelector('.inp-emp').value,
                FTE: CONSTANTS.FTE,
                StartDate: processedDate, 
                WebsiteURL_en: CONSTANTS.WebsiteURL_en,
                WebsiteURL_translated: CONSTANTS.WebsiteURL_translated,
                Primary: CONSTANTS.Primary,
                StaffType: CONSTANTS.StaffType,
                EndDate: CONSTANTS.EndDate,
                DirectPhoneNr: CONSTANTS.DirectPhoneNr,
                MobilePhoneNr: CONSTANTS.MobilePhoneNr,
                FaxNr: CONSTANTS.FaxNr,
                Email: email
            });
        }
    });

    if(mode === 'add' && staffArr.length === 0) return alert("New Persons must have at least one affiliation.");

    const rawIdx = document.getElementById('editStagedIndex').value;
    const editType = document.getElementById('editStagedType').value; 
    let newEntry = null; 
    let editEntry = null; 

    if(mode === 'add') {
         newEntry = {
            person: {
                PersonID: pid, Profiled: "yes", Username: email, Email: email,
                Title: "", Title_translated: "", PostNominals: document.getElementById('p_PostNominals').value,
                FirstNameKnownAs: document.getElementById('p_KnownFirst').value,
                LastNameKnownAs: document.getElementById('p_KnownLast').value,
                FirstNameSorting: "", LastNameSorting: "", FormerLastName: "", PriorAffiliations: "",
                Gender: CONSTANTS.Gender, Visibility: document.getElementById('p_Visibility').value,
                ProfilePhoto: photoName, ClientID_1: "", ClientID_2: "", ClientID_3: "", 
                ExternallyAuthenticated: "no",
                Firstname: first, Lastname: last, Nationality: "", ORCID: ""
            },
            staff: staffArr, photoBlob, photoName
        };
    } else {
        editEntry = {
            person: {
                PersonID: pid, Firstname: first, Lastname: last, Email: email,
                FirstNameKnownAs: document.getElementById('p_KnownFirst').value,
                LastNameKnownAs: document.getElementById('p_KnownLast').value,
                PostNominals: document.getElementById('p_PostNominals').value,
                Visibility: document.getElementById('p_Visibility').value,
                ProfilePhoto: (photoInput.files.length > 0) ? photoName : APP_STATE.currentOriginalPhotoName
            },
            staff: staffArr
        };
    }

    if(editType === 'new' && rawIdx !== "-1") {
        const idx = parseInt(rawIdx);
        APP_STATE.stagedNewPersons[idx] = newEntry; 
    } else if (editType === 'edit' && rawIdx !== "-1" && rawIdx !== "") {
        if(rawIdx !== pid) {
            APP_STATE.stagedEdits.delete(rawIdx);
            APP_STATE.stagedPhotos.delete(rawIdx);
        }
        APP_STATE.stagedEdits.set(pid, editEntry);
        if(photoInput.files.length > 0) {
             APP_STATE.stagedPhotos.set(pid, { 
                 file: photoBlob, 
                 filename: photoName, 
                 name: pid, 
                 originalFilename: APP_STATE.currentOriginalPhotoName 
             });
        }
    } else {
        if(mode === 'add') APP_STATE.stagedNewPersons.push(newEntry);
        else {
            APP_STATE.stagedEdits.set(pid, editEntry);
            if(photoInput.files.length > 0) {
                 APP_STATE.stagedPhotos.set(pid, { 
                     file: photoBlob, 
                     filename: photoName, 
                     name: pid,
                     originalFilename: APP_STATE.currentOriginalPhotoName
                 });
            }
        }
    }
    renderUnifiedStaging();
    resetEditor(true);
}