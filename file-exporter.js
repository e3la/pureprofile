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
    // We use this clone to look up original data for merging
    let outPersons = JSON.parse(JSON.stringify(APP_STATE.rawPersons));
    let outStaff = JSON.parse(JSON.stringify(APP_STATE.rawStaff));

    // Arrays to hold the data that will be appended to the bottom
    const newPersonRows = [];
    const newStaffRows = [];

    // --- HELPER: Create a blank object based on headers ---
    // This ensures the row exists in Excel but is visually empty, preserving row numbers.
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

    // 3. Handle EDITS 
    // Logic: Find original -> Clone it -> Merge changes -> Push to bottom -> Blank original
    APP_STATE.stagedEdits.forEach((data, pid) => {
        
        // --- A. PROCESS PERSON SHEET ---
        const pIdx = outPersons.findIndex(p => p.PersonID === pid);
        
        if(pIdx > -1) {
            // 1. Grab the exact original row (contains columns the UI might not show)
            const originalRow = outPersons[pIdx];

            // 2. Create a merge: Start with Original, Overwrite with UI Changes
            // data.person contains only the fields controlled by the Form.
            // Spreading (...originalRow) first ensures we keep unknown columns.
            const mergedRow = { ...originalRow, ...data.person };

            // 3. Blank out the original slot in the main array
            // This prevents "Spreadsheet Compare" from seeing every subsequent row shift up by 1.
            outPersons[pIdx] = createBlankRow(APP_STATE.sheetHeaders.person);
            
            // 4. Add the complete, merged data to the "New Rows" queue
            newPersonRows.push(mergedRow);
        } else {
            // Fallback: If ID not found in original (rare), just push the UI data
            newPersonRows.push(data.person);
        }

        // --- B. PROCESS STAFF SHEET ---
        // Staff rows are complex because 1 Person can have multiple rows.
        // It is safer to blank ALL old rows for this person and append the NEW configuration.
        let firstFound = false;
        for(let i = 0; i < outStaff.length; i++) {
            if(outStaff[i].PersonID === pid) {
                // Blank the original row
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

    // Save File
    XLSX.writeFile(wb, "Pure_Updated_Masterlist.xlsx");

    // 6. Generate Change Log
    setTimeout(() => {
        const logContent = generateChangeLog();
        const logBlob = new Blob([logContent], {type: "text/plain;charset=utf-8"});
        triggerDownload(logBlob, "change_log.txt");
    }, 500);

    // 7. Generate Photo Zip (only if needed)
    const hasNewPhotos = APP_STATE.stagedNewPersons.some(x => x.photoBlob);
    const hasPhotoEdits = APP_STATE.stagedPhotos.size > 0;
    
    if(hasNewPhotos || hasPhotoEdits) {
        setTimeout(async () => {
            try {
                const newZip = new JSZip();
                const filesToRemove = new Set();
                APP_STATE.stagedPhotos.forEach(val => { if(val.originalFilename) filesToRemove.add(val.originalFilename); });

                // Copy existing photos from original Zip
                if(APP_STATE.zipObject && APP_STATE.zipObject.files) {
                    for (const [filename, fileData] of Object.entries(APP_STATE.zipObject.files)) {
                        // Skip directories and files marked for replacement
                        if(!fileData.dir && !filesToRemove.has(filename)) {
                            // Check strict filename first
                            let shouldInclude = true;
                            
                            // Also check our "Case Insensitive" logic from before to ensure we don't duplicate
                            // e.g. if we are replacing "photo.JPG", don't copy "photo.jpg"
                            const lowerName = filename.split('/').pop().toLowerCase();
                            for(let removeName of filesToRemove) {
                                if(removeName.toLowerCase() === lowerName) shouldInclude = false;
                            }

                            if(shouldInclude) {
                                const content = await fileData.async('blob');
                                newZip.file(filename, content);
                            }
                        }
                    }
                }

                // Add New/Edited Photos
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