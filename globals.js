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
    FTE: '',
    WebsiteURL_en: '',
    WebsiteURL_translated: '',
    DirectPhoneNr: '',
    MobilePhoneNr: '',
    FaxNr: ''
};

const APP_STATE = {
    workbook: null,
    zipObject: null,
    hasExcel: false,
    hasZip: false,
    rawPersons: [],
    rawStaff: [],
    sheetHeaders: { person: [], staff: [] },
    allOrgs: [],
    topJobTitles: [],  
    allJobTitles: [], 
    employmentOptions: ['faculty', 'staff', 'emeritus', 'other'],
    settings: { showUUIDs: false, includeRestricted: false },
    stagedEdits: new Map(), 
    stagedNewPersons: [],   
    stagedPhotos: new Map(), 
    currentMode: 'add', 
    currentEditBlob: null, 
    currentOriginalPhotoName: null
};

window.onerror = function(msg, url, line, col, error) {
    console.error("Global Error:", msg, error);
    return false;
};