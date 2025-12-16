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