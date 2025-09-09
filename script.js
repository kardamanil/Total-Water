// js/script.js - Modular Water Quality Report App (JalGanana style)
const firebaseConfig = {
    apiKey: "AIzaSyDSqffmlg33rpMc6a2JUddg9pYQFvR8aXU",
    authDomain: "labcalc-cee5c.firebaseapp.com",
    databaseURL: "https://labcalc-cee5c-default-rtdb.firebaseio.com",
    projectId: "labcalc-cee5c",
    storageBucket: "labcalc-cee5c.firebasestorage.app",
    messagingSenderId: "1030109019271",
    appId: "1:1030109019271:web:e66f263a8a1c003b41cc22",
    measurementId: "G-R9KDMPJX3S"
};

firebase.initializeApp(firebaseConfig);
const db = firebase.firestore();

const tests = [
    { name: "Colour", bilingual_name: "रंग", max_limit: "-", desirable_limit: "Clear" },
    { name: "Odour", bilingual_name: "गंध", max_limit: "-", desirable_limit: "OK" },
    { name: "Turbidity", bilingual_name: "टर्बिडिटी", max_limit: "5 NTU", desirable_limit: "1 NTU" },
    { name: "TDS", bilingual_name: "टीडीएस", max_limit: "2000 mg/l", desirable_limit: "500 mg/l" },
    { name: "pH", bilingual_name: "पीएच", max_limit: "6.5 से 8.5", desirable_limit: "6.5 से 8.5" },
    { name: "T. Hardness", bilingual_name: "कुल कठोरता", max_limit: "600 mg/l", desirable_limit: "300 mg/l" },
    { name: "Calcium", bilingual_name: "कैल्शियम", max_limit: "200 mg/l", desirable_limit: "75 mg/l" },
    { name: "Magnesium", bilingual_name: "मैग्नीशियम", max_limit: "100 mg/l", desirable_limit: "30 mg/l" },
    { name: "Chloride", bilingual_name: "क्लोराइड", max_limit: "1000 mg/l", desirable_limit: "250 mg/l" },
    { name: "Alkalinity", bilingual_name: "क्षारीयता", max_limit: "600 mg/l", desirable_limit: "200 mg/l" }
];

let chiDetails = {};
let sampleDetails = [];
let chemicalResults = [];
let queryResults = [];

// Validation Functions
function validateDate(dateStr) {
    try {
        const parts = dateStr.split(/[-\/]/);
        const date = new Date(parts[2], parts[1] - 1, parts[0]);
        return date.getDate() == parts[0] && (date.getMonth() + 1) == parts[1] && date.getFullYear() == parts[2];
    } catch {
        return false;
    }
}

function validateLabNo(labNo) {
    return /^\d+\/\d{4}$/.test(labNo.trim());
}

function validateNumber(value, min = 0, max = Infinity) {
    const num = parseFloat(value);
    return !isNaN(num) && num >= min && num <= max;
}

function categorizeSample(testName, result, maxLimit, desLimit) {
    try {
        const val = parseFloat(result);
        if (isNaN(val)) return "unknown";
        if (testName === "pH") {
            const [maxLow, maxHigh] = maxLimit.split(' से ').map(parseFloat);
            const [desLow, desHigh] = desLimit.split(' से ').map(parseFloat);
            if (val < maxLow || val > maxHigh) return "unsuitable";
            if (val < desLow || val > desHigh) return "permissible";
            return "suitable";
        } else {
            const maxVal = parseFloat(maxLimit.replace(/ mg\/l| NTU/g, '')) || Infinity;
            const desVal = parseFloat(desLimit.replace(/ mg\/l| NTU/g, '')) || Infinity;
            if (val > maxVal) return "unsuitable";
            if (val > desVal) return "permissible";
            return "suitable";
        }
    } catch {
        if (["Colour", "Odour", "Turbidity"].includes(testName)) {
            const des = desLimit.toLowerCase();
            return result.toLowerCase() === des ? "suitable" : "unsuitable";
        }
        return "unknown";
    }
}

// CHI Functions
function loadChiCsv() {
    const file = document.getElementById('chi-csv').files[0];
    if (!file) return setStatus("Please select a CSV file.", "warning");
    const reader = new FileReader();
    reader.onload = (e) => {
        const lines = e.target.result.split('\n').filter(line => line.trim());
        if (lines.length < 2) return setStatus("CSV is empty or invalid.", "error");
        const headers = lines[0].split(',');
        const row = lines[1].split(',');
        const data = {};
        headers.forEach((h, i) => data[h.trim()] = row[i]?.trim());
        const allowedDivisions = ["अजमेर", "जोधपुर", "जयपुर", "बीकानेर"];
        if (!data["CHI Letter No."] || !data["CHI Address"] || !allowedDivisions.includes(data["Division"]) || !validateDate(data["Report Date"])) {
            return setStatus("Invalid CSV data. Check columns: CHI Letter No., CHI Address, Division, Report Date.", "error");
        }
        document.getElementById('chi-letter-no').value = data["CHI Letter No."];
        document.getElementById('chi-address').value = data["CHI Address"];
        document.getElementById('chi-division').value = data["Division"];
        document.getElementById('report-date').value = data["Report Date"];
        setStatus("CHI details loaded from CSV.", "success");
    };
    reader.readAsText(file, 'UTF-8');
}

function clearChiForm() {
    document.getElementById('chi-letter-no').value = '';
    document.getElementById('chi-address').value = '';
    document.getElementById('chi-division').value = '';
    document.getElementById('report-date').value = new Date().toLocaleDateString('en-GB').split('/').reverse().join('-');
    document.getElementById('chi-csv').value = '';
    setStatus("CHI form cleared.", "info");
}

function openSampleTab() {
    chiDetails = {
        letterNo: document.getElementById('chi-letter-no').value.trim(),
        address: document.getElementById('chi-address').value.trim(),
        division: document.getElementById('chi-division').value,
        reportDate: document.getElementById('report-date').value.trim()
    };
    const allowedDivisions = ["अजमेर", "जोधपुर", "जयपुर", "बीकानेर"];
    if (!chiDetails.letterNo || !chiDetails.address || !allowedDivisions.includes(chiDetails.division) || !validateDate(chiDetails.reportDate)) {
        return setStatus("Please enter valid CHI Letter No., Address, Division, and Report Date (DD-MM-YYYY).", "error");
    }
    new bootstrap.Tab(document.querySelector('#sample-tab')).show();
    setStatus("Moved to Sample Details tab.", "success");
}

// Sample Functions
function addSampleEntry(sample = { source: '', location: '', chiSampleNo: '', date: '', labNo: '', sender: '' }, index = sampleDetails.length) {
    const container = document.getElementById('sample-entries');
    const div = document.createElement('div');
    div.className = 'sample-entry row align-items-end mb-3 p-3 border rounded bg-light';
    div.id = `sample-entry-${index}`;
    div.innerHTML = `
        <div class="col-md-10">
            <h6 class="mb-2 text-primary">Sample ${index + 1}</h6>
            <div class="row g-2">
                <div class="col-md-2">
                    <label class="form-label small fw-bold">Source</label>
                    <input type="text" class="form-control sample-source" value="${sample.source}" placeholder="Enter Source">
                </div>
                <div class="col-md-2">
                    <label class="form-label small fw-bold">Location</label>
                    <input type="text" class="form-control sample-location" value="${sample.location}" placeholder="Enter Location">
                </div>
                <div class="col-md-2">
                    <label class="form-label small fw-bold">CHI Sample No.</label>
                    <input type="text" class="form-control sample-chi-sample-no" value="${sample.chiSampleNo}" placeholder="Enter CHI No.">
                </div>
                <div class="col-md-2">
                    <label class="form-label small fw-bold">Date (DD/MM/YYYY)</label>
                    <input type="text" class="form-control sample-date" value="${sample.date}" placeholder="DD/MM/YYYY">
                </div>
                <div class="col-md-2">
                    <label class="form-label small fw-bold">Lab No. (123/2025)</label>
                    <input type="text" class="form-control sample-lab-no" value="${sample.labNo}" placeholder="123/2025">
                </div>
                <div class="col-md-2">
                    <label class="form-label small fw-bold">Sender</label>
                    <input type="text" class="form-control sample-sender" value="${sample.sender}" placeholder="Enter Sender">
                </div>
            </div>
        </div>
        <div class="col-md-2 text-end">
            <button type="button" class="btn btn-danger delete-btn p-2" onclick="deleteSampleEntry(${index})" title="Delete Sample ${index + 1}">
                <i class="bi bi-x-circle-fill" style="font-size: 1.5rem; color: red;"></i>
            </button>
        </div>
    `;
    container.appendChild(div);
    sampleDetails[index] = { Source: sample.source, Location: sample.location, "CHI Sample No.": sample.chiSampleNo, Date: sample.date, "Lab No.": sample.labNo, Sender: sample.sender };
}

function addSamplesFromNum() {
    const num = parseInt(document.getElementById('num-samples').value);
    if (isNaN(num) || num < 1 || num > 20) return setStatus("Enter number between 1-20.", "warning");
    for (let i = sampleDetails.length; i < num; i++) {
        addSampleEntry({}, i);
    }
    setStatus(`${num} samples added. You can edit or delete any.`, "success");
}

function loadSampleCsv() {
    const file = document.getElementById('sample-csv').files[0];
    if (!file) return setStatus("Select CSV file.", "warning");
    const reader = new FileReader();
    reader.onload = (e) => {
        const lines = e.target.result.split('\n').filter(l => l.trim());
        if (lines.length < 2) return setStatus("Invalid CSV.", "error");
        const headers = lines[0].split(',');
        const required = ["Source", "Location", "CHI Sample No.", "Date", "Lab No.", "Sender"];
        if (!required.every(h => headers.some(head => head.trim() === h))) return setStatus("CSV must contain: " + required.join(", "), "error");
        renderSampleEntries(); // Clear first
        let loaded = 0;
        for (let i = 1; i < lines.length; i++) {
            const row = lines[i].split(',');
            const data = {};
            headers.forEach((h, j) => data[h.trim()] = row[j]?.trim());
            if (validateDate(data.Date) && validateLabNo(data["Lab No."])) {
                addSampleEntry({
                    source: data.Source || '',
                    location: data.Location || '',
                    chiSampleNo: data["CHI Sample No."] || '',
                    date: data.Date || '',
                    labNo: data["Lab No."] || '',
                    sender: data.Sender || ''
                }, loaded);
                loaded++;
            }
        }
        setStatus(`${loaded} samples loaded from CSV. You can edit or delete any.`, "success");
    };
    reader.readAsText(file, 'UTF-8');
}

function deleteSampleEntry(index) {
    if (confirm(`Delete Sample ${index + 1}?`)) {
        sampleDetails.splice(index, 1);
        renderSampleEntries();
        setStatus(`Sample ${index + 1} deleted. Total samples now: ${sampleDetails.length}. This will reflect in Chemical table.`, "warning");
    }
}

function renderSampleEntries() {
    document.getElementById('sample-entries').innerHTML = '';
    sampleDetails.forEach((sample, i) => addSampleEntry(sample, i));
}

function clearSampleForm() {
    document.getElementById('num-samples').value = '';
    document.getElementById('sample-csv').value = '';
    sampleDetails = [];
    document.getElementById('sample-entries').innerHTML = '';
    setStatus("Sample form cleared.", "info");
}

function generateReport() {
    sampleDetails = [];
    const entries = document.querySelectorAll('.sample-entry');
    let valid = 0;
    entries.forEach(entry => {
        const source = entry.querySelector('.sample-source').value.trim();
        const location = entry.querySelector('.sample-location').value.trim();
        const chiSampleNo = entry.querySelector('.sample-chi-sample-no').value.trim();
        const date = entry.querySelector('.sample-date').value.trim();
        const labNo = entry.querySelector('.sample-lab-no').value.trim();
        const sender = entry.querySelector('.sample-sender').value.trim();
        if (source && location && chiSampleNo && date && labNo && sender && validateDate(date) && validateLabNo(labNo)) {
            sampleDetails.push({ Source: source, Location: location, "CHI Sample No.": chiSampleNo, Date: date, "Lab No.": labNo, Sender: sender });
            valid++;
        }
    });
    if (sampleDetails.length === 0) return setStatus("No valid samples. Fill all fields and validate formats.", "error");
    populateChemicalResultsTab();
    new bootstrap.Tab(document.querySelector('#chemical-tab')).show();
    setStatus(`${valid} valid samples processed. Moved to Chemical Results tab.`, "success");
}

// Chemical Functions
async function populateChemicalResultsTab() {
    const container = document.getElementById('chemical-table-container');
    container.innerHTML = '<div class="text-center p-3"><div class="spinner-border" role="status"><span class="visually-hidden">Loading...</span></div></div>';
    let tableHTML = '<table class="table table-bordered table-sm"><thead><tr><th>Test</th>';
    sampleDetails.forEach(s => tableHTML += `<th colspan="3" class="text-center fw-bold">${s["Lab No."]}<br><small class="text-muted">Input | Final | Status</small></th>`);
    tableHTML += '</tr></thead><tbody>';
    for (const test of tests) {
        tableHTML += `<tr><td class="fw-bold">${test.name}:</td>`;
        for (const sample of sampleDetails) {
            const labNo = sample["Lab No."];
            const docId = labNo.replace('/', '-');
            let value = '';
            try {
                const doc = await db.collection('samples').doc(docId).get();
                if (doc.exists) {
                    const data = doc.data();
                    // Fetch TDS, TH, Ca, Mg, Chl, Alk (lowercase keys as per JalGanana)
                    const keyMap = {
                        "TDS": "tds",
                        "T. Hardness": "t_hardness",
                        "Calcium": "calcium",
                        "Magnesium": "magnesium",
                        "Chloride": "chloride",
                        "Alkalinity": "alkalinity"
                    };
                    value = data[keyMap[test.name]] || data[test.name.toLowerCase().replace('.', '_')] || '';
                } else {
                    setStatus(`No data in Firebase for ${labNo}. Manual entry enabled.`, "warning");
                }
            } catch (err) {
                console.error(err);
                setStatus(`Fetch error for ${labNo}.`, "error");
            }
            // Prefilled defaults (editable)
            if (test.name === "Colour") value = "Clear";
            else if (test.name === "Odour") value = "OK";
            else if (test.name === "Turbidity") value = "NO";
            else if (test.name === "pH" && !value) value = '';
            tableHTML += `<td><input type="text" class="form-control chemical-input d-inline-block me-1" data-lab="${labNo}" data-test="${test.name}" value="${value}" oninput="updateFinalAndStatus('${labNo}', '${test.name}', this.value)"></td>
                          <td><input type="text" class="form-control chemical-final d-inline-block me-1" data-lab="${labNo}" data-test="${test.name}" readonly value="${value}"></td>
                          <td><span class="chemical-status fw-bold d-inline-block" data-lab="${labNo}" data-test="${test.name}"></span></td>`;
            updateFinalAndStatus(labNo, test.name, value);
        }
        tableHTML += '</tr>';
    }
    tableHTML += '</tbody></table>';
    container.innerHTML = tableHTML;
    setStatus("Chemical Results loaded. Fetched from Firebase (TDS, TH, Ca, Mg, Chloride, Alkalinity). Defaults set for Colour, Odour, Turbidity, pH. Edit and check status.", "success");
}

function updateFinalAndStatus(labNo, testName, value) {
    const finalEl = document.querySelector(`.chemical-final[data-lab="${labNo}"][data-test="${testName}"]`);
    const statusEl = document.querySelector(`.chemical-status[data-lab="${labNo}"][data-test="${testName}"]`);
    if (finalEl) finalEl.value = value;
    const test = tests.find(t => t.name === testName);
    if (!test) return;
    let category = categorizeSample(testName, value, test.max_limit, test.desirable_limit);
    if (["TDS", "pH", "T. Hardness", "Calcium", "Magnesium", "Chloride", "Alkalinity"].includes(testName) && value && !validateNumber(value)) {
        category = "invalid";
    }
    if (statusEl) {
        const statusText = category === "suitable" ? "✅ Desirable" : category === "permissible" ? "⚠️ Permissible" : category === "unsuitable" ? "❌ Failed" : "Invalid Number";
        statusEl.textContent = statusText;
        statusEl.className = `chemical-status fw-bold ${category === "unsuitable" || category === "invalid" ? "text-danger" : category === "permissible" ? "text-warning" : "text-success"}`;
    }
}

function clearChemicalForm() {
    populateChemicalResultsTab();
    setStatus("Chemical results cleared. Defaults and Firebase data restored.", "info");
}

function submitChemicalResults() {
    if (!confirm("Submit results and go to preview?")) return;
    chemicalResults = [];
    let allValid = true;
    sampleDetails.forEach(sample => {
        const labNo = sample["Lab No."];
        const entries = {};
        let valid = true;
        document.querySelectorAll(`.chemical-input[data-lab="${labNo}"]`).forEach(input => {
            const testName = input.dataset.test;
            const value = input.value.trim();
            entries[testName] = value;
            if (!value) valid = false;
            if (["TDS", "pH", "T. Hardness", "Calcium", "Magnesium", "Chloride", "Alkalinity"].includes(testName) && value && !validateNumber(value)) valid = false;
        });
        if (valid) chemicalResults.push({ "Lab No.": labNo, ...entries });
        else allValid = false;
    });
    if (!allValid) return setStatus("Fix invalid/empty fields in chemical results.", "error");
    populatePreviewTab();
    new bootstrap.Tab(document.querySelector('#preview-tab')).show();
    setStatus("Results submitted. Preview ready.", "success");
}

// Query Functions
async function fetchByLabNo() {
    const labNo = document.getElementById('query-lab-no').value.trim();
    if (!validateLabNo(labNo)) return setStatus("Enter valid Lab No. (e.g., 123/2025).", "warning");
    const docId = labNo.replace('/', '-');
    try {
        const doc = await db.collection('samples').doc(docId).get();
        queryResults = doc.exists ? [doc.data()] : [];
        renderQueryTable();
        setStatus(queryResults.length ? `Found result for ${labNo}.` : `No data for ${labNo}.`, queryResults.length ? "success" : "warning");
    } catch (err) {
        setStatus(`Query error: ${err.message}`, "error");
    }
}

async function fetchBySentBy() {
    const sentBy = document.getElementById('query-sent-by').value.trim();
    if (!sentBy) return setStatus("Enter Sent By.", "warning");
    try {
        const snapshot = await db.collection('samples').where('Sender', '==', sentBy).get();
        queryResults = snapshot.empty ? [] : snapshot.docs.map(d => d.data());
        renderQueryTable();
        setStatus(`Found ${queryResults.length} results for "${sentBy}".`, "success");
    } catch (err) {
        setStatus(`Error: ${err.message}`, "error");
    }
}

async function fetchBySentByLocation() {
    const sentBy = document.getElementById('query-sent-by').value.trim();
    const location = document.getElementById('query-location').value.trim();
    if (!sentBy || !location) return setStatus("Enter both Sent By and Location.", "warning");
    try {
        const snapshot = await db.collection('samples').where('Sender', '==', sentBy).where('Location', '==', location).get();
        queryResults = snapshot.empty ? [] : snapshot.docs.map(d => d.data());
        renderQueryTable();
        setStatus(`Found ${queryResults.length} results for "${sentBy}" + "${location}".`, "success");
    } catch (err) {
        setStatus(`Error: ${err.message}`, "error");
    }
}

function renderQueryTable() {
    const tbody = document.querySelector('#query-table tbody');
    tbody.innerHTML = '';
    if (queryResults.length === 0) return;
    const columns = ["Lab No.", "Source", "Location", "CHI Sample No.", "Date", "Sender", "Colour", "Odour", "Turbidity", "TDS", "pH", "T. Hardness", "Calcium", "Magnesium", "Chloride", "Alkalinity"];
    queryResults.forEach(result => {
        const row = tbody.insertRow();
        columns.forEach(col => {
            const cell = row.insertCell();
            cell.textContent = result[col] || '-';
            cell.classList.add('text-center');
        });
    });
}

function generateQueryPdf() {
    if (queryResults.length === 0) return setStatus("No results to export.", "warning");
    const win = window.open('', '_blank');
    win.document.write(`
        <html><head><title>Query Report - ${new Date().toLocaleDateString()}</title>
        <style>body { font-family: Arial; } table { border-collapse: collapse; width: 100%; } th, td { border: 1px solid black; padding: 8px; text-align: center; } th { background-color: #f2f2f2; }</style>
        </head><body><h2>Database Query Report</h2><p>Generated on: ${new Date().toLocaleDateString()}</p>${document.getElementById('query-table').outerHTML}<script>window.print();</script></body></html>
    `);
    win.document.close();
    setStatus("PDF export opened. Use browser print to save as PDF.", "success");
}

// Preview Functions
function populatePreviewTab() {
    // Sample Particulars
    const sampleTbody = document.querySelector('#sample-preview-table tbody');
    sampleTbody.innerHTML = '';
    const sampleHeaders = ["क्र.सं.", "विवरण"].concat(sampleDetails.map((_, i) => `(${i+1})`));
    document.querySelector('#sample-preview-table thead tr').innerHTML = sampleHeaders.map(h => `<th>${h}</th>`).join('');
    const sampleRows = [
        ["1.1", "स्रोत (Source)"].concat(sampleDetails.map(s => s.Source)),
        ["1.2", "स्थान (Location)"].concat(sampleDetails.map(s => s.Location)),
        ["1.3", "मुख्य नि. नमूने की संख्या (CHI Sample No.)"].concat(sampleDetails.map(s => s["CHI Sample No."])),
        ["1.4", "नमूना संग्रह की तारीख (Date)"].concat(sampleDetails.map(s => s.Date)),
        ["1.5", "प्रयोगशाला संख्या (Lab No.)"].concat(sampleDetails.map(s => s["Lab No."]))
    ];
    sampleRows.forEach(row => {
        const tr = sampleTbody.insertRow();
        row.forEach(cell => {
            const td = tr.insertCell();
            td.textContent = cell;
            td.classList.add('text-center');
        });
    });

    // Chemical Analysis
    const chemicalTbody = document.querySelector('#chemical-preview-table tbody');
    chemicalTbody.innerHTML = '';
    const chemHeaders = ["क.सं.", "परीक्षण (Tests)", "निर्धारित मान (Max)", "निर्धारित मान (Desirable)"].concat(sampleDetails.map(s => s["Lab No."]));
    document.querySelector('#chemical-preview-table thead tr').innerHTML = chemHeaders.map(h => `<th>${h}</th>`).join('');
    tests.forEach((test, i) => {
        const tr = chemicalTbody.insertRow();
        tr.insertCell().textContent = `2.${i+1}`;
        tr.insertCell().textContent = `${test.name} (${test.bilingual_name})`;
        tr.insertCell().textContent = test.max_limit;
        tr.insertCell().textContent = test.desirable_limit;
        chemicalResults.forEach(r => {
            const td = tr.insertCell();
            td.textContent = r[test.name] || '-';
            td.classList.add('text-center');
        });
    });
    setStatus("Report preview updated. Check tables before generating DOCX.", "success");
}

function backToChemical() {
    new bootstrap.Tab(document.querySelector('#chemical-tab')).show();
    setStatus("Back to Chemical Results.", "info");
}

// Final Report
async function generateFinalReport() {
    if (!confirm("Generate final DOCX report? Data will be saved to Firebase (with overwrite prompt).")) return;
    // Save to Firebase
    for (let i = 0; i < sampleDetails.length; i++) {
        const sample = sampleDetails[i];
        const chemical = chemicalResults[i];
        const labNo = sample["Lab No."];
        const docId = labNo.replace('/', '-');
        const docRef = db.collection('samples').doc(docId);
        try {
            const existing = await docRef.get();
            if (existing.exists) {
                if (!confirm(`Lab No. ${labNo} exists. Overwrite?`)) continue;
            }
            await docRef.set({ ...sample, ...chemical });
        } catch (err) {
            setStatus(`Error saving ${labNo}: ${err.message}`, "error");
            return;
        }
    }
    setStatus("Data saved to Firebase. Generating DOCX...", "info");

    // DOCX (exact format)
    const { Document, Packer, Paragraph, Table, TableRow, TableCell, TextRun, AlignmentType } = docx;
    const doc = new Document({
        sections: [{
            properties: { page: { margin: { top: 720, bottom: 720, left: 1080, right: 1080 } } },
            children: [
                new Paragraph({ children: [new TextRun({ text: "उत्तर पश्चिम रेलवे", bold: true, size: 24, font: "Times New Roman" })] , alignment: AlignmentType.CENTER }),
                new Paragraph({ children: [new TextRun({ text: "कार्यालय\nउप मु.रसा.एवं धातुज्ञ\nकेन्द्रीय प्रयोगशाला, कैरिज, अजमेर", size: 18, font: "Times New Roman" })] , alignment: AlignmentType.RIGHT }),
                new Paragraph({ children: [new TextRun({ text: `संख्याः सी.एंड एम./सीएल/एफएलडब्ल्यू/वाटर/${formatLabNoRange()}                                                        दिनांक: ${chiDetails.reportDate}`, size: 18, font: "Times New Roman" })] }),
                new Paragraph({ children: [new TextRun({ text: `${chiDetails.address}`, bold: true, size: 18, font: "Times New Roman" })] }),
                new Paragraph({ children: [new TextRun({ text: "\t  विषय: पेयजल का रसायनिक विश्लेषण।", size: 18, font: "Times New Roman" })] }),
                new Paragraph({ children: [new TextRun({ text: `\t  संदर्भ: ${chiDetails.address} का पत्र संख्या ${chiDetails.letterNo}`, size: 18, font: "Times New Roman" })] }),
                new Paragraph({ children: [new TextRun({ text: "(1) नमूना विवरण (Sample Particulars)", bold: true, size: 18, font: "Times New Roman" })] }),
                createSampleDocxTable(),
                new Paragraph({ children: [new TextRun({ text: "(2) रसायनिक विश्लेषण (Chemical Analysis)", bold: true, size: 18, font: "Times New Roman" })] }),
                createChemicalDocxTable(),
                new Paragraph({ children: [new TextRun({ text: "टिप्पणी:", size: 18, font: "Times New Roman" })] }),
                ...generateRemarksDocx(),
                new Paragraph({ children: [new TextRun({ text: "\nरसायन एवं धातुकर्म अधीक्षक (एफएलडब्ल्यू)\nकेंद्रीय प्रयोगशाला, उ.प.रे., अजमेर", bold: true, size: 18, font: "Times New Roman" })] , alignment: AlignmentType.RIGHT }),
                new Paragraph({ children: [new TextRun({ text: `प्रतिलिपी: आवश्यक कार्यवाही हेतु - मंडल चिकित्सा अधिकारी (स्वास्थ्य)/${chiDetails.division}`, size: 18, font: "Times New Roman" })] })
            ]
        }]
    });
    Packer.toBlob(doc).then(blob => {
        saveAs(blob, `water_${formatLabNoRange().replace('/', '_')}_${new Date().toISOString().slice(0,10).replace(/-/g,'')}.docx`);
        setStatus("DOCX report generated and downloaded successfully!", "success");
    }).catch(err => setStatus(`DOCX generation error: ${err.message}`, "error"));
}

function formatLabNoRange() {
    if (sampleDetails.length === 0) return "N/A";
    const prefixes = sampleDetails.map(s => parseInt(s["Lab No."].split('/')[0]));
    const year = sampleDetails[0]["Lab No."].split('/')[1];
    const min = Math.min(...prefixes);
    const max = Math.max(...prefixes);
    return min === max ? `${min}/${year}` : `${min}-${max}/${year}`;
}

function createSampleDocxTable() {
    const headers = ["क्र.सं.", "विवरण"].concat(sampleDetails.map((_, i) => `(${i+1})`));
    const rows = [new TableRow({ children: headers.map(h => new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: h, size: 18 })] })] })) })];
    const dataRows = [
        ["1.1", "स्रोत (Source)"].concat(sampleDetails.map(s => s.Source)),
        ["1.2", "स्थान (Location)"].concat(sampleDetails.map(s => s.Location)),
        ["1.3", "मुख्य नि. नमूने की संख्या (CHI Sample No.)"].concat(sampleDetails.map(s => s["CHI Sample No."])),
        ["1.4", "नमूना संग्रह की तारीख (Date)"].concat(sampleDetails.map(s => s.Date)),
        ["1.5", "प्रयोगशाला संख्या (Lab No.)"].concat(sampleDetails.map(s => s["Lab No."]))
    ];
    dataRows.forEach(row => {
        rows.push(new TableRow({ children: row.map(cell => new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: cell, size: 18 })] })] })) }));
    });
    return new Table({ rows });
}

function createChemicalDocxTable() {
    // Header with merging
    const header1 = new TableRow({ children: [
        new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "क.सं.", bold: true, size: 18 })] })] }),
        new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "परीक्षण (Tests)", bold: true, size: 18 })] })] }),
        new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "निर्दिष्ट मान (Specified IS: 10500-2012)", bold: true, size: 18 })] })] , columnSpan: 2 }),
        new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "परिणाम (Results)", bold: true, size
