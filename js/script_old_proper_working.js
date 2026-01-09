// js/script.js - Modular Water Quality Report App (JalGanana style)
import { getFirestore, collection, doc, getDoc, setDoc, query, where, getDocs } from 'https://www.gstatic.com/firebasejs/10.13.2/firebase-firestore.js';

// Access global DBs from index.html
const totalWaterDb = window.totalWaterDb; // Total-Water Firebase
const jalGananaDb = window.jalGananaDb;   // JalGanana Firebase

const tests = [
    { name: "Colour", bilingual_name: "à¤°à¤‚à¤—", max_limit: "-", desirable_limit: "-" },
    { name: "Odour", bilingual_name: "à¤—à¤‚à¤§", max_limit: "-", desirable_limit: "-" },
    { name: "Turbidity", bilingual_name: "à¤Ÿà¤°à¥à¤¬à¤¿à¤¡à¤¿à¤Ÿà¥€", max_limit: "-", desirable_limit: "-" },
    { name: "TDS", bilingual_name: "à¤Ÿà¥€à¤¡à¥€à¤à¤¸", max_limit: "2000 mg/l", desirable_limit: "500 mg/l" },
    { name: "pH", bilingual_name: "à¤ªà¥€à¤à¤š", max_limit: "6.5 à¤¸à¥‡ 8.5", desirable_limit: "6.5 à¤¸à¥‡ 8.5" },
    { name: "T. Hardness", bilingual_name: "à¤•à¥à¤² à¤•à¤ à¥‹à¤°à¤¤à¤¾", max_limit: "600 mg/l", desirable_limit: "200 mg/l" },
    { name: "Calcium", bilingual_name: "à¤•à¥ˆà¤²à¥à¤¶à¤¿à¤¯à¤®", max_limit: "200 mg/l", desirable_limit: "75 mg/l" },
    { name: "Magnesium", bilingual_name: "à¤®à¥ˆà¤—à¥à¤¨à¥€à¤¶à¤¿à¤¯à¤®", max_limit: "100 mg/l", desirable_limit: "30 mg/l" },
    { name: "Chloride", bilingual_name: "à¤•à¥à¤²à¥‹à¤°à¤¾à¤‡à¤¡", max_limit: "1000 mg/l", desirable_limit: "250 mg/l" },
    { name: "Alkalinity", bilingual_name: "à¤•à¥à¤·à¤¾à¤°à¥€à¤¯à¤¤à¤¾", max_limit: "600 mg/l", desirable_limit: "200 mg/l" }
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
            const [maxLow, maxHigh] = maxLimit.split(' à¤¸à¥‡ ').map(parseFloat);
            const [desLow, desHigh] = desLimit.split(' à¤¸à¥‡ ').map(parseFloat);
            if (val < maxLow || val > maxHigh) return "unsuitable";
            if (val < desLow || val > desHigh) return "permissible";
            return "suitable";
        } else {
            const maxVal = parseFloat(maxLimit.replace(/ mg\/l| -/g, '')) || Infinity;
            const desVal = parseFloat(desLimit.replace(/ mg\/l| -/g, '')) || Infinity;
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

// Status Update Function
function setStatus(message, type = "info") {
    const statusDiv = document.getElementById('status-var');
    statusDiv.innerHTML = `<div class="alert alert-${type}">${message}</div>`;
}

// CHI Functions
function loadChiCsv() {
    const file = document.getElementById('chi-csv').files[0];
    if (!file) return setStatus("à¤•à¥ƒà¤ªà¤¯à¤¾ CSV à¤«à¤¼à¤¾à¤‡à¤² à¤šà¥à¤¨à¥‡à¤‚à¥¤", "warning");
    const reader = new FileReader();
    reader.onload = (e) => {
        const lines = e.target.result.split('\n').filter(line => line.trim());
        if (lines.length < 2) return setStatus("CSV à¤–à¤¾à¤²à¥€ à¤¯à¤¾ à¤…à¤®à¤¾à¤¨à¥à¤¯ à¤¹à¥ˆà¥¤", "danger");
        const headers = lines[0].split(',');
        const row = lines[1].split(',');
        const data = {};
        headers.forEach((h, i) => data[h.trim()] = row[i]?.trim());
        const allowedDivisions = ["à¤…à¤œà¤®à¥‡à¤°", "à¤œà¥‹à¤§à¤ªà¥à¤°", "à¤œà¤¯à¤ªà¥à¤°", "à¤¬à¥€à¤•à¤¾à¤¨à¥‡à¤°"];
        if (!data["CHI Letter No."] || !data["CHI Address"] || !allowedDivisions.includes(data["Division"]) || !validateDate(data["Report Date"])) {
            return setStatus("à¤…à¤®à¤¾à¤¨à¥à¤¯ CSV à¤¡à¥‡à¤Ÿà¤¾à¥¤ à¤•à¥‰à¤²à¤® à¤šà¥‡à¤• à¤•à¤°à¥‡à¤‚: CHI Letter No., CHI Address, Division, Report Dateà¥¤", "danger");
        }
        document.getElementById('chi-letter-no').value = data["CHI Letter No."];
        document.getElementById('chi-address').value = data["CHI Address"];
        document.getElementById('chi-division').value = data["Division"];
        document.getElementById('report-date').value = data["Report Date"];
        setStatus("CHI à¤µà¤¿à¤µà¤°à¤£ CSV à¤¸à¥‡ à¤²à¥‹à¤¡ à¤¹à¥‹ à¤—à¤à¥¤", "success");
    };
    reader.readAsText(file, 'UTF-8');
}

function clearChiForm() {
    document.getElementById('chi-letter-no').value = '';
    document.getElementById('chi-address').value = '';
    document.getElementById('chi-division').value = '';
    document.getElementById('report-date').value = new Date().toLocaleDateString('en-GB').split('/').reverse().join('-');
    document.getElementById('chi-csv').value = '';
    setStatus("CHI à¤«à¥‰à¤°à¥à¤® à¤¸à¤¾à¤«à¤¼ à¤•à¤¿à¤¯à¤¾ à¤—à¤¯à¤¾à¥¤", "info");
}

function openSampleTab() {
    chiDetails = {
        letterNo: document.getElementById('chi-letter-no').value.trim(),
        address: document.getElementById('chi-address').value.trim(),
        division: document.getElementById('chi-division').value,
        reportDate: document.getElementById('report-date').value.trim()
    };
    const allowedDivisions = ["à¤…à¤œà¤®à¥‡à¤°", "à¤œà¥‹à¤§à¤ªà¥à¤°", "à¤œà¤¯à¤ªà¥à¤°", "à¤¬à¥€à¤•à¤¾à¤¨à¥‡à¤°"];
    if (!chiDetails.letterNo || !chiDetails.address || !allowedDivisions.includes(chiDetails.division) || !validateDate(chiDetails.reportDate)) {
        return setStatus("à¤•à¥ƒà¤ªà¤¯à¤¾ à¤µà¥ˆà¤§ CHI Letter No., Address, Division, à¤”à¤° Report Date (DD-MM-YYYY) à¤¦à¤°à¥à¤œ à¤•à¤°à¥‡à¤‚à¥¤", "danger");
    }
    new bootstrap.Tab(document.querySelector('#sample-tab')).show();
    setStatus("Sample Details à¤Ÿà¥ˆà¤¬ à¤ªà¤° à¤—à¤à¥¤", "success");
}

// Sample Functions
function addSampleEntry(sample = { Source: '', Location: '', 'CHI Sample No.': '', Date: '', 'Lab No.': '', Sender: '' }, index = sampleDetails.length) {
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
                    <input type="text" class="form-control sample-source" value="${sample.Source || ''}" placeholder="Enter Source">
                </div>
                <div class="col-md-2">
                    <label class="form-label small fw-bold">Location</label>
                    <input type="text" class="form-control sample-location" value="${sample.Location || ''}" placeholder="Enter Location">
                </div>
                <div class="col-md-2">
                    <label class="form-label small fw-bold">CHI Sample No.</label>
                    <input type="text" class="form-control sample-chi-sample-no" value="${sample['CHI Sample No.'] || ''}" placeholder="Enter CHI No.">
                </div>
                <div class="col-md-2">
                    <label class="form-label small fw-bold">Date (DD/MM/YYYY)</label>
                    <input type="text" class="form-control sample-date" value="${sample.Date || ''}" placeholder="DD/MM/YYYY">
                </div>
                <div class="col-md-2">
                    <label class="form-label small fw-bold">Lab No. (123/2025)</label>
                    <input type="text" class="form-control sample-lab-no" value="${sample['Lab No.'] || ''}" placeholder="123/2025">
                </div>
                <div class="col-md-2">
                    <label class="form-label small fw-bold">Sender</label>
                    <input type="text" class="form-control sample-sender" value="${sample.Sender || ''}" placeholder="Enter Sender">
                </div>
            </div>
        </div>
        <div class="col-md-2 text-end">
            <button type="button" class="btn btn-danger delete-btn p-2" data-index="${index}" title="Delete Sample ${index + 1}">
                <i class="bi bi-x-circle-fill" style="font-size: 1.5rem; color: red;"></i>
            </button>
        </div>
    `;
    container.appendChild(div);
    sampleDetails[index] = { Source: sample.Source, Location: sample.Location, 'CHI Sample No.': sample['CHI Sample No.'], Date: sample.Date, 'Lab No.': sample['Lab No.'], Sender: sample.Sender };
}

function addSamplesFromNum() {
    const num = parseInt(document.getElementById('num-samples').value);
    if (isNaN(num) || num < 1 || num > 20) return setStatus("1-20 à¤•à¥‡ à¤¬à¥€à¤š à¤¨à¤‚à¤¬à¤° à¤¦à¤°à¥à¤œ à¤•à¤°à¥‡à¤‚à¥¤", "warning");
    for (let i = sampleDetails.length; i < sampleDetails.length + num; i++) {
        addSampleEntry({}, i);
    }
    setStatus(`${num} à¤¸à¥ˆà¤‚à¤ªà¤² à¤œà¥‹à¤¡à¤¼à¥‡ à¤—à¤à¥¤ à¤†à¤ª à¤‡à¤¨à¥à¤¹à¥‡à¤‚ à¤à¤¡à¤¿à¤Ÿ à¤¯à¤¾ à¤¡à¤¿à¤²à¥€à¤Ÿ à¤•à¤° à¤¸à¤•à¤¤à¥‡ à¤¹à¥ˆà¤‚à¥¤`, "success");
}

function loadSampleCsv() {
    const file = document.getElementById('sample-csv').files[0];
    if (!file) return setStatus("CSV à¤«à¤¼à¤¾à¤‡à¤² à¤šà¥à¤¨à¥‡à¤‚à¥¤", "warning");
    const reader = new FileReader();
    reader.onload = (e) => {
        const lines = e.target.result.split('\n').filter(l => l.trim());
        if (lines.length < 2) return setStatus("à¤…à¤®à¤¾à¤¨à¥à¤¯ CSVà¥¤", "danger");
        const headers = lines[0].split(',');
        const required = ["Source", "Location", "CHI Sample No.", "Date", "Lab No.", "Sender"];
        if (!required.every(h => headers.some(head => head.trim() === h))) return setStatus("CSV à¤®à¥‡à¤‚ à¤¯à¥‡ à¤•à¥‰à¤²à¤® à¤¹à¥‹à¤¨à¥‡ à¤šà¤¾à¤¹à¤¿à¤: " + required.join(", "), "danger");
        const maxLoaded = document.getElementById('num-samples').value ? parseInt(document.getElementById('num-samples').value) : lines.length - 1;
        sampleDetails = [];
        renderSampleEntries();
        let loaded = 0;
        for (let i = 1; i < lines.length && loaded < maxLoaded; i++) {
            const row = lines[i].split(',');
            const data = {};
            headers.forEach((h, j) => data[h.trim()] = row[j]?.trim());
            if (validateDate(data.Date) && validateLabNo(data["Lab No."])) {
                addSampleEntry({
                    Source: data["Source"] || '',
                    Location: data["Location"] || '',
                    'CHI Sample No.': data["CHI Sample No."] || '',
                    Date: data.Date || '',
                    'Lab No.': data["Lab No."] || '',
                    Sender: data.Sender || ''
                }, loaded);
                loaded++;
            }
        }
        setStatus(`${loaded} à¤¸à¥ˆà¤‚à¤ªà¤² CSV à¤¸à¥‡ à¤²à¥‹à¤¡ à¤¹à¥à¤ (à¤…à¤§à¤¿à¤•à¤¤à¤® ${maxLoaded})à¥¤ à¤†à¤ª à¤‡à¤¨à¥à¤¹à¥‡à¤‚ à¤à¤¡à¤¿à¤Ÿ à¤¯à¤¾ à¤¡à¤¿à¤²à¥€à¤Ÿ à¤•à¤° à¤¸à¤•à¤¤à¥‡ à¤¹à¥ˆà¤‚à¥¤`, "success");
    };
    reader.readAsText(file, 'UTF-8');
}

function deleteSampleEntry(index) {
    if (confirm(`Sample ${index + 1} à¤¡à¤¿à¤²à¥€à¤Ÿ à¤•à¤°à¥‡à¤‚?`)) {
        sampleDetails.splice(index, 1);
        renderSampleEntries();
        setStatus(`Sample ${index + 1} à¤¡à¤¿à¤²à¥€à¤Ÿ à¤¹à¥‹ à¤—à¤¯à¤¾à¥¤ à¤•à¥à¤² à¤¸à¥ˆà¤‚à¤ªà¤² à¤…à¤¬: ${sampleDetails.length}à¥¤ à¤¯à¤¹ Chemical à¤Ÿà¥‡à¤¬à¤² à¤®à¥‡à¤‚ à¤¦à¤¿à¤–à¥‡à¤—à¤¾à¥¤`, "warning");
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
    setStatus("à¤¸à¥ˆà¤‚à¤ªà¤² à¤«à¥‰à¤°à¥à¤® à¤¸à¤¾à¤«à¤¼ à¤•à¤¿à¤¯à¤¾ à¤—à¤¯à¤¾à¥¤", "info");
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
            sampleDetails.push({ Source: source, Location: location, 'CHI Sample No.': chiSampleNo, Date: date, 'Lab No.': labNo, Sender: sender });
            valid++;
        }
    });
    if (sampleDetails.length === 0) return setStatus("à¤•à¥‹à¤ˆ à¤µà¥ˆà¤§ à¤¸à¥ˆà¤‚à¤ªà¤² à¤¨à¤¹à¥€à¤‚à¥¤ à¤¸à¤­à¥€ à¤«à¤¼à¥€à¤²à¥à¤¡ à¤­à¤°à¥‡à¤‚ à¤”à¤° à¤«à¥‰à¤°à¥à¤®à¥‡à¤Ÿ à¤šà¥‡à¤• à¤•à¤°à¥‡à¤‚à¥¤", "danger");
    populateChemicalResultsTab();
    new bootstrap.Tab(document.querySelector('#chemical-tab')).show();
    setStatus(`${valid} à¤µà¥ˆà¤§ à¤¸à¥ˆà¤‚à¤ªà¤² à¤ªà¥à¤°à¥‹à¤¸à¥‡à¤¸ à¤¹à¥à¤à¥¤ Chemical Results à¤Ÿà¥ˆà¤¬ à¤ªà¤° à¤—à¤à¥¤`, "success");
}

// Chemical Functions
async function populateChemicalResultsTab() {
    const container = document.getElementById('chemical-table-container');
    container.innerHTML = '<div class="text-center p-3"><div class="spinner-border" role="status"><span class="visually-hidden">Loading...</span></div></div>';
    let tableHTML = '<table class="table table-bordered table-sm"><thead><tr><th>Test</th>';
    sampleDetails.forEach(s => tableHTML += `<th colspan="3" class="text-center fw-bold">${s['Lab No.']}<br><small class="text-muted">Input | Final | Status</small></th>`);
    tableHTML += '</tr></thead><tbody>';
    for (const test of tests) {
        tableHTML += `<tr><td class="fw-bold">${test.name}:</td>`;
        for (const sample of sampleDetails) {
            const labNo = sample['Lab No.'];
            const docId = labNo.replace('/', '-');
            let value = '';
            try {
                // JalGanana (labcalc-cee5c) à¤¸à¥‡ lab_calculations à¤•à¤²à¥‡à¤•à¥à¤¶à¤¨ à¤¸à¥‡ à¤¡à¥‡à¤Ÿà¤¾ fetch
                const docRef = doc(collection(jalGananaDb, 'lab_calculations'), docId);
                const docSnap = await getDoc(docRef);
                if (docSnap.exists()) {
                    const data = docSnap.data();
                    const keyMap = {
                        "TDS": "tds",
                        "T. Hardness": "th",
                        "Calcium": "ca",
                        "Magnesium": "mg",
                        "Chloride": "chl",
                        "Alkalinity": "alk",
                        "pH": "ph" // pH à¤•à¥‡ à¤²à¤¿à¤ key à¤œà¥‹à¤¡à¤¼à¤¾, à¤…à¤—à¤° JalGanana à¤®à¥‡à¤‚ ph à¤•à¥‡ à¤²à¤¿à¤ à¤…à¤²à¤— key à¤¹à¥‹ à¤¤à¥‹ à¤…à¤ªà¤¡à¥‡à¤Ÿ à¤•à¤°à¥‹
                    };
                    value = data[keyMap[test.name]] || '';
                    // Whole number à¤®à¥‡à¤‚ à¤•à¤¨à¥à¤µà¤°à¥à¤Ÿ à¤•à¤°à¥‡à¤‚
                    if (value && ["TDS", "T. Hardness", "Calcium", "Magnesium", "Chloride", "Alkalinity"].includes(test.name)) {
                        value = Math.round(parseFloat(value)).toString();
                    }
                } else {
                    setStatus(`JalGanana (lab_calculations) à¤®à¥‡à¤‚ ${labNo} à¤•à¥‡ à¤²à¤¿à¤ à¤¡à¥‡à¤Ÿà¤¾ à¤¨à¤¹à¥€à¤‚ à¤®à¤¿à¤²à¤¾à¥¤ à¤®à¥ˆà¤¨à¥à¤…à¤² à¤à¤‚à¤Ÿà¥à¤°à¥€ à¤•à¤°à¥‡à¤‚à¥¤`, "warning");
                }
            } catch (err) {
                console.error(err);
                setStatus(`${labNo} à¤•à¥‡ à¤²à¤¿à¤ JalGanana à¤¸à¥‡ fetch error: ${err.message}. labcalc-cee5c à¤•à¥‡ Firebase permissions à¤šà¥‡à¤• à¤•à¤°à¥‡à¤‚à¥¤`, "danger");
            }
            // à¤¡à¤¿à¤«à¥‰à¤²à¥à¤Ÿ à¤µà¥ˆà¤²à¥à¤¯à¥‚à¤œ à¤¸à¥‡à¤Ÿ à¤•à¤°à¥‡à¤‚
            if (test.name === "Colour") value = value || "Clear";
            else if (test.name === "Odour") value = value || "OK";
            else if (test.name === "Turbidity") value = value || "NO";
            else if (test.name === "pH") value = value || "8.0"; // pH à¤¡à¤¿à¤«à¥‰à¤²à¥à¤Ÿ 8.0
            tableHTML += `<td><input type="text" class="form-control chemical-input d-inline-block me-1" data-lab="${labNo}" data-test="${test.name}" value="${value}" oninput="updateFinalAndStatus('${labNo}', '${test.name}', this.value)"></td>
                          <td><input type="text" class="form-control chemical-final d-inline-block me-1" data-lab="${labNo}" data-test="${test.name}" readonly value="${value}"></td>
                          <td><span class="chemical-status fw-bold d-inline-block" data-lab="${labNo}" data-test="${test.name}"></span></td>`;
            updateFinalAndStatus(labNo, test.name, value);
        }
        tableHTML += '</tr>';
    }
    tableHTML += '</tbody></table>';
    container.innerHTML = tableHTML;
    setStatus(`Chemical Results à¤²à¥‹à¤¡ à¤¹à¥‹ à¤—à¤à¥¤ JalGanana (labcalc-cee5c, lab_calculations) à¤¸à¥‡ TDS (${sampleDetails.map(s => s["Lab No."]).join(', ')}) à¤•à¥‡ à¤²à¤¿à¤ à¤¡à¥‡à¤Ÿà¤¾ à¤²à¤¾à¤ à¤—à¤à¥¤ pH à¤¡à¤¿à¤«à¥‰à¤²à¥à¤Ÿ 8.0 à¤¸à¥‡à¤Ÿà¥¤ Colour, Odour, Turbidity à¤¡à¤¿à¤«à¥‰à¤²à¥à¤Ÿà¥¤ à¤¬à¤¾à¤•à¥€ à¤«à¥€à¤²à¥à¤¡à¥à¤¸ à¤à¤¡à¤¿à¤Ÿ à¤•à¤°à¥‡à¤‚à¥¤`, "success");
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
        const statusText = category === "suitable" ? "âœ… Desirable" : category === "permissible" ? "âš ï¸ Permissible" : category === "unsuitable" ? "âŒ Failed" : "Invalid Number";
        statusEl.textContent = statusText;
        statusEl.className = `chemical-status fw-bold ${category === "unsuitable" || category === "invalid" ? "text-danger" : category === "permissible" ? "text-warning" : "text-success"}`;
    }
}

function clearChemicalForm() {
    populateChemicalResultsTab();
    setStatus("Chemical results à¤¸à¤¾à¤«à¤¼ à¤•à¤¿à¤ à¤—à¤à¥¤ à¤¡à¤¿à¤«à¤¼à¥‰à¤²à¥à¤Ÿ à¤”à¤° JalGanana à¤¡à¥‡à¤Ÿà¤¾ à¤¬à¤¹à¤¾à¤²à¥¤", "info");
}

function submitChemicalResults() {
    if (!confirm("à¤°à¤¿à¤œà¤²à¥à¤Ÿ à¤¸à¤¬à¤®à¤¿à¤Ÿ à¤•à¤°à¥‡à¤‚ à¤”à¤° à¤ªà¥à¤°à¥€à¤µà¥à¤¯à¥‚ à¤¦à¥‡à¤–à¥‡à¤‚?")) return;
    chemicalResults = [];
    let allValid = true;
    sampleDetails.forEach(sample => {
        const labNo = sample['Lab No.'];
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
    if (!allValid) return setStatus("Chemical results à¤®à¥‡à¤‚ à¤…à¤®à¤¾à¤¨à¥à¤¯/à¤–à¤¾à¤²à¥€ à¤«à¤¼à¥€à¤²à¥à¤¡ à¤ à¥€à¤• à¤•à¤°à¥‡à¤‚à¥¤", "danger");
    populatePreviewTab();
    new bootstrap.Tab(document.querySelector('#preview-tab')).show();
    setStatus("à¤°à¤¿à¤œà¤²à¥à¤Ÿ à¤¸à¤¬à¤®à¤¿à¤Ÿ à¤¹à¥‹ à¤—à¤à¥¤ à¤ªà¥à¤°à¥€à¤µà¥à¤¯à¥‚ à¤¤à¥ˆà¤¯à¤¾à¤° à¤¹à¥ˆà¥¤", "success");
}

// Query Functions
async function fetchByLabNo() {
    const labNo = document.getElementById('query-lab-no').value.trim();
    if (!validateLabNo(labNo)) return setStatus("à¤µà¥ˆà¤§ Lab No. à¤¦à¤°à¥à¤œ à¤•à¤°à¥‡à¤‚ (à¤œà¥ˆà¤¸à¥‡, 123/2025)à¥¤", "warning");
    const docId = labNo.replace('/', '-');
    try {
        const docRef = doc(collection(totalWaterDb, 'samples'), docId);
        const docSnap = await getDoc(docRef);
        queryResults = docSnap.exists() ? [docSnap.data()] : [];
        renderQueryTable();
        setStatus(queryResults.length ? `${labNo} à¤•à¥‡ à¤²à¤¿à¤ à¤°à¤¿à¤œà¤²à¥à¤Ÿ à¤®à¤¿à¤²à¤¾à¥¤` : `${labNo} à¤•à¥‡ à¤²à¤¿à¤ à¤¡à¥‡à¤Ÿà¤¾ à¤¨à¤¹à¥€à¤‚ à¤®à¤¿à¤²à¤¾à¥¤`, queryResults.length ? "success" : "warning");
    } catch (err) {
        setStatus(`Query error: ${err.message}. Total-Water Firebase permissions à¤šà¥‡à¤• à¤•à¤°à¥‡à¤‚à¥¤`, "danger");
    }
}

async function fetchBySentBy() {
    const sentBy = document.getElementById('query-sent-by').value.trim();
    if (!sentBy) return setStatus("Sent By à¤¦à¤°à¥à¤œ à¤•à¤°à¥‡à¤‚à¥¤", "warning");
    try {
        const q = query(collection(totalWaterDb, 'samples'), where('Sender', '==', sentBy));
        const snapshot = await getDocs(q);
        queryResults = snapshot.empty ? [] : snapshot.docs.map(d => d.data());
        renderQueryTable();
        setStatus(`"${sentBy}" à¤•à¥‡ à¤²à¤¿à¤ ${queryResults.length} à¤°à¤¿à¤œà¤²à¥à¤Ÿ à¤®à¤¿à¤²à¥‡à¥¤`, "success");
    } catch (err) {
        setStatus(`Error: ${err.message}. Total-Water Firebase permissions à¤šà¥‡à¤• à¤•à¤°à¥‡à¤‚à¥¤`, "danger");
    }
}

async function fetchBySentByLocation() {
    const sentBy = document.getElementById('query-sent-by').value.trim();
    const location = document.getElementById('query-location').value.trim();
    if (!sentBy || !location) return setStatus("Sent By à¤”à¤° Location à¤¦à¥‹à¤¨à¥‹à¤‚ à¤¦à¤°à¥à¤œ à¤•à¤°à¥‡à¤‚à¥¤", "warning");
    try {
        const q = query(collection(totalWaterDb, 'samples'), where('Sender', '==', sentBy), where('Location', '==', location));
        const snapshot = await getDocs(q);
        queryResults = snapshot.empty ? [] : snapshot.docs.map(d => d.data());
        renderQueryTable();
        setStatus(`"${sentBy}" + "${location}" à¤•à¥‡ à¤²à¤¿à¤ ${queryResults.length} à¤°à¤¿à¤œà¤²à¥à¤Ÿ à¤®à¤¿à¤²à¥‡à¥¤`, "success");
    } catch (err) {
        setStatus(`Error: ${err.message}. Total-Water Firebase permissions à¤šà¥‡à¤• à¤•à¤°à¥‡à¤‚à¥¤`, "danger");
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
    if (queryResults.length === 0) return setStatus("à¤à¤•à¥à¤¸à¤ªà¥‹à¤°à¥à¤Ÿ à¤•à¤°à¤¨à¥‡ à¤•à¥‡ à¤²à¤¿à¤ à¤•à¥‹à¤ˆ à¤°à¤¿à¤œà¤²à¥à¤Ÿ à¤¨à¤¹à¥€à¤‚à¥¤", "warning");
    const win = window.open('', '_blank');
    win.document.write(`
        <html><head><title>Query Report - ${new Date().toLocaleDateString()}</title>
        <style>body { font-family: Arial; } table { border-collapse: collapse; width: 100%; } th, td { border: 1px solid black; padding: 8px; text-align: center; } th { background-color: #f2f2f2; }</style>
        </head><body><h2>Database Query Report</h2><p>Generated on: ${new Date().toLocaleDateString()}</p>${document.getElementById('query-table').outerHTML}<script>window.print();</script></body></html>
    `);
    win.document.close();
    setStatus("PDF à¤à¤•à¥à¤¸à¤ªà¥‹à¤°à¥à¤Ÿ à¤–à¥à¤² à¤—à¤¯à¤¾à¥¤ à¤¬à¥à¤°à¤¾à¤‰à¤œà¤¼à¤° à¤ªà¥à¤°à¤¿à¤‚à¤Ÿ à¤¸à¥‡ PDF à¤¸à¥‡à¤µ à¤•à¤°à¥‡à¤‚à¥¤", "success");
}

// Preview Functions
function populatePreviewTab() {
    const sampleTbody = document.querySelector('#sample-preview-table tbody');
    sampleTbody.innerHTML = '';
    const sampleHeaders = ["à¤•à¥à¤°.à¤¸à¤‚.", "à¤µà¤¿à¤µà¤°à¤£"].concat(sampleDetails.map((_, i) => `(${i+1})`));
    document.querySelector('#sample-preview-table thead tr').innerHTML = sampleHeaders.map(h => `<th>${h}</th>`).join('');
    const sampleRows = [
        ["1.1", "à¤¸à¥à¤°à¥‹à¤¤ (Source)"].concat(sampleDetails.map(s => s.Source)),
        ["1.2", "à¤¸à¥à¤¥à¤¾à¤¨ (Location)"].concat(sampleDetails.map(s => s.Location)),
        ["1.3", "à¤¸à¥€à¤à¤šà¤†à¤ˆ à¤¨à¤®à¥‚à¤¨à¤¾ à¤¸à¤‚à¤–à¥à¤¯à¤¾ (CHI Sample No.)"].concat(sampleDetails.map(s => s['CHI Sample No.'])),
        ["1.4", "à¤¨à¤®à¥‚à¤¨à¤¾ à¤¸à¤‚à¤—à¥à¤°à¤¹ à¤•à¥€ à¤¤à¤¾à¤°à¥€à¤– (Date)"].concat(sampleDetails.map(s => s.Date)),
        ["1.5", "à¤ªà¥à¤°à¤¯à¥‹à¤—à¤¶à¤¾à¤²à¤¾ à¤¸à¤‚à¤–à¥à¤¯à¤¾ (Lab No.)"].concat(sampleDetails.map(s => s['Lab No.']))
    ];
    sampleRows.forEach(row => {
        const tr = sampleTbody.insertRow();
        row.forEach(cell => {
            const td = tr.insertCell();
            td.textContent = cell;
            td.classList.add('text-center');
        });
    });

    const chemicalTbody = document.querySelector('#chemical-preview-table tbody');
    chemicalTbody.innerHTML = '';
    const chemHeaders = ["à¤•.à¤¸à¤‚.", "à¤ªà¤°à¥€à¤•à¥à¤·à¤£ (Tests)", "à¤¨à¤¿à¤°à¥à¤§à¤¾à¤°à¤¿à¤¤ à¤®à¤¾à¤¨ (Max)", "à¤¨à¤¿à¤°à¥à¤§à¤¾à¤°à¤¿à¤¤ à¤®à¤¾à¤¨ (Desirable)"].concat(sampleDetails.map(s => s["Lab No."]));
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
    setStatus("à¤°à¤¿à¤ªà¥‹à¤°à¥à¤Ÿ à¤ªà¥à¤°à¥€à¤µà¥à¤¯à¥‚ à¤…à¤ªà¤¡à¥‡à¤Ÿ à¤¹à¥‹ à¤—à¤¯à¤¾à¥¤ DOCX à¤œà¤¨à¤°à¥‡à¤Ÿ à¤•à¤°à¤¨à¥‡ à¤¸à¥‡ à¤ªà¤¹à¤²à¥‡ à¤Ÿà¥‡à¤¬à¤² à¤šà¥‡à¤• à¤•à¤°à¥‡à¤‚à¥¤", "success");
}

function backToChemical() {
    new bootstrap.Tab(document.querySelector('#chemical-tab')).show();
    setStatus("Chemical Results à¤ªà¤° à¤µà¤¾à¤ªà¤¸ à¤—à¤à¥¤", "info");
}

// Final Report
async function generateFinalReport() {
    if (!confirm("à¤…à¤‚à¤¤à¤¿à¤® DOCX à¤°à¤¿à¤ªà¥‹à¤°à¥à¤Ÿ à¤œà¤¨à¤°à¥‡à¤Ÿ à¤•à¤°à¥‡à¤‚? à¤¡à¥‡à¤Ÿà¤¾ Total-Water Firebase à¤®à¥‡à¤‚ à¤¸à¥‡à¤µ à¤¹à¥‹à¤—à¤¾ (overwrite prompt à¤•à¥‡ à¤¸à¤¾à¤¥)à¥¤")) return;
    // Save to Total-Water Firebase
    for (let i = 0; i < sampleDetails.length; i++) {
        const sample = sampleDetails[i];
        const chemical = chemicalResults[i];
        const labNo = sample['Lab No.'];
        const docId = labNo.replace('/', '-');
        const docRef = doc(collection(totalWaterDb, 'samples'), docId);
        try {
            console.log('Attempting to save to Firebase for Lab No.:', labNo);
            const existing = await getDoc(docRef);
            if (existing.exists()) {
                if (!confirm(`Lab No. ${labNo} à¤ªà¤¹à¤²à¥‡ à¤¸à¥‡ à¤®à¥Œà¤œà¥‚à¤¦ à¤¹à¥ˆà¥¤ à¤“à¤µà¤°à¤°à¤¾à¤‡à¤Ÿ à¤•à¤°à¥‡à¤‚?`)) continue;
            }
            await setDoc(docRef, { ...sample, ...chemical, docId: docId });
            console.log('Successfully saved to Firebase for Lab No.:', labNo);
        } catch (err) {
            console.error('Save error for Lab No.', labNo, ':', err);
            setStatus(` ${labNo} à¤•à¥‹ Total-Water à¤®à¥‡à¤‚ à¤¸à¥‡à¤µ à¤•à¤°à¤¨à¥‡ à¤®à¥‡à¤‚ à¤¤à¥à¤°à¥à¤Ÿà¤¿: ${err.message}. Firebase permissions à¤¯à¤¾ auth à¤šà¥‡à¤• à¤•à¤°à¥‡à¤‚à¥¤`, "danger");
            return;
        }
    }
    setStatus("Total-Water Firebase à¤®à¥‡à¤‚ à¤¡à¥‡à¤Ÿà¤¾ à¤¸à¥‡à¤µ à¤¹à¥‹ à¤—à¤¯à¤¾à¥¤ DOCX à¤œà¤¨à¤°à¥‡à¤Ÿ à¤¹à¥‹ à¤°à¤¹à¤¾ à¤¹à¥ˆ...", "info");

    // DOCX Generation
    try {
        if (typeof window.docx === "undefined") {
            setStatus("docx.js load à¤¨à¤¹à¥€à¤‚ à¤¹à¥à¤†à¥¤ Internet connection à¤¯à¤¾ script link à¤šà¥‡à¤• à¤•à¤°à¥‡à¤‚à¥¤", "danger");
            return;
        }
        const {
            Document,
            Packer,
            Paragraph,
            Table,
            TableRow,
            TableCell,
            TextRun,
            AlignmentType
        } = window.docx;
        
        console.log('Starting DOCX generation');
        
        const doc = new Document({
            sections: [{
                properties: { page: { margin: { top: 720, bottom: 720, left: 1080, right: 1080 } } },
                children: [
                    new Paragraph({ children: [new TextRun({ text: "à¤‰à¤¤à¥à¤¤à¤° à¤ªà¤¶à¥à¤šà¤¿à¤® à¤°à¥‡à¤²à¤µà¥‡", bold: true, size: 24, font: "Times New Roman" })] , alignment: AlignmentType.CENTER }),
                    new Paragraph({ children: [new TextRun({ text: "à¤•à¤¾à¤°à¥à¤¯à¤¾à¤²à¤¯", size: 18, font: "Times New Roman" })] , alignment: AlignmentType.RIGHT }),
                    new Paragraph({ children: [new TextRun({ text: "à¤‰à¤ª à¤®à¥. à¤°à¤¸à¤¾. à¤à¤µà¤‚ à¤§à¤¾à¤¤à¥à¤œà¥à¤ž", size: 18, font: "Times New Roman" })] , alignment: AlignmentType.RIGHT }),
                    new Paragraph({ children: [new TextRun({ text: "à¤•à¥‡à¤¨à¥à¤¦à¥à¤°à¥€à¤¯ à¤ªà¥à¤°à¤¯à¥‹à¤—à¤¶à¤¾à¤²à¤¾, à¤•à¥ˆà¤°à¤¿à¤œ, à¤…à¤œà¤®à¥‡à¤°", size: 18, font: "Times New Roman" })] , alignment: AlignmentType.RIGHT }),
                                        
                    new Paragraph({ children: [new TextRun({ text: `à¤¸à¤‚à¤–à¥à¤¯à¤¾à¤ƒ à¤¸à¥€.à¤à¤‚à¤¡ à¤à¤®./à¤¸à¥€à¤à¤²/à¤à¤«à¤à¤²à¤¡à¤¬à¥à¤²à¥à¤¯à¥‚/à¤µà¤¾à¤Ÿà¤°/${formatLabNoRange()}                                                        à¤¦à¤¿à¤¨à¤¾à¤‚à¤•: ${chiDetails.reportDate}`, size: 18, font: "Times New Roman" })] }),
                    
                    
                    new Paragraph({ children: [new TextRun({ text: `${chiDetails.address}`, bold: true, size: 18, font: "Times New Roman" })] }),
                    new Paragraph({ children: [new TextRun({ text: "\t  à¤µà¤¿à¤·à¤¯: à¤ªà¥‡à¤¯à¤œà¤² à¤•à¤¾ à¤°à¤¾à¤¸à¤¾à¤¯à¤¨à¤¿à¤• à¤µà¤¿à¤¶à¥à¤²à¥‡à¤·à¤£à¥¤", size: 18, font: "Times New Roman" })] }),
                    new Paragraph({ children: [new TextRun({ text: `\t  à¤¸à¤‚à¤¦à¤°à¥à¤­: ${chiDetails.address} à¤•à¤¾ à¤ªà¤¤à¥à¤° à¤¸à¤‚à¤–à¥à¤¯à¤¾ ${chiDetails.letterNo}`, size: 18, font: "Times New Roman" })] }),
                    new Paragraph({ children: [new TextRun({ text: "(1) à¤¨à¤®à¥‚à¤¨à¤¾ à¤µà¤¿à¤µà¤°à¤£ (Sample Particulars)", bold: true, size: 18, font: "Times New Roman" })] }),
                    createSampleDocxTable({ Table, TableRow, TableCell, Paragraph, TextRun }),
                    new Paragraph({ children: [new TextRun({ text: "(2) à¤°à¤¾à¤¸à¤¾à¤¯à¤¨à¤¿à¤• à¤µà¤¿à¤¶à¥à¤²à¥‡à¤·à¤£ (Chemical Analysis)", bold: true, size: 18, font: "Times New Roman" })] }),
                    createChemicalDocxTable({ Table, TableRow, TableCell, Paragraph, TextRun }),
                    new Paragraph({ children: [new TextRun({ text: "à¤Ÿà¤¿à¤ªà¥à¤ªà¤£à¥€:", size: 18, font: "Times New Roman" })] }),
                    ...generateRemarksDocx({ Paragraph, TextRun }),
                    
                    new Paragraph({ children: [new TextRun({ text: "\n(à¤…à¤¨à¤¿à¤² à¤•à¥à¤®à¤¾à¤° à¤•à¤°à¥à¤¦à¤®)", bold: true, size: 18, font: "Times New Roman" })] , alignment: AlignmentType.RIGHT }),
                    new Paragraph({ children: [new TextRun({ text: "\nà¤°à¤¸à¤¾à¤¯à¤¨ à¤à¤µà¤‚ à¤§à¤¾à¤¤à¥à¤•à¤°à¥à¤® à¤…à¤§à¥€à¤•à¥à¤·à¤• (à¤à¤«à¤à¤²à¤¡à¤¬à¥à¤²à¥à¤¯à¥‚)", bold: true, size: 18, font: "Times New Roman" })] , alignment: AlignmentType.RIGHT }),
                    new Paragraph({ children: [new TextRun({ text: "\nà¤•à¥‡à¤‚à¤¦à¥à¤°à¥€à¤¯ à¤ªà¥à¤°à¤¯à¥‹à¤—à¤¶à¤¾à¤²à¤¾, à¤‰.à¤ª.à¤°à¥‡., à¤…à¤œà¤®à¥‡à¤°", bold: true, size: 18, font: "Times New Roman" })] , alignment: AlignmentType.RIGHT }),
                    
                    new Paragraph({ children: [new TextRun({ text: `à¤ªà¥à¤°à¤¤à¤¿à¤²à¤¿à¤ªà¥€: à¤†à¤µà¤¶à¥à¤¯à¤• à¤•à¤¾à¤°à¥à¤¯à¤µà¤¾à¤¹à¥€ à¤¹à¥‡à¤¤à¥ - à¤®à¤‚à¤¡à¤² à¤šà¤¿à¤•à¤¿à¤¤à¥à¤¸à¤¾ à¤…à¤§à¤¿à¤•à¤¾à¤°à¥€ (à¤¸à¥à¤µà¤¾à¤¸à¥à¤¥à¥à¤¯)/${chiDetails.division}`, size: 18, font: "Times New Roman" })] })
                ]
            }]
        });
        const blob = await Packer.toBlob(doc);
        saveAs(blob, `water_${formatLabNoRange().replace('/', '_')}_${new Date().toISOString().slice(0,10).replace(/-/g,'')}.docx`);
        setStatus("DOCX à¤°à¤¿à¤ªà¥‹à¤°à¥à¤Ÿ à¤œà¤¨à¤°à¥‡à¤Ÿ à¤”à¤° à¤¡à¤¾à¤‰à¤¨à¤²à¥‹à¤¡ à¤¹à¥‹ à¤—à¤ˆ!", "success");
        console.log('DOCX generation successful');
    } catch (err) {
        console.error('DOCX generation error:', err);
        setStatus(`DOCX à¤œà¤¨à¤°à¥‡à¤¶à¤¨ à¤®à¥‡à¤‚ à¤¤à¥à¤°à¥à¤Ÿà¤¿: ${err.message}. docx.js à¤¯à¤¾ FileSaver.js à¤šà¥‡à¤• à¤•à¤°à¥‡à¤‚à¥¤`, "danger");
    }
}

function formatLabNoRange() {
    if (sampleDetails.length === 0) return "N/A";
    const prefixes = sampleDetails.map(s => parseInt(s['Lab No.'].split('/')[0]));
    const year = sampleDetails[0]['Lab No.'].split('/')[1];
    const min = Math.min(...prefixes);
    const max = Math.max(...prefixes);
    return min === max ? `${min}/${year}` : `${min}-${max}/${year}`;
}
function createSampleDocxTable({ Table, TableRow, TableCell, Paragraph, TextRun }) {
    const headers = ["à¤•à¥à¤°.à¤¸à¤‚.", "à¤µà¤¿à¤µà¤°à¤£"].concat(sampleDetails.map((_, i) => `(${i+1})`));
    const rows = [
        new TableRow({
            children: headers.map(h =>
                new TableCell({
                    children: [
                        new Paragraph({
                            children: [new TextRun({ text: h, size: 18 })]
                        })
                    ]
                })
            )
        })
    ];

    const dataRows = [
        ["1.1", "à¤¸à¥à¤°à¥‹à¤¤ (Source)"].concat(sampleDetails.map(s => s.Source)),
        ["1.2", "à¤¸à¥à¤¥à¤¾à¤¨ (Location)"].concat(sampleDetails.map(s => s.Location)),
        ["1.3", "à¤¸à¥€à¤à¤šà¤†à¤ˆ à¤¨à¤®à¥‚à¤¨à¤¾ à¤¸à¤‚. (CHI Sample No.)"].concat(sampleDetails.map(s => s['CHI Sample No.'])),
        ["1.4", "à¤¨à¤®à¥‚à¤¨à¤¾ à¤¸à¤‚à¤—à¥à¤°à¤¹ à¤•à¥€ à¤¤à¤¾à¤°à¥€à¤– (Date)"].concat(sampleDetails.map(s => s.Date)),
        ["1.5", "à¤ªà¥à¤°à¤¯à¥‹à¤—à¤¶à¤¾à¤²à¤¾ à¤¸à¤‚à¤–à¥à¤¯à¤¾ (Lab No.)"].concat(sampleDetails.map(s => s['Lab No.']))
    ];

    dataRows.forEach(row => {
        rows.push(
            new TableRow({
                children: row.map(cell =>
                    new TableCell({
                        children: [
                            new Paragraph({
                                children: [new TextRun({ text: cell, size: 18 })]
                            })
                        ]
                    })
                )
            })
        );
    });

    return new Table({ rows });
}

function createChemicalDocxTable({ Table, TableRow, TableCell, Paragraph, TextRun }) {
    const sampleCount = sampleDetails.length;

const topHeaderRow = new TableRow({
  children: [
    new TableCell({
      rowSpan: 2,
      children: [new Paragraph({ children: [new TextRun({ text: "à¤•à¥à¤°.à¤¸à¤‚.", size: 18 })] })]
    }),
    new TableCell({
      rowSpan: 2,
      children: [new Paragraph({ children: [new TextRun({ text: "à¤ªà¤°à¥€à¤•à¥à¤·à¤£ (Tests)", size: 18 })] })]
    }),
    new TableCell({
      columnSpan: 2,
      children: [new Paragraph({ children: [new TextRun({ text: "IS 10500:2012", size: 18 })] })]
    }),
    new TableCell({
      columnSpan: sampleCount,
      children: [new Paragraph({ children: [new TextRun({ text: "Result", size: 18 })] })]
    })
  ]
});

const secondHeaderCells = [
  new TableCell({
    children: [new Paragraph({ children: [new TextRun({ text: "à¤¨à¤¿à¤°à¥à¤§à¤¾à¤°à¤¿à¤¤ à¤®à¤¾à¤¨ (Max)", size: 18 })] })]
  }),
  new TableCell({
    children: [new Paragraph({ children: [new TextRun({ text: "à¤¨à¤¿à¤°à¥à¤§à¤¾à¤°à¤¿à¤¤ à¤®à¤¾à¤¨ (Desirable)", size: 18 })] })]
  })
];

sampleDetails.forEach(s => {
  secondHeaderCells.push(
    new TableCell({
      children: [new Paragraph({ children: [new TextRun({ text: s["Lab No."], size: 18 })] })]
    })
  );
});

const secondHeaderRow = new TableRow({
  children: secondHeaderCells
});

const rows = [topHeaderRow, secondHeaderRow];


    tests.forEach((test, i) => {
        const row = new TableRow({
            children: [
                new TableCell({
                    children: [new Paragraph({ children: [new TextRun({ text: `2.${i+1}`, size: 18 })] })]
                }),
                new TableCell({
                    children: [new Paragraph({ children: [new TextRun({ text: `${test.name} (${test.bilingual_name})`, size: 18 })] })]
                }),
                new TableCell({
                    children: [new Paragraph({ children: [new TextRun({ text: test.max_limit, size: 18 })] })]
                }),
                new TableCell({
                    children: [new Paragraph({ children: [new TextRun({ text: test.desirable_limit, size: 18 })] })]
                }),
                ...chemicalResults.map(r =>
                    new TableCell({
                        children: [
                            new Paragraph({
                                children: [new TextRun({ text: r[test.name] || '-', size: 18 })]
                            })
                        ]
                    })
                )
            ]
        });
        rows.push(row);
    });

    return new Table({ rows });
}

function generateRemarksDocx({ Paragraph, TextRun }) {
    const failedLabs = [];          // (unsuitable / invalid)
    const permissibleOnlyLabs = []; // (permissible, but no fail)
    const desirableLabs = [];       // (all desirable)

    chemicalResults.forEach((result) => {
        const labNo = result["Lab No."];
        let hasUnsuitableOrInvalid = false;
        let hasPermissible = false;
        let allDesirable = true;

        tests.forEach(test => {
            const value = result[test.name];
            const category = categorizeSample(
                test.name,
                value,
                test.max_limit,
                test.desirable_limit
            );

            if (category === "unsuitable" || category === "invalid") {
                hasUnsuitableOrInvalid = true;
                allDesirable = false;
            } else if (category === "permissible") {
                hasPermissible = true;
                allDesirable = false;
            }
        });

        if (hasUnsuitableOrInvalid) {
            failedLabs.push(labNo);
        } else if (hasPermissible) {
            permissibleOnlyLabs.push(labNo);
        } else if (allDesirable) {
            desirableLabs.push(labNo);
        }
    });

    const remarks = [];
    let lineNo = 1;

    // fail group
    if (failedLabs.length > 0) {
        const text = `(${lineNo}) à¤¨à¤®à¥‚à¤¨à¤¾ à¤ªà¥à¤°à¤¯à¥‹à¤—à¤¶à¤¾à¤²à¤¾ à¤¸à¤‚à¤–à¥à¤¯à¤¾ ${failedLabs.join(", ")} à¤…à¤§à¤¿à¤•à¤¤à¤® à¤…à¤¨à¥à¤®à¥‡à¤¯ à¤¶à¥à¤°à¥‡à¤£à¥€ à¤•à¥€ à¤¨à¤¿à¤°à¥à¤§à¤¾à¤°à¤¿à¤¤ à¤†à¤µà¤¶à¥à¤¯à¤•à¤¤à¤¾à¤“à¤‚ à¤•à¥€ à¤ªà¥‚à¤°à¥à¤¤à¤¿ à¤¨à¤¹à¥€à¤‚ à¤•à¤°à¤¤à¤¾ à¤¹à¥ˆ, à¤…à¤¤à¤ƒ à¤¯à¤¹ à¤‰à¤ªà¤¯à¥à¤•à¥à¤¤ à¤¨à¤¹à¥€à¤‚ à¤¹à¥ˆà¥¤`;
        remarks.push(
            new Paragraph({
                children: [
                    new TextRun({
                        text: text,
                        size: 18,
                        font: "Times New Roman"
                    })
                ]
            })
        );
        lineNo++;
    }

    // permissible-only group
    if (permissibleOnlyLabs.length > 0) {
        const text = `(${lineNo}) à¤¨à¤®à¥‚à¤¨à¤¾ à¤ªà¥à¤°à¤¯à¥‹à¤—à¤¶à¤¾à¤²à¤¾ à¤¸à¤‚à¤–à¥à¤¯à¤¾ ${permissibleOnlyLabs.join(", ")} à¤‰à¤šà¥à¤šà¤¤à¤® à¤…à¤­à¥€à¤·à¥à¤Ÿ à¤¶à¥à¤°à¥‡à¤£à¥€ à¤•à¥€ à¤¨à¤¿à¤°à¥à¤¦à¤¿à¤·à¥à¤Ÿ à¤†à¤µà¤¶à¥à¤¯à¤•à¤¤à¤¾à¤“à¤‚ à¤•à¥€ à¤ªà¥‚à¤°à¥à¤¤à¤¿ à¤¨à¤¹à¥€à¤‚ à¤•à¤°à¤¤à¤¾ à¤¹à¥ˆ à¤•à¤¿à¤¨à¥à¤¤à¥ à¤…à¤§à¤¿à¤•à¤¤à¤® à¤…à¤¨à¥à¤œà¥à¤žà¥‡à¤¯ à¤¶à¥à¤°à¥‡à¤£à¥€ à¤•à¥‡ à¤¨à¤¿à¤°à¥à¤¦à¤¿à¤·à¥à¤Ÿ à¤†à¤µà¤¶à¥à¤¯à¤•à¤¤à¤¾à¤“à¤‚ à¤•à¥€ à¤ªà¥‚à¤°à¥à¤¤à¤¿ à¤•à¤°à¤¤à¤¾ à¤¹à¥ˆà¤‚ à¤…à¤¤: à¤‰à¤šà¥à¤šà¤¤à¤° à¤µà¥ˆà¤•à¤²à¥à¤ªà¤¿à¤• à¤¸à¥à¤¤à¥à¤°à¥‹à¤¤ à¤•à¥‡ à¤…à¤­à¤¾à¤µ à¤®à¥‡à¤‚ à¤œà¤² à¤•à¤¾ à¤‰à¤ªà¤¯à¥‹à¤— à¤•à¤° à¤¸à¤•à¤¤à¥‡ à¤¹à¥ˆà¤‚ à¤”à¤° à¤•à¥à¤·à¥‡à¤¤à¥à¤° à¤®à¥‡à¤‚ à¤‰à¤ªà¤²à¤¬à¥à¤§ à¤œà¤² à¤•à¥€ à¤¸à¤¾à¤®à¤¾à¤¨à¥à¤¯ à¤µà¤¿à¤¶à¤¿à¤·à¥à¤Ÿà¤¤à¤¾à¤“ à¤•à¥‡ à¤†à¤§à¤¾à¤° à¤ªà¤° à¤œà¤² à¤•à¥‹ à¤¸à¤¾à¤§à¤¾à¤°à¤£à¤¤à¤¯à¤¾ à¤¸à¥à¤µà¥€à¤•à¥ƒà¤¤ à¤•à¤¿à¤¯à¤¾ à¤œà¤¾ à¤¸à¤•à¤¤à¤¾ à¤¹à¥ˆà¥¤`;
        remarks.push(
            new Paragraph({
                children: [
                    new TextRun({
                        text: text,
                        size: 18,
                        font: "Times New Roman"
                    })
                ]
            })
        );
        lineNo++;
    }

    // all-desirable group
    if (desirableLabs.length > 0) {
        const text = `(${lineNo}) à¤¨à¤®à¥‚à¤¨à¤¾ à¤ªà¥à¤°à¤¯à¥‹à¤—à¤¶à¤¾à¤²à¤¾ à¤¸à¤‚à¤–à¥à¤¯à¤¾ ${desirableLabs.join(", ")} à¤‰à¤šà¥à¤šà¤¤à¤® à¤…à¤­à¥€à¤·à¥à¤Ÿ à¤¶à¥à¤°à¥‡à¤£à¥€ à¤•à¥€ à¤¨à¤¿à¤°à¥à¤§à¤¾à¤°à¤¿à¤¤ à¤†à¤µà¤¶à¥à¤¯à¤•à¤¤à¤¾à¤“à¤‚ à¤•à¥€ à¤ªà¥‚à¤°à¥à¤¤à¤¿ à¤•à¤°à¤¤à¤¾ à¤¹à¥ˆ, à¤…à¤¤à¤ƒ à¤‰à¤ªà¤¯à¥à¤•à¥à¤¤ à¤¹à¥ˆà¥¤`;
        remarks.push(
            new Paragraph({
                children: [
                    new TextRun({
                        text: text,
                        size: 18,
                        font: "Times New Roman"
                    })
                ]
            })
        );
    }

    return remarks;
}




// Event Listeners
document.addEventListener('DOMContentLoaded', () => {
    document.getElementById('load-chi-csv').addEventListener('click', loadChiCsv);
    document.getElementById('clear-chi').addEventListener('click', clearChiForm);
    document.getElementById('next-to-samples').addEventListener('click', openSampleTab);
    document.getElementById('add-samples-btn').addEventListener('click', addSamplesFromNum);
    document.getElementById('load-sample-csv').addEventListener('click', loadSampleCsv);
    document.getElementById('next-chemical').addEventListener('click', generateReport);
    document.getElementById('clear-samples').addEventListener('click', clearSampleForm);
    document.getElementById('submit-chemical').addEventListener('click', submitChemicalResults);
    document.getElementById('clear-chemical').addEventListener('click', clearChemicalForm);
    document.getElementById('fetch-lab-no').addEventListener('click', fetchByLabNo);
    document.getElementById('fetch-sent-by').addEventListener('click', fetchBySentBy);
    document.getElementById('fetch-sent-location').addEventListener('click', fetchBySentByLocation);
    document.getElementById('export-query-pdf').addEventListener('click', generateQueryPdf);
    document.getElementById('generate-final-report').addEventListener('click', generateFinalReport);
    document.getElementById('back-to-chemical').addEventListener('click', backToChemical);

    // Delete à¤¬à¤Ÿà¤¨à¥‹à¤‚ à¤•à¥‹ à¤²à¤¿à¤¸à¤¨ à¤•à¤°à¥‹ (deleteSampleEntry à¤«à¤¿à¤•à¥à¤¸)
    document.getElementById('sample-entries').addEventListener('click', function(event) {
        if (event.target.closest('.delete-btn')) {
            const btn = event.target.closest('.delete-btn');
            const index = parseInt(btn.dataset.index);
            deleteSampleEntry(index);
        }
    });
});
