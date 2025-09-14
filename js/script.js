// js/script.js - Modular Water Quality Report App (JalGanana style)
import { getFirestore, collection, doc, getDoc, setDoc, query, where, getDocs } from 'https://www.gstatic.com/firebasejs/10.13.2/firebase-firestore.js';

// Access global DBs from index.html
const totalWaterDb = window.totalWaterDb; // Total-Water Firebase
const jalGananaDb = window.jalGananaDb;   // JalGanana Firebase

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

// Status Update Function
function setStatus(message, type = "info") {
    const statusDiv = document.getElementById('status-var');
    statusDiv.innerHTML = `<div class="alert alert-${type}">${message}</div>`;
}

// CHI Functions
function loadChiCsv() {
    const file = document.getElementById('chi-csv').files[0];
    if (!file) return setStatus("कृपया CSV फ़ाइल चुनें।", "warning");
    const reader = new FileReader();
    reader.onload = (e) => {
        const lines = e.target.result.split('\n').filter(line => line.trim());
        if (lines.length < 2) return setStatus("CSV खाली या अमान्य है।", "danger");
        const headers = lines[0].split(',');
        const row = lines[1].split(',');
        const data = {};
        headers.forEach((h, i) => data[h.trim()] = row[i]?.trim());
        const allowedDivisions = ["अजमेर", "जोधपुर", "जयपुर", "बीकानेर"];
        if (!data["CHI Letter No."] || !data["CHI Address"] || !allowedDivisions.includes(data["Division"]) || !validateDate(data["Report Date"])) {
            return setStatus("अमान्य CSV डेटा। कॉलम चेक करें: CHI Letter No., CHI Address, Division, Report Date।", "danger");
        }
        document.getElementById('chi-letter-no').value = data["CHI Letter No."];
        document.getElementById('chi-address').value = data["CHI Address"];
        document.getElementById('chi-division').value = data["Division"];
        document.getElementById('report-date').value = data["Report Date"];
        setStatus("CHI विवरण CSV से लोड हो गए।", "success");
    };
    reader.readAsText(file, 'UTF-8');
}

function clearChiForm() {
    document.getElementById('chi-letter-no').value = '';
    document.getElementById('chi-address').value = '';
    document.getElementById('chi-division').value = '';
    document.getElementById('report-date').value = new Date().toLocaleDateString('en-GB').split('/').reverse().join('-');
    document.getElementById('chi-csv').value = '';
    setStatus("CHI फॉर्म साफ़ किया गया।", "info");
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
        return setStatus("कृपया वैध CHI Letter No., Address, Division, और Report Date (DD-MM-YYYY) दर्ज करें।", "danger");
    }
    new bootstrap.Tab(document.querySelector('#sample-tab')).show();
    setStatus("Sample Details टैब पर गए।", "success");
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
    if (isNaN(num) || num < 1 || num > 20) return setStatus("1-20 के बीच नंबर दर्ज करें।", "warning");
    for (let i = sampleDetails.length; i < sampleDetails.length + num; i++) {
        addSampleEntry({}, i);
    }
    setStatus(`${num} सैंपल जोड़े गए। आप इन्हें एडिट या डिलीट कर सकते हैं।`, "success");
}

function loadSampleCsv() {
    const file = document.getElementById('sample-csv').files[0];
    if (!file) return setStatus("CSV फ़ाइल चुनें।", "warning");
    const reader = new FileReader();
    reader.onload = (e) => {
        const lines = e.target.result.split('\n').filter(l => l.trim());
        if (lines.length < 2) return setStatus("अमान्य CSV।", "danger");
        const headers = lines[0].split(',');
        const required = ["Source", "Location", "CHI Sample No.", "Date", "Lab No.", "Sender"];
        if (!required.every(h => headers.some(head => head.trim() === h))) return setStatus("CSV में ये कॉलम होने चाहिए: " + required.join(", "), "danger");
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
        setStatus(`${loaded} सैंपल CSV से लोड हुए (अधिकतम ${maxLoaded})। आप इन्हें एडिट या डिलीट कर सकते हैं।`, "success");
    };
    reader.readAsText(file, 'UTF-8');
}

function deleteSampleEntry(index) {
    if (confirm(`Sample ${index + 1} डिलीट करें?`)) {
        sampleDetails.splice(index, 1);
        renderSampleEntries();
        setStatus(`Sample ${index + 1} डिलीट हो गया। कुल सैंपल अब: ${sampleDetails.length}। यह Chemical टेबल में दिखेगा।`, "warning");
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
    setStatus("सैंपल फॉर्म साफ़ किया गया।", "info");
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
    if (sampleDetails.length === 0) return setStatus("कोई वैध सैंपल नहीं। सभी फ़ील्ड भरें और फॉर्मेट चेक करें।", "danger");
    populateChemicalResultsTab();
    new bootstrap.Tab(document.querySelector('#chemical-tab')).show();
    setStatus(`${valid} वैध सैंपल प्रोसेस हुए। Chemical Results टैब पर गए।`, "success");
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
                // JalGanana (labcalc-cee5c) से lab_calculations कलेक्शन से डेटा fetch
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
                        "pH": "ph" // pH के लिए key जोड़ा, अगर JalGanana में ph के लिए अलग key हो तो अपडेट करो
                    };
                    value = data[keyMap[test.name]] || '';
                    // Whole number में कन्वर्ट करें
                    if (value && ["TDS", "T. Hardness", "Calcium", "Magnesium", "Chloride", "Alkalinity"].includes(test.name)) {
                        value = Math.round(parseFloat(value)).toString();
                    }
                } else {
                    setStatus(`JalGanana (lab_calculations) में ${labNo} के लिए डेटा नहीं मिला। मैनुअल एंट्री करें।`, "warning");
                }
            } catch (err) {
                console.error(err);
                setStatus(`${labNo} के लिए JalGanana से fetch error: ${err.message}. labcalc-cee5c के Firebase permissions चेक करें।`, "danger");
            }
            // डिफॉल्ट वैल्यूज सेट करें
            if (test.name === "Colour") value = value || "Clear";
            else if (test.name === "Odour") value = value || "OK";
            else if (test.name === "Turbidity") value = value || "NO";
            else if (test.name === "pH") value = value || "8.0"; // pH डिफॉल्ट 8.0
            tableHTML += `<td><input type="text" class="form-control chemical-input d-inline-block me-1" data-lab="${labNo}" data-test="${test.name}" value="${value}" oninput="updateFinalAndStatus('${labNo}', '${test.name}', this.value)"></td>
                          <td><input type="text" class="form-control chemical-final d-inline-block me-1" data-lab="${labNo}" data-test="${test.name}" readonly value="${value}"></td>
                          <td><span class="chemical-status fw-bold d-inline-block" data-lab="${labNo}" data-test="${test.name}"></span></td>`;
            updateFinalAndStatus(labNo, test.name, value);
        }
        tableHTML += '</tr>';
    }
    tableHTML += '</tbody></table>';
    container.innerHTML = tableHTML;
    setStatus(`Chemical Results लोड हो गए। JalGanana (labcalc-cee5c, lab_calculations) से TDS (${sampleDetails.map(s => s["Lab No."]).join(', ')}) के लिए डेटा लाए गए। pH डिफॉल्ट 8.0 सेट। Colour, Odour, Turbidity डिफॉल्ट। बाकी फील्ड्स एडिट करें।`, "success");
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
    setStatus("Chemical results साफ़ किए गए। डिफ़ॉल्ट और JalGanana डेटा बहाल।", "info");
}

function submitChemicalResults() {
    if (!confirm("रिजल्ट सबमिट करें और प्रीव्यू देखें?")) return;
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
    if (!allValid) return setStatus("Chemical results में अमान्य/खाली फ़ील्ड ठीक करें।", "danger");
    populatePreviewTab();
    new bootstrap.Tab(document.querySelector('#preview-tab')).show();
    setStatus("रिजल्ट सबमिट हो गए। प्रीव्यू तैयार है।", "success");
}

// Query Functions
async function fetchByLabNo() {
    const labNo = document.getElementById('query-lab-no').value.trim();
    if (!validateLabNo(labNo)) return setStatus("वैध Lab No. दर्ज करें (जैसे, 123/2025)।", "warning");
    const docId = labNo.replace('/', '-');
    try {
        const docRef = doc(collection(totalWaterDb, 'samples'), docId);
        const docSnap = await getDoc(docRef);
        queryResults = docSnap.exists() ? [docSnap.data()] : [];
        renderQueryTable();
        setStatus(queryResults.length ? `${labNo} के लिए रिजल्ट मिला।` : `${labNo} के लिए डेटा नहीं मिला।`, queryResults.length ? "success" : "warning");
    } catch (err) {
        setStatus(`Query error: ${err.message}. Total-Water Firebase permissions चेक करें।`, "danger");
    }
}

async function fetchBySentBy() {
    const sentBy = document.getElementById('query-sent-by').value.trim();
    if (!sentBy) return setStatus("Sent By दर्ज करें।", "warning");
    try {
        const q = query(collection(totalWaterDb, 'samples'), where('Sender', '==', sentBy));
        const snapshot = await getDocs(q);
        queryResults = snapshot.empty ? [] : snapshot.docs.map(d => d.data());
        renderQueryTable();
        setStatus(`"${sentBy}" के लिए ${queryResults.length} रिजल्ट मिले।`, "success");
    } catch (err) {
        setStatus(`Error: ${err.message}. Total-Water Firebase permissions चेक करें।`, "danger");
    }
}

async function fetchBySentByLocation() {
    const sentBy = document.getElementById('query-sent-by').value.trim();
    const location = document.getElementById('query-location').value.trim();
    if (!sentBy || !location) return setStatus("Sent By और Location दोनों दर्ज करें।", "warning");
    try {
        const q = query(collection(totalWaterDb, 'samples'), where('Sender', '==', sentBy), where('Location', '==', location));
        const snapshot = await getDocs(q);
        queryResults = snapshot.empty ? [] : snapshot.docs.map(d => d.data());
        renderQueryTable();
        setStatus(`"${sentBy}" + "${location}" के लिए ${queryResults.length} रिजल्ट मिले।`, "success");
    } catch (err) {
        setStatus(`Error: ${err.message}. Total-Water Firebase permissions चेक करें।`, "danger");
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
    if (queryResults.length === 0) return setStatus("एक्सपोर्ट करने के लिए कोई रिजल्ट नहीं।", "warning");
    const win = window.open('', '_blank');
    win.document.write(`
        <html><head><title>Query Report - ${new Date().toLocaleDateString()}</title>
        <style>body { font-family: Arial; } table { border-collapse: collapse; width: 100%; } th, td { border: 1px solid black; padding: 8px; text-align: center; } th { background-color: #f2f2f2; }</style>
        </head><body><h2>Database Query Report</h2><p>Generated on: ${new Date().toLocaleDateString()}</p>${document.getElementById('query-table').outerHTML}<script>window.print();</script></body></html>
    `);
    win.document.close();
    setStatus("PDF एक्सपोर्ट खुल गया। ब्राउज़र प्रिंट से PDF सेव करें।", "success");
}

// Preview Functions
function populatePreviewTab() {
    const sampleTbody = document.querySelector('#sample-preview-table tbody');
    sampleTbody.innerHTML = '';
    const sampleHeaders = ["क्र.सं.", "विवरण"].concat(sampleDetails.map((_, i) => `(${i+1})`));
    document.querySelector('#sample-preview-table thead tr').innerHTML = sampleHeaders.map(h => `<th>${h}</th>`).join('');
    const sampleRows = [
        ["1.1", "स्रोत (Source)"].concat(sampleDetails.map(s => s.Source)),
        ["1.2", "स्थान (Location)"].concat(sampleDetails.map(s => s.Location)),
        ["1.3", "मुख्य नि. नमूने की संख्या (CHI Sample No.)"].concat(sampleDetails.map(s => s['CHI Sample No.'])),
        ["1.4", "नमूना संग्रह की तारीख (Date)"].concat(sampleDetails.map(s => s.Date)),
        ["1.5", "प्रयोगशाला संख्या (Lab No.)"].concat(sampleDetails.map(s => s['Lab No.']))
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
    setStatus("रिपोर्ट प्रीव्यू अपडेट हो गया। DOCX जनरेट करने से पहले टेबल चेक करें।", "success");
}

function backToChemical() {
    new bootstrap.Tab(document.querySelector('#chemical-tab')).show();
    setStatus("Chemical Results पर वापस गए।", "info");
}

// Final Report
async function generateFinalReport() {
    if (!confirm("अंतिम DOCX रिपोर्ट जनरेट करें? डेटा Total-Water Firebase में सेव होगा (overwrite prompt के साथ)।")) return;
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
                if (!confirm(`Lab No. ${labNo} पहले से मौजूद है। ओवरराइट करें?`)) continue;
            }
            await setDoc(docRef, { ...sample, ...chemical, docId: docId });
            console.log('Successfully saved to Firebase for Lab No.:', labNo);
        } catch (err) {
            console.error('Save error for Lab No.', labNo, ':', err);
            setStatus(` ${labNo} को Total-Water में सेव करने में त्रुटि: ${err.message}. Firebase permissions या auth चेक करें।`, "danger");
            return;
        }
    }
    setStatus("Total-Water Firebase में डेटा सेव हो गया। DOCX जनरेट हो रहा है...", "info");

    // DOCX Generation
    try {
        console.log('Starting DOCX generation');
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
        const blob = await Packer.toBlob(doc);
        saveAs(blob, `water_${formatLabNoRange().replace('/', '_')}_${new Date().toISOString().slice(0,10).replace(/-/g,'')}.docx`);
        setStatus("DOCX रिपोर्ट जनरेट और डाउनलोड हो गई!", "success");
        console.log('DOCX generation successful');
    } catch (err) {
        console.error('DOCX generation error:', err);
        setStatus(`DOCX जनरेशन में त्रुटि: ${err.message}. docx.js या FileSaver.js चेक करें।`, "danger");
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

function createSampleDocxTable() {
    const headers = ["क्र.सं.", "विवरण"].concat(sampleDetails.map((_, i) => `(${i+1})`));
    const rows = [new TableRow({ children: headers.map(h => new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: h, size: 18 })] })] })) })];
    const dataRows = [
        ["1.1", "स्रोत (Source)"].concat(sampleDetails.map(s => s.Source)),
        ["1.2", "स्थान (Location)"].concat(sampleDetails.map(s => s.Location)),
        ["1.3", "मुख्य नि. नमूने की संख्या (CHI Sample No.)"].concat(sampleDetails.map(s => s['CHI Sample No.'])),
        ["1.4", "नमूना संग्रह की तारीख (Date)"].concat(sampleDetails.map(s => s.Date)),
        ["1.5", "प्रयोगशाला संख्या (Lab No.)"].concat(sampleDetails.map(s => s['Lab No.']))
    ];
    dataRows.forEach(row => {
        rows.push(new TableRow({ children: row.map(cell => new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: cell, size: 18 })] })] })) }));
    });
    return new Table({ rows });
}

function createChemicalDocxTable() {
    const headers = ["क.सं.", "परीक्षण (Tests)", "निर्धारित मान (Max)", "निर्धारित मान (Desirable)"].concat(sampleDetails.map(s => s['Lab No.']));
    const rows = [new TableRow({ children: headers.map(h => new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: h, size: 18 })] })] })) })];
    tests.forEach((test, i) => {
        const row = new TableRow({
            children: [
                new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: `2.${i+1}`, size: 18 })] })] }),
                new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: `${test.name} (${test.bilingual_name})`, size: 18 })] })] }),
                new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: test.max_limit, size: 18 })] })] }),
                new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: test.desirable_limit, size: 18 })] })] }),
                ...chemicalResults.map(r => new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: r[test.name] || '-', size: 18 })] })] }))
            ]
        });
        rows.push(row);
    });
    return new Table({ rows });
}

function generateRemarksDocx() {
    const remarks = [];
    chemicalResults.forEach((result, i) => {
        const labNo = result["Lab No."];
        let allDesirable = true;
        let hasPermissible = false;
        tests.forEach(test => {
            const value = result[test.name];
            const category = categorizeSample(test.name, value, test.max_limit, test.desirable_limit);
            if (category === "unsuitable" || category === "invalid") allDesirable = false;
            if (category === "permissible") hasPermissible = true;
        });
        const status = allDesirable ? "पेयजल के लिए उपयुक्त" : hasPermissible ? "पेयजल के लिए स्वीकार्य" : "पेयजल के लिए अनुपयुक्त";
        remarks.push(new Paragraph({ children: [new TextRun({ text: `Sample ${i+1} (${labNo}): ${status}`, size: 18, font: "Times New Roman" })] }));
    });
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

    // Delete बटनों को लिसन करो (deleteSampleEntry फिक्स)
    document.getElementById('sample-entries').addEventListener('click', function(event) {
        if (event.target.closest('.delete-btn')) {
            const btn = event.target.closest('.delete-btn');
            const index = parseInt(btn.dataset.index);
            deleteSampleEntry(index);
        }
    });
});
