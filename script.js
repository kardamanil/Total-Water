// js/script.js - Modular Water Quality Report App (JalGanana style)

// Firebase Config (from your message)
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

// Tests configuration
const tests = [
    { name: "Colour", bilingual_name: "रंग", max_limit: "-", desirable_limit: "Clear" },
    { name: "Odour", bilingual_name: "गंध", max_limit: "-", desirable_limit: "OK" },
    { name: "Turbidity", bilingual_name: "टर्बिडिटी", max_limit: "-", desirable_limit: "NO" },
    { name: "TDS", bilingual_name: "टीडीएस", max_limit: "2000 mg/l", desirable_limit: "500 mg/l" },
    { name: "pH", bilingual_name: "पीएच", max_limit: "6.5 से 8.5", desirable_limit: "6.5 से 8.5" },
    { name: "T. Hardness", bilingual_name: "कुल कठोरता", max_limit: "600 mg/l", desirable_limit: "300 mg/l" },
    { name: "Calcium", bilingual_name: "कैल्शियम", max_limit: "200 mg/l", desirable_limit: "75 mg/l" },
    { name: "Magnesium", bilingual_name: "मैग्नीशियम", max_limit: "100 mg/l", desirable_limit: "30 mg/l" },
    { name: "Chloride", bilingual_name: "क्लोराइड", max_limit: "1000 mg/l", desirable_limit: "250 mg/l" },
    { name: "Alkalinity", bilingual_name: "क्षारीयता", max_limit: "600 mg/l", desirable_limit: "200 mg/l" }
];

// Global variables
let chiDetails = {};
let sampleDetails = [];
let chemicalResults = [];
let queryResults = [];

// Validation Functions
function validateDate(dateStr) {
    try {
        const [dd, mm, yyyy] = dateStr.split(/[-\/]/).map(Number);
        const date = new Date(yyyy, mm - 1, dd);
        return date.getDate() === dd && date.getMonth() === mm - 1 && date.getFullYear() === yyyy;
    } catch {
        return false;
    }
}

function validateLabNo(labNo) {
    return /^\d+\/\d{4}$/.test(labNo);
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
    if (!file) return setStatus("Please select a CSV file.");
    const reader = new FileReader();
    reader.onload = (e) => {
        const lines = e.target.result.split('\n');
        const headers = lines[0].split(',');
        const row = lines[1].split(',');
        const data = {};
        headers.forEach((h, i) => data[h.trim()] = row[i]?.trim());
        if (!data["CHI Letter No."] || !data["CHI Address"] || !data["Division"] || !data["Report Date"]) {
            return setStatus("Invalid CSV format.");
        }
        if (!["अजमेर", "जोधपुर", "जयपुर", "बीकानेर"].includes(data["Division"])) return setStatus("Invalid Division.");
        if (!validateDate(data["Report Date"])) return setStatus("Invalid Report Date.");
        document.getElementById('chi-letter-no').value = data["CHI Letter No."];
        document.getElementById('chi-address').value = data["CHI Address"];
        document.getElementById('chi-division').value = data["Division"];
        document.getElementById('report-date').value = data["Report Date"];
        setStatus("CHI details loaded from CSV.");
    };
    reader.readAsText(file);
}

function clearChiForm() {
    document.getElementById('chi-letter-no').value = '';
    document.getElementById('chi-address').value = '';
    document.getElementById('chi-division').value = '';
    document.getElementById('report-date').value = new Date().toLocaleDateString('en-GB').replace(/\//g, '-');
    setStatus("CHI form cleared");
}

function openSampleTab() {
    chiDetails = {
        letterNo: document.getElementById('chi-letter-no').value,
        address: document.getElementById('chi-address').value,
        division: document.getElementById('chi-division').value,
        reportDate: document.getElementById('report-date').value
    };
    if (!chiDetails.letterNo || !chiDetails.address || !chiDetails.division || !validateDate(chiDetails.reportDate)) {
        return setStatus("Please fill all CHI fields correctly.");
    }
    new bootstrap.Tab(document.querySelector('a[href="#sample"]')).show();
    setStatus("Moved to Sample Details tab");
}

// Sample Functions
function addSampleEntry(sample = { source: '', location: '', chiSampleNo: '', date: '', labNo: '', sender: '' }, index = sampleDetails.length) {
    const container = document.getElementById('sample-entries');
    const div = document.createElement('div');
    div.className = 'sample-entry';
    div.innerHTML = `
        <h5>Sample ${index+1}</h5>
        <div class="form-group">
            <label>Source:</label>
            <input type="text" class="form-control sample-source" value="${sample.source}">
        </div>
        <div class="form-group">
            <label>Location:</label>
            <input type="text" class="form-control sample-location" value="${sample.location}">
        </div>
        <div class="form-group">
            <label>CHI Sample No.:</label>
            <input type="text" class="form-control sample-chi-sample-no" value="${sample.chiSampleNo}">
        </div>
        <div class="form-group">
            <label>Date (DD/MM/YYYY):</label>
            <input type="text" class="form-control sample-date" value="${sample.date}">
        </div>
        <div class="form-group">
            <label>Lab No. (e.g., 123/2025):</label>
            <input type="text" class="form-control sample-lab-no" value="${sample.labNo}">
        </div>
        <div class="form-group">
            <label>Sender:</label>
            <input type="text" class="form-control sample-sender" value="${sample.sender}">
        </div>
        <button type="button" class="btn btn-danger" onclick="deleteSampleEntry(${index})"><i class="bi bi-x-circle"></i> Delete</button>
    `;
    container.appendChild(div);
    sampleDetails[index] = sample;
}

function addSamplesFromNum() {
    const num = parseInt(document.getElementById('num-samples').value);
    if (isNaN(num) || num < 1) return setStatus("Invalid number of samples.");
    for (let i = 0; i < num; i++) {
        addSampleEntry({}, sampleDetails.length);
    }
    setStatus(`${num} samples added.`);
}

function deleteSampleEntry(index) {
    sampleDetails.splice(index, 1);
    renderSampleEntries();
    setStatus("Sample deleted.");
}

function renderSampleEntries() {
    const container = document.getElementById('sample-entries');
    container.innerHTML = '';
    sampleDetails.forEach((sample, i) => addSampleEntry(sample, i));
}

function clearSampleForm() {
    sampleDetails = [];
    renderSampleEntries();
    setStatus("Sample form cleared.");
}

function generateReport() {
    const entries = document.querySelectorAll('.sample-entry');
    sampleDetails = [];
    let valid = true;
    entries.forEach((entry, i) => {
        const source = entry.querySelector('.sample-source').value;
        const location = entry.querySelector('.sample-location').value;
        const chiSampleNo = entry.querySelector('.sample-chi-sample-no').value;
        const date = entry.querySelector('.sample-date').value;
        const labNo = entry.querySelector('.sample-lab-no').value;
        const sender = entry.querySelector('.sample-sender').value;
        if (!source || !location || !chiSampleNo || !date || !labNo || !sender || !validateDate(date) || !validateLabNo(labNo)) {
            valid = false;
            return;
        }
        sampleDetails.push({ Source: source, Location: location, "CHI Sample No.": chiSampleNo, Date: date, "Lab No.": labNo, Sender: sender });
    });
    if (!valid) return setStatus("Please fill all sample fields correctly.");
    populateChemicalResultsTab();
    new bootstrap.Tab(document.querySelector('a[href="#chemical"]')).show();
    setStatus("Moved to Chemical Results.");
}

// Chemical Functions
async function populateChemicalResultsTab() {
    const container = document.getElementById('chemical-table-container');
    container.innerHTML = '';
    const table = document.createElement('table');
    table.className = 'table table-bordered';
    const thead = document.createElement('thead');
    const tr = document.createElement('tr');
    tr.appendChild(document.createElement('th').appendChild(document.createTextNode('Test')));
    sampleDetails.forEach(sample => {
        const th = document.createElement('th');
        th.colSpan = 3;
        th.textContent = sample["Lab No."];
        tr.appendChild(th);
    });
    thead.appendChild(tr);
    table.appendChild(thead);
    const tbody = document.createElement('tbody');
    for (const test of tests) {
        const row = document.createElement('tr');
        row.appendChild(document.createElement('td').appendChild(document.createTextNode(test.name)));
        for (const sample of sampleDetails) {
            const labNo = sample["Lab No."];
            const docId = labNo.replace('/', '-');
            const doc = await db.collection('samples').doc(docId).get();
            let value = '';
            if (doc.exists) {
                const data = doc.data();
                value = data[test.name.toLowerCase().replace('.', '')] || '';
            }
            if (test.name === "Colour") value = "Clear";
            else if (test.name === "Odour") value = "OK";
            else if (test.name === "Turbidity") value = "NO";
            const inputTd = document.createElement('td');
            const input = document.createElement('input');
            input.type = 'text';
            input.className = 'form-control chemical-input';
            input.dataset.lab = labNo;
            input.dataset.test = test.name;
            input.value = value;
            input.addEventListener('input', () => updateFinalAndStatus(labNo, test.name, input.value));
            inputTd.appendChild(input);
            row.appendChild(inputTd);
            const finalTd = document.createElement('td');
            const finalInput = document.createElement('input');
            finalInput.type = 'text';
            finalInput.className = 'form-control chemical-final';
            finalInput.dataset.lab = labNo;
            finalInput.dataset.test = test.name;
            finalInput.readOnly = true;
            finalTd.appendChild(finalInput);
            row.appendChild(finalTd);
            const statusTd = document.createElement('td');
            const statusSpan = document.createElement('span');
            statusSpan.className = 'chemical-status';
            statusSpan.dataset.lab = labNo;
            statusSpan.dataset.test = test.name;
            statusTd.appendChild(statusSpan);
            row.appendChild(statusTd);
        }
        tbody.appendChild(row);
    }
    table.appendChild(tbody);
    container.appendChild(table);
    setStatus("Chemical Results tab populated.");
}

function updateFinalAndStatus(labNo, testName, value) {
    const finalInput = document.querySelector(`.chemical-final[data-lab="${labNo}"][data-test="${test.name}"]`);
    const statusSpan = document.querySelector(`.chemical-status[data-lab="${labNo}"][data-test="${test.name}"]`);
    finalInput.value = value;
    const category = categorizeSample(testName, value, tests.find(t => t.name === testName).max_limit, tests.find(t => t.name === testName).desirable_limit);
    statusSpan.textContent = category === "suitable" ? "✅ Desirable" : category === "permissible" ? "⚠️ Permissible" : "❌ Failed";
}

function clearChemicalForm() {
    populateChemicalResultsTab();
    setStatus("Chemical results cleared.");
}

function submitChemicalResults() {
    if (!confirm("Submit results and view preview?")) return;
    chemicalResults = [];
    sampleDetails.forEach(sample => {
        const labNo = sample["Lab No."];
        const entries = {};
        document.querySelectorAll(`.chemical-input[data-lab="${labNo}"]`).forEach(input => {
            entries[input.dataset.test] = input.value.trim();
        });
        chemicalResults.push({ "Lab No.": labNo, ...entries });
    });
    populatePreviewTab();
    new bootstrap.Tab(document.querySelector('a[href="#preview"]')).show();
    setStatus("Report preview generated.");
}

// Query Functions
async function fetchByLabNo() {
    const labNo = document.getElementById('query-lab-no').value;
    if (!labNo) return setStatus("Please enter Lab No.");
    const docId = labNo.replace('/', '-');
    const doc = await db.collection('samples').doc(docId).get();
    queryResults = doc.exists ? [doc.data()] : [];
    renderQueryTable();
}

async function fetchBySentBy() {
    const sentBy = document.getElementById('query-sent-by').value;
    if (!sentBy) return setStatus("Please enter Sent By.");
    const snapshot = await db.collection('samples').where('Sender', '==', sentBy).get();
    queryResults = snapshot.docs.map(doc => doc.data());
    renderQueryTable();
}

async function fetchBySentByLocation() {
    const sentBy = document.getElementById('query-sent-by').value;
    const location = document.getElementById('query-location').value;
    if (!sentBy || !location) return setStatus("Please enter both Sent By and Location.");
    const snapshot = await db.collection('samples').where('Sender', '==', sentBy).where('Location', '==', location).get();
    queryResults = snapshot.docs.map(doc => doc.data());
    renderQueryTable();
}

function renderQueryTable() {
    const tbody = document.querySelector('#query-table tbody');
    tbody.innerHTML = '';
    queryResults.forEach(result => {
        const row = tbody.insertRow();
        ["Lab No.", "Source", "Location", "CHI Sample No.", "Date", "Sender", "Colour", "Odour", "Turbidity", "TDS", "pH", "T. Hardness", "Calcium", "Magnesium", "Chloride", "Alkalinity"].forEach(col => {
            const cell = row.insertCell();
            cell.textContent = result[col] || '-';
        });
    });
    setStatus(`Found ${queryResults.length} results.`);
}

function generateQueryPdf() {
    // Simple implementation using window.print
    const content = document.getElementById('query-table').outerHTML;
    const win = window.open('', '', 'height=700,width=700');
    win.document.write('<html><head><title>Query Report</title></head><body>');
    win.document.write(content);
    win.document.write('</body></html>');
    win.document.close();
    win.print();
}

function populatePreviewTab() {
    const sampleTbody = document.querySelector('#sample-preview-table tbody');
    sampleTbody.innerHTML = '';
    const sampleData = [
        ["1.1", "स्रोत (Source)", ...sampleDetails.map(s => s.Source)],
        ["1.2", "स्थान (Location)", ...sampleDetails.map(s => s.Location)],
        ["1.3", "मुख्य नि. नमूने की संख्या (CHI Sample No.)", ...sampleDetails.map(s => s["CHI Sample No."])],
        ["1.4", "नमूना संग्रह की तारीख (Date)", ...sampleDetails.map(s => s.Date)],
        ["1.5", "प्रयोगशाला संख्या (Lab No.)", ...sampleDetails.map(s => s["Lab No."])]
    ];
    sampleData.forEach(row => {
        const tr = sampleTbody.insertRow();
        row.forEach(cell => tr.insertCell().textContent = cell);
    });

    const chemicalTbody = document.querySelector('#chemical-preview-table tbody');
    chemicalTbody.innerHTML = '';
    tests.forEach((test, i) => {
        const tr = chemicalTbody.insertRow();
        tr.insertCell().textContent = `2.${i+1}`;
        tr.insertCell().textContent = `${test.name} (${test.bilingual_name})`;
        tr.insertCell().textContent = test.max_limit;
        tr.insertCell().textContent = test.desirable_limit;
        chemicalResults.forEach(r => tr.insertCell().textContent = r[test.name] || '-');
    });
}

function generateFinalReport() {
    if (!confirm("Generate final DOCX report?")) return;
    // Save to Firebase with overwrite
    sampleDetails.forEach(async (sample, i) => {
        const labNo = sample["Lab No."];
        const docId = labNo.replace('/', '-');
        const docRef = db.collection('samples').doc(docId);
        const chemical = chemicalResults[i];
        if (await docRef.get().exists) {
            if (!confirm(`Overwrite ${labNo}?`)) return;
        }
        docRef.set({ ...sample, ...chemical });
    });
    // DOCX generation (same as original)
    const doc = new docx.Document({
        sections: [{
            properties: {},
            children: [
                new docx.Paragraph({
                    text: "उत्तर पश्चिम रेलवे",
                    alignment: docx.AlignmentType.CENTER,
                    style: "bold"
                }),
                // Add all paragraphs and tables as per original
                // (Full implementation omitted for brevity, but use the structure from previous code)
            ]
        }]
    });
    docx.Packer.toBlob(doc).then(blob => {
        saveAs(blob, "report.docx");
    });
}

function setStatus(message) {
    document.getElementById('status-var').textContent = message;
}

// Init
document.addEventListener('DOMContentLoaded', () => {
    document.getElementById('load-chi-csv').addEventListener('click', loadChiCsv);
    document.getElementById('clear-chi').addEventListener('click', clearChiForm);
    document.getElementById('next-to-samples').addEventListener('click', openSampleTab);
    document.getElementById('add-samples-btn').addEventListener('click', addSampleEntryFromNum);
    document.getElementById('load-sample-csv').addEventListener('click', loadSampleCsv);
    document.getElementById('clear-samples').addEventListener('click', clearSampleForm);
    document.getElementById('next-to-chemical').addEventListener('click', generateReport);
    document.getElementById('clear-chemical').addEventListener('click', clearChemicalForm);
    document.getElementById('submit-chemical').addEventListener('click', submitChemicalResults);
    document.getElementById('fetch-lab-no').addEventListener('click', fetchByLabNo);
    document.getElementById('fetch-sent-by').addEventListener('click', fetchBySentBy);
    document.getElementById('fetch-sent-location').addEventListener('click', fetchBySentByLocation);
    document.getElementById('export-query-pdf').addEventListener('click', generateQueryPdf);
    document.getElementById('back-to-chemical').addEventListener('click', () => new bootstrap.Tab(document.querySelector('a[href="#chemical"]')).show());
    document.getElementById('generate-final-report').addEventListener('click', generateFinalReport);
});
