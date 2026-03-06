/**
 * seed_demo.js
 * Run: node seed_demo.js
 * Creates a demo bar entry under data/Surya_Bar_Restaurant/FY_2025-26/
 */

const fs = require('fs');
const path = require('path');

const demoBar = {
    barName: "Surya Bar Restaurant",
    shopType: "FL-III",
    entityType: "Proprietorship",
    tradeName: "Surya Bar",
    financialYear: "01/04/2025 to 31/03/2026",

    ownerName: "Ramesh Balaji Shinde",
    contactPerson: "Mahesh Shinde",
    phone: "9876543210",
    altPhone: "02240001234",
    email: "surya.bar@example.com",

    address: "Shop No. 12, Shivaji Market, MG Road",
    area: "Shivaji Nagar",
    city: "Pune",
    district: "Pune",
    state: "Maharashtra",
    pinCode: "411005",

    gstin: "27AABCS1234C1Z5",
    pan: "AABCS1234C",
    aadhar: "XXXX XXXX 5678",
    fssaiNo: "11325999000563",

    licenseNoInput: "0042",
    licenseNo: "FL-III/0042",
    licenseType: "FL-III",
    licenseIssueDate: "2025-04-01",
    licenseExpiryDate: "2026-03-31",
    exciseDivision: "Pune District Excise Office",
    shopActNo: "PNQ/2025/00789",

    bankName: "State Bank of India",
    bankBranch: "MG Road Pune",
    accountNo: "32100123456789",
    ifsc: "SBIN0001234",

    fireNoc: "FIRE/PNQ/2025/00321",
    policeVerification: "PUNE/PV/2025/00456",
    remarks: "Demo bar entry for testing",

    bar_id: "B_DEMO_001",
    created_at: new Date().toISOString()
};

// ── Build folder path ─────────────────────────────────────
function sanitize(name) {
    return name.trim().replace(/[<>:"/\\|?*\x00-\x1F]/g, '').replace(/\s+/g, '_').substring(0, 60);
}

function fyToFolder(fyLabel) {
    const match = fyLabel.match(/(\d{4})/g);
    if (match && match.length >= 2) {
        return `FY_${match[0]}-${match[1].slice(2)}`;
    }
    return 'FY_Unknown';
}

const barFolder = sanitize(demoBar.barName);
const fyFolder = fyToFolder(demoBar.financialYear);
const dirPath = path.join(__dirname, 'data', barFolder, fyFolder);

fs.mkdirSync(dirPath, { recursive: true });
fs.writeFileSync(
    path.join(dirPath, 'bar_master.json'),
    JSON.stringify(demoBar, null, 2)
);

// ── Update global index ───────────────────────────────────
const indexPath = path.join(__dirname, 'data', 'bars_index.json');
let index = [];
if (fs.existsSync(indexPath)) {
    const raw = fs.readFileSync(indexPath, 'utf8');
    if (raw) index = JSON.parse(raw);
}

// Remove existing demo entry if re-running
index = index.filter(b => b.bar_id !== demoBar.bar_id);
index.push({
    bar_id: demoBar.bar_id,
    barName: demoBar.barName,
    financialYear: demoBar.financialYear,
    folderPath: path.join('data', barFolder, fyFolder),
    created_at: demoBar.created_at
});
fs.writeFileSync(indexPath, JSON.stringify(index, null, 2));

console.log('✅ Demo bar created!');
console.log(`   → data/${barFolder}/${fyFolder}/bar_master.json`);
console.log(`   → data/bars_index.json updated`);
