/**
* Menampilkan halaman Website utama saat URL Web App diakses
*/
function doGet() {
return HtmlService.createTemplateFromFile('Index')
.evaluate()
.setTitle('Portal Manajemen & Screening CV HRD')
.setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
.addMetaTag('viewport', 'width=device-width, initial-scale=1');
}
/**
* Mengambil data pelamar dan statistik KPI secara real-time untuk ditampilkan di website
*/
function getDashboardData() {
const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Database Pelamar");
const lastRow = sheet.getLastRow();
if (lastRow < 5) {
return {
kpis: { total: 0, lolos: 0, avg: 0 },
applicants: []
};
}
const data = sheet.getRange(5, 1, lastRow - 4, 12).getValues();
let total = 0;
let lolos = 0;
let totalScore = 0;
let validScoresCount = 0;
const applicants = data.map((row, index) => {
const score = parseInt(row);
const validScore = !isNaN(score) && row !== "";
if (row) { // Pastikan kolom nama tidak kosong
total++;
if (validScore) {
totalScore += score;
validScoresCount++;
if (score >= 80) lolos++;
}
}
return {
rowNum: index + 5, // Simpan nomor baris asli di Sheets
timestamp: row ? Utilities.formatDate(new Date(row), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm") : "-",
nama: row || "-",
email: row || "-",
hp: row || "-",
posisi: row || "-",
cvUrl: row || "#",
score: validScore ? score : null,
kategori: row || "Belum Dinilai",
summary: row || "Sedang memproses analisis AI...",
skills: row || "-",
tindakan: row || "Belum Diperiksa",
statusEmail: row || "Pending"
};
}).filter(app => app.nama !== "-");
// Urutkan pelamar berdasarkan skor tertinggi (descending) untuk tampilan website
applicants.sort((a, b) => (b.score || 0) - (a.score || 0));
const avg = validScoresCount > 0 ? (totalScore / validScoresCount).toFixed(1) : 0;
return {
kpis: { total, lolos, avg },
applicants: applicants
};
}
/**
* Memproses tindakan HRD (Undang/Tolak) langsung dari Website
*/
function processHRAction(rowNum, action) {
try {
const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Database Pelamar");
// Tulis tindakan ke kolom K (kolom 11)
sheet.getRange(rowNum, 11).setValue(action);
SpreadsheetApp.flush();
// Trigger onEdit manual untuk mengirimkan email otomatis
const emailCell = sheet.getRange(rowNum, 12);
if (emailCell.getValue() !== "Sent") {
const nama = sheet.getRange(rowNum, 2).getValue();
const email = sheet.getRange(rowNum, 3).getValue();
const posisi = sheet.getRange(rowNum, 5).getValue();
if (action === "Undang Wawancara") {
sendInterviewEmail(email, nama, posisi);
emailCell.setValue("Sent");
} else if (action === "Tolak") {
sendRejectionEmail(email, nama, posisi);
emailCell.setValue("Sent");
}
}
return { success: true };
} catch (err) {
return { success: false, error: err.toString() };
}
}