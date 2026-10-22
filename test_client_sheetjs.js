const XLSX = require('./static/libs/xlsx.full.min.js');
const fs = require('fs');

console.log("Testing SheetJS buffer export...");

const wb = XLSX.utils.book_new();
const data = [
    ["Sản phẩm", "Số lượng", "Đơn giá", "Thành tiền"],
    ["Laptop Dell XPS", 5, 25000000, { f: "B2*C2" }],
    ["Bàn phím cơ", 20, 1500000, { f: "B3*C3" }],
    ["Màn hình 4K", 8, 8500000, { f: "B4*C4" }],
    ["TỔNG CỘNG", { f: "SUM(B2:B4)" }, "", { f: "SUM(D2:D4)" }]
];

const ws = XLSX.utils.aoa_to_sheet(data);
ws['!cols'] = [{ wch: 20 }, { wch: 12 }, { wch: 16 }, { wch: 18 }];
XLSX.utils.book_append_sheet(wb, ws, "BanHang");

// Generate XLSX buffer
const buffer = XLSX.write(wb, { bookType: 'xlsx', type: 'buffer' });
const filename = "test_sheetjs_output.xlsx";
fs.writeFileSync(filename, buffer);

console.log("File test_sheetjs_output.xlsx written successfully.");
const stats = fs.statSync(filename);
console.log(`Generated file size: ${stats.size} bytes.`);

if (stats.size > 1000) {
    console.log(">>> SUCCESS: Client-side SheetJS XLSX export test passed! <<<");
} else {
    console.error("FAIL: File too small or corrupt");
    process.exit(1);
}
