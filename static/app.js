// ==========================================================================
// Excel Web Pro - Core Application Script
// Features: Full Formula Engine, Grid Manipulation, Formatting, Floating Charts,
// Multi-sheet, Autocomplete, Drag-Fill, Direct XLSX Import & Export
// ==========================================================================

const appState = {
    filename: "Bao_Cao_Kinh_Doanh.xlsx",
    filepath: null,
    sheets: {},
    styles: {},
    colWidths: {},
    rowHeights: {},
    merges: {},
    images: {}, // { [sheetName]: [ { id, src, name, x, y, width, height, col, row, zIndex } ] }
    selectedImageId: null,
    imageClipboard: null, // { src, name, width, height }
    activeSheetName: "DoanhThu",
    selectedCell: { row: 0, col: 0 },
    selectionRange: { startRow: 0, startCol: 0, endRow: 0, endCol: 0 },
    isSelecting: false,
    isEditing: false,
    dragFill: { active: false, startRow: 0, startCol: 0, endRow: 0, endCol: 0 },
    clipboard: null, // { type: 'copy'|'cut', sheet: '', range: {}, data: [[]], styles: {} }
    undoStack: [],
    redoStack: [],
    charts: [],
    theme: "dark",
    zoom: 100,
    filters: {}, // { sheetName: { colIdx: [allowedValues] } }
    filterActiveCol: null,
    filterMode: true,
    resizingCol: null,
    resizingRow: null
};

// ==========================================================================
// Supported Formulas Definitions & Descriptions
// ==========================================================================
const FORMULA_DEFINITIONS = [
    { name: "SUM", syntax: "=SUM(number1, [number2], ...)", category: "math", desc: "Tính tổng tất cả các số trong vùng dữ liệu được chọn." },
    { name: "AVERAGE", syntax: "=AVERAGE(number1, [number2], ...)", category: "stat", desc: "Tính giá trị trung bình cộng của các ô chứa số." },
    { name: "MIN", syntax: "=MIN(number1, [number2], ...)", category: "stat", desc: "Tìm giá trị nhỏ nhất trong tập hợp dữ liệu." },
    { name: "MAX", syntax: "=MAX(number1, [number2], ...)", category: "stat", desc: "Tìm giá trị lớn nhất trong tập hợp dữ liệu." },
    { name: "COUNT", syntax: "=COUNT(value1, [value2], ...)", category: "stat", desc: "Đếm số lượng ô chứa giá trị dạng số." },
    { name: "COUNTA", syntax: "=COUNTA(value1, [value2], ...)", category: "stat", desc: "Đếm số lượng ô không rỗng (chứa bất kỳ dữ liệu nào)." },
    { name: "COUNTBLANK", syntax: "=COUNTBLANK(range)", category: "stat", desc: "Đếm số lượng ô trống trong vùng chỉ định." },
    { name: "COUNTIF", syntax: "=COUNTIF(range, criteria)", category: "stat", desc: "Đếm số ô thỏa mãn một điều kiện nhất định (ví dụ: >50, 'Nam')." },
    { name: "SUMIF", syntax: "=SUMIF(range, criteria, [sum_range])", category: "math", desc: "Tính tổng các ô thỏa mãn điều kiện chỉ định." },
    { name: "PRODUCT", syntax: "=PRODUCT(number1, [number2], ...)", category: "math", desc: "Nhân tất cả các số trong các đối số và trả về tích." },
    { name: "ROUND", syntax: "=ROUND(number, num_digits)", category: "math", desc: "Làm tròn số đến số chữ số thập phân chỉ định." },
    { name: "ROUNDUP", syntax: "=ROUNDUP(number, num_digits)", category: "math", desc: "Làm tròn số lên, ra xa số 0." },
    { name: "ROUNDDOWN", syntax: "=ROUNDDOWN(number, num_digits)", category: "math", desc: "Làm tròn số xuống, về phía số 0." },
    { name: "ABS", syntax: "=ABS(number)", category: "math", desc: "Trả về giá trị tuyệt đối của một số." },
    { name: "SQRT", syntax: "=SQRT(number)", category: "math", desc: "Trả về căn bậc hai của một số dương." },
    { name: "POWER", syntax: "=POWER(number, power)", category: "math", desc: "Tính lũy thừa số mũ của một số (tương đương ^)." },
    { name: "MOD", syntax: "=MOD(number, divisor)", category: "math", desc: "Trả về số dư sau khi chia một số cho số chia." },
    { name: "MEDIAN", syntax: "=MEDIAN(number1, [number2], ...)", category: "stat", desc: "Trả về số trung vị (ở giữa) của một tập dữ liệu." },
    { name: "STDEV", syntax: "=STDEV(number1, [number2], ...)", category: "stat", desc: "Ước tính độ lệch chuẩn dựa trên mẫu dữ liệu." },
    { name: "INT", syntax: "=INT(number)", category: "math", desc: "Làm tròn một số xuống số nguyên gần nhất." },
    { name: "IF", syntax: "=IF(logical_test, value_if_true, [value_if_false])", category: "logical", desc: "Kiểm tra điều kiện logic và trả về giá trị tương ứng." },
    { name: "IFS", syntax: "=IFS(test1, val1, [test2, val2], ...)", category: "logical", desc: "Kiểm tra nhiều điều kiện và trả về giá trị cho điều kiện TRUE đầu tiên." },
    { name: "AND", syntax: "=AND(logical1, [logical2], ...)", category: "logical", desc: "Trả về TRUE nếu tất cả các đối số là TRUE." },
    { name: "OR", syntax: "=OR(logical1, [logical2], ...)", category: "logical", desc: "Trả về TRUE nếu có ít nhất một đối số là TRUE." },
    { name: "NOT", syntax: "=NOT(logical)", category: "logical", desc: "Đảo ngược giá trị logic (TRUE thành FALSE và ngược lại)." },
    { name: "CONCAT", syntax: "=CONCAT(text1, [text2], ...)", category: "text", desc: "Nối nhiều chuỗi văn bản hoặc ô thành một chuỗi." },
    { name: "CONCATENATE", syntax: "=CONCATENATE(text1, [text2], ...)", category: "text", desc: "Nối hai hoặc nhiều chuỗi văn bản thành một chuỗi." },
    { name: "TEXTJOIN", syntax: "=TEXTJOIN(delimiter, ignore_empty, text1, ...)", category: "text", desc: "Nối văn bản với dấu phân cách và tùy chọn bỏ qua ô trống." },
    { name: "UPPER", syntax: "=UPPER(text)", category: "text", desc: "Chuyển đổi tất cả các ký tự trong chuỗi văn bản thành chữ hoa." },
    { name: "LOWER", syntax: "=LOWER(text)", category: "text", desc: "Chuyển đổi tất cả các ký tự trong chuỗi văn bản thành chữ thường." },
    { name: "PROPER", syntax: "=PROPER(text)", category: "text", desc: "Viết hoa chữ cái đầu tiên của mỗi từ trong chuỗi văn bản." },
    { name: "LEN", syntax: "=LEN(text)", category: "text", desc: "Trả về số lượng ký tự trong chuỗi văn bản." },
    { name: "TRIM", syntax: "=TRIM(text)", category: "text", desc: "Xóa tất cả các khoảng trắng thừa ở đầu, cuối và giữa các từ." },
    { name: "LEFT", syntax: "=LEFT(text, [num_chars])", category: "text", desc: "Trả về số ký tự đầu tiên tính từ bên trái của chuỗi." },
    { name: "RIGHT", syntax: "=RIGHT(text, [num_chars])", category: "text", desc: "Trả về số ký tự cuối cùng tính từ bên phải của chuỗi." },
    { name: "MID", syntax: "=MID(text, start_num, num_chars)", category: "text", desc: "Trả về số ký tự bắt đầu từ vị trí chỉ định trong chuỗi." },
    { name: "VLOOKUP", syntax: "=VLOOKUP(lookup_value, table_array, col_index, [range_lookup])", category: "lookup", desc: "Tìm kiếm giá trị ở cột đầu tiên bên trái và trả về giá trị cùng hàng ở cột chỉ định." },
    { name: "INDEX", syntax: "=INDEX(array, row_num, [column_num])", category: "lookup", desc: "Trả về giá trị của một phần tử trong bảng hoặc mảng tại hàng và cột chỉ định." },
    { name: "MATCH", syntax: "=MATCH(lookup_value, lookup_array, [match_type])", category: "lookup", desc: "Trả về vị trí tương đối của mục trong một mảng khớp với giá trị chỉ định." },
    { name: "TODAY", syntax: "=TODAY()", category: "date", desc: "Trả về ngày hiện tại của hệ thống." },
    { name: "NOW", syntax: "=NOW()", category: "date", desc: "Trả về ngày và giờ hiện tại của hệ thống." },
    { name: "DATE", syntax: "=DATE(year, month, day)", category: "date", desc: "Tạo giá trị ngày tháng từ năm, tháng và ngày riêng biệt." },
    { name: "YEAR", syntax: "=YEAR(serial_number)", category: "date", desc: "Trả về năm từ một giá trị ngày tháng." },
    { name: "MONTH", syntax: "=MONTH(serial_number)", category: "date", desc: "Trả về tháng (từ 1 đến 12) từ một giá trị ngày tháng." },
    { name: "DAY", syntax: "=DAY(serial_number)", category: "date", desc: "Trả về ngày trong tháng (từ 1 đến 31) từ một giá trị ngày tháng." }
];

// ==========================================================================
// Initialization
// ==========================================================================
document.addEventListener("DOMContentLoaded", () => {
    initApp();
});

async function initApp() {
    setupRibbonTabs();
    setupFormulaBar();
    setupFormattingButtons();
    setupModals();
    setupKeyboardShortcuts();
    setupContextMenu();
    setupZoomAndTheme();
    setupInfiniteGridControls();
    setupGridEventDelegation();
    setupGlobalDragAndDrop();
    setupSaveStatusBadge();
    setupImageControls();
    setupFilterPopoverGlobalEvents();
    document.addEventListener("paste", handlePasteEvent);
    
    // 1. Instant 0ms load if initial data was pre-injected into HTML by server
    const payloadEl = document.getElementById("initial-data-payload");
    if (payloadEl && payloadEl.textContent.trim()) {
        try {
            const data = JSON.parse(payloadEl.textContent);
            if (data && data.sheets && Object.keys(data.sheets).length > 0) {
                applyLoadedWorkbook(data);
                return;
            }
        } catch (e) {
            console.warn("Lỗi đọc initial-data-payload:", e);
        }
    }

    // 2. Fallback to API check
    const fileLoaded = await checkAndLoadInitialFile();
    if (!fileLoaded) {
        // Load default demo sales template only if no initial file was opened
        loadTemplate("sales");
    }
}

async function checkAndLoadInitialFile() {
    try {
        const res = await fetch("/api/initial-file");
        if (!res.ok) return false;
        const info = await res.json();
        if (!info.has_file) return false;

        showLoading(`Đang mở file: ${info.filename}...`);
        const loadRes = await fetch("/api/load-initial-file");
        if (!loadRes.ok) {
            hideLoading();
            return false;
        }
        const data = await loadRes.json();

        if (data.fallback_to_raw && data.raw_url) {
            const blobRes = await fetch(data.raw_url);
            const blob = await blobRes.blob();
            const file = new File([blob], data.filename, { type: blob.type });
            importExcelFile(file);
            if (data.filepath) appState.filepath = data.filepath;
            hideLoading();
            return true;
        }

        if (data.sheets && Object.keys(data.sheets).length > 0) {
            applyLoadedWorkbook(data);
            hideLoading();
            showToast(`Đã mở thành công file: ${data.filename}`, "success");
            return true;
        }
    } catch (err) {
        console.error("Lỗi khi tải file ban đầu:", err);
        hideLoading();
    }
    return false;
}

function applyLoadedWorkbook(data) {
    pushHistory();
    appState.filename = data.filename;
    appState.filepath = data.filepath || null;
    appState.sheets = data.sheets;
    appState.styles = data.styles || {};
    appState.colWidths = data.colWidths || {};
    appState.rowHeights = data.rowHeights || {};
    appState.merges = data.merges || {};
    appState.images = data.images || {};
    appState.filters = {};
    appState.activeSheetName = (data.sheetNames && data.sheetNames[0]) || Object.keys(data.sheets)[0];

    const nameInput = document.getElementById("workbook-filename");
    if (nameInput) nameInput.value = data.filename;
    document.title = `${data.filename} - Excel Web Pro`;
    setSaveStatus(true, "Đã mở");

    renderTabs();
    renderGrid();
    renderImages();
    selectCell(0, 0);
}

function setSaveStatus(isSaved, message) {
    const badge = document.getElementById("save-badge");
    if (!badge) return;
    if (isSaved) {
        badge.className = "save-status-badge saved";
        badge.textContent = message || "Đã lưu";
    } else {
        badge.className = "save-status-badge unsaved";
        badge.textContent = "Chưa lưu";
    }
}

function setupSaveStatusBadge() {
    const badge = document.getElementById("save-badge");
    if (badge) {
        badge.style.cursor = "pointer";
        badge.title = "Nhấn để lưu file (Ctrl+S)";
        badge.addEventListener("click", () => {
            saveCurrentFile();
        });
    }
}

async function saveCurrentFile() {
    if (appState.isEditing) commitInlineEdit();

    if (appState.filepath) {
        showLoading("Đang lưu trực tiếp vào file...");
        const payload = buildExportPayload(appState.filename);
        payload.filepath = appState.filepath;

        try {
            const res = await fetch("/api/save_in_place", {
                method: "POST",
                headers: { "Content-Type": "application/json" },
                body: JSON.stringify(payload)
            });
            const result = await res.json();
            hideLoading();
            if (result.success) {
                setSaveStatus(true, "Đã lưu");
                showToast(`✅ Đã lưu trực tiếp vào: ${result.filename}!`, "success");
                return;
            }
        } catch (err) {
            console.warn("Lưu trực tiếp thất bại, chuyển sang xuất file:", err);
            hideLoading();
        }
    }

    exportDirectXLSX();
}

function setupGlobalDragAndDrop() {
    window.addEventListener("dragover", (e) => {
        e.preventDefault();
        e.stopPropagation();
    });
    window.addEventListener("drop", (e) => {
        e.preventDefault();
        e.stopPropagation();
        if (e.dataTransfer && e.dataTransfer.files && e.dataTransfer.files.length > 0) {
            const file = e.dataTransfer.files[0];
            const ext = file.name.substring(file.name.lastIndexOf(".")).toLowerCase();
            if ([".xlsx", ".xls", ".xlsm", ".csv"].includes(ext)) {
                importExcelFile(file);
            } else if (file.type.startsWith("image/") || [".png", ".jpg", ".jpeg", ".webp", ".gif", ".svg", ".bmp"].includes(ext)) {
                const gridEl = document.getElementById("grid-element");
                const rect = gridEl ? gridEl.getBoundingClientRect() : { left: 0, top: 0 };
                const x = Math.max(50, e.clientX - rect.left);
                const y = Math.max(40, e.clientY - rect.top);
                handleImageFileUpload(file, x, y);
            } else {
                showToast("Vui lòng kéo thả file Excel (.xlsx, .xls) hoặc file ảnh (.png, .jpg).", "warning");
            }
        }
    });
}

// ==========================================================================
// Helpers: Coordinate conversions (A1 <-> {row, col})
// ==========================================================================
function colIndexToLetter(idx) {
    let letter = "";
    let temp = idx;
    while (temp >= 0) {
        letter = String.fromCharCode((temp % 26) + 65) + letter;
        temp = Math.floor(temp / 26) - 1;
    }
    return letter;
}

function colLetterToIndex(letter) {
    let idx = 0;
    const upper = letter.toUpperCase();
    for (let i = 0; i < upper.length; i++) {
        idx = idx * 26 + (upper.charCodeAt(i) - 64);
    }
    return idx - 1;
}

function coordToA1(r, c) {
    return colIndexToLetter(c) + (r + 1);
}

function parseA1(a1) {
    const match = a1.trim().match(/^([A-Za-z]+)(\d+)$/);
    if (!match) return null;
    return {
        col: colLetterToIndex(match[1]),
        row: parseInt(match[2], 10) - 1
    };
}

function parseRange(rangeStr) {
    const parts = rangeStr.trim().split(":");
    if (parts.length === 1) {
        const c = parseA1(parts[0]);
        if (!c) return null;
        return { startRow: c.row, startCol: c.col, endRow: c.row, endCol: c.col };
    }
    const c1 = parseA1(parts[0]);
    const c2 = parseA1(parts[1]);
    if (!c1 || !c2) return null;
    return {
        startRow: Math.min(c1.row, c2.row),
        startCol: Math.min(c1.col, c2.col),
        endRow: Math.max(c1.row, c2.row),
        endCol: Math.max(c1.col, c2.col)
    };
}

function cellToString(val) {
    if (val === null || val === undefined) return "";
    return String(val);
}

// ==========================================================================
// Merged Cells Management (Merge & Center / Unmerge)
// ==========================================================================
let _activeMergesIndex = null;

function normalizeRange(r) {
    if (!r) return null;
    if (typeof r === "string") {
        return parseRange(r);
    }
    if (r.s && r.e) {
        return {
            startRow: Math.min(r.s.r, r.e.r),
            startCol: Math.min(r.s.c, r.e.c),
            endRow: Math.max(r.s.r, r.e.r),
            endCol: Math.max(r.s.c, r.e.c)
        };
    }
    if (r.startRow !== undefined && r.startCol !== undefined) {
        return {
            startRow: Math.min(r.startRow, r.endRow !== undefined ? r.endRow : r.startRow),
            startCol: Math.min(r.startCol, r.endCol !== undefined ? r.endCol : r.startCol),
            endRow: Math.max(r.startRow, r.endRow !== undefined ? r.endRow : r.startRow),
            endCol: Math.max(r.startCol, r.endCol !== undefined ? r.endCol : r.startCol)
        };
    }
    return null;
}

function getActiveSheetMerges() {
    if (!appState.merges) appState.merges = {};
    if (!appState.merges[appState.activeSheetName]) {
        appState.merges[appState.activeSheetName] = [];
    }
    return appState.merges[appState.activeSheetName];
}

function buildMergesIndex(sheetName = appState.activeSheetName) {
    const map = new Map();
    const list = (appState.merges && appState.merges[sheetName]) || [];
    for (const item of list) {
        const norm = normalizeRange(item);
        if (!norm) continue;
        const rangeStr = coordToA1(norm.startRow, norm.startCol) + ":" + coordToA1(norm.endRow, norm.endCol);
        const info = {
            raw: item,
            rangeStr,
            startRow: norm.startRow,
            startCol: norm.startCol,
            endRow: norm.endRow,
            endCol: norm.endCol
        };
        for (let r = norm.startRow; r <= norm.endRow; r++) {
            for (let c = norm.startCol; c <= norm.endCol; c++) {
                map.set(`${r},${c}`, {
                    ...info,
                    isMaster: (r === norm.startRow && c === norm.startCol)
                });
            }
        }
    }
    return map;
}

function findMergeForCell(r, c, sheetName = appState.activeSheetName) {
    if (sheetName === appState.activeSheetName && _activeMergesIndex) {
        return _activeMergesIndex.get(`${r},${c}`) || null;
    }
    const list = (appState.merges && appState.merges[sheetName]) || [];
    for (const item of list) {
        const norm = normalizeRange(item);
        if (!norm) continue;
        if (r >= norm.startRow && r <= norm.endRow && c >= norm.startCol && c <= norm.endCol) {
            const rangeStr = coordToA1(norm.startRow, norm.startCol) + ":" + coordToA1(norm.endRow, norm.endCol);
            return {
                raw: item,
                rangeStr,
                startRow: norm.startRow,
                startCol: norm.startCol,
                endRow: norm.endRow,
                endCol: norm.endCol,
                isMaster: (r === norm.startRow && c === norm.startCol)
            };
        }
    }
    return null;
}

function rangesOverlap(r1, r2) {
    return !(
        r1.endRow < r2.startRow ||
        r1.startRow > r2.endRow ||
        r1.endCol < r2.startCol ||
        r1.startCol > r2.endCol
    );
}

function updateMergeButtonState() {
    const btn = document.getElementById("btn-merge-cells");
    const ctxText = document.getElementById("ctx-merge-text");
    const { startRow, startCol, endRow, endCol } = appState.selectionRange;
    const merge = findMergeForCell(startRow, startCol);
    const isSingleCell = (startRow === endRow && startCol === endCol);
    const isMerged = !!(merge && (
        isSingleCell ||
        (merge.startRow === startRow && merge.startCol === startCol && merge.endRow === endRow && merge.endCol === endCol)
    ));

    if (btn) {
        if (isMerged) {
            btn.classList.add("active");
            btn.setAttribute("title", "Hủy gộp ô (Unmerge Cells)");
        } else {
            btn.classList.remove("active");
            btn.setAttribute("title", "Hợp nhất ô & Căn giữa (Merge & Center)");
        }
    }
    if (ctxText) {
        ctxText.textContent = isMerged ? "Hủy gộp ô (Unmerge)" : "Gộp ô & Căn giữa (Merge & Center)";
    }
}

function toggleMergeCells() {
    if (appState.isEditing) commitInlineEdit();

    const { startRow, startCol, endRow, endCol } = appState.selectionRange;
    const sheetName = appState.activeSheetName;
    const currentMerges = getActiveSheetMerges();
    const isSingleCellSelected = (startRow === endRow && startCol === endCol);

    // 1. If currently inside an existing merge -> UNMERGE
    const existingMerge = findMergeForCell(startRow, startCol, sheetName);
    if (existingMerge) {
        const isExactMatch = (
            existingMerge.startRow === startRow &&
            existingMerge.startCol === startCol &&
            existingMerge.endRow === endRow &&
            existingMerge.endCol === endCol
        );

        if (isSingleCellSelected || isExactMatch) {
            pushHistory();
            const idx = currentMerges.findIndex(m => {
                const norm = normalizeRange(m);
                return norm && norm.startRow === existingMerge.startRow && norm.startCol === existingMerge.startCol &&
                               norm.endRow === existingMerge.endRow && norm.endCol === existingMerge.endCol;
            });
            if (idx !== -1) {
                currentMerges.splice(idx, 1);
            }
            renderGrid();
            selectCell(existingMerge.startRow, existingMerge.startCol);
            showToast("Đã hủy gộp ô thành công", "info");
            return;
        }
    }

    // 2. If single cell and not merged, require selecting at least 2 cells
    if (isSingleCellSelected) {
        showToast("Vui lòng kéo chọn từ 2 ô trở lên để thực hiện gộp ô.", "warning");
        return;
    }

    // 3. Multi-cell selection -> MERGE & CENTER
    pushHistory();

    const newNorm = {
        startRow: Math.min(startRow, endRow),
        startCol: Math.min(startCol, endCol),
        endRow: Math.max(startRow, endRow),
        endCol: Math.max(startCol, endCol)
    };

    // Remove any overlapping merges
    for (let i = currentMerges.length - 1; i >= 0; i--) {
        const norm = normalizeRange(currentMerges[i]);
        if (norm && rangesOverlap(norm, newNorm)) {
            currentMerges.splice(i, 1);
        }
    }

    // Value handling: keep master value, clear other cells
    const sheetData = appState.sheets[sheetName];
    if (sheetData) {
        let masterVal = sheetData[newNorm.startRow] ? sheetData[newNorm.startRow][newNorm.startCol] : "";
        if (masterVal === "" || masterVal === null || masterVal === undefined) {
            findLoop: for (let r = newNorm.startRow; r <= newNorm.endRow; r++) {
                for (let c = newNorm.startCol; c <= newNorm.endCol; c++) {
                    const v = sheetData[r] ? sheetData[r][c] : "";
                    if (v !== "" && v !== null && v !== undefined) {
                        masterVal = v;
                        break findLoop;
                    }
                }
            }
        }

        for (let r = newNorm.startRow; r <= newNorm.endRow; r++) {
            for (let c = newNorm.startCol; c <= newNorm.endCol; c++) {
                if (r === newNorm.startRow && c === newNorm.startCol) {
                    sheetData[r][c] = masterVal;
                } else {
                    if (sheetData[r]) sheetData[r][c] = "";
                }
            }
        }
    }

    // Apply "Center & Middle" alignment (Merge & Center standard)
    if (!appState.styles[sheetName]) appState.styles[sheetName] = {};
    const sheetStyles = appState.styles[sheetName];
    const masterKey = `${newNorm.startRow},${newNorm.startCol}`;
    if (!sheetStyles[masterKey]) sheetStyles[masterKey] = {};
    sheetStyles[masterKey].align = "center";
    sheetStyles[masterKey].valign = "middle";

    const rangeStr = coordToA1(newNorm.startRow, newNorm.startCol) + ":" + coordToA1(newNorm.endRow, newNorm.endCol);
    currentMerges.push(rangeStr);

    renderGrid();
    selectCell(newNorm.startRow, newNorm.startCol);
    showToast(`Đã gộp ô ${rangeStr} & căn giữa thành công!`, "success");
}

function shiftMergesOnRowInsert(sheetName, rowIdx) {
    const list = (appState.merges && appState.merges[sheetName]) || [];
    const updated = [];
    for (const m of list) {
        const norm = normalizeRange(m);
        if (!norm) continue;
        if (norm.startRow >= rowIdx) {
            norm.startRow++;
            norm.endRow++;
        } else if (norm.endRow >= rowIdx) {
            norm.endRow++;
        }
        updated.push(coordToA1(norm.startRow, norm.startCol) + ":" + coordToA1(norm.endRow, norm.endCol));
    }
    appState.merges[sheetName] = updated;
}

function shiftMergesOnColInsert(sheetName, colIdx) {
    const list = (appState.merges && appState.merges[sheetName]) || [];
    const updated = [];
    for (const m of list) {
        const norm = normalizeRange(m);
        if (!norm) continue;
        if (norm.startCol >= colIdx) {
            norm.startCol++;
            norm.endCol++;
        } else if (norm.endCol >= colIdx) {
            norm.endCol++;
        }
        updated.push(coordToA1(norm.startRow, norm.startCol) + ":" + coordToA1(norm.endRow, norm.endCol));
    }
    appState.merges[sheetName] = updated;
}

function shiftMergesOnRowDelete(sheetName, rowIdx) {
    const list = (appState.merges && appState.merges[sheetName]) || [];
    const updated = [];
    for (const m of list) {
        const norm = normalizeRange(m);
        if (!norm) continue;
        if (norm.startRow === rowIdx && norm.endRow === rowIdx) {
            continue;
        }
        if (norm.startRow > rowIdx) {
            norm.startRow--;
            norm.endRow--;
        } else if (norm.endRow >= rowIdx) {
            norm.endRow--;
        }
        if (norm.startRow <= norm.endRow && norm.startCol <= norm.endCol && (norm.startRow < norm.endRow || norm.startCol < norm.endCol)) {
            updated.push(coordToA1(norm.startRow, norm.startCol) + ":" + coordToA1(norm.endRow, norm.endCol));
        }
    }
    appState.merges[sheetName] = updated;
}

function shiftMergesOnColDelete(sheetName, colIdx) {
    const list = (appState.merges && appState.merges[sheetName]) || [];
    const updated = [];
    for (const m of list) {
        const norm = normalizeRange(m);
        if (!norm) continue;
        if (norm.startCol === colIdx && norm.endCol === colIdx) {
            continue;
        }
        if (norm.startCol > colIdx) {
            norm.startCol--;
            norm.endCol--;
        } else if (norm.endCol >= colIdx) {
            norm.endCol--;
        }
        if (norm.startRow <= norm.endRow && norm.startCol <= norm.endCol && (norm.startRow < norm.endRow || norm.startCol < norm.endCol)) {
            updated.push(coordToA1(norm.startRow, norm.startCol) + ":" + coordToA1(norm.endRow, norm.endCol));
        }
    }
    appState.merges[sheetName] = updated;
}

function showToast(message, type = "info") {
    const container = document.getElementById("toast-container");
    const toast = document.createElement("div");
    toast.className = `toast ${type}`;
    toast.innerHTML = `<span>${message}</span>`;
    container.appendChild(toast);
    setTimeout(() => {
        toast.style.opacity = "0";
        setTimeout(() => toast.remove(), 300);
    }, 3500);
}

function showLoading(msg = "Đang xử lý...") {
    document.getElementById("loading-message").textContent = msg;
    document.getElementById("loading-overlay").classList.remove("hidden");
}

function hideLoading() {
    document.getElementById("loading-overlay").classList.add("hidden");
}

// ==========================================================================
// History (Undo / Redo)
// ==========================================================================
function pushHistory() {
    const snapshot = {
        sheets: JSON.parse(JSON.stringify(appState.sheets)),
        styles: JSON.parse(JSON.stringify(appState.styles)),
        colWidths: JSON.parse(JSON.stringify(appState.colWidths)),
        rowHeights: JSON.parse(JSON.stringify(appState.rowHeights)),
        merges: JSON.parse(JSON.stringify(appState.merges || {})),
        images: JSON.parse(JSON.stringify(appState.images || {})),
        filters: JSON.parse(JSON.stringify(appState.filters || {})),
        activeSheetName: appState.activeSheetName
    };
    appState.undoStack.push(snapshot);
    if (appState.undoStack.length > 30) appState.undoStack.shift();
    appState.redoStack = [];
    updateUndoRedoButtons();
    setSaveStatus(false);
}

function undo() {
    if (appState.undoStack.length === 0) return;
    const current = {
        sheets: JSON.parse(JSON.stringify(appState.sheets)),
        styles: JSON.parse(JSON.stringify(appState.styles)),
        colWidths: JSON.parse(JSON.stringify(appState.colWidths)),
        rowHeights: JSON.parse(JSON.stringify(appState.rowHeights)),
        merges: JSON.parse(JSON.stringify(appState.merges || {})),
        images: JSON.parse(JSON.stringify(appState.images || {})),
        filters: JSON.parse(JSON.stringify(appState.filters || {})),
        activeSheetName: appState.activeSheetName
    };
    appState.redoStack.push(current);
    const prev = appState.undoStack.pop();
    applyStateSnapshot(prev);
    updateUndoRedoButtons();
    showToast("Đã hoàn tác thao tác", "info");
}

function redo() {
    if (appState.redoStack.length === 0) return;
    const current = {
        sheets: JSON.parse(JSON.stringify(appState.sheets)),
        styles: JSON.parse(JSON.stringify(appState.styles)),
        colWidths: JSON.parse(JSON.stringify(appState.colWidths)),
        rowHeights: JSON.parse(JSON.stringify(appState.rowHeights)),
        merges: JSON.parse(JSON.stringify(appState.merges || {})),
        images: JSON.parse(JSON.stringify(appState.images || {})),
        filters: JSON.parse(JSON.stringify(appState.filters || {})),
        activeSheetName: appState.activeSheetName
    };
    appState.undoStack.push(current);
    const next = appState.redoStack.pop();
    applyStateSnapshot(next);
    updateUndoRedoButtons();
    showToast("Đã làm lại thao tác", "info");
}

function applyStateSnapshot(snap) {
    appState.sheets = snap.sheets;
    appState.styles = snap.styles;
    appState.colWidths = snap.colWidths;
    appState.rowHeights = snap.rowHeights;
    appState.merges = snap.merges || {};
    appState.images = snap.images || {};
    appState.filters = snap.filters || {};
    appState.activeSheetName = snap.activeSheetName;
    renderTabs();
    renderGrid();
    renderImages();
    updateFormulaBar();
    updateCharts();
}

function updateUndoRedoButtons() {
    const btnUndo = document.getElementById("btn-undo");
    const btnRedo = document.getElementById("btn-redo");
    if (btnUndo) btnUndo.disabled = appState.undoStack.length === 0;
    if (btnRedo) btnRedo.disabled = appState.redoStack.length === 0;
}

// ==========================================================================
// Templates Engine
// ==========================================================================
function loadTemplate(name) {
    pushHistory();
    if (name === "blank") {
        appState.filename = "Bang_Tinh_Moi.xlsx";
        appState.sheets = { "Sheet1": createEmptyMatrix(100, 26) };
        appState.styles = { "Sheet1": {} };
        appState.colWidths = { "Sheet1": {} };
        appState.rowHeights = { "Sheet1": {} };
        appState.merges = { "Sheet1": [] };
        appState.activeSheetName = "Sheet1";
    } else if (name === "sales") {
        appState.filename = "Bao_Cao_Doanh_Thu_2026.xlsx";
        const rows = createEmptyMatrix(100, 26);
        const styles = {};
        
        // Header title
        rows[0][0] = "BÁO CÁO DOANH THU & LỢI NHUẬN NĂM 2026";
        styles["0,0"] = { bold: true, fontSize: 14, color: "#107c41", align: "center", valign: "middle" };
        
        // Table Columns
        const headers = ["Tháng", "Doanh Thu (Tr.đ)", "Chi Phí (Tr.đ)", "Lợi Nhuận (Tr.đ)", "Tỷ Suất LN (%)", "Đánh Giá"];
        headers.forEach((h, idx) => {
            rows[2][idx] = h;
            styles[`2,${idx}`] = { bold: true, bg: "#107c41", color: "#ffffff", align: "center", border: true };
        });
        
        // Sample Month Data
        const monthsData = [
            ["Tháng 1", 120, 85],
            ["Tháng 2", 145, 92],
            ["Tháng 3", 190, 110],
            ["Tháng 4", 210, 125],
            ["Tháng 5", 240, 130],
            ["Tháng 6", 310, 160],
            ["Tháng 7", 280, 150],
            ["Tháng 8", 295, 155],
            ["Tháng 9", 330, 170],
            ["Tháng 10", 380, 190],
            ["Tháng 11", 420, 210],
            ["Tháng 12", 500, 240]
        ];
        
        monthsData.forEach((item, i) => {
            const r = 3 + i;
            rows[r][0] = item[0];
            rows[r][1] = item[1];
            rows[r][2] = item[2];
            rows[r][3] = `=B${r+1}-C${r+1}`;
            rows[r][4] = `=(D${r+1}/B${r+1})*100`;
            rows[r][5] = `=IF(D${r+1}>150,"Đạt mục tiêu","Cần cải thiện")`;
            
            for (let c = 0; c < 6; c++) {
                styles[`${r},${c}`] = {
                    border: true,
                    align: c === 0 ? "left" : (c === 5 ? "center" : "right")
                };
            }
        });
        
        // Total row
        const totR = 3 + monthsData.length;
        rows[totR][0] = "TỔNG CỘNG";
        rows[totR][1] = `=SUM(B4:B${totR})`;
        rows[totR][2] = `=SUM(C4:C${totR})`;
        rows[totR][3] = `=SUM(D4:D${totR})`;
        rows[totR][4] = `=AVERAGE(E4:E${totR})`;
        rows[totR][5] = `XUẤT SẮC`;
        
        for (let c = 0; c < 6; c++) {
            styles[`${totR},${c}`] = {
                bold: true,
                bg: "#1e293b",
                color: "#38bdf8",
                border: true,
                align: c === 0 ? "left" : "right"
            };
        }
        
        // Summary Card
        rows[2][7] = "THỐNG KÊ NHANH";
        styles["2,7"] = { bold: true, bg: "#334155", color: "#ffffff", align: "center", border: true };
        rows[3][7] = "Tháng cao nhất"; rows[3][8] = "=MAX(B4:B15)";
        rows[4][7] = "Tháng thấp nhất"; rows[4][8] = "=MIN(B4:B15)";
        rows[5][7] = "Doanh thu TB"; rows[5][8] = "=AVERAGE(B4:B15)";
        rows[6][7] = "Số tháng theo dõi"; rows[6][8] = "=COUNT(B4:B15)";
        
        for (let r = 3; r <= 6; r++) {
            styles[`${r},7`] = { border: true, bold: true };
            styles[`${r},8`] = { border: true, align: "right" };
        }
        
        appState.sheets = { "DoanhThu": rows, "GhiChu": createEmptyMatrix(25, 12) };
        appState.styles = { "DoanhThu": styles, "GhiChu": {} };
        appState.colWidths = { "DoanhThu": { "A": 110, "B": 130, "C": 130, "D": 130, "E": 120, "F": 130, "H": 130, "I": 110 } };
        appState.rowHeights = { "DoanhThu": {} };
        appState.merges = { "DoanhThu": ["A1:F1"], "GhiChu": [] };
        appState.activeSheetName = "DoanhThu";
        
        // Auto create an initial demonstration floating chart!
        setTimeout(() => {
            createFloatingChart({
                type: "bar",
                title: "Biểu đồ Doanh Thu & Lợi Nhuận 2026",
                range: "A3:D9",
                palette: "emerald",
                x: 620,
                y: 120,
                width: 480,
                height: 300
            });
        }, 150);

    } else if (name === "budget") {
        appState.filename = "Ngan_Sach_Chi_Tieu_Gia_Dinh.xlsx";
        const rows = createEmptyMatrix(100, 26);
        const styles = {};
        
        rows[0][0] = "KẾ HOẠCH NGÂN SÁCH THÁNG 10/2026";
        styles["0,0"] = { bold: true, fontSize: 13, color: "#2563eb" };
        
        const headers = ["Khoản Mục Chi Tiêu", "Ngân Sách (₫)", "Thực Chi (₫)", "Chênh Lệch (₫)", "Tỷ Lệ (%)"];
        headers.forEach((h, idx) => {
            rows[2][idx] = h;
            styles[`2,${idx}`] = { bold: true, bg: "#2563eb", color: "#ffffff", align: "center", border: true };
        });
        
        const items = [
            ["Tiền nhà & Tiện ích", 8000000, 7800000],
            ["Ăn uống & Siêu thị", 7000000, 7500000],
            ["Học phí con cái", 5000000, 5000000],
            ["Giao thông & Xăng xe", 2000000, 1800000],
            ["Giải trí & Mua sắm", 3000000, 3400000],
            ["Tiết kiệm đầu tư", 10000000, 10000000],
            ["Quỹ dự phòng y tế", 2000000, 1200000]
        ];
        
        items.forEach((item, i) => {
            const r = 3 + i;
            rows[r][0] = item[0];
            rows[r][1] = item[1];
            rows[r][2] = item[2];
            rows[r][3] = `=B${r+1}-C${r+1}`;
            rows[r][4] = `=(C${r+1}/B${r+1})*100`;
            for (let c = 0; c < 5; c++) styles[`${r},${c}`] = { border: true, align: c === 0 ? "left" : "right" };
        });
        
        const tot = 3 + items.length;
        rows[tot][0] = "TỔNG CHI PHÍ";
        rows[tot][1] = `=SUM(B4:B${tot})`;
        rows[tot][2] = `=SUM(C4:C${tot})`;
        rows[tot][3] = `=B${tot+1}-C${tot+1}`;
        rows[tot][4] = `=AVERAGE(E4:E${tot})`;
        for (let c = 0; c < 5; c++) styles[`${tot},${c}`] = { bold: true, bg: "#1e293b", color: "#38bdf8", border: true, align: c === 0 ? "left" : "right" };
        
        appState.sheets = { "ChiTieu": rows };
        appState.styles = { "ChiTieu": styles };
        appState.colWidths = { "ChiTieu": { "A": 160, "B": 130, "C": 130, "D": 130, "E": 110 } };
        appState.rowHeights = { "ChiTieu": {} };
        appState.activeSheetName = "ChiTieu";
        
        setTimeout(() => {
            createFloatingChart({
                type: "doughnut",
                title: "Tỷ Trọng Thực Chi Tiêu",
                range: "A3:C10",
                palette: "modern",
                x: 620,
                y: 100,
                width: 440,
                height: 310
            });
        }, 150);

    } else if (name === "grades") {
        appState.filename = "Bang_Diem_Hoc_Ky.xlsx";
        const rows = createEmptyMatrix(100, 26);
        const styles = {};
        
        rows[0][0] = "BẢNG ĐIỂM TỔNG HỢP HỌC KỲ I - LỚP K65";
        styles["0,0"] = { bold: true, fontSize: 13, color: "#9333ea" };
        
        const headers = ["Mã SV", "Họ và Tên", "Toán", "Lý", "Hóa", "Điểm TB", "Xếp Loại"];
        headers.forEach((h, idx) => {
            rows[2][idx] = h;
            styles[`2,${idx}`] = { bold: true, bg: "#7e22ce", color: "#ffffff", align: "center", border: true };
        });
        
        const students = [
            ["SV001", "Nguyễn Văn An", 8.5, 7.5, 9.0],
            ["SV002", "Trần Thị Bình", 9.0, 9.5, 8.5],
            ["SV003", "Lê Hoàng Cường", 6.5, 7.0, 6.0],
            ["SV004", "Phạm Hải Đăng", 8.0, 8.5, 8.0],
            ["SV005", "Đặng Thùy Dung", 9.5, 9.0, 9.5],
            ["SV006", "Hoàng Gia Huy", 5.5, 6.0, 5.0],
            ["SV007", "Vũ Minh Khôi", 7.5, 8.0, 7.5]
        ];
        
        students.forEach((stu, i) => {
            const r = 3 + i;
            rows[r][0] = stu[0];
            rows[r][1] = stu[1];
            rows[r][2] = stu[2];
            rows[r][3] = stu[3];
            rows[r][4] = stu[4];
            rows[r][5] = `=AVERAGE(C${r+1}:E${r+1})`;
            rows[r][6] = `=IF(F${r+1}>=8.5,"Xuất sắc",IF(F${r+1}>=7.0,"Khá","Trung bình"))`;
            for (let c = 0; c < 7; c++) styles[`${r},${c}`] = { border: true, align: c < 2 ? "left" : (c === 6 ? "center" : "right") };
        });
        
        appState.sheets = { "BangDiem": rows };
        appState.styles = { "BangDiem": styles };
        appState.colWidths = { "BangDiem": { "A": 90, "B": 160, "C": 80, "D": 80, "E": 80, "F": 90, "G": 110 } };
        appState.rowHeights = { "BangDiem": {} };
        appState.activeSheetName = "BangDiem";
    }

    document.getElementById("workbook-filename").value = appState.filename;
    renderTabs();
    renderGrid();
    selectCell(3, 1);
    showToast(`Đã tải mẫu: ${appState.filename}`, "success");
}

function createEmptyMatrix(rows, cols) {
    return Array.from({ length: rows }, () => Array(cols).fill(""));
}

// ==========================================================================
// Ribbon Tabs Switching
// ==========================================================================
function setupRibbonTabs() {
    const tabBtns = document.querySelectorAll(".ribbon-tab-btn");
    tabBtns.forEach(btn => {
        btn.addEventListener("click", () => {
            tabBtns.forEach(b => b.classList.remove("active"));
            btn.classList.add("active");
            
            const targetPaneId = btn.getAttribute("data-tab");
            document.querySelectorAll(".ribbon-pane").forEach(pane => {
                pane.classList.add("hidden");
                pane.classList.remove("active");
            });
            const targetPane = document.getElementById(targetPaneId);
            if (targetPane) {
                targetPane.classList.remove("hidden");
                targetPane.classList.add("active");
            }
        });
    });

    // File tab action buttons
    document.getElementById("btn-file-new").addEventListener("click", () => loadTemplate("blank"));
    document.getElementById("btn-file-open").addEventListener("click", () => openImportModal());
    document.getElementById("btn-template-sales").addEventListener("click", () => loadTemplate("sales"));
    document.getElementById("btn-template-budget").addEventListener("click", () => loadTemplate("budget"));
    document.getElementById("btn-template-grades").addEventListener("click", () => loadTemplate("grades"));
    
    // Quick buttons in topbar
    document.getElementById("btn-quick-import").addEventListener("click", () => openImportModal());
    document.getElementById("btn-quick-chart").addEventListener("click", () => openChartModal());

    // Direct 1-Click Export and Options Dropdown
    const btnQuickExport = document.getElementById("btn-quick-export");
    const btnExportCaret = document.getElementById("btn-export-dropdown-caret");
    
    // Direct 1-click download when clicking the main "Xuất .xlsx" button
    if (btnQuickExport) {
        btnQuickExport.addEventListener("click", () => {
            exportDirectXLSX();
        });
    }

    // Direct save / download button in Ribbon > File tab
    const btnFileSaveDownload = document.getElementById("btn-file-save-download");
    if (btnFileSaveDownload) {
        btnFileSaveDownload.addEventListener("click", () => {
            saveCurrentFile();
        });
    }

    // Toggle dropdown options when clicking the small caret
    if (btnExportCaret) {
        btnExportCaret.addEventListener("click", (e) => {
            e.stopPropagation();
            btnExportCaret.parentElement.classList.toggle("open");
        });
    }

    document.addEventListener("click", (e) => {
        if (!e.target.closest(".dropdown-container")) {
            document.querySelectorAll(".dropdown-container").forEach(c => c.classList.remove("open"));
        }
    });

    // Export options dropdown items
    document.getElementById("btn-export-direct").addEventListener("click", () => {
        if (btnExportCaret) btnExportCaret.parentElement.classList.remove("open");
        exportDirectXLSX();
    });
    document.getElementById("btn-export-server").addEventListener("click", () => {
        if (btnExportCaret) btnExportCaret.parentElement.classList.remove("open");
        exportServerXLSX();
    });
    document.getElementById("btn-export-csv").addEventListener("click", () => {
        if (btnExportCaret) btnExportCaret.parentElement.classList.remove("open");
        exportCSV();
    });

    // Auto-fit rows button in Ribbon
    const btnAutofitRows = document.getElementById("btn-autofit-rows");
    if (btnAutofitRows) {
        btnAutofitRows.addEventListener("click", autoFitAllRows);
    }

    // Undo / Redo
    document.getElementById("btn-undo").addEventListener("click", undo);
    document.getElementById("btn-redo").addEventListener("click", redo);
    
    // Recalculate all button
    const btnRecalc = document.getElementById("btn-recalculate-all");
    if (btnRecalc) {
        btnRecalc.addEventListener("click", () => {
            renderGrid();
            updateCharts();
            showToast("Đã tính toán lại toàn bộ bảng tính!", "success");
        });
    }

    // Workbook rename
    const nameInput = document.getElementById("workbook-filename");
    nameInput.addEventListener("change", () => {
        let val = nameInput.value.trim();
        if (!val) val = "Untitled.xlsx";
        if (!val.endsWith(".xlsx")) val += ".xlsx";
        appState.filename = val;
        nameInput.value = val;
        showToast(`Đổi tên thành công: ${val}`, "info");
    });
}

// ==========================================================================
// Grid Rendering
// ==========================================================================
// Grid Rendering & Infinite Grid Engine
// ==========================================================================
function isRowMatchingFilters(r, sheetData, sheetFilters) {
    if (r === 0) return true; // Row 0 is header row, never filtered out
    if (!sheetFilters) return true;
    const filterCols = Object.keys(sheetFilters);
    if (filterCols.length === 0) return true;

    for (const colStr of filterCols) {
        const allowed = sheetFilters[colStr];
        if (!Array.isArray(allowed)) continue;
        const col = parseInt(colStr, 10);
        const cellRaw = sheetData[r] ? sheetData[r][col] : "";
        const cellVal = cellToString(cellRaw).trim();
        if (!allowed.includes(cellVal)) {
            return false;
        }
    }
    return true;
}

function createRowElement(r, colCount, sheetData, sheetStyles, colWidths, rowHeights) {
    const rowDiv = document.createElement("div");
    rowDiv.className = "grid-row";
    const height = rowHeights[r] || 26;
    rowDiv.style.height = `${height}px`;

    // Filter check: hide row if it doesn't match active sheet filters
    const sheetFilters = appState.filters[appState.activeSheetName];
    if (r > 0 && sheetFilters && !isRowMatchingFilters(r, sheetData, sheetFilters)) {
        rowDiv.style.display = "none";
        rowDiv.classList.add("row-filtered-out");
    }

    // Row Header (Number)
    const rowHeader = document.createElement("div");
    rowHeader.className = "grid-header-cell row-header";
    rowHeader.id = `row-header-${r}`;
    rowHeader.setAttribute("data-row", r);
    rowHeader.innerHTML = `<span>${r + 1}</span>`;

    // Row Resize Handle
    const rowResizeHandle = document.createElement("div");
    rowResizeHandle.className = "row-resize-handle";
    setupRowResize(rowResizeHandle, r);
    rowHeader.appendChild(rowResizeHandle);

    rowHeader.addEventListener("click", () => selectEntireRow(r));
    rowDiv.appendChild(rowHeader);

    // Data Cells in Row
    for (let c = 0; c < colCount; c++) {
        const cellDiv = document.createElement("div");
        cellDiv.className = "grid-cell";
        cellDiv.id = `cell-${r}-${c}`;
        cellDiv.setAttribute("data-row", r);
        cellDiv.setAttribute("data-col", c);

        const width = colWidths[colIndexToLetter(c)] || 110;
        const merge = findMergeForCell(r, c, appState.activeSheetName);

        if (merge) {
            if (merge.isMaster) {
                // Master cell: span width and height across merged region
                let totalWidth = 0;
                for (let col = merge.startCol; col <= merge.endCol; col++) {
                    totalWidth += (colWidths[colIndexToLetter(col)] || 110);
                }
                let totalHeight = 0;
                for (let row = merge.startRow; row <= merge.endRow; row++) {
                    totalHeight += (rowHeights[row] || 26);
                }
                cellDiv.style.width = `${totalWidth}px`;
                cellDiv.style.minWidth = `${totalWidth}px`;
                cellDiv.style.height = `${totalHeight}px`;
                cellDiv.classList.add("merged-master-cell");
                cellDiv.setAttribute("data-merged", "true");
                cellDiv.setAttribute("data-merge-range", merge.rangeStr);

                const rawVal = sheetData[r] ? sheetData[r][c] : "";
                const computedVal = displayValue(rawVal, r, c);
                cellDiv.textContent = computedVal;

                const styleKey = `${r},${c}`;
                if (sheetStyles[styleKey]) {
                    applyStyleToElement(cellDiv, sheetStyles[styleKey]);
                }
                rowDiv.appendChild(cellDiv);
            } else if (r === merge.startRow) {
                // Covered cell in the same row: hidden completely
                cellDiv.style.display = "none";
                cellDiv.classList.add("merged-hidden-cell");
                cellDiv.setAttribute("data-master-row", merge.startRow);
                cellDiv.setAttribute("data-master-col", merge.startCol);
                rowDiv.appendChild(cellDiv);
            } else {
                // Covered cell in subsequent rows: takes width for flex alignment, but invisible
                cellDiv.style.width = `${width}px`;
                cellDiv.style.minWidth = `${width}px`;
                cellDiv.style.height = `${height}px`;
                cellDiv.style.visibility = "hidden";
                cellDiv.style.pointerEvents = "none";
                cellDiv.classList.add("merged-covered-cell");
                cellDiv.setAttribute("data-master-row", merge.startRow);
                cellDiv.setAttribute("data-master-col", merge.startCol);
                rowDiv.appendChild(cellDiv);
            }
        } else {
            // Normal unmerged cell
            cellDiv.style.width = `${width}px`;
            cellDiv.style.minWidth = `${width}px`;
            cellDiv.style.height = `${height}px`;

            const rawVal = sheetData[r] ? sheetData[r][c] : "";
            const computedVal = displayValue(rawVal, r, c);
            cellDiv.textContent = computedVal;

            const styleKey = `${r},${c}`;
            if (sheetStyles[styleKey]) {
                applyStyleToElement(cellDiv, sheetStyles[styleKey]);
            }
            rowDiv.appendChild(cellDiv);
        }
    }

    return rowDiv;
}

function renderGrid() {
    const gridEl = document.getElementById("grid-element");
    gridEl.innerHTML = "";

    const sheetData = appState.sheets[appState.activeSheetName];
    if (!sheetData || sheetData.length === 0) return;

    _activeMergesIndex = buildMergesIndex(appState.activeSheetName);

    const rowCount = sheetData.length;
    const colCount = sheetData[0] ? sheetData[0].length : 0;
    const sheetStyles = appState.styles[appState.activeSheetName] || {};
    const colWidths = appState.colWidths[appState.activeSheetName] || {};
    const rowHeights = appState.rowHeights[appState.activeSheetName] || {};

    // 1. Column Header Row (Top)
    const headerRow = document.createElement("div");
    headerRow.className = "grid-row col-headers-row";
    
    // Corner select-all cell
    const corner = document.createElement("div");
    corner.className = "grid-header-cell corner";
    corner.innerHTML = '<span style="font-size: 10px; opacity: 0.6;">&#9698;</span>';
    corner.addEventListener("click", selectAllCells);
    headerRow.appendChild(corner);

    for (let c = 0; c < colCount; c++) {
        const colLetter = colIndexToLetter(c);
        const colHeader = document.createElement("div");
        colHeader.className = "grid-header-cell col-header";
        colHeader.id = `col-header-${c}`;
        colHeader.setAttribute("data-col", c);
        
        // Custom width if set
        const width = colWidths[colLetter] || 110;
        colHeader.style.width = `${width}px`;
        colHeader.style.minWidth = `${width}px`;
        
        colHeader.innerHTML = `<span>${colLetter}</span>`;

        // Filter button indicator
        const isColFiltered = !!(appState.filters[appState.activeSheetName]?.[c]);
        const filterBtn = document.createElement("button");
        filterBtn.className = "col-filter-btn" + (isColFiltered ? " active" : "");
        if (appState.filterMode || isColFiltered) {
            filterBtn.style.display = "block";
        }
        filterBtn.innerHTML = isColFiltered ? "&#9660;" : "&#9662;";
        filterBtn.title = isColFiltered ? `Cột ${colLetter} đang được lọc (Nhấn để chỉnh sửa hoặc bỏ lọc)` : `Lọc cột ${colLetter}`;
        filterBtn.addEventListener("click", (e) => {
            e.stopPropagation();
            openFilterPopover(c, colHeader);
        });
        colHeader.appendChild(filterBtn);

        // Column Resize Handle
        const colResizeHandle = document.createElement("div");
        colResizeHandle.className = "col-resize-handle";
        setupColResize(colResizeHandle, c);
        colHeader.appendChild(colResizeHandle);

        colHeader.addEventListener("click", () => selectEntireColumn(c));
        headerRow.appendChild(colHeader);
    }
    gridEl.appendChild(headerRow);

    // 2. Data Rows
    const fragment = document.createDocumentFragment();
    for (let r = 0; r < rowCount; r++) {
        fragment.appendChild(createRowElement(r, colCount, sheetData, sheetStyles, colWidths, rowHeights));
    }
    gridEl.appendChild(fragment);

    updateGridFooterInfo();
    highlightSelection();
    renderImages();
}

function appendRowsToActiveSheet(count = 50) {
    const sheetData = appState.sheets[appState.activeSheetName];
    if (!sheetData) return;
    const colCount = sheetData[0] ? sheetData[0].length : 26;
    const sheetStyles = appState.styles[appState.activeSheetName] || {};
    const colWidths = appState.colWidths[appState.activeSheetName] || {};
    const rowHeights = appState.rowHeights[appState.activeSheetName] || {};
    const gridEl = document.getElementById("grid-element");

    const oldRowCount = sheetData.length;
    for (let i = 0; i < count; i++) {
        sheetData.push(new Array(colCount).fill(""));
    }

    const fragment = document.createDocumentFragment();
    for (let r = oldRowCount; r < sheetData.length; r++) {
        fragment.appendChild(createRowElement(r, colCount, sheetData, sheetStyles, colWidths, rowHeights));
    }
    gridEl.appendChild(fragment);
    updateGridFooterInfo();
}

function appendColsToActiveSheet(count = 10) {
    const sheetData = appState.sheets[appState.activeSheetName];
    if (!sheetData || sheetData.length === 0) return;
    const oldColCount = sheetData[0].length;
    const newColCount = oldColCount + count;
    for (let r = 0; r < sheetData.length; r++) {
        for (let c = oldColCount; c < newColCount; c++) {
            sheetData[r].push("");
        }
    }
    renderGrid();
    updateGridFooterInfo();
}

function updateGridFooterInfo() {
    const sheetData = appState.sheets[appState.activeSheetName];
    if (!sheetData) return;
    const rowCount = sheetData.length;
    const colCount = sheetData[0] ? sheetData[0].length : 0;
    const infoEl = document.getElementById("grid-rows-cols-count");
    if (infoEl) {
        const lastCol = colIndexToLetter(Math.max(0, colCount - 1));
        const sheetFilters = appState.filters[appState.activeSheetName] || {};
        const activeFilterCols = Object.keys(sheetFilters).filter(c => Array.isArray(sheetFilters[c]));
        if (activeFilterCols.length > 0) {
            let visibleCount = 0;
            for (let r = 0; r < rowCount; r++) {
                if (isRowMatchingFilters(r, sheetData, sheetFilters)) visibleCount++;
            }
            infoEl.textContent = `Đang lọc: ${visibleCount}/${rowCount} hàng × ${colCount} cột (${lastCol})`;
        } else {
            infoEl.textContent = `Tổng: ${rowCount} hàng × ${colCount} cột (${lastCol})`;
        }
    }
}

function setupInfiniteGridControls() {
    const viewport = document.getElementById("grid-scroll-viewport");
    if (viewport) {
        let expandTimer = null;
        viewport.addEventListener("scroll", () => {
            if (expandTimer) return;
            expandTimer = setTimeout(() => {
                expandTimer = null;
                // Infinite vertical scroll: when within 300px of bottom, append 50 rows
                if (viewport.scrollTop + viewport.clientHeight >= viewport.scrollHeight - 300) {
                    appendRowsToActiveSheet(50);
                }
                // Infinite horizontal scroll: when within 300px of right edge, append 10 columns
                if (viewport.scrollLeft + viewport.clientWidth >= viewport.scrollWidth - 300) {
                    appendColsToActiveSheet(10);
                }
            }, 60);
        }, { passive: true });
    }

    const btnAdd50 = document.getElementById("btn-add-50-rows");
    if (btnAdd50) {
        btnAdd50.addEventListener("click", () => {
            appendRowsToActiveSheet(50);
            showToast("Đã thêm 50 hàng mới!", "info");
        });
    }

    const btnAdd100 = document.getElementById("btn-add-100-rows");
    if (btnAdd100) {
        btnAdd100.addEventListener("click", () => {
            appendRowsToActiveSheet(100);
            showToast("Đã thêm 100 hàng mới!", "info");
        });
    }

    const btnAdd10Cols = document.getElementById("btn-add-10-cols");
    if (btnAdd10Cols) {
        btnAdd10Cols.addEventListener("click", () => {
            appendColsToActiveSheet(10);
            showToast("Đã thêm 10 cột mới!", "info");
        });
    }
}

function applyStyleToElement(el, s) {
    if (s.bold) el.style.fontWeight = "700";
    if (s.italic) el.style.fontStyle = "italic";
    if (s.underline) el.style.textDecoration = "underline";
    if (s.strikethrough) el.style.textDecoration = (el.style.textDecoration ? el.style.textDecoration + " " : "") + "line-through";
    if (s.fontSize) el.style.fontSize = `${s.fontSize}px`;
    if (s.fontFamily) el.style.fontFamily = s.fontFamily;
    if (s.color) el.style.color = s.color;
    if (s.bg) el.style.backgroundColor = s.bg;
    if (s.align) el.style.justifyContent = s.align === "center" ? "center" : (s.align === "right" ? "flex-end" : "flex-start");
    if (s.valign) el.style.alignItems = s.valign === "top" ? "flex-start" : (s.valign === "bottom" ? "flex-end" : "center");
    if (s.border) el.style.border = "1px solid var(--border-color)";
    if (s.wrap) {
        el.style.whiteSpace = "normal";
        el.style.wordBreak = "break-word";
    }
}

// ==========================================================================
// Cell Selection & Drag-to-Select (High Performance Event Delegation)
// ==========================================================================
function setupGridEventDelegation() {
    const gridEl = document.getElementById("grid-element");
    if (!gridEl) return;

    gridEl.addEventListener("mousedown", (e) => {
        const cell = e.target.closest(".grid-cell");
        if (cell) {
            const r = parseInt(cell.getAttribute("data-row"), 10);
            const c = parseInt(cell.getAttribute("data-col"), 10);
            if (!isNaN(r) && !isNaN(c)) onCellMouseDown(e, r, c);
        }
    });

    gridEl.addEventListener("mouseover", (e) => {
        if (!appState.isSelecting) return;
        const cell = e.target.closest(".grid-cell");
        if (cell) {
            const r = parseInt(cell.getAttribute("data-row"), 10);
            const c = parseInt(cell.getAttribute("data-col"), 10);
            if (!isNaN(r) && !isNaN(c)) onCellMouseEnter(r, c);
        }
    });

    gridEl.addEventListener("dblclick", (e) => {
        const cell = e.target.closest(".grid-cell");
        if (cell) {
            const r = parseInt(cell.getAttribute("data-row"), 10);
            const c = parseInt(cell.getAttribute("data-col"), 10);
            if (!isNaN(r) && !isNaN(c)) startInlineEditing(r, c);
        }
    });
}

function onCellMouseDown(e, r, c) {
    if (e.button !== 0 && e.button !== 2) return; // Left or Right click
    if (appState.isEditing) commitInlineEdit();

    const merge = findMergeForCell(r, c);

    if (e.shiftKey && appState.selectedCell) {
        // Range select with Shift
        const targetEndR = merge ? (r >= appState.selectedCell.row ? merge.endRow : merge.startRow) : r;
        const targetEndC = merge ? (c >= appState.selectedCell.col ? merge.endCol : merge.startCol) : c;
        appState.selectionRange = {
            startRow: Math.min(appState.selectedCell.row, targetEndR, merge ? merge.startRow : r),
            startCol: Math.min(appState.selectedCell.col, targetEndC, merge ? merge.startCol : c),
            endRow: Math.max(appState.selectedCell.row, targetEndR, merge ? merge.endRow : r),
            endCol: Math.max(appState.selectedCell.col, targetEndC, merge ? merge.endCol : c)
        };
        highlightSelection();
        updateFormulaBar();
        updateStatusBarStats();
        return;
    }

    if (merge) {
        appState.selectedCell = { row: merge.startRow, col: merge.startCol };
        appState.selectionRange = {
            startRow: merge.startRow,
            startCol: merge.startCol,
            endRow: merge.endRow,
            endCol: merge.endCol
        };
    } else {
        appState.selectedCell = { row: r, col: c };
        appState.selectionRange = { startRow: r, startCol: c, endRow: r, endCol: c };
    }
    appState.isSelecting = true;

    highlightSelection();
    updateFormulaBar();
    updateStatusBarStats();
}

function onCellMouseEnter(r, c) {
    if (!appState.isSelecting) return;
    const merge = findMergeForCell(r, c);
    const targetEndR = merge ? (r >= appState.selectedCell.row ? merge.endRow : merge.startRow) : r;
    const targetEndC = merge ? (c >= appState.selectedCell.col ? merge.endCol : merge.startCol) : c;

    appState.selectionRange = {
        startRow: Math.min(appState.selectedCell.row, targetEndR, merge ? merge.startRow : r),
        startCol: Math.min(appState.selectedCell.col, targetEndC, merge ? merge.startCol : c),
        endRow: Math.max(appState.selectedCell.row, targetEndR, merge ? merge.endRow : r),
        endCol: Math.max(appState.selectedCell.col, targetEndC, merge ? merge.endCol : c)
    };
    highlightSelection();
    updateFormulaBar();
    updateStatusBarStats();
}

document.addEventListener("mouseup", () => {
    if (appState.isSelecting) {
        appState.isSelecting = false;
    }
});

function selectCell(r, c) {
    if (appState.isEditing) commitInlineEdit();
    const merge = findMergeForCell(r, c);
    if (merge) {
        appState.selectedCell = { row: merge.startRow, col: merge.startCol };
        appState.selectionRange = {
            startRow: merge.startRow,
            startCol: merge.startCol,
            endRow: merge.endRow,
            endCol: merge.endCol
        };
    } else {
        appState.selectedCell = { row: r, col: c };
        appState.selectionRange = { startRow: r, startCol: c, endRow: r, endCol: c };
    }
    highlightSelection();
    updateFormulaBar();
    updateStatusBarStats();
    scrollCellIntoView(r, c);
}

function scrollCellIntoView(r, c) {
    const cellEl = document.getElementById(`cell-${r}-${c}`);
    const viewport = document.getElementById("grid-scroll-viewport");
    if (!cellEl || !viewport) return;

    const cellRect = cellEl.getBoundingClientRect();
    const viewRect = viewport.getBoundingClientRect();
    const headerHeight = 32;
    const rowHeaderWidth = 46;

    if (cellRect.bottom > viewRect.bottom) {
        viewport.scrollTop += (cellRect.bottom - viewRect.bottom) + 10;
    } else if (cellRect.top < viewRect.top + headerHeight) {
        viewport.scrollTop -= (viewRect.top + headerHeight - cellRect.top) + 10;
    }

    if (cellRect.right > viewRect.right) {
        viewport.scrollLeft += (cellRect.right - viewRect.right) + 20;
    } else if (cellRect.left < viewRect.left + rowHeaderWidth) {
        viewport.scrollLeft -= (viewRect.left + rowHeaderWidth - cellRect.left) + 20;
    }
}

function selectEntireColumn(c) {
    const rowCount = appState.sheets[appState.activeSheetName].length;
    appState.selectedCell = { row: 0, col: c };
    appState.selectionRange = { startRow: 0, startCol: c, endRow: rowCount - 1, endCol: c };
    highlightSelection();
    updateFormulaBar();
    updateStatusBarStats();
}

function selectEntireRow(r) {
    const colCount = appState.sheets[appState.activeSheetName][0].length;
    appState.selectedCell = { row: r, col: 0 };
    appState.selectionRange = { startRow: r, startCol: 0, endRow: r, endCol: colCount - 1 };
    highlightSelection();
    updateFormulaBar();
    updateStatusBarStats();
}

function selectAllCells() {
    const sheet = appState.sheets[appState.activeSheetName];
    appState.selectedCell = { row: 0, col: 0 };
    appState.selectionRange = { startRow: 0, startCol: 0, endRow: sheet.length - 1, endCol: sheet[0].length - 1 };
    highlightSelection();
    updateFormulaBar();
    updateStatusBarStats();
}

function highlightSelection() {
    // Clear previous selection classes
    document.querySelectorAll(".grid-cell.selected").forEach(el => el.classList.remove("selected"));
    document.querySelectorAll(".grid-cell.active-focus").forEach(el => el.classList.remove("active-focus"));
    document.querySelectorAll(".fill-handle").forEach(el => el.remove());
    document.querySelectorAll(".grid-header-cell.active-header").forEach(el => el.classList.remove("active-header"));

    const { startRow, startCol, endRow, endCol } = appState.selectionRange;
    const focusRow = appState.selectedCell ? appState.selectedCell.row : startRow;
    const focusCol = appState.selectedCell ? appState.selectedCell.col : startCol;

    for (let r = startRow; r <= endRow; r++) {
        const rowH = document.getElementById(`row-header-${r}`);
        if (rowH) rowH.classList.add("active-header");

        for (let c = startCol; c <= endCol; c++) {
            const cellEl = document.getElementById(`cell-${r}-${c}`);
            if (cellEl) {
                cellEl.classList.add("selected");
                if (r === focusRow && c === focusCol) {
                    cellEl.classList.add("active-focus");
                }
            }
        }
    }

    for (let c = startCol; c <= endCol; c++) {
        const colH = document.getElementById(`col-header-${c}`);
        if (colH) colH.classList.add("active-header");
    }

    // Append Drag-to-Fill handle at the bottom-right corner cell of selection
    let brCell = document.getElementById(`cell-${endRow}-${endCol}`);
    if (!brCell || brCell.style.display === "none" || brCell.style.visibility === "hidden") {
        const focusCellEl = document.getElementById(`cell-${focusRow}-${focusCol}`);
        if (focusCellEl && focusCellEl.classList.contains("merged-master-cell")) {
            brCell = focusCellEl;
        }
    }
    if (brCell) {
        const handle = document.createElement("div");
        handle.className = "fill-handle";
        handle.title = "Kéo để sao chép công thức hoặc chuỗi số tự động";
        handle.addEventListener("mousedown", (e) => {
            e.stopPropagation();
            startDragFill();
        });
        brCell.appendChild(handle);
    }
    updateMergeButtonState();
}

// Smart Drag to Fill Functionality (Handles numbers, formulas, and strings like ref-retro-1 -> ref-retro-2)
function startDragFill(e) {
    if (e) {
        e.stopPropagation();
        e.preventDefault();
    }
    appState.dragFill.active = true;
    appState.dragFill.startRow = appState.selectionRange.startRow;
    appState.dragFill.startCol = appState.selectionRange.startCol;
    appState.dragFill.endRow = appState.selectionRange.endRow;
    appState.dragFill.endCol = appState.selectionRange.endCol;
    appState.dragFill.targetRow = appState.selectionRange.endRow;
    appState.dragFill.targetCol = appState.selectionRange.endCol;

    function onFillMouseMove(moveEvt) {
        if (!appState.dragFill.active) return;
        
        // Auto-scroll viewport when dragging near edges
        const viewport = document.getElementById("grid-scroll-viewport");
        if (viewport) {
            const vRect = viewport.getBoundingClientRect();
            if (moveEvt.clientY > vRect.bottom - 40) {
                viewport.scrollTop += 25;
                if (viewport.scrollTop + viewport.clientHeight >= viewport.scrollHeight - 150) {
                    appendRowsToActiveSheet(25);
                }
            } else if (moveEvt.clientY < vRect.top + 40 && viewport.scrollTop > 0) {
                viewport.scrollTop -= 25;
            }
        }

        const el = document.elementFromPoint(moveEvt.clientX, moveEvt.clientY);
        const cell = el ? el.closest(".grid-cell") : null;
        let curR = appState.dragFill.targetRow;
        let curC = appState.dragFill.targetCol;

        if (cell) {
            const r = parseInt(cell.getAttribute("data-row"), 10);
            const c = parseInt(cell.getAttribute("data-col"), 10);
            if (!isNaN(r)) curR = r;
            if (!isNaN(c)) curC = c;
        }

        appState.dragFill.targetRow = curR;
        appState.dragFill.targetCol = curC;

        // Clear previous fill previews
        document.querySelectorAll(".grid-cell.fill-preview").forEach(c => c.classList.remove("fill-preview"));

        const minR = Math.min(appState.dragFill.startRow, curR);
        const maxR = Math.max(appState.dragFill.endRow, curR);
        const minC = Math.min(appState.dragFill.startCol, curC);
        const maxC = Math.max(appState.dragFill.endCol, curC);

        for (let r = minR; r <= maxR; r++) {
            for (let c = minC; c <= maxC; c++) {
                const cEl = document.getElementById(`cell-${r}-${c}`);
                if (cEl) cEl.classList.add("fill-preview");
            }
        }
    }

    function onFillMouseUp() {
        document.removeEventListener("mousemove", onFillMouseMove);
        document.removeEventListener("mouseup", onFillMouseUp);
        document.querySelectorAll(".grid-cell.fill-preview").forEach(c => c.classList.remove("fill-preview"));
        if (appState.dragFill.active) {
            appState.dragFill.active = false;
            completeSmartDragFill(appState.dragFill.targetRow, appState.dragFill.targetCol);
        }
    }

    document.addEventListener("mousemove", onFillMouseMove);
    document.addEventListener("mouseup", onFillMouseUp);
}

function completeSmartDragFill(targetRow, targetCol) {
    const src = {
        startRow: appState.dragFill.startRow,
        startCol: appState.dragFill.startCol,
        endRow: appState.dragFill.endRow,
        endCol: appState.dragFill.endCol
    };
    const sheet = appState.sheets[appState.activeSheetName];
    if (!sheet) return;

    // Check if dragging DOWN
    if (targetRow > src.endRow) {
        pushHistory();
        const numNewRows = targetRow - src.endRow;

        // Auto expand matrix if dragging past existing rows
        while (sheet.length <= targetRow) {
            sheet.push(Array(sheet[0].length).fill(""));
        }

        for (let c = src.startCol; c <= src.endCol; c++) {
            const srcValues = [];
            for (let r = src.startRow; r <= src.endRow; r++) {
                srcValues.push(sheet[r][c]);
            }

            for (let i = 1; i <= numNewRows; i++) {
                const fillRow = src.endRow + i;
                sheet[fillRow][c] = generateNextSeriesValue(srcValues, i);
            }
        }

        appState.selectionRange.endRow = targetRow;
        renderGrid();
        updateFormulaBar();
        updateCharts();
        showToast("Đã tự động điền chuỗi dữ liệu (Auto-fill) thành công!", "success");
    }
    // Check if dragging RIGHT
    else if (targetCol > src.endCol) {
        pushHistory();
        const numNewCols = targetCol - src.endCol;

        sheet.forEach(row => {
            while (row.length <= targetCol) {
                row.push("");
            }
        });

        for (let r = src.startRow; r <= src.endRow; r++) {
            const srcValues = [];
            for (let c = src.startCol; c <= src.endCol; c++) {
                srcValues.push(sheet[r][c]);
            }

            for (let i = 1; i <= numNewCols; i++) {
                const fillCol = src.endCol + i;
                sheet[r][fillCol] = generateNextSeriesValue(srcValues, i);
            }
        }

        appState.selectionRange.endCol = targetCol;
        renderGrid();
        updateFormulaBar();
        updateCharts();
        showToast("Đã tự động điền dữ liệu sang ngang thành công!", "success");
    }
}

// Smart Value Progression (ref-retro-1 -> ref-retro-2, ref-code1 -> ref-code2, 1 -> 2, formulas offset)
function generateNextSeriesValue(sourceVals, step) {
    if (!sourceVals || sourceVals.length === 0) return "";
    const lastRaw = sourceVals[sourceVals.length - 1];
    const lastStr = cellToString(lastRaw);

    // 1. Formula reference offset
    if (lastStr.startsWith("=")) {
        return offsetFormulaRows(lastStr, step);
    }

    // 2. Alphanumeric series ending with digits or digits before file extension (e.g. ref-retro-1, ref-code1, Image_01.webp)
    const match = lastStr.match(/^(.*?)(\d+)(\.[a-zA-Z0-9]+)?$/);
    if (match) {
        const prefix = match[1];
        const numStr = match[2];
        const ext = match[3] || "";

        let numStep = 1;
        if (sourceVals.length >= 2) {
            const prevStr = cellToString(sourceVals[sourceVals.length - 2]);
            const prevMatch = prevStr.match(/^(.*?)(\d+)(\.[a-zA-Z0-9]+)?$/);
            if (prevMatch && prevMatch[1] === prefix) {
                numStep = parseInt(numStr, 10) - parseInt(prevMatch[2], 10);
            }
        }

        const nextNum = Math.max(0, parseInt(numStr, 10) + numStep * step);
        const padded = String(nextNum).padStart(numStr.length, "0");
        return prefix + padded + ext;
    }

    // 3. Pure Number progression
    const num = parseFloat(lastStr);
    if (!isNaN(num) && String(num) === lastStr.trim()) {
        let diff = 1;
        if (sourceVals.length >= 2) {
            const prevNum = parseFloat(sourceVals[sourceVals.length - 2]);
            if (!isNaN(prevNum)) diff = num - prevNum;
        }
        return num + diff * step;
    }

    // 4. Default: replicate value
    return lastRaw;
}

function offsetFormulaRows(formulaStr, offset) {
    return formulaStr.replace(/([A-Z]+)(\d+)/g, (match, col, row) => {
        return col + (parseInt(row, 10) + offset);
    });
}

// ==========================================================================
// Column and Row Resizing
// ==========================================================================
function setupColResize(handle, colIdx) {
    handle.addEventListener("mousedown", (e) => {
        e.stopPropagation();
        e.preventDefault();
        const startX = e.clientX;
        const colLetter = colIndexToLetter(colIdx);
        const colHeader = document.getElementById(`col-header-${colIdx}`);
        const startWidth = colHeader ? colHeader.offsetWidth : 110;
        const guide = document.getElementById("col-resize-guide");

        guide.classList.remove("hidden");
        guide.style.left = `${e.clientX}px`;

        function onMouseMove(moveEvent) {
            guide.style.left = `${moveEvent.clientX}px`;
        }

        function onMouseUp(upEvent) {
            document.removeEventListener("mousemove", onMouseMove);
            document.removeEventListener("mouseup", onMouseUp);
            guide.classList.add("hidden");

            const delta = upEvent.clientX - startX;
            const newWidth = Math.max(40, startWidth + delta);

            pushHistory();
            if (!appState.colWidths[appState.activeSheetName]) appState.colWidths[appState.activeSheetName] = {};
            appState.colWidths[appState.activeSheetName][colLetter] = newWidth;
            renderGrid();
        }

        document.addEventListener("mousemove", onMouseMove);
        document.addEventListener("mouseup", onMouseUp);
    });
}

function setupRowResize(handle, rowIdx) {
    // Drag row height
    handle.addEventListener("mousedown", (e) => {
        e.stopPropagation();
        e.preventDefault();
        const startY = e.clientY;
        const rowHeaderEl = document.getElementById(`row-header-${rowIdx}`);
        const rowDiv = rowHeaderEl ? rowHeaderEl.parentElement : null;
        const startHeight = rowDiv ? rowDiv.offsetHeight : 26;
        const guide = document.getElementById("row-resize-guide");

        guide.classList.remove("hidden");
        guide.style.top = `${e.clientY}px`;

        function onMouseMove(moveEvent) {
            guide.style.top = `${moveEvent.clientY}px`;
        }

        function onMouseUp(upEvent) {
            document.removeEventListener("mousemove", onMouseMove);
            document.removeEventListener("mouseup", onMouseUp);
            guide.classList.add("hidden");

            const delta = upEvent.clientY - startY;
            const newHeight = Math.max(22, startHeight + delta);

            pushHistory();
            if (!appState.rowHeights[appState.activeSheetName]) appState.rowHeights[appState.activeSheetName] = {};
            appState.rowHeights[appState.activeSheetName][rowIdx] = newHeight;
            renderGrid();
        }

        document.addEventListener("mousemove", onMouseMove);
        document.addEventListener("mouseup", onMouseUp);
    });

    // Double-click row boundary to auto-fit row height
    handle.addEventListener("dblclick", (e) => {
        e.stopPropagation();
        autoFitSingleRowHeight(rowIdx);
    });
}

function autoFitSingleRowHeight(rowIdx) {
    const sheet = appState.sheets[appState.activeSheetName];
    if (!sheet || !sheet[rowIdx]) return;
    pushHistory();

    let maxLen = 0;
    for (let c = 0; c < sheet[rowIdx].length; c++) {
        const str = cellToString(sheet[rowIdx][c]);
        if (str.length > maxLen) maxLen = str.length;
    }

    const lines = Math.max(1, Math.ceil(maxLen / 35));
    const newHeight = Math.min(Math.max(26, lines * 20 + 8), 500);

    if (!appState.rowHeights[appState.activeSheetName]) appState.rowHeights[appState.activeSheetName] = {};
    appState.rowHeights[appState.activeSheetName][rowIdx] = newHeight;
    renderGrid();
    showToast(`Đã dãn hàng ${rowIdx + 1} thành ${newHeight}px`, "info");
}

function autoFitAllRows() {
    const sheet = appState.sheets[appState.activeSheetName];
    if (!sheet) return;
    pushHistory();
    if (!appState.rowHeights[appState.activeSheetName]) appState.rowHeights[appState.activeSheetName] = {};

    const { startRow, endRow } = appState.selectionRange;
    const isSingle = startRow === endRow;
    const targetStart = isSingle ? 0 : startRow;
    const targetEnd = isSingle ? sheet.length - 1 : endRow;

    for (let r = targetStart; r <= targetEnd; r++) {
        let maxLen = 0;
        for (let c = 0; c < sheet[r].length; c++) {
            const str = cellToString(sheet[r][c]);
            if (str.length > maxLen) maxLen = str.length;
        }
        if (maxLen > 30) {
            const lines = Math.ceil(maxLen / 35);
            appState.rowHeights[appState.activeSheetName][r] = Math.min(Math.max(26, lines * 20 + 8), 500);
        } else {
            appState.rowHeights[appState.activeSheetName][r] = 26;
        }
    }
    renderGrid();
    showToast("Đã tự động dãn độ cao các dòng theo nội dung!", "success");
}

// ==========================================================================
// Inline & Formula Bar Editing
// ==========================================================================
function startInlineEditing(r, c, initialChar = null) {
    if (appState.isEditing) commitInlineEdit();
    const merge = findMergeForCell(r, c);
    const targetR = merge ? merge.startRow : r;
    const targetC = merge ? merge.startCol : c;

    appState.isEditing = true;
    document.getElementById("status-mode-indicator").textContent = "CHỈNH SỬA";

    const cellEl = document.getElementById(`cell-${targetR}-${targetC}`);
    if (!cellEl) return;

    const sheet = appState.sheets[appState.activeSheetName];
    const currentRawVal = cellToString(sheet[targetR][targetC]);

    const input = document.createElement("input");
    input.type = "text";
    input.className = "grid-cell-input";
    
    if (initialChar !== null) {
        input.value = initialChar;
    } else {
        input.value = currentRawVal;
    }

    cellEl.appendChild(input);
    input.focus();
    if (initialChar !== null) {
        input.setSelectionRange(input.value.length, input.value.length);
    } else {
        input.select();
    }

    input.addEventListener("keydown", (e) => {
        if (e.key === "Enter") {
            e.preventDefault();
            e.stopPropagation();
            e.stopImmediatePropagation();
            const sheet = appState.sheets[appState.activeSheetName];
            const nextRow = e.shiftKey ? Math.max(0, (merge ? merge.startRow : targetR) - 1) : (merge ? merge.endRow : targetR) + 1;
            if (nextRow >= sheet.length) {
                appendRowsToActiveSheet(25);
            }
            commitInlineEdit({ row: nextRow, col: targetC });
        } else if (e.key === "Tab") {
            e.preventDefault();
            e.stopPropagation();
            e.stopImmediatePropagation();
            const sheet = appState.sheets[appState.activeSheetName];
            const nextCol = e.shiftKey ? Math.max(0, (merge ? merge.startCol : targetC) - 1) : (merge ? merge.endCol : targetC) + 1;
            if (nextCol >= sheet[0].length) {
                appendColsToActiveSheet(10);
            }
            commitInlineEdit({ row: targetR, col: nextCol });
        } else if (e.key === "ArrowDown") {
            e.preventDefault();
            e.stopPropagation();
            e.stopImmediatePropagation();
            const sheet = appState.sheets[appState.activeSheetName];
            const nextRow = (merge ? merge.endRow : targetR) + 1;
            if (nextRow >= sheet.length) {
                appendRowsToActiveSheet(25);
            }
            commitInlineEdit({ row: nextRow, col: targetC });
        } else if (e.key === "ArrowUp") {
            e.preventDefault();
            e.stopPropagation();
            e.stopImmediatePropagation();
            const nextRow = Math.max(0, (merge ? merge.startRow : targetR) - 1);
            commitInlineEdit({ row: nextRow, col: targetC });
        } else if (e.key === "Escape") {
            e.preventDefault();
            e.stopPropagation();
            e.stopImmediatePropagation();
            input.onblur = null;
            if (input.parentNode) input.parentNode.removeChild(input);
            appState.isEditing = false;
            document.getElementById("status-mode-indicator").textContent = "SẴN SÀNG";
            updateFormulaBar();
        }
    });

    input.addEventListener("blur", () => {
        if (appState.isEditing) {
            commitInlineEdit();
        }
    });

    input.addEventListener("input", () => {
        document.getElementById("formula-input-field").value = input.value;
        checkFormulaAutocomplete(input.value);
    });
}

function commitInlineEdit(andNavigate = null) {
    if (!appState.isEditing || appState.isCommitting) return;
    appState.isCommitting = true;

    const { row, col } = appState.selectedCell;
    const cellEl = document.getElementById(`cell-${row}-${col}`);
    const input = cellEl ? cellEl.querySelector(".grid-cell-input") : null;

    if (input) {
        const newVal = input.value;
        input.onblur = null;
        if (input.parentNode) {
            input.parentNode.removeChild(input);
        }
        pushHistory();
        appState.sheets[appState.activeSheetName][row][col] = newVal;
        if (cellEl) {
            cellEl.textContent = displayValue(newVal, row, col);
            const sheetStyles = appState.styles[appState.activeSheetName] || {};
            if (sheetStyles[`${row},${col}`]) {
                applyStyleToElement(cellEl, sheetStyles[`${row},${col}`]);
            }
        }
    }

    appState.isEditing = false;
    appState.isCommitting = false;
    document.getElementById("status-mode-indicator").textContent = "SẴN SÀNG";
    updateFormulaBar();
    updateCharts();
    hideAutocomplete();

    if (andNavigate) {
        selectCell(andNavigate.row, andNavigate.col);
    }
}

function setupFormulaBar() {
    const formulaInput = document.getElementById("formula-input-field");
    const nameBox = document.getElementById("name-box-input");

    formulaInput.addEventListener("keydown", (e) => {
        if (e.key === "Enter") {
            e.preventDefault();
            const { row, col } = appState.selectedCell;
            pushHistory();
            appState.sheets[appState.activeSheetName][row][col] = formulaInput.value;
            renderGrid();
            updateCharts();
            formulaInput.blur();
            hideAutocomplete();
        } else if (e.key === "Escape") {
            updateFormulaBar();
            formulaInput.blur();
            hideAutocomplete();
        }
    });

    formulaInput.addEventListener("input", () => {
        checkFormulaAutocomplete(formulaInput.value);
    });

    // Jump to coordinate via Name Box
    nameBox.addEventListener("keydown", (e) => {
        if (e.key === "Enter") {
            e.preventDefault();
            const coord = parseA1(nameBox.value);
            if (coord) {
                const sheet = appState.sheets[appState.activeSheetName];
                if (coord.row >= 0 && coord.row < sheet.length && coord.col >= 0 && coord.col < sheet[0].length) {
                    selectCell(coord.row, coord.col);
                    const cellEl = document.getElementById(`cell-${coord.row}-${coord.col}`);
                    if (cellEl) cellEl.scrollIntoView({ behavior: "smooth", block: "center", inline: "center" });
                }
            }
        }
    });

    document.getElementById("btn-formula-cancel").addEventListener("click", () => {
        updateFormulaBar();
        hideAutocomplete();
    });

    document.getElementById("btn-formula-confirm").addEventListener("click", () => {
        const { row, col } = appState.selectedCell;
        pushHistory();
        appState.sheets[appState.activeSheetName][row][col] = formulaInput.value;
        renderGrid();
        updateCharts();
        hideAutocomplete();
    });

    document.getElementById("btn-formula-wizard-trigger").addEventListener("click", openFuncWizard);
    document.getElementById("btn-formula-fx").addEventListener("click", openFuncWizard);
    document.getElementById("btn-insert-func-wizard").addEventListener("click", openFuncWizard);
}

function updateFormulaBar() {
    const { startRow, startCol, endRow, endCol } = appState.selectionRange;
    const nameBox = document.getElementById("name-box-input");
    const formulaInput = document.getElementById("formula-input-field");

    if (!nameBox || !formulaInput) return;

    const sheet = appState.sheets[appState.activeSheetName];
    if (!sheet || !sheet[startRow]) {
        nameBox.value = coordToA1(startRow, startCol);
        formulaInput.value = "";
        updateMergeButtonState();
        return;
    }

    const merge = findMergeForCell(startRow, startCol);
    const isSingleMerged = merge && merge.startRow === startRow && merge.startCol === startCol && merge.endRow === endRow && merge.endCol === endCol;

    if ((startRow === endRow && startCol === endCol) || isSingleMerged) {
        nameBox.value = coordToA1(startRow, startCol);
        const rawVal = sheet[startRow][startCol];
        formulaInput.value = cellToString(rawVal);
    } else {
        nameBox.value = `${coordToA1(startRow, startCol)}:${coordToA1(endRow, endCol)}`;
        formulaInput.value = "";
    }
    updateMergeButtonState();
}

// ==========================================================================
// Formula Autocomplete
// ==========================================================================
function checkFormulaAutocomplete(text) {
    const box = document.getElementById("formula-autocomplete");
    if (!text.startsWith("=")) {
        box.classList.add("hidden");
        return;
    }

    const match = text.match(/=([A-Z]*)$/i);
    if (!match) {
        box.classList.add("hidden");
        return;
    }

    const query = match[1].toUpperCase();
    const suggestions = FORMULA_DEFINITIONS.filter(f => f.name.startsWith(query));

    if (suggestions.length === 0) {
        box.classList.add("hidden");
        return;
    }

    box.innerHTML = "";
    suggestions.slice(0, 8).forEach(s => {
        const item = document.createElement("div");
        item.className = "autocomplete-item";
        item.innerHTML = `<b>${s.name}</b> <span class="desc">${s.syntax}</span>`;
        item.addEventListener("click", () => {
            insertFormulaSuggestion(s.name);
        });
        box.appendChild(item);
    });

    box.classList.remove("hidden");
}

function insertFormulaSuggestion(funcName) {
    const formulaInput = document.getElementById("formula-input-field");
    formulaInput.value = `=${funcName}(`;
    formulaInput.focus();
    hideAutocomplete();
}

function hideAutocomplete() {
    const box = document.getElementById("formula-autocomplete");
    if (box) box.classList.add("hidden");
}

// ==========================================================================
// Status Bar Selection Statistics
// ==========================================================================
function updateStatusBarStats() {
    const { startRow, startCol, endRow, endCol } = appState.selectionRange;
    const sheet = appState.sheets[appState.activeSheetName];
    if (!sheet) return;
    
    let cellCount = 0;
    const numbers = [];

    for (let r = startRow; r <= endRow; r++) {
        for (let c = startCol; c <= endCol; c++) {
            cellCount++;
            const val = sheet[r][c];
            const disp = displayValue(val, r, c);
            const num = parseFloat(disp);
            if (!isNaN(num)) numbers.push(num);
        }
    }

    document.getElementById("stat-count").innerHTML = `Số ô: <b>${cellCount}</b>`;

    if (numbers.length > 0) {
        const sum = numbers.reduce((a, b) => a + b, 0);
        const avg = sum / numbers.length;
        document.getElementById("stat-avg").innerHTML = `Trung bình: <b>${avg.toLocaleString("vi-VN", { maximumFractionDigits: 2 })}</b>`;
        document.getElementById("stat-sum").innerHTML = `Tổng: <b>${sum.toLocaleString("vi-VN", { maximumFractionDigits: 2 })}</b>`;
    } else {
        document.getElementById("stat-avg").innerHTML = `Trung bình: <b>-</b>`;
        document.getElementById("stat-sum").innerHTML = `Tổng: <b>-</b>`;
    }
}

// ==========================================================================
// Formula Engine Evaluation
// ==========================================================================
function displayValue(val, r, c) {
    const str = cellToString(val);
    if (!str.startsWith("=")) return str;
    try {
        const evaluated = evaluateFormula(str, r, c);
        return evaluated !== null && evaluated !== undefined ? String(evaluated) : "#VALUE!";
    } catch (e) {
        return "#ERROR!";
    }
}

let evaluationCallStack = [];

function evaluateFormula(formulaStr, currentR, currentC) {
    const cellKey = `${appState.activeSheetName}!${currentR},${currentC}`;
    if (evaluationCallStack.includes(cellKey)) {
        throw new Error("Circular reference");
    }
    evaluationCallStack.push(cellKey);

    try {
        let expr = formulaStr.substring(1).trim();

        // 1. Check for Top-Level Functions (e.g. SUM(A1:A5), IF(A1>0, "Yes", "No"))
        const funcMatch = expr.match(/^([A-Z_]+)\s*\((.*)\)$/i);
        if (funcMatch) {
            const funcName = funcMatch[1].toUpperCase();
            const rawArgs = funcMatch[2].trim();
            const result = executeFunction(funcName, rawArgs, currentR, currentC);
            if (result !== undefined) {
                evaluationCallStack.pop();
                return result;
            }
        }

        // 2. Arithmetic expressions with cell coordinates (e.g. A1 + B2 * 10 - C3)
        const cellRegex = /([A-Za-z]+)(\d+)/g;
        let mathExpr = expr.replace(cellRegex, (match, colLetter, rowNum) => {
            const cIdx = colLetterToIndex(colLetter);
            const rIdx = parseInt(rowNum, 10) - 1;
            const sheet = appState.sheets[appState.activeSheetName];
            if (rIdx < 0 || rIdx >= sheet.length || cIdx < 0 || cIdx >= sheet[0].length) return 0;
            let val = sheet[rIdx][cIdx];
            if (cellToString(val).startsWith("=")) val = evaluateFormula(val, rIdx, cIdx);
            const num = parseFloat(val);
            return isNaN(num) ? 0 : num;
        });

        const sanitized = mathExpr.replace(/[^0-9+\-*/().^% \s]/g, "").replace(/\^/g, "**");
        // eslint-disable-next-line no-new-func
        const finalVal = Function('"use strict"; return (' + sanitized + ')')();

        evaluationCallStack.pop();
        return isNaN(finalVal) ? "#VALUE!" : (Number.isInteger(finalVal) ? finalVal : parseFloat(finalVal.toFixed(4)));

    } catch (e) {
        evaluationCallStack.pop();
        return "#ERROR!";
    }
}

function executeFunction(name, rawArgs, curR, curC) {
    const args = parseFunctionArguments(rawArgs);

    switch (name) {
        case "SUM": {
            const vals = resolveRangeOrArgs(args, curR, curC, true);
            return vals.reduce((a, b) => a + b, 0);
        }
        case "AVERAGE":
        case "AVG": {
            const vals = resolveRangeOrArgs(args, curR, curC, true);
            return vals.length ? vals.reduce((a, b) => a + b, 0) / vals.length : 0;
        }
        case "MIN": {
            const vals = resolveRangeOrArgs(args, curR, curC, true);
            return vals.length ? Math.min(...vals) : 0;
        }
        case "MAX": {
            const vals = resolveRangeOrArgs(args, curR, curC, true);
            return vals.length ? Math.max(...vals) : 0;
        }
        case "COUNT": {
            const vals = resolveRangeOrArgs(args, curR, curC, true);
            return vals.length;
        }
        case "COUNTA": {
            const vals = resolveRangeOrArgs(args, curR, curC, false);
            return vals.filter(v => v !== "" && v !== null && v !== undefined).length;
        }
        case "COUNTBLANK": {
            const vals = resolveRangeOrArgs(args, curR, curC, false);
            return vals.filter(v => v === "" || v === null || v === undefined).length;
        }
        case "PRODUCT": {
            const vals = resolveRangeOrArgs(args, curR, curC, true);
            return vals.reduce((a, b) => a * b, 1);
        }
        case "ROUND": {
            const num = evaluateSingleArg(args[0], curR, curC);
            const digits = args[1] ? parseInt(evaluateSingleArg(args[1], curR, curC), 10) : 0;
            return Number(Math.round(Number(num + "e" + digits)) + "e-" + digits);
        }
        case "ROUNDUP": {
            const num = evaluateSingleArg(args[0], curR, curC);
            const digits = args[1] ? parseInt(evaluateSingleArg(args[1], curR, curC), 10) : 0;
            return Number(Math.ceil(Number(num + "e" + digits)) + "e-" + digits);
        }
        case "ROUNDDOWN": {
            const num = evaluateSingleArg(args[0], curR, curC);
            const digits = args[1] ? parseInt(evaluateSingleArg(args[1], curR, curC), 10) : 0;
            return Number(Math.floor(Number(num + "e" + digits)) + "e-" + digits);
        }
        case "ABS": {
            return Math.abs(parseFloat(evaluateSingleArg(args[0], curR, curC)) || 0);
        }
        case "SQRT": {
            return Math.sqrt(parseFloat(evaluateSingleArg(args[0], curR, curC)) || 0);
        }
        case "POWER": {
            const base = parseFloat(evaluateSingleArg(args[0], curR, curC)) || 0;
            const exp = parseFloat(evaluateSingleArg(args[1], curR, curC)) || 1;
            return Math.pow(base, exp);
        }
        case "MOD": {
            const n = parseFloat(evaluateSingleArg(args[0], curR, curC)) || 0;
            const d = parseFloat(evaluateSingleArg(args[1], curR, curC)) || 1;
            return n % d;
        }
        case "IF": {
            const cond = evaluateCondition(args[0], curR, curC);
            return cond ? evaluateSingleArg(args[1], curR, curC) : (args[2] ? evaluateSingleArg(args[2], curR, curC) : "");
        }
        case "IFS": {
            for (let i = 0; i < args.length; i += 2) {
                if (evaluateCondition(args[i], curR, curC)) {
                    return evaluateSingleArg(args[i + 1], curR, curC);
                }
            }
            return "#N/A";
        }
        case "AND": {
            return args.every(a => evaluateCondition(a, curR, curC));
        }
        case "OR": {
            return args.some(a => evaluateCondition(a, curR, curC));
        }
        case "NOT": {
            return !evaluateCondition(args[0], curR, curC);
        }
        case "CONCAT":
        case "CONCATENATE": {
            const vals = resolveRangeOrArgs(args, curR, curC, false);
            return vals.join("");
        }
        case "UPPER": {
            return String(evaluateSingleArg(args[0], curR, curC)).toUpperCase();
        }
        case "LOWER": {
            return String(evaluateSingleArg(args[0], curR, curC)).toLowerCase();
        }
        case "PROPER": {
            const str = String(evaluateSingleArg(args[0], curR, curC));
            return str.replace(/\b\w/g, l => l.toUpperCase());
        }
        case "LEN": {
            return String(evaluateSingleArg(args[0], curR, curC)).length;
        }
        case "TRIM": {
            return String(evaluateSingleArg(args[0], curR, curC)).trim().replace(/\s+/g, " ");
        }
        case "LEFT": {
            const str = String(evaluateSingleArg(args[0], curR, curC));
            const n = args[1] ? parseInt(evaluateSingleArg(args[1], curR, curC), 10) : 1;
            return str.substring(0, n);
        }
        case "RIGHT": {
            const str = String(evaluateSingleArg(args[0], curR, curC));
            const n = args[1] ? parseInt(evaluateSingleArg(args[1], curR, curC), 10) : 1;
            return str.substring(str.length - n);
        }
        case "MID": {
            const str = String(evaluateSingleArg(args[0], curR, curC));
            const start = parseInt(evaluateSingleArg(args[1], curR, curC), 10) - 1;
            const len = parseInt(evaluateSingleArg(args[2], curR, curC), 10);
            return str.substring(start, start + len);
        }
        case "VLOOKUP": {
            const lookupVal = evaluateSingleArg(args[0], curR, curC);
            const range = parseRange(args[1]);
            const colIndex = parseInt(args[2], 10) - 1;
            if (!range) return "#REF!";
            const sheet = appState.sheets[appState.activeSheetName];
            for (let r = range.startRow; r <= range.endRow; r++) {
                const rowKeyVal = displayValue(sheet[r][range.startCol], r, range.startCol);
                if (String(rowKeyVal).trim().toLowerCase() === String(lookupVal).trim().toLowerCase()) {
                    const targetCol = range.startCol + colIndex;
                    return displayValue(sheet[r][targetCol], r, targetCol);
                }
            }
            return "#N/A";
        }
        case "INDEX": {
            const range = parseRange(args[0]);
            const rOffset = parseInt(evaluateSingleArg(args[1], curR, curC), 10) - 1;
            const cOffset = args[2] ? parseInt(evaluateSingleArg(args[2], curR, curC), 10) - 1 : 0;
            if (!range) return "#REF!";
            const sheet = appState.sheets[appState.activeSheetName];
            const targetR = range.startRow + rOffset;
            const targetC = range.startCol + cOffset;
            return displayValue(sheet[targetR][targetC], targetR, targetC);
        }
        case "MATCH": {
            const lookupVal = String(evaluateSingleArg(args[0], curR, curC)).trim().toLowerCase();
            const range = parseRange(args[1]);
            if (!range) return "#REF!";
            const sheet = appState.sheets[appState.activeSheetName];
            let pos = 1;
            for (let r = range.startRow; r <= range.endRow; r++) {
                for (let c = range.startCol; c <= range.endCol; c++) {
                    const v = String(displayValue(sheet[r][c], r, c)).trim().toLowerCase();
                    if (v === lookupVal) return pos;
                    pos++;
                }
            }
            return "#N/A";
        }
        case "TODAY": {
            const d = new Date();
            return `${d.getDate()}/${d.getMonth() + 1}/${d.getFullYear()}`;
        }
        case "NOW": {
            const d = new Date();
            return `${d.getHours()}:${d.getMinutes()} ${d.getDate()}/${d.getMonth() + 1}/${d.getFullYear()}`;
        }
        case "YEAR": {
            const val = evaluateSingleArg(args[0], curR, curC);
            const d = new Date(val);
            return isNaN(d.getFullYear()) ? new Date().getFullYear() : d.getFullYear();
        }
        default:
            return undefined;
    }
}

function parseFunctionArguments(rawArgs) {
    const args = [];
    let depth = 0;
    let inQuotes = false;
    let current = "";

    for (let i = 0; i < rawArgs.length; i++) {
        const char = rawArgs[i];
        if (char === '"' && rawArgs[i - 1] !== '\\') inQuotes = !inQuotes;
        if (!inQuotes) {
            if (char === "(") depth++;
            else if (char === ")") depth--;
            else if (char === "," && depth === 0) {
                args.push(current.trim());
                current = "";
                continue;
            }
        }
        current += char;
    }
    if (current.trim()) args.push(current.trim());
    return args;
}

function resolveRangeOrArgs(args, curR, curC, numericOnly = true) {
    const sheet = appState.sheets[appState.activeSheetName];
    const results = [];

    args.forEach(arg => {
        if (arg.includes(":")) {
            const range = parseRange(arg);
            if (range) {
                for (let r = range.startRow; r <= range.endRow; r++) {
                    for (let c = range.startCol; c <= range.endCol; c++) {
                        let val = sheet[r][c];
                        if (cellToString(val).startsWith("=")) val = evaluateFormula(val, r, c);
                        if (numericOnly) {
                            const num = parseFloat(val);
                            if (!isNaN(num)) results.push(num);
                        } else {
                            results.push(val);
                        }
                    }
                }
            }
        } else {
            const single = evaluateSingleArg(arg, curR, curC);
            if (numericOnly) {
                const num = parseFloat(single);
                if (!isNaN(num)) results.push(num);
            } else {
                results.push(single);
            }
        }
    });

    return results;
}

function evaluateSingleArg(arg, curR, curC) {
    if (!arg) return "";
    arg = arg.trim();

    // Quoted string literal
    if ((arg.startsWith('"') && arg.endsWith('"')) || (arg.startsWith("'") && arg.endsWith("'"))) {
        return arg.substring(1, arg.length - 1);
    }

    // Direct cell reference (A1)
    const coord = parseA1(arg);
    if (coord) {
        const sheet = appState.sheets[appState.activeSheetName];
        let val = sheet[coord.row][coord.col];
        if (cellToString(val).startsWith("=")) val = evaluateFormula(val, coord.row, coord.col);
        return val;
    }

    // Number literal
    if (!isNaN(parseFloat(arg)) && isFinite(arg)) {
        return parseFloat(arg);
    }

    // Nested Formula function (e.g. AVERAGE(B2:B5))
    if (arg.startsWith("=")) {
        return evaluateFormula(arg, curR, curC);
    }

    return arg;
}

function evaluateCondition(condStr, curR, curC) {
    const operators = [">=", "<=", "<>", "!=", "=", ">", "<"];
    for (const op of operators) {
        if (condStr.includes(op)) {
            const [leftPart, rightPart] = condStr.split(op);
            const leftVal = evaluateSingleArg(leftPart, curR, curC);
            const rightVal = evaluateSingleArg(rightPart, curR, curC);

            const leftNum = parseFloat(leftVal);
            const rightNum = parseFloat(rightVal);
            const isNumeric = !isNaN(leftNum) && !isNaN(rightNum);

            const l = isNumeric ? leftNum : String(leftVal).trim();
            const r = isNumeric ? rightNum : String(rightVal).trim();

            if (op === "=") return l === r;
            if (op === "<>" || op === "!=") return l !== r;
            if (op === ">") return l > r;
            if (op === "<") return l < r;
            if (op === ">=") return l >= r;
            if (op === "<=") return l <= r;
        }
    }
    return Boolean(evaluateSingleArg(condStr, curR, curC));
}

// ==========================================================================
// Floating Draggable Chart Engine (Chart.js)
// ==========================================================================
function openChartModal() {
    const { startRow, startCol, endRow, endCol } = appState.selectionRange;
    document.getElementById("chart-range-input").value = `${coordToA1(startRow, startCol)}:${coordToA1(endRow, endCol)}`;
    document.getElementById("chart-creator-modal").classList.remove("hidden");
    updateChartPreview();
}

let previewChartInstance = null;

function updateChartPreview() {
    const canvas = document.getElementById("chart-preview-canvas");
    const rangeStr = document.getElementById("chart-range-input").value;
    const type = document.getElementById("chart-type-select").value;
    const palette = document.getElementById("chart-palette-select").value;
    const firstRowHeader = document.getElementById("chart-first-row-header").checked;

    const chartData = extractChartDataFromRange(rangeStr, firstRowHeader, palette);

    if (previewChartInstance) previewChartInstance.destroy();
    previewChartInstance = new Chart(canvas, {
        type: type,
        data: chartData,
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: { legend: { position: "top" } }
        }
    });
}

function createFloatingChart(config) {
    const container = document.getElementById("floating-charts-container");
    const chartId = "chart-" + Date.now();

    const card = document.createElement("div");
    card.className = "floating-chart-card";
    card.id = chartId;
    card.style.left = `${config.x || 100}px`;
    card.style.top = `${config.y || 80}px`;
    card.style.width = `${config.width || 440}px`;
    card.style.height = `${config.height || 300}px`;

    card.innerHTML = `
        <div class="chart-card-header">
            <div class="chart-title-area">
                <span class="chart-card-title">${config.title}</span>
                <span class="chart-range-tag">${config.range}</span>
            </div>
            <div class="chart-card-actions">
                <button class="btn-chart-action btn-download-chart" title="Tải ảnh PNG">&#128247;</button>
                <button class="btn-chart-action btn-close-chart" title="Xóa biểu đồ">&times;</button>
            </div>
        </div>
        <div class="chart-canvas-container">
            <canvas id="canvas-${chartId}"></canvas>
        </div>
    `;

    container.appendChild(card);

    // Make chart card draggable
    makeElementDraggable(card, card.querySelector(".chart-card-header"));

    // Render Chart.js
    const canvas = card.querySelector(`#canvas-${chartId}`);
    const chartData = extractChartDataFromRange(config.range, true, config.palette || "emerald");
    const chartInstance = new Chart(canvas, {
        type: config.type || "bar",
        data: chartData,
        options: {
            responsive: true,
            maintainAspectRatio: false,
            animation: { duration: 400 },
            plugins: {
                legend: { position: "top" },
                title: { display: false }
            }
        }
    });

    const chartObj = {
        id: chartId,
        sheetName: appState.activeSheetName,
        range: config.range,
        type: config.type,
        palette: config.palette || "emerald",
        instance: chartInstance
    };
    appState.charts.push(chartObj);

    // Card buttons
    card.querySelector(".btn-close-chart").addEventListener("click", () => {
        chartInstance.destroy();
        card.remove();
        appState.charts = appState.charts.filter(c => c.id !== chartId);
        showToast("Đã xóa biểu đồ", "info");
    });

    card.querySelector(".btn-download-chart").addEventListener("click", () => {
        const link = document.createElement("a");
        link.download = `${config.title.replace(/\s+/g, "_")}.png`;
        link.href = canvas.toDataURL("image/png");
        link.click();
        showToast("Đã tải biểu đồ dạng ảnh PNG!", "success");
    });

    showToast(`Đã tạo biểu đồ: ${config.title}`, "success");
}

function updateCharts() {
    appState.charts.forEach(ch => {
        if (ch.sheetName === appState.activeSheetName && ch.instance) {
            const newData = extractChartDataFromRange(ch.range, true, ch.palette);
            ch.instance.data = newData;
            ch.instance.update();
        }
    });
}

function extractChartDataFromRange(rangeStr, firstRowHeader, paletteName) {
    const range = parseRange(rangeStr);
    if (!range) return { labels: [], datasets: [] };

    const sheet = appState.sheets[appState.activeSheetName];
    const labels = [];
    const datasets = [];

    const colors = getPaletteColors(paletteName);

    const startR = firstRowHeader ? range.startRow + 1 : range.startRow;

    // Labels from first column
    for (let r = startR; r <= range.endRow; r++) {
        const rowLabel = displayValue(sheet[r][range.startCol], r, range.startCol);
        labels.push(rowLabel || `Hàng ${r + 1}`);
    }

    // Series from remaining columns
    let colorIdx = 0;
    for (let c = range.startCol + 1; c <= range.endCol; c++) {
        const seriesName = firstRowHeader ? displayValue(sheet[range.startRow][c], range.startRow, c) : `Cột ${colIndexToLetter(c)}`;
        const dataPoints = [];

        for (let r = startR; r <= range.endRow; r++) {
            const raw = displayValue(sheet[r][c], r, c);
            const num = parseFloat(raw);
            dataPoints.push(isNaN(num) ? 0 : num);
        }

        const color = colors[colorIdx % colors.length];
        datasets.push({
            label: seriesName,
            data: dataPoints,
            backgroundColor: color.bg,
            borderColor: color.border,
            borderWidth: 2,
            tension: 0.35,
            fill: true
        });
        colorIdx++;
    }

    // If only 1 column in range, use values directly
    if (datasets.length === 0) {
        const singleData = [];
        for (let r = startR; r <= range.endRow; r++) {
            const raw = displayValue(sheet[r][range.startCol], r, range.startCol);
            const num = parseFloat(raw);
            singleData.push(isNaN(num) ? 0 : num);
        }
        datasets.push({
            label: "Dữ liệu",
            data: singleData,
            backgroundColor: colors.map(c => c.bg),
            borderColor: colors.map(c => c.border),
            borderWidth: 1.5
        });
    }

    return { labels, datasets };
}

function getPaletteColors(name) {
    switch (name) {
        case "modern":
            return [
                { bg: "rgba(59, 130, 246, 0.7)", border: "#3b82f6" },
                { bg: "rgba(249, 115, 22, 0.7)", border: "#f97316" },
                { bg: "rgba(16, 185, 129, 0.7)", border: "#10b981" },
                { bg: "rgba(236, 72, 153, 0.7)", border: "#ec4899" }
            ];
        case "neon":
            return [
                { bg: "rgba(6, 182, 212, 0.8)", border: "#06b6d4" },
                { bg: "rgba(244, 63, 94, 0.8)", border: "#f43f5e" },
                { bg: "rgba(234, 179, 8, 0.8)", border: "#eab308" },
                { bg: "rgba(168, 85, 247, 0.8)", border: "#a855f7" }
            ];
        case "pastel":
            return [
                { bg: "rgba(147, 197, 253, 0.8)", border: "#93c5fd" },
                { bg: "rgba(252, 165, 165, 0.8)", border: "#fca5a5" },
                { bg: "rgba(167, 243, 208, 0.8)", border: "#a7f3d0" },
                { bg: "rgba(253, 230, 138, 0.8)", border: "#fde68a" }
            ];
        case "emerald":
        default:
            return [
                { bg: "rgba(16, 124, 65, 0.75)", border: "#107c41" },
                { bg: "rgba(56, 189, 248, 0.75)", border: "#38bdf8" },
                { bg: "rgba(245, 158, 11, 0.75)", border: "#f59e0b" },
                { bg: "rgba(167, 139, 250, 0.75)", border: "#a78bfa" }
            ];
    }
}

function makeElementDraggable(elmnt, dragHandle) {
    let pos1 = 0, pos2 = 0, pos3 = 0, pos4 = 0;
    dragHandle.onmousedown = dragMouseDown;

    function dragMouseDown(e) {
        if (e.target.tagName === "BUTTON") return;
        e.preventDefault();
        pos3 = e.clientX;
        pos4 = e.clientY;
        document.onmouseup = closeDragElement;
        document.onmousemove = elementDrag;
    }

    function elementDrag(e) {
        e.preventDefault();
        pos1 = pos3 - e.clientX;
        pos2 = pos4 - e.clientY;
        pos3 = e.clientX;
        pos4 = e.clientY;
        elmnt.style.top = `${elmnt.offsetTop - pos2}px`;
        elmnt.style.left = `${elmnt.offsetLeft - pos1}px`;
    }

    function closeDragElement() {
        document.onmouseup = null;
        document.onmousemove = null;
    }
}

// ==========================================================================
// Formatting Toolbar Actions
// ==========================================================================
function setupFormattingButtons() {
    function applyStyleToSelection(updater) {
        pushHistory();
        const { startRow, startCol, endRow, endCol } = appState.selectionRange;
        if (!appState.styles[appState.activeSheetName]) appState.styles[appState.activeSheetName] = {};
        const sheetStyles = appState.styles[appState.activeSheetName];

        for (let r = startRow; r <= endRow; r++) {
            for (let c = startCol; c <= endCol; c++) {
                const key = `${r},${c}`;
                if (!sheetStyles[key]) sheetStyles[key] = {};
                updater(sheetStyles[key]);
            }
        }
        renderGrid();
    }

    document.getElementById("btn-bold").addEventListener("click", () => {
        applyStyleToSelection(s => { s.bold = !s.bold; });
    });

    document.getElementById("btn-italic").addEventListener("click", () => {
        applyStyleToSelection(s => { s.italic = !s.italic; });
    });

    document.getElementById("btn-underline").addEventListener("click", () => {
        applyStyleToSelection(s => { s.underline = !s.underline; });
    });

    document.getElementById("btn-strikethrough").addEventListener("click", () => {
        applyStyleToSelection(s => { s.strikethrough = !s.strikethrough; });
    });

    document.getElementById("select-font-family").addEventListener("change", (e) => {
        applyStyleToSelection(s => { s.fontFamily = e.target.value; });
    });

    document.getElementById("select-font-size").addEventListener("change", (e) => {
        applyStyleToSelection(s => { s.fontSize = e.target.value; });
    });

    document.getElementById("btn-font-increase").addEventListener("click", () => {
        const select = document.getElementById("select-font-size");
        const nextIdx = Math.min(select.selectedIndex + 1, select.options.length - 1);
        select.selectedIndex = nextIdx;
        applyStyleToSelection(s => { s.fontSize = select.value; });
    });

    document.getElementById("btn-font-decrease").addEventListener("click", () => {
        const select = document.getElementById("select-font-size");
        const nextIdx = Math.max(select.selectedIndex - 1, 0);
        select.selectedIndex = nextIdx;
        applyStyleToSelection(s => { s.fontSize = select.value; });
    });

    document.getElementById("color-cell-bg").addEventListener("input", (e) => {
        document.getElementById("fill-color-bar").style.backgroundColor = e.target.value;
        applyStyleToSelection(s => { s.bg = e.target.value; });
    });

    document.getElementById("color-cell-text").addEventListener("input", (e) => {
        document.getElementById("text-color-bar").style.backgroundColor = e.target.value;
        applyStyleToSelection(s => { s.color = e.target.value; });
    });

    document.getElementById("btn-align-left").addEventListener("click", () => {
        applyStyleToSelection(s => { s.align = "left"; });
    });
    document.getElementById("btn-align-center").addEventListener("click", () => {
        applyStyleToSelection(s => { s.align = "center"; });
    });
    document.getElementById("btn-align-right").addEventListener("click", () => {
        applyStyleToSelection(s => { s.align = "right"; });
    });

    document.getElementById("btn-valign-top").addEventListener("click", () => {
        applyStyleToSelection(s => { s.valign = "top"; });
    });
    document.getElementById("btn-valign-middle").addEventListener("click", () => {
        applyStyleToSelection(s => { s.valign = "middle"; });
    });
    document.getElementById("btn-valign-bottom").addEventListener("click", () => {
        applyStyleToSelection(s => { s.valign = "bottom"; });
    });

    document.getElementById("btn-wrap-text").addEventListener("click", () => {
        applyStyleToSelection(s => { s.wrap = !s.wrap; });
    });

    const btnMergeCells = document.getElementById("btn-merge-cells");
    if (btnMergeCells) {
        btnMergeCells.addEventListener("click", toggleMergeCells);
    }

    // Borders dropdown
    const btnBorderMenu = document.getElementById("btn-border-menu");
    btnBorderMenu.addEventListener("click", (e) => {
        e.stopPropagation();
        btnBorderMenu.parentElement.classList.toggle("open");
    });

    document.querySelectorAll("#border-dropdown .dropdown-item").forEach(item => {
        item.addEventListener("click", () => {
            const type = item.getAttribute("data-border");
            btnBorderMenu.parentElement.classList.remove("open");
            applyStyleToSelection(s => { s.border = type !== "none"; });
        });
    });

    // Number formatting
    document.getElementById("select-number-format").addEventListener("change", (e) => {
        applyNumberFormat(e.target.value);
    });
    document.getElementById("btn-format-currency").addEventListener("click", () => applyNumberFormat("currency-vnd"));
    document.getElementById("btn-format-percent").addEventListener("click", () => applyNumberFormat("percent"));

    // Insert / Delete row & col
    document.getElementById("btn-insert-row-above").addEventListener("click", insertRowAbove);
    document.getElementById("btn-insert-col-left").addEventListener("click", insertColLeft);
    document.getElementById("btn-delete-row-active").addEventListener("click", deleteActiveRow);
    document.getElementById("btn-delete-col-active").addEventListener("click", deleteActiveCol);

    // AutoSum trigger
    const btnAutoSum = document.getElementById("btn-autosum-trigger");
    btnAutoSum.addEventListener("click", (e) => {
        e.stopPropagation();
        btnAutoSum.parentElement.classList.toggle("open");
    });
    document.querySelectorAll("#autosum-menu .dropdown-item").forEach(item => {
        item.addEventListener("click", () => {
            const func = item.getAttribute("data-func");
            btnAutoSum.parentElement.classList.remove("open");
            applyAutoSum(func);
        });
    });

    // Sort & Filter
    document.getElementById("btn-sort-asc").addEventListener("click", () => sortActiveColumn(true));
    document.getElementById("btn-sort-desc").addEventListener("click", () => sortActiveColumn(false));
    document.getElementById("btn-data-sort-asc").addEventListener("click", () => sortActiveColumn(true));
    document.getElementById("btn-data-sort-desc").addEventListener("click", () => sortActiveColumn(false));
    document.getElementById("btn-filter-toggle").addEventListener("click", toggleFilterMode);
    document.getElementById("btn-data-filter").addEventListener("click", toggleFilterMode);

    const btnClearFilterHome = document.getElementById("btn-clear-filter-home");
    if (btnClearFilterHome) btnClearFilterHome.addEventListener("click", clearAllFilters);
    const btnClearFilterData = document.getElementById("btn-data-clear-filter");
    if (btnClearFilterData) btnClearFilterData.addEventListener("click", clearAllFilters);

    // Clear selection
    document.getElementById("btn-clear-selection").addEventListener("click", () => {
        pushHistory();
        const { startRow, startCol, endRow, endCol } = appState.selectionRange;
        const sheet = appState.sheets[appState.activeSheetName];
        for (let r = startRow; r <= endRow; r++) {
            for (let c = startCol; c <= endCol; c++) {
                sheet[r][c] = "";
            }
        }
        renderGrid();
        updateFormulaBar();
        updateCharts();
        showToast("Đã xóa nội dung các ô được chọn", "info");
    });

    // Clipboard Cut, Copy, Paste
    document.getElementById("btn-copy").addEventListener("click", copySelection);
    document.getElementById("btn-cut").addEventListener("click", cutSelection);
    document.getElementById("btn-paste").addEventListener("click", pasteClipboard);
}

function applyNumberFormat(type) {
    pushHistory();
    const { startRow, startCol, endRow, endCol } = appState.selectionRange;
    const sheet = appState.sheets[appState.activeSheetName];

    for (let r = startRow; r <= endRow; r++) {
        for (let c = startCol; c <= endCol; c++) {
            const val = sheet[r][c];
            if (cellToString(val).startsWith("=")) continue; // Keep formula
            const num = parseFloat(val);
            if (isNaN(num)) continue;

            if (type === "currency-vnd") {
                sheet[r][c] = num.toLocaleString("vi-VN") + " ₫";
            } else if (type === "currency-usd") {
                sheet[r][c] = "$" + num.toLocaleString("en-US", { minimumFractionDigits: 2 });
            } else if (type === "percent") {
                sheet[r][c] = (num).toFixed(1) + "%";
            } else if (type === "number") {
                sheet[r][c] = num.toLocaleString("en-US", { minimumFractionDigits: 2 });
            }
        }
    }
    renderGrid();
}

function applyAutoSum(funcName) {
    const { startRow, startCol, endRow, endCol } = appState.selectionRange;
    const sheet = appState.sheets[appState.activeSheetName];
    pushHistory();

    const targetRow = endRow + 1;
    if (targetRow >= sheet.length) {
        sheet.push(Array(sheet[0].length).fill(""));
    }

    for (let c = startCol; c <= endCol; c++) {
        const rangeA1 = `${coordToA1(startRow, c)}:${coordToA1(endRow, c)}`;
        sheet[targetRow][c] = `=${funcName}(${rangeA1})`;
    }

    renderGrid();
    selectCell(targetRow, startCol);
    showToast(`Đã áp dụng ${funcName} tại hàng ${targetRow + 1}`, "success");
}

function sortActiveColumn(ascending = true, targetCol = null) {
    const colIdx = (targetCol !== null && targetCol !== undefined) ? targetCol : appState.selectedCell.col;
    const sheet = appState.sheets[appState.activeSheetName];
    if (!sheet || sheet.length <= 1) return;
    pushHistory();

    const sheetStyles = appState.styles[appState.activeSheetName] || {};
    const colCount = sheet[0] ? sheet[0].length : 0;

    // Preserve first row if header
    const hasHeader = isNaN(parseFloat(sheet[0][colIdx])) && sheet[0][colIdx] !== "";
    const startR = hasHeader ? 1 : 0;

    const indexedRows = [];
    for (let r = startR; r < sheet.length; r++) {
        indexedRows.push({
            data: sheet[r],
            oldRow: r
        });
    }

    indexedRows.sort((a, b) => {
        const valA = a.data[colIdx];
        const valB = b.data[colIdx];
        const numA = parseFloat(valA);
        const numB = parseFloat(valB);

        if (!isNaN(numA) && !isNaN(numB)) {
            return ascending ? numA - numB : numB - numA;
        }
        return ascending ? String(valA || "").localeCompare(String(valB || ""), "vi") : String(valB || "").localeCompare(String(valA || ""), "vi");
    });

    const newRows = hasHeader ? [sheet[0]] : [];
    const newStyles = {};

    if (hasHeader) {
        for (let c = 0; c < colCount; c++) {
            if (sheetStyles[`0,${c}`]) {
                newStyles[`0,${c}`] = sheetStyles[`0,${c}`];
            }
        }
    }

    indexedRows.forEach((item, newIdx) => {
        const newR = startR + newIdx;
        newRows.push(item.data);
        for (let c = 0; c < colCount; c++) {
            const oldKey = `${item.oldRow},${c}`;
            if (sheetStyles[oldKey]) {
                newStyles[`${newR},${c}`] = sheetStyles[oldKey];
            }
        }
    });

    appState.sheets[appState.activeSheetName] = newRows;
    appState.styles[appState.activeSheetName] = newStyles;
    renderGrid();
    updateCharts();
    showToast(`Đã sắp xếp cột ${colIndexToLetter(colIdx)} ${ascending ? "A -> Z" : "Z -> A"}`, "success");
}

function toggleFilterMode() {
    appState.filterMode = !appState.filterMode;
    const filterBtns = document.querySelectorAll(".col-filter-btn");

    if (!appState.filterMode) {
        const sheetFilters = appState.filters[appState.activeSheetName];
        const hasFilters = sheetFilters && Object.keys(sheetFilters).length > 0;
        if (hasFilters) {
            pushHistory();
            appState.filters[appState.activeSheetName] = {};
        }
        filterBtns.forEach(btn => btn.style.display = "none");
        showToast("Đã tắt chế độ lọc", "info");
    } else {
        filterBtns.forEach(btn => btn.style.display = "block");
        showToast("Đã bật bộ lọc: Nhấn vào mũi tên trên tiêu đề cột để lọc!", "info");
    }
    renderGrid();
}

function openFilterPopover(colIdx, colHeaderEl) {
    appState.filterActiveCol = colIdx;
    const popover = document.getElementById("filter-popover");
    if (!popover) return;
    const colLetter = colIndexToLetter(colIdx);
    document.getElementById("filter-col-title").textContent = `Lọc cột ${colLetter}`;

    const sheet = appState.sheets[appState.activeSheetName];
    if (!sheet || sheet.length === 0) return;

    let hasBlanks = false;
    const valSet = new Set();
    for (let r = 1; r < sheet.length; r++) {
        const raw = sheet[r]?.[colIdx];
        const v = cellToString(raw).trim();
        if (v === "") {
            hasBlanks = true;
        } else {
            valSet.add(v);
        }
    }

    const sortedVals = Array.from(valSet).sort((a, b) => {
        const na = parseFloat(a);
        const nb = parseFloat(b);
        if (!isNaN(na) && !isNaN(nb)) return na - nb;
        return a.localeCompare(b, "vi", { sensitivity: "base" });
    });

    const currentFilter = appState.filters[appState.activeSheetName]?.[colIdx];
    const isFiltered = Array.isArray(currentFilter);
    const allowedSet = isFiltered ? new Set(currentFilter) : null;

    const listEl = document.getElementById("filter-items-list");
    listEl.innerHTML = "";

    // 1. Select All Checkbox
    const selectAllRow = document.createElement("label");
    selectAllRow.className = "filter-item-row filter-select-all-row";
    selectAllRow.style.fontWeight = "600";
    selectAllRow.style.borderBottom = "1px solid var(--border-color)";
    selectAllRow.style.paddingBottom = "5px";
    selectAllRow.style.marginBottom = "5px";

    const selectAllCb = document.createElement("input");
    selectAllCb.type = "checkbox";
    selectAllCb.id = "filter-select-all";
    const selectAllSpan = document.createElement("span");
    selectAllSpan.textContent = "(Chọn tất cả)";
    selectAllRow.appendChild(selectAllCb);
    selectAllRow.appendChild(selectAllSpan);
    listEl.appendChild(selectAllRow);

    // 2. Blank cells item
    if (hasBlanks) {
        const blankRow = document.createElement("label");
        blankRow.className = "filter-item-row";
        const blankCb = document.createElement("input");
        blankCb.type = "checkbox";
        blankCb.className = "filter-item-cb";
        blankCb.dataset.val = "";
        blankCb.checked = !isFiltered || allowedSet.has("");
        const blankSpan = document.createElement("span");
        blankSpan.textContent = "(Ô trống)";
        blankSpan.style.fontStyle = "italic";
        blankSpan.style.opacity = "0.75";
        blankRow.appendChild(blankCb);
        blankRow.appendChild(blankSpan);
        listEl.appendChild(blankRow);
    }

    // 3. Regular items
    sortedVals.forEach(val => {
        const itemRow = document.createElement("label");
        itemRow.className = "filter-item-row";
        const cb = document.createElement("input");
        cb.type = "checkbox";
        cb.className = "filter-item-cb";
        cb.dataset.val = val;
        cb.checked = !isFiltered || allowedSet.has(val);
        const span = document.createElement("span");
        span.textContent = val;
        itemRow.appendChild(cb);
        itemRow.appendChild(span);
        listEl.appendChild(itemRow);
    });

    function updateSelectAllState() {
        const allVisibleCbs = Array.from(listEl.querySelectorAll(".filter-item-cb")).filter(cb => {
            return cb.closest(".filter-item-row").style.display !== "none";
        });
        if (allVisibleCbs.length === 0) {
            selectAllCb.checked = false;
            selectAllCb.indeterminate = false;
            return;
        }
        const checkedCount = allVisibleCbs.filter(cb => cb.checked).length;
        if (checkedCount === allVisibleCbs.length) {
            selectAllCb.checked = true;
            selectAllCb.indeterminate = false;
        } else if (checkedCount === 0) {
            selectAllCb.checked = false;
            selectAllCb.indeterminate = false;
        } else {
            selectAllCb.checked = false;
            selectAllCb.indeterminate = true;
        }
    }

    selectAllCb.addEventListener("change", () => {
        const isChecked = selectAllCb.checked;
        const allVisibleCbs = Array.from(listEl.querySelectorAll(".filter-item-cb")).filter(cb => {
            return cb.closest(".filter-item-row").style.display !== "none";
        });
        allVisibleCbs.forEach(cb => { cb.checked = isChecked; });
    });

    listEl.addEventListener("change", (e) => {
        if (e.target.classList.contains("filter-item-cb")) {
            updateSelectAllState();
        }
    });

    updateSelectAllState();

    // Search input live filtering
    const searchInput = document.getElementById("filter-search-input");
    searchInput.value = "";
    searchInput.oninput = () => {
        const query = searchInput.value.toLowerCase().trim();
        const rows = listEl.querySelectorAll(".filter-item-row:not(.filter-select-all-row)");
        rows.forEach(r => {
            const txt = r.textContent.toLowerCase();
            r.style.display = txt.includes(query) ? "" : "none";
        });
        updateSelectAllState();
    };

    // Position popover
    const rect = colHeaderEl.getBoundingClientRect();
    popover.style.top = `${rect.bottom + 4}px`;
    popover.style.left = `${Math.max(10, Math.min(rect.left, window.innerWidth - 270))}px`;
    popover.classList.remove("hidden");

    // Close button
    document.getElementById("btn-close-filter").onclick = () => popover.classList.add("hidden");

    // Sort buttons
    document.getElementById("btn-filter-sort-asc").onclick = () => {
        sortActiveColumn(true, colIdx);
        popover.classList.add("hidden");
    };
    document.getElementById("btn-filter-sort-desc").onclick = () => {
        sortActiveColumn(false, colIdx);
        popover.classList.add("hidden");
    };

    // "Áp dụng" button
    document.getElementById("btn-filter-apply").onclick = () => {
        applyColumnFilter(colIdx);
    };

    // "Bỏ lọc" button
    document.getElementById("btn-filter-clear").onclick = () => {
        clearColumnFilter(colIdx);
    };
}

function applyColumnFilter(colIdx) {
    const popover = document.getElementById("filter-popover");
    const listEl = document.getElementById("filter-items-list");
    const allCbs = Array.from(listEl.querySelectorAll(".filter-item-cb"));
    const checkedCbs = allCbs.filter(cb => cb.checked);

    if (!appState.filters[appState.activeSheetName]) {
        appState.filters[appState.activeSheetName] = {};
    }

    const colLetter = colIndexToLetter(colIdx);
    pushHistory();

    if (checkedCbs.length === allCbs.length) {
        delete appState.filters[appState.activeSheetName][colIdx];
        showToast(`Đã hiển thị tất cả dữ liệu cột ${colLetter}`, "info");
    } else {
        const allowedVals = checkedCbs.map(cb => cb.dataset.val);
        appState.filters[appState.activeSheetName][colIdx] = allowedVals;
        showToast(`Đã áp dụng bộ lọc cột ${colLetter} (${checkedCbs.length} giá trị)`, "success");
    }

    if (popover) popover.classList.add("hidden");
    renderGrid();
}

function clearColumnFilter(colIdx) {
    const popover = document.getElementById("filter-popover");
    const colLetter = colIndexToLetter(colIdx);
    if (appState.filters[appState.activeSheetName]) {
        delete appState.filters[appState.activeSheetName][colIdx];
    }
    pushHistory();
    if (popover) popover.classList.add("hidden");
    renderGrid();
    showToast(`Đã bỏ lọc cột ${colLetter}`, "info");
}

function clearAllFilters() {
    const sheetFilters = appState.filters[appState.activeSheetName];
    const hasFilters = sheetFilters && Object.keys(sheetFilters).some(k => Array.isArray(sheetFilters[k]));
    if (hasFilters) {
        pushHistory();
        appState.filters[appState.activeSheetName] = {};
        const popover = document.getElementById("filter-popover");
        if (popover) popover.classList.add("hidden");
        renderGrid();
        showToast("Đã xóa tất cả bộ lọc trên trang tính!", "success");
    } else {
        showToast("Không có bộ lọc nào đang được áp dụng.", "info");
    }
}

function setupFilterPopoverGlobalEvents() {
    document.addEventListener("click", (e) => {
        const popover = document.getElementById("filter-popover");
        if (!popover || popover.classList.contains("hidden")) return;
        if (e.target.closest("#filter-popover") || e.target.closest(".col-filter-btn")) return;
        popover.classList.add("hidden");
    });
}

// ==========================================================================
// Insert & Delete Rows / Columns
// ==========================================================================
function insertRowAbove() {
    const rowIdx = appState.selectedCell.row;
    const sheet = appState.sheets[appState.activeSheetName];
    pushHistory();
    const cols = sheet[0] ? sheet[0].length : 14;
    sheet.splice(rowIdx, 0, Array(cols).fill(""));
    shiftMergesOnRowInsert(appState.activeSheetName, rowIdx);
    renderGrid();
    selectCell(rowIdx, appState.selectedCell.col);
    showToast(`Đã chèn hàng mới tại vị trí ${rowIdx + 1}`, "success");
}

function insertColLeft() {
    const colIdx = appState.selectedCell.col;
    const sheet = appState.sheets[appState.activeSheetName];
    pushHistory();
    sheet.forEach(row => row.splice(colIdx, 0, ""));
    shiftMergesOnColInsert(appState.activeSheetName, colIdx);
    renderGrid();
    selectCell(appState.selectedCell.row, colIdx);
    showToast(`Đã chèn cột mới tại vị trí ${colIndexToLetter(colIdx)}`, "success");
}

function deleteActiveRow() {
    const rowIdx = appState.selectedCell.row;
    const sheet = appState.sheets[appState.activeSheetName];
    if (sheet.length <= 1) return showToast("Không thể xóa hàng duy nhất còn lại", "error");
    pushHistory();
    sheet.splice(rowIdx, 1);
    shiftMergesOnRowDelete(appState.activeSheetName, rowIdx);
    appState.selectedCell.row = Math.min(rowIdx, sheet.length - 1);
    renderGrid();
    updateCharts();
    showToast(`Đã xóa hàng ${rowIdx + 1}`, "info");
}

function deleteActiveCol() {
    const colIdx = appState.selectedCell.col;
    const sheet = appState.sheets[appState.activeSheetName];
    if (sheet[0].length <= 1) return showToast("Không thể xóa cột duy nhất còn lại", "error");
    pushHistory();
    sheet.forEach(row => row.splice(colIdx, 1));
    shiftMergesOnColDelete(appState.activeSheetName, colIdx);
    appState.selectedCell.col = Math.min(colIdx, sheet[0].length - 1);
    renderGrid();
    updateCharts();
    showToast(`Đã xóa cột ${colIndexToLetter(colIdx)}`, "info");
}

// ==========================================================================
// Clipboard (Cut, Copy, Paste) & Sheet Image Manipulation
// ==========================================================================

function copySelection() {
    if (appState.selectedImageId) {
        copySelectedImage();
        return;
    }

    const { startRow, startCol, endRow, endCol } = appState.selectionRange;
    const sheet = appState.sheets[appState.activeSheetName];
    if (!sheet) return;

    const data = [];
    const tsvLines = [];

    for (let r = startRow; r <= endRow; r++) {
        const rowData = [];
        const tsvCells = [];
        for (let c = startCol; c <= endCol; c++) {
            const val = sheet[r] ? sheet[r][c] : "";
            rowData.push(val);
            const disp = displayValue(val, r, c);
            tsvCells.push(disp);
        }
        data.push(rowData);
        tsvLines.push(tsvCells.join("\t"));
    }

    appState.clipboard = {
        type: "copy",
        sheet: appState.activeSheetName,
        range: { ...appState.selectionRange },
        data: data
    };

    // Write TSV to system clipboard for Excel / Google Sheets
    const tsvText = tsvLines.join("\n");
    if (navigator.clipboard && navigator.clipboard.writeText) {
        navigator.clipboard.writeText(tsvText).catch(() => {});
    }

    // Add marching ants animation to copied cells
    document.querySelectorAll(".grid-cell.clipboard-copy").forEach(el => el.classList.remove("clipboard-copy"));
    for (let r = startRow; r <= endRow; r++) {
        for (let c = startCol; c <= endCol; c++) {
            const el = document.getElementById(`cell-${r}-${c}`);
            if (el) el.classList.add("clipboard-copy");
        }
    }

    showToast("Đã sao chép vào bộ nhớ tạm (Ctrl+C)", "info");
}

function cutSelection() {
    if (appState.selectedImageId) {
        copySelectedImage();
        deleteSelectedImage();
        showToast("Đã cắt hình ảnh (Ctrl+X)", "info");
        return;
    }

    copySelection();
    if (appState.clipboard) {
        appState.clipboard.type = "cut";
    }
    showToast("Đã cắt vùng chọn (Ctrl+X)", "info");
}

function copySelectedImage() {
    if (!appState.selectedImageId) return;
    const list = (appState.images && appState.images[appState.activeSheetName]) || [];
    const imgObj = list.find(img => img.id === appState.selectedImageId);
    if (!imgObj) return;

    appState.imageClipboard = {
        src: imgObj.src,
        name: imgObj.name,
        width: imgObj.width,
        height: imgObj.height
    };

    showToast(`Đã sao chép hình ảnh: ${imgObj.name || 'Image'} (Ctrl+C)`, "info");
}

async function pasteClipboard() {
    if (appState.isEditing) commitInlineEdit();

    // 1. Check internal image clipboard first
    if (appState.imageClipboard) {
        const img = appState.imageClipboard;
        const targetPos = getTargetPosForActiveCell();
        insertImageOnSheet(img.src, img.name, targetPos.x + 20, targetPos.y + 20, img.width, img.height);
        return;
    }

    // 2. Try reading system clipboard for image or text
    if (navigator.clipboard && navigator.clipboard.read) {
        try {
            const items = await navigator.clipboard.read();
            for (const item of items) {
                const imgType = item.types.find(t => t.startsWith("image/"));
                if (imgType) {
                    const blob = await item.getType(imgType);
                    const reader = new FileReader();
                    reader.onload = (loadEvt) => {
                        const targetPos = getTargetPosForActiveCell();
                        insertImageOnSheet(loadEvt.target.result, "Pasted_Image.png", targetPos.x, targetPos.y);
                    };
                    reader.readAsDataURL(blob);
                    return;
                }
            }
        } catch (err) {}
    }

    // 3. Try reading system clipboard text
    if (navigator.clipboard && navigator.clipboard.readText) {
        try {
            const text = await navigator.clipboard.readText();
            if (text && text.trim()) {
                pasteTextIntoGrid(text);
                return;
            }
        } catch (err) {}
    }

    // 4. Fallback to appState.clipboard
    if (appState.clipboard && appState.clipboard.data) {
        applyInternalClipboardData();
    }
}

function handlePasteEvent(e) {
    const activeEl = document.activeElement;
    if (activeEl && (activeEl.tagName === "INPUT" || activeEl.tagName === "TEXTAREA")) {
        if (!activeEl.classList.contains("grid-cell-input") && activeEl.id !== "formula-input-field") {
            return;
        }
    }

    const clipboardData = e.clipboardData || window.clipboardData;
    if (!clipboardData) return;

    // 1. Detect Image paste from clipboard (e.g. screenshot Win+Shift+S, copied image)
    const items = clipboardData.items;
    if (items) {
        for (let i = 0; i < items.length; i++) {
            if (items[i].type && items[i].type.indexOf("image") !== -1) {
                e.preventDefault();
                const file = items[i].getAsFile();
                if (file) {
                    const reader = new FileReader();
                    reader.onload = (loadEvt) => {
                        const targetPos = getTargetPosForActiveCell();
                        insertImageOnSheet(loadEvt.target.result, file.name || "Pasted_Image.png", targetPos.x, targetPos.y);
                    };
                    reader.readAsDataURL(file);
                    return;
                }
            }
        }
    }

    // 2. Detect TSV / Text paste
    const text = clipboardData.getData("text/plain");
    if (text) {
        if (appState.isEditing) return; // Allow typing paste inside cell input
        e.preventDefault();
        pasteTextIntoGrid(text);
    }
}

function pasteTextIntoGrid(text) {
    if (!text) return;
    const cleanText = text.replace(/\r\n/g, "\n").replace(/\r/g, "\n");
    const rawRows = cleanText.split("\n");
    if (rawRows.length > 0 && rawRows[rawRows.length - 1] === "") {
        rawRows.pop();
    }
    if (rawRows.length === 0) return;

    const parsedMatrix = rawRows.map(rowStr => rowStr.split("\t"));
    const destR = appState.selectedCell.row;
    const destC = appState.selectedCell.col;
    const sheet = appState.sheets[appState.activeSheetName];
    if (!sheet) return;

    pushHistory();

    const rows = parsedMatrix.length;
    const cols = Math.max(...parsedMatrix.map(r => r.length), 1);

    for (let r = 0; r < rows; r++) {
        const targetR = destR + r;
        while (targetR >= sheet.length) {
            appendRowsToActiveSheet(25);
        }
        for (let c = 0; c < parsedMatrix[r].length; c++) {
            const targetC = destC + c;
            while (targetC >= sheet[0].length) {
                appendColsToActiveSheet(10);
            }
            sheet[targetR][targetC] = parsedMatrix[r][c];
        }
    }

    renderGrid();
    updateCharts();
    showToast(`Đã dán ${rows} hàng × ${cols} cột vào bảng tính!`, "success");
}

function applyInternalClipboardData() {
    if (!appState.clipboard || !appState.clipboard.data) return;
    const destR = appState.selectedCell.row;
    const destC = appState.selectedCell.col;
    const sheet = appState.sheets[appState.activeSheetName];
    if (!sheet) return;

    pushHistory();

    const clipData = appState.clipboard.data;
    const rows = clipData.length;
    const cols = clipData[0].length;

    for (let r = 0; r < rows; r++) {
        const targetR = destR + r;
        while (targetR >= sheet.length) {
            appendRowsToActiveSheet(25);
        }
        for (let c = 0; c < cols; c++) {
            const targetC = destC + c;
            while (targetC >= sheet[0].length) {
                appendColsToActiveSheet(10);
            }
            sheet[targetR][targetC] = clipData[r][c];
        }
    }

    if (appState.clipboard.type === "cut" && appState.clipboard.sheet === appState.activeSheetName) {
        const src = appState.clipboard.range;
        for (let r = src.startRow; r <= src.endRow; r++) {
            for (let c = src.startCol; c <= src.endCol; c++) {
                if (r < destR || r >= destR + rows || c < destC || c >= destC + cols) {
                    sheet[r][c] = "";
                }
            }
        }
        appState.clipboard = null;
    }

    document.querySelectorAll(".grid-cell.clipboard-copy").forEach(el => el.classList.remove("clipboard-copy"));
    renderGrid();
    updateCharts();
    showToast("Đã dán dữ liệu thành công (Ctrl+V)", "success");
}

// ==========================================================================
// Sheet Image Management & Upload Engine
// ==========================================================================

function setupImageControls() {
    const btnInsertImage = document.getElementById("btn-insert-image");
    const imageFileInput = document.getElementById("image-file-input");

    if (btnInsertImage && imageFileInput) {
        btnInsertImage.addEventListener("click", () => {
            imageFileInput.click();
        });

        imageFileInput.addEventListener("change", (e) => {
            if (e.target.files && e.target.files.length > 0) {
                Array.from(e.target.files).forEach(file => {
                    handleImageFileUpload(file);
                });
                e.target.value = "";
            }
        });
    }

    // Deselect image when clicking elsewhere
    document.addEventListener("mousedown", (e) => {
        if (!e.target.closest(".sheet-image-card") && !e.target.closest(".sheet-image-toolbar")) {
            deselectSheetImage();
        }
    });
}

function handleImageFileUpload(file, targetX = null, targetY = null) {
    if (!file) return;
    const validTypes = ["image/png", "image/jpeg", "image/jpg", "image/webp", "image/gif", "image/svg+xml", "image/bmp"];
    const ext = file.name.substring(file.name.lastIndexOf(".")).toLowerCase();
    const validExts = [".png", ".jpg", ".jpeg", ".webp", ".gif", ".svg", ".bmp"];

    if (!validTypes.includes(file.type) && !validExts.includes(ext)) {
        showToast("Vui lòng chọn định dạng file ảnh hợp lệ (.png, .jpg, .webp, .gif).", "warning");
        return;
    }

    const reader = new FileReader();
    reader.onload = (e) => {
        const dataUrl = e.target.result;
        insertImageOnSheet(dataUrl, file.name, targetX, targetY);
    };
    reader.onerror = () => {
        showToast("Lỗi khi đọc file ảnh từ máy tính.", "error");
    };
    reader.readAsDataURL(file);
}

function getTargetPosForActiveCell() {
    const { row, col } = appState.selectedCell || { row: 0, col: 0 };
    const cellEl = document.getElementById(`cell-${row}-${col}`);
    if (cellEl) {
        return {
            x: Math.max(40, cellEl.offsetLeft + 10),
            y: Math.max(30, cellEl.offsetTop + 10)
        };
    }
    return { x: col * 110 + 50, y: row * 26 + 35 };
}

function insertImageOnSheet(dataUrl, name = "Image.png", targetX = null, targetY = null, width = 240, height = 180) {
    if (!appState.images) appState.images = {};
    if (!appState.images[appState.activeSheetName]) {
        appState.images[appState.activeSheetName] = [];
    }

    if (targetX === null || targetY === null) {
        const pos = getTargetPosForActiveCell();
        targetX = pos.x;
        targetY = pos.y;
    }

    const tempImg = new Image();
    tempImg.onload = () => {
        let finalW = width;
        let finalH = height;
        if (tempImg.width && tempImg.height) {
            const aspect = tempImg.width / tempImg.height;
            if (aspect >= 1) {
                finalW = Math.min(Math.max(tempImg.width, 160), 380);
                finalH = Math.round(finalW / aspect);
            } else {
                finalH = Math.min(Math.max(tempImg.height, 140), 320);
                finalW = Math.round(finalH * aspect);
            }
        }

        const id = "img-" + Date.now() + "-" + Math.floor(Math.random() * 1000);
        const colIdx = Math.max(0, Math.floor((targetX - 40) / 110));
        const rowIdx = Math.max(0, Math.floor((targetY - 30) / 26));

        const imgObj = {
            id: id,
            src: dataUrl,
            name: name,
            x: Math.round(targetX),
            y: Math.round(targetY),
            width: finalW,
            height: finalH,
            col: colIdx,
            row: rowIdx,
            zIndex: (appState.images[appState.activeSheetName].length + 1) * 5 + 10
        };

        pushHistory();
        appState.images[appState.activeSheetName].push(imgObj);
        renderImages();
        selectSheetImage(id);
        showToast(`Đã tải và chèn ảnh: ${name}`, "success");
    };
    tempImg.onerror = () => {
        showToast("Không thể tải định dạng hình ảnh này.", "error");
    };
    tempImg.src = dataUrl;
}

function renderImages() {
    const gridEl = document.getElementById("grid-element");
    if (!gridEl) return;

    let container = document.getElementById("sheet-images-container");
    if (!container) {
        container = document.createElement("div");
        container.id = "sheet-images-container";
        container.className = "sheet-images-container";
        gridEl.appendChild(container);
    } else if (container.parentElement !== gridEl) {
        gridEl.appendChild(container);
    }

    container.innerHTML = "";

    const activeList = (appState.images && appState.images[appState.activeSheetName]) || [];
    activeList.forEach(imgData => {
        const card = createSheetImageElement(imgData);
        container.appendChild(card);
    });
}

function createSheetImageElement(imgData) {
    const card = document.createElement("div");
    card.className = "sheet-image-card" + (appState.selectedImageId === imgData.id ? " selected" : "");
    card.id = imgData.id;
    card.style.left = `${imgData.x}px`;
    card.style.top = `${imgData.y}px`;
    card.style.width = `${imgData.width}px`;
    card.style.height = `${imgData.height}px`;
    card.style.zIndex = imgData.zIndex || 10;

    card.innerHTML = `
        <div class="sheet-image-toolbar">
            <button class="btn-img-tool" data-act="copy" title="Sao chép ảnh (Ctrl+C)">📋 Sao chép</button>
            <button class="btn-img-tool" data-act="download" title="Tải ảnh về máy">💾 Lưu ảnh</button>
            <button class="btn-img-tool" data-act="forward" title="Đưa lên trên">🔼 Lên trên</button>
            <button class="btn-img-tool" data-act="backward" title="Đưa xuống dưới">🔽 Xuống dưới</button>
            <button class="btn-img-tool btn-danger" data-act="delete" title="Xóa ảnh (Delete)">🗑️ Xóa</button>
        </div>
        <div class="sheet-image-inner">
            <img src="${imgData.src}" alt="${imgData.name || 'Image'}" class="sheet-image-content" draggable="false">
        </div>
        <div class="img-resize-handle handle-nw" data-handle="nw"></div>
        <div class="img-resize-handle handle-ne" data-handle="ne"></div>
        <div class="img-resize-handle handle-sw" data-handle="sw"></div>
        <div class="img-resize-handle handle-se" data-handle="se"></div>
        <div class="img-resize-handle handle-n" data-handle="n"></div>
        <div class="img-resize-handle handle-s" data-handle="s"></div>
        <div class="img-resize-handle handle-w" data-handle="w"></div>
        <div class="img-resize-handle handle-e" data-handle="e"></div>
    `;

    setupImageInteractions(card, imgData);

    return card;
}

function setupImageInteractions(card, imgData) {
    // 1. Click to select
    card.addEventListener("mousedown", (e) => {
        if (e.target.closest(".sheet-image-toolbar")) return;
        if (e.target.classList.contains("img-resize-handle")) return;
        e.stopPropagation();
        selectSheetImage(imgData.id);

        // Drag to move
        let startX = e.clientX;
        let startY = e.clientY;
        let origLeft = card.offsetLeft;
        let origTop = card.offsetTop;
        let moved = false;

        function onMouseMove(moveEvt) {
            moveEvt.preventDefault();
            const dx = moveEvt.clientX - startX;
            const dy = moveEvt.clientY - startY;
            if (Math.abs(dx) > 3 || Math.abs(dy) > 3) moved = true;

            const newLeft = Math.max(0, origLeft + dx);
            const newTop = Math.max(0, origTop + dy);
            card.style.left = `${newLeft}px`;
            card.style.top = `${newTop}px`;
        }

        function onMouseUp() {
            document.removeEventListener("mousemove", onMouseMove);
            document.removeEventListener("mouseup", onMouseUp);

            if (moved) {
                imgData.x = Math.round(card.offsetLeft);
                imgData.y = Math.round(card.offsetTop);
                imgData.col = Math.max(0, Math.floor((imgData.x - 40) / 110));
                imgData.row = Math.max(0, Math.floor((imgData.y - 30) / 26));
                pushHistory();
                setSaveStatus(false);
            }
        }

        document.addEventListener("mousemove", onMouseMove);
        document.addEventListener("mouseup", onMouseUp);
    });

    // 2. Toolbar buttons
    card.querySelectorAll(".btn-img-tool").forEach(btn => {
        btn.addEventListener("click", (e) => {
            e.stopPropagation();
            const act = btn.getAttribute("data-act");
            if (act === "copy") copySelectedImage();
            else if (act === "download") downloadSheetImage(imgData);
            else if (act === "forward") bringImageForward(imgData.id);
            else if (act === "backward") sendImageBackward(imgData.id);
            else if (act === "delete") deleteSelectedImage();
        });
    });

    // 3. Resizing handles
    card.querySelectorAll(".img-resize-handle").forEach(handle => {
        handle.addEventListener("mousedown", (e) => {
            e.stopPropagation();
            e.preventDefault();
            const handleType = handle.getAttribute("data-handle");
            let startX = e.clientX;
            let startY = e.clientY;
            let origW = card.offsetWidth;
            let origH = card.offsetHeight;
            let origLeft = card.offsetLeft;
            let origTop = card.offsetTop;
            let resized = false;

            function onResizeMove(moveEvt) {
                moveEvt.preventDefault();
                resized = true;
                const dx = moveEvt.clientX - startX;
                const dy = moveEvt.clientY - startY;

                let newW = origW;
                let newH = origH;
                let newLeft = origLeft;
                let newTop = origTop;

                if (handleType.includes("e")) newW = Math.max(40, origW + dx);
                if (handleType.includes("s")) newH = Math.max(40, origH + dy);
                if (handleType.includes("w")) {
                    newW = Math.max(40, origW - dx);
                    newLeft = origLeft + (origW - newW);
                }
                if (handleType.includes("n")) {
                    newH = Math.max(40, origH - dy);
                    newTop = origTop + (origH - newH);
                }

                card.style.width = `${newW}px`;
                card.style.height = `${newH}px`;
                card.style.left = `${newLeft}px`;
                card.style.top = `${newTop}px`;
            }

            function onResizeUp() {
                document.removeEventListener("mousemove", onResizeMove);
                document.removeEventListener("mouseup", onResizeUp);

                if (resized) {
                    imgData.width = Math.round(card.offsetWidth);
                    imgData.height = Math.round(card.offsetHeight);
                    imgData.x = Math.round(card.offsetLeft);
                    imgData.y = Math.round(card.offsetTop);
                    imgData.col = Math.max(0, Math.floor((imgData.x - 40) / 110));
                    imgData.row = Math.max(0, Math.floor((imgData.y - 30) / 26));
                    pushHistory();
                    setSaveStatus(false);
                }
            }

            document.addEventListener("mousemove", onResizeMove);
            document.addEventListener("mouseup", onResizeUp);
        });
    });
}

function selectSheetImage(id) {
    appState.selectedImageId = id;
    document.querySelectorAll(".sheet-image-card").forEach(c => {
        if (c.id === id) c.classList.add("selected");
        else c.classList.remove("selected");
    });
}

function deselectSheetImage() {
    appState.selectedImageId = null;
    document.querySelectorAll(".sheet-image-card.selected").forEach(c => {
        c.classList.remove("selected");
    });
}

function deleteSelectedImage() {
    if (!appState.selectedImageId) return;
    const list = appState.images && appState.images[appState.activeSheetName];
    if (!list) return;

    pushHistory();
    appState.images[appState.activeSheetName] = list.filter(img => img.id !== appState.selectedImageId);
    appState.selectedImageId = null;
    renderImages();
    setSaveStatus(false);
    showToast("Đã xóa hình ảnh khỏi trang tính", "info");
}

function nudgeSelectedImage(direction, amount = 2) {
    if (!appState.selectedImageId) return;
    const list = (appState.images && appState.images[appState.activeSheetName]) || [];
    const imgObj = list.find(img => img.id === appState.selectedImageId);
    if (!imgObj) return;

    if (direction === "ArrowUp") imgObj.y = Math.max(0, imgObj.y - amount);
    else if (direction === "ArrowDown") imgObj.y += amount;
    else if (direction === "ArrowLeft") imgObj.x = Math.max(0, imgObj.x - amount);
    else if (direction === "ArrowRight") imgObj.x += amount;

    const el = document.getElementById(imgObj.id);
    if (el) {
        el.style.left = `${imgObj.x}px`;
        el.style.top = `${imgObj.y}px`;
    }
}

function bringImageForward(id) {
    const list = (appState.images && appState.images[appState.activeSheetName]) || [];
    const imgObj = list.find(img => img.id === id);
    if (!imgObj) return;
    const maxZ = Math.max(...list.map(i => i.zIndex || 10), 10);
    imgObj.zIndex = maxZ + 2;
    const el = document.getElementById(id);
    if (el) el.style.zIndex = imgObj.zIndex;
}

function sendImageBackward(id) {
    const list = (appState.images && appState.images[appState.activeSheetName]) || [];
    const imgObj = list.find(img => img.id === id);
    if (!imgObj) return;
    const minZ = Math.min(...list.map(i => i.zIndex || 10), 10);
    imgObj.zIndex = Math.max(2, minZ - 2);
    const el = document.getElementById(id);
    if (el) el.style.zIndex = imgObj.zIndex;
}

function downloadSheetImage(imgData) {
    if (!imgData || !imgData.src) return;
    const link = document.createElement("a");
    link.href = imgData.src;
    link.download = imgData.name || "sheet_image.png";
    link.click();
    showToast(`Đang tải ảnh: ${link.download}`, "success");
}

// ==========================================================================
// Keyboard Navigation & Shortcuts
// ==========================================================================
function setupKeyboardShortcuts() {
    document.addEventListener("keydown", (e) => {
        if (appState.isEditing || appState.isCommitting) return;
        if (e.target && (e.target.classList.contains("grid-cell-input") || (e.target.closest && e.target.closest(".grid-cell-input")))) return;
        const activeTag = document.activeElement ? document.activeElement.tagName : "";
        if (activeTag === "INPUT" && document.activeElement.id !== "name-box-input") return;

        // If an image on sheet is selected:
        if (appState.selectedImageId) {
            if (e.key === "Escape") {
                deselectSheetImage();
                return;
            }
            if (e.key === "Delete" || e.key === "Backspace") {
                e.preventDefault();
                deleteSelectedImage();
                return;
            }
            if (["ArrowUp", "ArrowDown", "ArrowLeft", "ArrowRight"].includes(e.key)) {
                e.preventDefault();
                nudgeSelectedImage(e.key, e.shiftKey ? 10 : 2);
                return;
            }
            if ((e.ctrlKey || e.metaKey) && (e.key === "c" || e.key === "C")) {
                e.preventDefault();
                copySelectedImage();
                return;
            }
        }

        const { row, col } = appState.selectedCell;
        const maxRow = appState.sheets[appState.activeSheetName].length;
        const maxCol = appState.sheets[appState.activeSheetName][0].length;

        // Ctrl Shortcuts
        if (e.ctrlKey || e.metaKey) {
            if (e.key === "z" || e.key === "Z") {
                e.preventDefault();
                if (e.shiftKey) redo(); else undo();
                return;
            }
            if (e.key === "y" || e.key === "Y") {
                e.preventDefault();
                redo();
                return;
            }
            if (e.key === "c" || e.key === "C") {
                e.preventDefault();
                copySelection();
                return;
            }
            if (e.key === "x" || e.key === "X") {
                e.preventDefault();
                cutSelection();
                return;
            }
            if (e.key === "v" || e.key === "V") {
                // Let native 'paste' event fire so handlePasteEvent gets e.clipboardData with image/text
                return;
            }
            if (e.key === "a" || e.key === "A") {
                e.preventDefault();
                selectAllCells();
                return;
            }
            if (e.key === "f" || e.key === "F") {
                e.preventDefault();
                openFindReplaceModal();
                return;
            }
            if (e.key === "b" || e.key === "B") {
                e.preventDefault();
                document.getElementById("btn-bold").click();
                return;
            }
            if (e.key === "i" || e.key === "I") {
                e.preventDefault();
                document.getElementById("btn-italic").click();
                return;
            }
            if (e.key === "u" || e.key === "U") {
                e.preventDefault();
                document.getElementById("btn-underline").click();
                return;
            }
            if (e.key === "s" || e.key === "S") {
                e.preventDefault();
                saveCurrentFile();
                return;
            }
        }

        const curMerge = findMergeForCell(row, col);
        let targetRow = row;
        let targetCol = col;
        let handled = false;

        switch (e.key) {
            case "ArrowUp":
                targetRow = (curMerge ? curMerge.startRow : row) - 1;
                if (targetRow < 0) targetRow = 0;
                handled = true; break;
            case "ArrowDown":
                targetRow = (curMerge ? curMerge.endRow : row) + 1;
                if (targetRow >= maxRow - 1) {
                    appendRowsToActiveSheet(25);
                }
                handled = true; break;
            case "ArrowLeft":
                targetCol = (curMerge ? curMerge.startCol : col) - 1;
                if (targetCol < 0) targetCol = 0;
                handled = true; break;
            case "ArrowRight":
                targetCol = (curMerge ? curMerge.endCol : col) + 1;
                if (targetCol >= maxCol - 1) {
                    appendColsToActiveSheet(10);
                }
                handled = true; break;
            case "Tab":
                e.preventDefault();
                if (e.shiftKey) { 
                    targetCol = (curMerge ? curMerge.startCol : col) - 1;
                    if (targetCol < 0) targetCol = 0;
                } else {
                    targetCol = (curMerge ? curMerge.endCol : col) + 1;
                    if (targetCol >= maxCol - 1) {
                        appendColsToActiveSheet(10);
                    }
                }
                handled = true; break;
            case "Enter":
                e.preventDefault();
                const sheet = appState.sheets[appState.activeSheetName];
                const nextR = e.shiftKey ? Math.max(0, (curMerge ? curMerge.startRow : row) - 1) : (curMerge ? curMerge.endRow : row) + 1;
                if (nextR >= sheet.length) {
                    appendRowsToActiveSheet(25);
                }
                selectCell(nextR, curMerge ? curMerge.startCol : col);
                return;
            case "F2":
                e.preventDefault();
                startInlineEditing(row, col);
                return;
            case "Delete":
            case "Backspace":
                e.preventDefault();
                document.getElementById("btn-clear-selection").click();
                return;
            default:
                // Start typing immediately on active cell
                if (e.key.length === 1 && !e.ctrlKey && !e.metaKey && !e.altKey) {
                    e.preventDefault();
                    startInlineEditing(row, col, e.key);
                    return;
                }
                break;
        }

        if (handled && (targetRow !== row || targetCol !== col)) {
            e.preventDefault();
            const destMerge = findMergeForCell(targetRow, targetCol);
            if (destMerge) {
                targetRow = destMerge.startRow;
                targetCol = destMerge.startCol;
            }
            selectCell(targetRow, targetCol);
        }
    });
}

// ==========================================================================
// Sheet Tabs Navigation
// ==========================================================================
function renderTabs() {
    const container = document.getElementById("sheet-tabs-container");
    container.innerHTML = "";
    const sheetNames = Object.keys(appState.sheets);

    sheetNames.forEach(name => {
        const tab = document.createElement("div");
        tab.className = "sheet-tab-item" + (name === appState.activeSheetName ? " active" : "");
        tab.setAttribute("data-sheet-name", name);
        tab.innerHTML = `<span>${name}</span>`;

        if (sheetNames.length > 1) {
            const closeBtn = document.createElement("button");
            closeBtn.className = "btn-tab-close";
            closeBtn.innerHTML = "&times;";
            closeBtn.title = `Xóa trang tính ${name}`;
            closeBtn.addEventListener("mousedown", (e) => {
                e.stopPropagation();
            });
            closeBtn.addEventListener("click", (e) => {
                e.preventDefault();
                e.stopPropagation();
                deleteSheet(name);
            });
            tab.appendChild(closeBtn);
        }

        tab.addEventListener("click", () => {
            if (appState.activeSheetName !== name) switchSheet(name);
        });

        // Double click to rename tab
        tab.addEventListener("dblclick", () => renameSheetInline(name, tab));

        // Right click context menu on tab
        tab.addEventListener("contextmenu", (e) => {
            e.preventDefault();
            openTabContextMenu(e, name);
        });

        container.appendChild(tab);
    });

    // Add sheet button
    document.getElementById("btn-add-sheet").onclick = () => {
        const count = Object.keys(appState.sheets).length + 1;
        let newName = `Sheet${count}`;
        while (appState.sheets[newName]) newName = `Sheet${Math.floor(Math.random() * 1000)}`;
        pushHistory();
        appState.sheets[newName] = createEmptyMatrix(100, 26);
        appState.styles[newName] = {};
        appState.colWidths[newName] = {};
        appState.rowHeights[newName] = {};
        switchSheet(newName);
        showToast(`Đã thêm trang tính mới: ${newName}`, "success");
    };
}

function switchSheet(name) {
    if (appState.isEditing) commitInlineEdit();
    appState.activeSheetName = name;
    appState.selectedCell = { row: 0, col: 0 };
    appState.selectionRange = { startRow: 0, startCol: 0, endRow: 0, endCol: 0 };
    renderTabs();
    renderGrid();
    updateFormulaBar();
    updateCharts();

    const viewport = document.getElementById("grid-scroll-viewport");
    if (viewport) {
        viewport.scrollTop = 0;
        viewport.scrollLeft = 0;
    }
}

function isSheetEmpty(sheetName) {
    const matrix = appState.sheets[sheetName];
    if (!matrix || matrix.length === 0) return true;
    for (let r = 0; r < matrix.length; r++) {
        if (!matrix[r]) continue;
        for (let c = 0; c < matrix[r].length; c++) {
            if (cellToString(matrix[r][c]).trim() !== "") {
                return false;
            }
        }
    }
    return true;
}

let pendingSheetToDelete = null;

function openDeleteSheetModal(sheetName) {
    pendingSheetToDelete = sheetName;
    const modal = document.getElementById("delete-sheet-modal");
    const msg = document.getElementById("delete-sheet-modal-msg");
    if (msg) {
        msg.innerHTML = `Trang tính <strong>"${sheetName}"</strong> có chứa dữ liệu. Bạn có chắc chắn muốn xóa trang tính này không?`;
    }
    if (modal) modal.classList.remove("hidden");
}

function closeDeleteSheetModal() {
    pendingSheetToDelete = null;
    const modal = document.getElementById("delete-sheet-modal");
    if (modal) modal.classList.add("hidden");
}

function deleteSheet(name, force = false) {
    const sheetNames = Object.keys(appState.sheets);
    if (sheetNames.length <= 1) {
        showToast("Bảng tính phải có ít nhất 1 trang tính!", "warning");
        return;
    }

    // If sheet has data and delete not forced, open custom modal
    if (!force && !isSheetEmpty(name)) {
        openDeleteSheetModal(name);
        return;
    }

    if (appState.isEditing) commitInlineEdit();
    pushHistory();

    delete appState.sheets[name];
    delete appState.styles[name];
    delete appState.colWidths[name];
    delete appState.rowHeights[name];
    delete appState.merges[name];
    delete appState.filters[name];

    const remaining = Object.keys(appState.sheets);
    if (appState.activeSheetName === name) {
        switchSheet(remaining[0]);
    } else {
        renderTabs();
        renderGrid();
        updateFormulaBar();
        updateCharts();
    }

    const viewport = document.getElementById("grid-scroll-viewport");
    if (viewport) {
        viewport.scrollTop = 0;
        viewport.scrollLeft = 0;
    }

    showToast(`Đã xóa trang tính: ${name}`, "info");
}

function renameSheetInline(oldName, tabEl) {
    const span = tabEl.querySelector("span");
    const input = document.createElement("input");
    input.type = "text";
    input.value = oldName;
    input.style.width = "80px";
    input.style.fontSize = "11px";
    input.style.border = "1px solid var(--excel-green)";
    input.style.borderRadius = "3px";
    input.style.background = "var(--bg-app)";
    input.style.color = "var(--text-primary)";

    tabEl.replaceChild(input, span);
    input.focus();
    input.select();

    function commitRename() {
        const newName = input.value.trim() || oldName;
        if (newName !== oldName && !appState.sheets[newName]) {
            pushHistory();
            appState.sheets[newName] = appState.sheets[oldName];
            delete appState.sheets[oldName];
            appState.styles[newName] = appState.styles[oldName] || {};
            delete appState.styles[oldName];
            if (appState.merges && appState.merges[oldName]) {
                appState.merges[newName] = appState.merges[oldName];
                delete appState.merges[oldName];
            }
            if (appState.activeSheetName === oldName) appState.activeSheetName = newName;
            showToast(`Đã đổi tên sheet thành: ${newName}`, "success");
        }
        renderTabs();
    }

    input.addEventListener("keydown", (e) => {
        if (e.key === "Enter") commitRename();
        else if (e.key === "Escape") renderTabs();
    });
    input.addEventListener("blur", commitRename);
}

// ==========================================================================
// DIRECT EXCEL (.XLSX) EXPORT & IMPORT
// Critical requirement: "xuất được file xlsx trực tiếp về máy (test kỹ chức năng này nhé)"
// ==========================================================================

function sanitizeFilenameForExcel(rawName) {
    if (!rawName) return "Bang_Tinh_Ket_Qua.xlsx";
    let name = rawName.trim();
    // Normalize Unicode and Vietnamese accents (e.g. đ -> d)
    name = name.normalize("NFD").replace(/[\u0300-\u036f]/g, "");
    name = name.replace(/đ/g, "d").replace(/Đ/g, "D");
    // Replace invalid Windows filename characters
    name = name.replace(/[<>:"/\\|?*\x00-\x1F]/g, "_");
    // Strip trailing dots and spaces
    name = name.replace(/[\s.]+$/, "").trim();
    if (!name.toLowerCase().endsWith(".xlsx")) {
        name += ".xlsx";
    }
    return name || "Bang_Tinh_Ket_Qua.xlsx";
}

function triggerNativeBrowserDownload(downloadUrl, filename) {
    const a = document.createElement("a");
    a.href = downloadUrl;
    if (filename) {
        a.setAttribute("download", filename);
        a.download = filename;
    }
    document.body.appendChild(a);
    a.click();
    setTimeout(() => {
        try { a.remove(); } catch (e) {}
    }, 2000);
}

function buildExportPayload(cleanFilename) {
    const trimmedSheets = {};
    Object.keys(appState.sheets).forEach(sheetName => {
        const matrix = appState.sheets[sheetName];
        let maxR = -1;
        let maxC = -1;
        for (let r = 0; r < matrix.length; r++) {
            for (let c = 0; c < matrix[r].length; c++) {
                const str = cellToString(matrix[r][c]).trim();
                if (str !== "") {
                    if (r > maxR) maxR = r;
                    if (c > maxC) maxC = c;
                }
            }
        }

        // Include merged cell ranges to prevent trimming merged regions
        const sheetMerges = (appState.merges && appState.merges[sheetName]) || [];
        sheetMerges.forEach(mStr => {
            const rng = normalizeRange(mStr);
            if (rng) {
                if (rng.endRow > maxR) maxR = rng.endRow;
                if (rng.endCol > maxC) maxC = rng.endCol;
            }
        });

        const rowLimit = Math.max(maxR + 1, 15);
        const colLimit = Math.max(maxC + 1, 8);
        const sheetRows = [];
        for (let r = 0; r < rowLimit; r++) {
            const rowData = [];
            for (let c = 0; c < colLimit; c++) {
                rowData.push(matrix[r] && matrix[r][c] !== undefined ? matrix[r][c] : "");
            }
            sheetRows.push(rowData);
        }
        trimmedSheets[sheetName] = sheetRows;
    });

    return {
        filename: cleanFilename,
        sheets: trimmedSheets,
        styles: appState.styles,
        colWidths: appState.colWidths,
        merges: appState.merges,
        images: appState.images || {}
    };
}

// 1. Direct Instant Export with Genuine .xlsx File Download
async function exportDirectXLSX() {
    if (appState.isEditing) commitInlineEdit();
    showLoading("Đang xuất file Excel (.xlsx)...");

    // Get latest edited filename from input box
    const nameInput = document.getElementById("workbook-filename");
    if (nameInput && nameInput.value.trim()) {
        appState.filename = nameInput.value.trim();
    }
    const cleanFilename = sanitizeFilenameForExcel(appState.filename);
    const payload = buildExportPayload(cleanFilename);

    try {
        // 1. Server generates full .xlsx and saves directly into user Downloads folder
        const response = await fetch("/prepare_export", {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify(payload)
        });

        if (response.ok) {
            const result = await response.json();
            if (result.success) {
                // Trigger browser native download tray as well
                if (result.download_url) {
                    triggerNativeBrowserDownload(result.download_url, result.filename);
                }
                hideLoading();
                showToast(`✅ Đã lưu file: ${result.filename} về thư mục Downloads!`, "success");
                return;
            }
        }
        throw new Error("Prepare export failed");
    } catch (err) {
        console.warn("Direct server prepare failed, submitting native form:", err);
        // Fallback: Submit form to /download
        let form = document.getElementById("native-download-form");
        if (!form) {
            form = document.createElement("form");
            form.id = "native-download-form";
            form.method = "POST";
            form.action = "/download";
            form.style.display = "none";
            document.body.appendChild(form);
        }
        form.innerHTML = "";
        const input = document.createElement("input");
        input.type = "hidden";
        input.name = "payload";
        input.value = JSON.stringify(payload);
        form.appendChild(input);
        form.submit();
        hideLoading();
        showToast(`✅ Đã lưu file: ${cleanFilename} về thư mục Downloads!`, "success");
    }
}

function exportViaSheetJS(payload, cleanFilename) {
    try {
        if (typeof XLSX === "undefined") {
            throw new Error("Không thể xuất file: SheetJS chưa được tải.");
        }
        const wb = XLSX.utils.book_new();
        Object.keys(payload.sheets).forEach(sheetName => {
            const sheetRows = payload.sheets[sheetName];
            const aoa = sheetRows.map((row, r) => row.map((val, c) => {
                const str = cellToString(val);
                if (str.startsWith("=")) {
                    return { f: str.substring(1), v: displayValue(str, r, c) };
                } else if (str === "") {
                    return null;
                } else {
                    const num = parseFloat(str);
                    return (!isNaN(num) && isFinite(str)) ? num : str;
                }
            }));
            const ws = XLSX.utils.aoa_to_sheet(aoa);
            const safeTitle = sheetName.replace(/[\\/*?:\[\]]/g, "").substring(0, 31) || "Sheet";
            XLSX.utils.book_append_sheet(wb, ws, safeTitle);
        });
        XLSX.writeFile(wb, cleanFilename);
        showToast(`✅ Đã tải file: ${cleanFilename} về thư mục Downloads!`, "success");
    } catch (e) {
        showToast(`Lỗi khi xuất file: ${e.message}`, "error");
    }
}

// 2. Server-side Export using Flask + OpenPyXL
function exportServerXLSX() {
    exportDirectXLSX();
}

// 3. Export to CSV
function exportCSV() {
    const sheetData = appState.sheets[appState.activeSheetName];
    let csvContent = "";

    sheetData.forEach(row => {
        const rowStr = row.map(cell => {
            const str = cellToString(cell).replace(/"/g, '""');
            return `"${str}"`;
        }).join(",");
        csvContent += rowStr + "\r\n";
    });

    const blob = new Blob(["\uFEFF" + csvContent], { type: "text/csv;charset=utf-8;" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `${appState.filename.replace(/\.[^/.]+$/, "")}_${appState.activeSheetName}.csv`;
    a.click();
    URL.revokeObjectURL(url);
    showToast("Đã xuất file CSV thành công!", "success");
}

// ==========================================================================
// Excel File Import Handler
// ==========================================================================
function openImportModal() {
    document.getElementById("upload-modal").classList.remove("hidden");
}

function setupModals() {
    // Close modals
    document.getElementById("btn-close-upload-modal").addEventListener("click", () => {
        document.getElementById("upload-modal").classList.add("hidden");
    });
    document.getElementById("btn-close-func-modal").addEventListener("click", () => {
        document.getElementById("func-wizard-modal").classList.add("hidden");
    });
    document.getElementById("btn-cancel-func").addEventListener("click", () => {
        document.getElementById("func-wizard-modal").classList.add("hidden");
    });
    document.getElementById("btn-close-chart-modal").addEventListener("click", () => {
        document.getElementById("chart-creator-modal").classList.add("hidden");
    });
    document.getElementById("btn-cancel-chart").addEventListener("click", () => {
        document.getElementById("chart-creator-modal").classList.add("hidden");
    });
    document.getElementById("btn-close-find-modal").addEventListener("click", () => {
        document.getElementById("find-replace-modal").classList.add("hidden");
    });

    // Delete sheet modal
    const closeDeleteSheetBtn = document.getElementById("btn-close-delete-sheet-modal");
    if (closeDeleteSheetBtn) closeDeleteSheetBtn.addEventListener("click", closeDeleteSheetModal);
    const cancelDeleteSheetBtn = document.getElementById("btn-cancel-delete-sheet");
    if (cancelDeleteSheetBtn) cancelDeleteSheetBtn.addEventListener("click", closeDeleteSheetModal);
    const confirmDeleteSheetBtn = document.getElementById("btn-confirm-delete-sheet");
    if (confirmDeleteSheetBtn) {
        confirmDeleteSheetBtn.addEventListener("click", () => {
            if (pendingSheetToDelete) {
                const sheetToDelete = pendingSheetToDelete;
                closeDeleteSheetModal();
                deleteSheet(sheetToDelete, true);
            }
        });
    }
    const deleteSheetModalEl = document.getElementById("delete-sheet-modal");
    if (deleteSheetModalEl) {
        deleteSheetModalEl.addEventListener("click", (e) => {
            if (e.target === deleteSheetModalEl) closeDeleteSheetModal();
        });
    }

    // File Drop Zone
    const dropZone = document.getElementById("file-drop-zone");
    const fileInput = document.getElementById("excel-file-input");

    document.getElementById("btn-browse-file").addEventListener("click", () => fileInput.click());

    fileInput.addEventListener("change", (e) => {
        if (e.target.files.length > 0) importExcelFile(e.target.files[0]);
    });

    dropZone.addEventListener("dragover", (e) => {
        e.preventDefault();
        dropZone.classList.add("dragover");
    });
    dropZone.addEventListener("dragleave", () => dropZone.classList.remove("dragover"));
    dropZone.addEventListener("drop", (e) => {
        e.preventDefault();
        dropZone.classList.remove("dragover");
        if (e.dataTransfer.files.length > 0) importExcelFile(e.dataTransfer.files[0]);
    });

    // Template chips inside upload modal
    document.querySelectorAll(".template-chips .chip").forEach(chip => {
        chip.addEventListener("click", () => {
            const tmpl = chip.getAttribute("data-template");
            document.getElementById("upload-modal").classList.add("hidden");
            loadTemplate(tmpl);
        });
    });

    // Chart Modal controls
    document.getElementById("chart-range-input").addEventListener("input", updateChartPreview);
    document.getElementById("chart-type-select").addEventListener("change", updateChartPreview);
    document.getElementById("chart-palette-select").addEventListener("change", updateChartPreview);
    document.getElementById("chart-first-row-header").addEventListener("change", updateChartPreview);

    document.getElementById("btn-create-chart-confirm").addEventListener("click", () => {
        const range = document.getElementById("chart-range-input").value;
        const type = document.getElementById("chart-type-select").value;
        const title = document.getElementById("chart-title-input").value || "Biểu đồ dữ liệu";
        const palette = document.getElementById("chart-palette-select").value;

        createFloatingChart({
            title: title,
            range: range,
            type: type,
            palette: palette,
            x: 200,
            y: 120,
            width: 460,
            height: 310
        });

        document.getElementById("chart-creator-modal").classList.add("hidden");
    });

    // Function Wizard
    setupFuncWizardEvents();

    // Find and Replace
    setupFindReplaceEvents();
}

function importExcelFile(file) {
    showLoading(`Đang đọc file: ${file.name}...`);
    document.getElementById("upload-modal").classList.add("hidden");

    // Try SheetJS client reader first for maximum speed
    const reader = new FileReader();
    reader.onload = function(e) {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: "array", cellFormula: true, cellStyles: true });

            const newSheets = {};
            const newMerges = {};
            workbook.SheetNames.forEach(sheetName => {
                const ws = workbook.Sheets[sheetName];
                const range = XLSX.utils.decode_range(ws["!ref"] || "A1:Z50");
                const rowCount = Math.max(range.e.r + 50, 100);
                const colCount = Math.max(range.e.c + 10, 26);

                const matrix = createEmptyMatrix(rowCount, colCount);

                for (let r = 0; r < rowCount; r++) {
                    for (let c = 0; c < colCount; c++) {
                        const cellAddr = XLSX.utils.encode_cell({ r, c });
                        const cell = ws[cellAddr];
                        if (cell) {
                            if (cell.f) {
                                matrix[r][c] = `=${cell.f}`;
                            } else if (cell.v !== undefined) {
                                matrix[r][c] = cell.v;
                            }
                        }
                    }
                }
                newSheets[sheetName] = matrix;

                if (Array.isArray(ws["!merges"])) {
                    newMerges[sheetName] = ws["!merges"].map(m => XLSX.utils.encode_range(m));
                } else {
                    newMerges[sheetName] = [];
                }
            });

            pushHistory();
            appState.filename = file.name;
            appState.filepath = null;
            appState.sheets = newSheets;
            appState.styles = {};
            appState.colWidths = {};
            appState.rowHeights = {};
            appState.merges = newMerges;
            appState.filters = {};
            appState.activeSheetName = workbook.SheetNames[0];

            document.getElementById("workbook-filename").value = file.name;
            document.title = `${file.name} - Excel Web Pro`;
            renderTabs();
            renderGrid();
            selectCell(0, 0);

            hideLoading();
            setSaveStatus(true, "Đã mở");
            showToast(`Đã mở thành công file: ${file.name} (${workbook.SheetNames.length} sheet)`, "success");

        } catch (clientErr) {
            // Fallback to server upload
            uploadViaServer(file);
        }
    };
    reader.onerror = () => uploadViaServer(file);
    reader.readAsArrayBuffer(file);
}

function uploadViaServer(file) {
    const formData = new FormData();
    formData.append("file", file);

    fetch("/upload", { method: "POST", body: formData })
    .then(res => {
        if (!res.ok) return res.json().then(e => { throw new Error(e.error || "Lỗi tải file"); });
        return res.json();
    })
    .then(data => {
        applyLoadedWorkbook(data);
        hideLoading();
        showToast(`Đã tải lên và mở file thành công: ${data.filename}`, "success");
    })
    .catch(err => {
        hideLoading();
        showToast("Lỗi khi đọc file: " + err.message, "error");
    });
}

// ==========================================================================
// Function Wizard Modal (fx)
// ==========================================================================
function openFuncWizard() {
    renderFuncList();
    document.getElementById("func-wizard-modal").classList.remove("hidden");
}

function setupFuncWizardEvents() {
    const searchInput = document.getElementById("func-search-input");
    const categorySelect = document.getElementById("func-category-select");

    searchInput.addEventListener("input", renderFuncList);
    categorySelect.addEventListener("change", renderFuncList);

    document.getElementById("btn-apply-func").addEventListener("click", () => {
        const activeItem = document.querySelector(".func-item.active");
        if (!activeItem) return;
        const funcName = activeItem.getAttribute("data-func");
        const { row, col } = appState.selectedCell;

        pushHistory();
        appState.sheets[appState.activeSheetName][row][col] = `=${funcName}()`;
        renderGrid();
        updateFormulaBar();
        document.getElementById("func-wizard-modal").classList.add("hidden");

        // Focus formula bar inside parenthesis
        const fInput = document.getElementById("formula-input-field");
        fInput.focus();
        fInput.setSelectionRange(funcName.length + 2, funcName.length + 2);
    });
}

function renderFuncList() {
    const container = document.getElementById("func-list-container");
    container.innerHTML = "";
    const q = document.getElementById("func-search-input").value.trim().toUpperCase();
    const cat = document.getElementById("func-category-select").value;

    const filtered = FORMULA_DEFINITIONS.filter(f => {
        const matchQ = f.name.includes(q) || f.desc.toLowerCase().includes(q.toLowerCase());
        const matchCat = cat === "all" || f.category === cat;
        return matchQ && matchCat;
    });

    filtered.forEach((f, idx) => {
        const item = document.createElement("div");
        item.className = "func-item" + (idx === 0 ? " active" : "");
        item.setAttribute("data-func", f.name);
        item.innerHTML = `<b>${f.name}</b> <span style="font-size: 11px; opacity: 0.7;">(${f.category})</span>`;

        item.addEventListener("click", () => {
            document.querySelectorAll(".func-item").forEach(i => i.classList.remove("active"));
            item.classList.add("active");
            displayFuncDetail(f);
        });

        container.appendChild(item);
    });

    if (filtered.length > 0) displayFuncDetail(filtered[0]);
}

function displayFuncDetail(f) {
    document.getElementById("func-detail-name").textContent = f.name;
    document.getElementById("func-detail-syntax").textContent = f.syntax;
    document.getElementById("func-detail-desc").textContent = f.desc;
}

// ==========================================================================
// Find & Replace Modal
// ==========================================================================
function openFindReplaceModal() {
    document.getElementById("find-replace-modal").classList.remove("hidden");
    document.getElementById("find-search-input").focus();
}

function setupFindReplaceEvents() {
    document.getElementById("btn-find-replace-open").addEventListener("click", openFindReplaceModal);

    document.getElementById("btn-find-next").addEventListener("click", () => {
        const query = document.getElementById("find-search-input").value;
        const matchCase = document.getElementById("find-match-case").checked;
        if (!query) return;

        const sheet = appState.sheets[appState.activeSheetName];
        const curR = appState.selectedCell.row;
        const curC = appState.selectedCell.col;

        let found = false;
        for (let r = 0; r < sheet.length; r++) {
            for (let c = 0; c < sheet[0].length; c++) {
                if (r === curR && c <= curC) continue; // Search forward
                const val = cellToString(sheet[r][c]);
                const match = matchCase ? val.includes(query) : val.toLowerCase().includes(query.toLowerCase());
                if (match) {
                    selectCell(r, c);
                    const el = document.getElementById(`cell-${r}-${c}`);
                    if (el) el.scrollIntoView({ behavior: "smooth", block: "center", inline: "center" });
                    found = true;
                    return;
                }
            }
        }
        if (!found) showToast("Không tìm thấy kết quả nào tiếp theo", "info");
    });

    document.getElementById("btn-replace-one").addEventListener("click", () => {
        const query = document.getElementById("find-search-input").value;
        const replacement = document.getElementById("find-replace-input").value;
        const { row, col } = appState.selectedCell;
        const sheet = appState.sheets[appState.activeSheetName];
        const val = cellToString(sheet[row][col]);

        if (val.toLowerCase().includes(query.toLowerCase())) {
            pushHistory();
            sheet[row][col] = val.replace(new RegExp(query, "i"), replacement);
            renderGrid();
            document.getElementById("btn-find-next").click();
            showToast("Đã thay thế 1 ô", "success");
        }
    });

    document.getElementById("btn-replace-all").addEventListener("click", () => {
        const query = document.getElementById("find-search-input").value;
        const replacement = document.getElementById("find-replace-input").value;
        if (!query) return;

        pushHistory();
        let count = 0;
        const sheet = appState.sheets[appState.activeSheetName];

        for (let r = 0; r < sheet.length; r++) {
            for (let c = 0; c < sheet[0].length; c++) {
                const val = cellToString(sheet[r][c]);
                if (val.toLowerCase().includes(query.toLowerCase())) {
                    sheet[r][c] = val.replaceAll(query, replacement);
                    count++;
                }
            }
        }
        renderGrid();
        showToast(`Đã thay thế tất cả ${count} ô phù hợp!`, "success");
    });
}

// ==========================================================================
// Right-Click Context Menu
// ==========================================================================
function setupContextMenu() {
    const contextMenu = document.getElementById("grid-context-menu");
    const workspace = document.getElementById("spreadsheet-workspace");

    workspace.addEventListener("contextmenu", (e) => {
        const cell = e.target.closest(".grid-cell");
        if (cell) {
            e.preventDefault();
            const r = parseInt(cell.getAttribute("data-row"), 10);
            const c = parseInt(cell.getAttribute("data-col"), 10);
            if (!cell.classList.contains("selected")) {
                selectCell(r, c);
            }
            contextMenu.style.top = `${Math.min(e.clientY, window.innerHeight - 300)}px`;
            contextMenu.style.left = `${Math.min(e.clientX, window.innerWidth - 240)}px`;
            contextMenu.classList.remove("hidden");
        }
    });

    document.addEventListener("click", () => {
        contextMenu.classList.add("hidden");
        document.getElementById("tab-context-menu").classList.add("hidden");
    });

    document.getElementById("ctx-cut").onclick = cutSelection;
    document.getElementById("ctx-copy").onclick = copySelection;
    document.getElementById("ctx-paste").onclick = pasteClipboard;
    const ctxInsertImg = document.getElementById("ctx-insert-image");
    if (ctxInsertImg) {
        ctxInsertImg.onclick = () => {
            const input = document.getElementById("image-file-input");
            if (input) input.click();
        };
    }
    const ctxMerge = document.getElementById("ctx-merge");
    if (ctxMerge) {
        ctxMerge.onclick = toggleMergeCells;
    }
    document.getElementById("ctx-insert-row").onclick = insertRowAbove;
    document.getElementById("ctx-insert-col").onclick = insertColLeft;
    document.getElementById("ctx-delete-row").onclick = deleteActiveRow;
    document.getElementById("ctx-delete-col").onclick = deleteActiveCol;
    document.getElementById("ctx-create-chart").onclick = openChartModal;
    document.getElementById("ctx-clear").onclick = () => document.getElementById("btn-clear-selection").click();
}

function openTabContextMenu(e, sheetName) {
    const tabMenu = document.getElementById("tab-context-menu");
    tabMenu.style.top = `${e.clientY - 80}px`;
    tabMenu.style.left = `${e.clientX}px`;
    tabMenu.classList.remove("hidden");

    document.getElementById("ctx-tab-rename").onclick = () => {
        const tabEl = document.querySelector(`.sheet-tab-item[data-sheet-name="${sheetName}"]`);
        if (tabEl) renameSheetInline(sheetName, tabEl);
    };

    document.getElementById("ctx-tab-duplicate").onclick = () => {
        pushHistory();
        const dupName = `${sheetName}_Copy`;
        appState.sheets[dupName] = JSON.parse(JSON.stringify(appState.sheets[sheetName]));
        appState.styles[dupName] = JSON.parse(JSON.stringify(appState.styles[sheetName] || {}));
        if (appState.merges && appState.merges[sheetName]) {
            appState.merges[dupName] = JSON.parse(JSON.stringify(appState.merges[sheetName]));
        }
        switchSheet(dupName);
        showToast(`Đã nhân bản: ${dupName}`, "success");
    };

    document.getElementById("ctx-tab-delete").onclick = () => deleteSheet(sheetName);
}

// ==========================================================================
// Zoom & Theme Controls
// ==========================================================================
function setupZoomAndTheme() {
    const slider = document.getElementById("zoom-slider");
    const label = document.getElementById("zoom-label");
    const gridEl = document.getElementById("grid-element");

    function applyZoom(val) {
        appState.zoom = val;
        slider.value = val;
        label.textContent = `${val}%`;
        gridEl.style.transform = `scale(${val / 100})`;
        gridEl.style.transformOrigin = "top left";
    }

    slider.addEventListener("input", (e) => applyZoom(parseInt(e.target.value, 10)));
    document.getElementById("btn-zoom-in").addEventListener("click", () => applyZoom(Math.min(appState.zoom + 10, 200)));
    document.getElementById("btn-zoom-out").addEventListener("click", () => applyZoom(Math.max(appState.zoom - 10, 50)));

    document.querySelectorAll(".btn-zoom-preset").forEach(btn => {
        btn.addEventListener("click", () => {
            const z = parseInt(btn.getAttribute("data-zoom"), 10);
            applyZoom(z);
            document.querySelectorAll(".btn-zoom-preset").forEach(b => b.classList.remove("active"));
            btn.classList.add("active");
        });
    });

    // Theme toggle
    const themeBtn = document.getElementById("btn-theme-toggle");
    const sunIcon = document.getElementById("theme-icon-sun");
    const moonIcon = document.getElementById("theme-icon-moon");

    themeBtn.addEventListener("click", () => {
        const nextTheme = appState.theme === "dark" ? "light" : "dark";
        appState.theme = nextTheme;
        document.documentElement.setAttribute("data-theme", nextTheme);

        if (nextTheme === "light") {
            sunIcon.classList.remove("hidden");
            moonIcon.classList.add("hidden");
        } else {
            sunIcon.classList.add("hidden");
            moonIcon.classList.remove("hidden");
        }
        showToast(`Chuyển sang giao diện: ${nextTheme === "dark" ? "Tối (Dark Mode)" : "Sáng (Office Light)"}`, "info");
    });
}
