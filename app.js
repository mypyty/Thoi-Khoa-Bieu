/* ============================================================
   LOAD PAGE
============================================================ */
function loadPage(name) {
    const container = document.getElementById("section-content");
    if (!container) return console.error("Không tìm thấy #section-content");

    // Highlight menu đang chọn
    try{
        document.querySelectorAll('.app-nav .btn[data-page]').forEach(b=>{
            b.classList.toggle('nav-active', (b.dataset.page || '') === name);
        });
    }catch(_){ /* ignore */ }

    if (name === "pccm") {
        container.innerHTML = renderPCCM();
        return;
    }

    // Menu "Môn học" là môn tổng hợp (mon.xlsx) => dùng section monhoc
    if (name === "mon") {
        renderSectionInto("monhoc", "section-content", document);
        return;
    }

    // Menu "Tiết chuẩn" (nhập Excel riêng) => hiển thị theo Khối + Môn + Số tiết/tuần + Giới hạn/buổi
    if (name === "tietchuan") {
        container.innerHTML = renderTietChuanPage();
        return;
    }

    renderSectionInto(name, "section-content", document);
}

/* ============================================================
   LOCAL STORAGE INIT
============================================================ */
// ============================================================
// B1 STORAGE (SQLite/sql.js + IndexedDB) — tách theo TRƯỜNG
// Mỗi trường = 1 DB riêng trong IndexedDB (không lẫn dữ liệu)
// URL: index.html?school=TruongA
// ============================================================
let __kv = null;

// LocalStorage backup per school (đảm bảo dữ liệu không bị "mất trắng" khi KVDB/sql.js lỗi)
function _lsKey(schoolId){
    return `TKB_STORE::${_sanitizeSchoolId(schoolId || CTX.schoolId || getSchoolId())}`;
}

function _safeParseJSON(raw, fallback){
    try{
        return raw ? JSON.parse(raw) : (fallback ?? {});
    }catch(e){
        console.warn("JSON parse failed; reset to empty", e);
        return (fallback ?? {});
    }
}

function _sanitizeSchoolId(x){
    x = (x || "default").toString().trim();
    // cho phép chữ, số, gạch dưới, gạch ngang
    x = x.replace(/[^0-9a-zA-Z_\-]/g, "_");
    if (!x) x = "default";
    return x;
}

// Map lưu TÊN HIỂN THỊ (có dấu) theo mã trường (đã sanitize)
function _schoolNameMap(){
    return _safeParseJSON(localStorage.getItem("TKB_SCHOOL_NAMES"), {});
}

function _getSchoolName(sid){
    const key = _sanitizeSchoolId(sid);
    const map = _schoolNameMap();
    const v = map ? map[key] : "";
    return (v == null) ? "" : String(v);
}

function _setSchoolName(sid, name){
    const key = _sanitizeSchoolId(sid);
    const n = (name||"").toString().trim();
    if(!key || !n) return;
    const map = _schoolNameMap();
    map[key] = n;
    try{ localStorage.setItem("TKB_SCHOOL_NAMES", JSON.stringify(map)); }catch(e){ /* ignore */ }
}

function _deleteSchoolName(sid){
    const key = _sanitizeSchoolId(sid);
    if(!key) return;
    const map = _schoolNameMap();
    if(map && Object.prototype.hasOwnProperty.call(map, key)){
        delete map[key];
        try{ localStorage.setItem("TKB_SCHOOL_NAMES", JSON.stringify(map)); }catch(e){ /* ignore */ }
    }
}

function _prettySchoolLabel(x){
    return (x||"").toString().trim()
        .replace(/_/g, " ")
        .replace(/\s+/g, " ")
        .trim();
}

function _getSchoolFromURL(){
    try{
        const u = new URL(window.location.href);
        const s = u.searchParams.get("school");
        // Trả về RAW để có thể hiển thị đúng dấu. Mã dùng cho storage sẽ sanitize riêng.
        return s ? String(s).trim() : "";
    }catch(e){
        return "";
    }
}

let CTX = { schoolId: "", schoolLabel: "" };

function getSchoolId(){
    const fromUrlRaw = _getSchoolFromURL();
    const fromLSKey  = _sanitizeSchoolId(localStorage.getItem("TKB_LAST_SCHOOL"));
    const schoolId   = _sanitizeSchoolId(fromUrlRaw || fromLSKey || "default");

    // Label ưu tiên:
    // - Nếu URL có chứa ký tự "bị sanitize" (có dấu/space/...) => coi là TÊN HIỂN THỊ và lưu mapping.
    // - Nếu URL chỉ là mã đã sanitize => dùng mapping đã lưu hoặc fallback.
    let label = "";
    if(fromUrlRaw){
        const raw = String(fromUrlRaw).trim();
        if(raw && _sanitizeSchoolId(raw) !== raw){
            label = raw;
            _setSchoolName(schoolId, label);
        }
    }

    if(!label){
        label = _getSchoolName(schoolId) || localStorage.getItem("TKB_LAST_SCHOOL_LABEL") || schoolId;
    }

    CTX.schoolId = schoolId;
    CTX.schoolLabel = _prettySchoolLabel(label || schoolId) || schoolId;

    try{
        localStorage.setItem("TKB_LAST_SCHOOL", schoolId);
        localStorage.setItem("TKB_LAST_SCHOOL_LABEL", CTX.schoolLabel);
    }catch(e){ /* ignore */ }

    // đảm bảo trường hiện tại luôn có trong danh sách trường
    try{ if(typeof addSchoolToList === "function") addSchoolToList(schoolId); }catch(_){ /* ignore */ }
    return schoolId;
}

function getSchoolLabel(){
    // Đảm bảo CTX được set
    if(!CTX.schoolId) getSchoolId();
    return CTX.schoolLabel || _prettySchoolLabel(CTX.schoolId) || CTX.schoolId || "default";
}

function getDbName(){
    return `TKB::SCHOOL::${getSchoolId()}`;
}

// DATA sẽ được nạp trong appBoot()
let DATA = {};

/* ============================================================
   EXCEL-LIKE TABLE UX (ROW SELECT + INLINE EDIT)
   - Click: select
   - Ctrl/Cmd + Click: toggle multi-select
   - Shift + Click: range select
   - Double click: edit row inline
============================================================ */
let TABLE_SELECTION = {};   // { section: Set(index) }
let TABLE_LAST_INDEX = {};  // { section: lastClickedIndex }
let INLINE_EDIT = { section: "", index: -1 };

function ensureDataShape(){
    ["khoi","lop","giaovien","monhoc","mon","phong"].forEach(sec=>{
        if (!Array.isArray(DATA[sec])) DATA[sec]=[];
    });
    if (!DATA.pccmMatrix) DATA.pccmMatrix={};
    if (!DATA.pccmRoomMatrix) DATA.pccmRoomMatrix={}; // key: "Lop|Mon" -> "P101"

    // (NEW) Lưu riêng số tiết/giới hạn theo Lớp|Môn để màn "Sắp xếp TKB" lấy đúng từ "Bảng phân công"
    // key: "Lop|Mon" -> "4" (string)
    if (!DATA.pccmTietMatrix) DATA.pccmTietMatrix = {};
    if (!DATA.pccmGioihanMatrix) DATA.pccmGioihanMatrix = {};
}

/* ============================================================
   ĐỒNG BỘ DỮ LIỆU (LIÊN KẾT CÁC BẢNG)
   - Khi Môn / Lớp / Giáo viên bị xoá → các bảng phụ thuộc phải trống
   - Khi chỉ xoá 1 phần → tự động dọn rác (orphan) theo dữ liệu còn lại
============================================================ */
function syncDerivedDataIntegrity(){
    let changed = false;

    const hasMonHoc = Array.isArray(DATA.monhoc) && DATA.monhoc.length > 0;
    const hasLop    = Array.isArray(DATA.lop) && DATA.lop.length > 0;
    const hasGV     = Array.isArray(DATA.giaovien) && DATA.giaovien.length > 0;

    // ===== valid sets =====
    const classNormSet = new Set((DATA.lop || [])
        .map(canonTen2FromLop)
        .filter(Boolean));

    const classIdSet = new Set((DATA.lop || [])
        .map(l => String(l.id))
        .filter(Boolean));

    const monAliasSet = new Set();
    (DATA.monhoc || []).forEach(m => {
        [m.ten, m.ma, m.ma2].forEach(v => {
            const s = _normText(v).toLowerCase();
            if (s) monAliasSet.add(s);
        });
    });
    // Thêm cả các key mà PCCM có thể lưu (mã nếu có)
    try{
        (buildPCCMMonList() || []).forEach(x => {
            [x.key, x.ten, x.code, x.ma, x.ma2].forEach(v => {
                const s = _normText(v).toLowerCase();
                if (s) monAliasSet.add(s);
            });
        });
    }catch(_){ /* ignore */ }

    const gvCodeSet = new Set((DATA.giaovien || [])
        .map(g => _normText(g.magv).toUpperCase())
        .filter(Boolean));

    const pruneObj = (obj, keepFn) => {
        if (!obj || typeof obj !== "object") return false;
        let ch = false;
        Object.keys(obj).forEach(k => {
            try{
                if (!keepFn(k, obj[k])){
                    delete obj[k];
                    ch = true;
                }
            }catch(_){
                delete obj[k];
                ch = true;
            }
        });
        return ch;
    };

    const clearObjIfNotEmpty = (field) => {
        if (DATA[field] && typeof DATA[field] === "object" && Object.keys(DATA[field]).length){
            DATA[field] = {};
            changed = true;
        }
    };

    // ===== 1) PCCM & phụ thuộc =====
    // Nếu thiếu 1 trong 3 dữ liệu cốt lõi → phân công phải trống
    if (!hasMonHoc || !hasLop || !hasGV){
        clearObjIfNotEmpty("pccmMatrix");
        clearObjIfNotEmpty("pccmRoomMatrix");
        clearObjIfNotEmpty("pccmTietMatrix");
        clearObjIfNotEmpty("pccmGioihanMatrix");
    } else {
        const keepKeyByClassMon = (key) => {
            const parts = String(key).split("|");
            if (parts.length < 2) return false;
            const cls = normalizeClassName(parts[0]);
            const mon = _normText(parts.slice(1).join("|")).toLowerCase();
            if (!cls || !classNormSet.has(cls)) return false;
            if (!mon || !monAliasSet.has(mon)) return false;
            return true;
        };

        changed = pruneObj(DATA.pccmMatrix, (k, v) => {
            if (!keepKeyByClassMon(k)) return false;
            const gv = _normText(v).toUpperCase();
            if (!gv) return false; // rỗng thì coi như chưa phân công → xoá key
            return gvCodeSet.has(gv);
        }) || changed;

        changed = pruneObj(DATA.pccmRoomMatrix, (k, v) => {
            if (!keepKeyByClassMon(k)) return false;
            const room = _normText(v);
            return room !== ""; // rỗng thì bỏ
        }) || changed;

        changed = pruneObj(DATA.pccmTietMatrix, (k, v) => {
            if (!keepKeyByClassMon(k)) return false;
            const n = Number(String(v).trim());
            return Number.isFinite(n) && n > 0;
        }) || changed;

        changed = pruneObj(DATA.pccmGioihanMatrix, (k, v) => {
            if (!keepKeyByClassMon(k)) return false;
            const n = Number(String(v).trim());
            return Number.isFinite(n) && n > 0;
        }) || changed;
    }

    // ===== 2) Tiết chuẩn =====
    // Nếu không có Môn hoặc Lớp → tiết chuẩn trở về mặc định (trống, sẽ auto-sync lại khi nạp mới)
    if (!hasMonHoc || !hasLop){
        if (Array.isArray(DATA.mon) && DATA.mon.length){
            DATA.mon = [];
            changed = true;
        }
    } else {
        // Dọn các dòng tiết chuẩn không còn thuộc môn hiện tại
        if (Array.isArray(DATA.mon) && DATA.mon.length){
            const before = DATA.mon.length;
            DATA.mon = DATA.mon.filter(r => {
                const ten = _normText(r.ten).toLowerCase();
                return ten && monAliasSet.has(ten);
            });
            if (DATA.mon.length !== before) changed = true;
        }
    }

    // ===== 3) TKB theo lớp (id) =====
    if (!hasLop){
        if (DATA.tkb && typeof DATA.tkb === "object" && Object.keys(DATA.tkb).length){
            DATA.tkb = {};
            changed = true;
        }
    } else {
        if (!DATA.tkb || typeof DATA.tkb !== "object") DATA.tkb = {};
        Object.keys(DATA.tkb).forEach(id => {
            if (!classIdSet.has(String(id))){
                delete DATA.tkb[id];
                changed = true;
            }
        });
    }

    return changed;
}

function saveStore(){
    // Luôn backup vào localStorage theo trường để tránh mất dữ liệu khi KVDB/sql.js/IndexedDB lỗi.
    const sid = CTX.schoolId || getSchoolId();
    const json = JSON.stringify(DATA);

    try{
        localStorage.setItem(_lsKey(sid), json);
    }catch(e){
        console.warn("saveStore localStorage failed", e);
    }

    // KVDB.set trả Promise, nhưng ta không cần await để UI mượt
    try{
        if (__kv) __kv.set("DATA_JSON", json);
    }catch(e){
        console.warn("saveStore KVDB failed", e);
    }
}

function updateSchoolBadge(){
    const badge = document.getElementById("schoolBadgeText");
    if (badge) badge.textContent = `Trường: ${getSchoolLabel()}`;
}

// Nút đổi trường trên UI
function changeSchool(){
    const cur = CTX.schoolId || getSchoolId();
    const next = prompt("Nhập mã trường (ví dụ: TruongA, TruongB). Mỗi mã là 1 dữ liệu riêng:", cur);
    if (next === null) return;
    const sid = _sanitizeSchoolId(next);
    // lưu tên hiển thị (có thể có dấu)
    try{ if(_sanitizeSchoolId(next) !== String(next).trim()) _setSchoolName(sid, String(next).trim()); }catch(_){ }
    // giữ lại path, đổi param school
    const u = new URL(window.location.href);
    u.searchParams.set("school", sid);
    window.location.href = u.toString();
}

// Mở trang Sắp xếp TKB (sapxep.html). Tự mang theo schoolId để dùng đúng dữ liệu trường.
function openTKBPlanner(){
    const u = new URL("pages/sapxep.html", window.location.href);

    // sid: mã trường ổn định (dùng cho storage)
    const sid = getSchoolId();

    // school: tên hiển thị (có dấu) để URL dễ đọc
    const label = (typeof getSchoolLabel === "function")
        ? getSchoolLabel()
        : (CTX.schoolLabel || sid);

    u.searchParams.set("sid", sid);
    u.searchParams.set("school", label);
    window.location.href = u.toString();
}



async function appBoot(){
    // mở DB theo schoolId
    const schoolId = getSchoolId();
    CTX.schoolId = schoolId;

    // Ưu tiên lấy từ KVDB (nếu có), nhưng luôn có backup localStorage
    const rawLS = localStorage.getItem(_lsKey(schoolId));

    if (window.KVDB) {
        try{
            __kv = await window.KVDB.open(getDbName());
            let raw = await __kv.get("DATA_JSON");

            // Nếu KVDB trống nhưng localStorage có dữ liệu => nạp từ backup và seed lại KVDB
            if (!raw && rawLS) {
                raw = rawLS;
                try{ await __kv.set("DATA_JSON", raw); }catch(e){ /* ignore */ }
            }

            DATA = _safeParseJSON(raw, {});
        }catch(e){
            console.warn("KVDB init/load failed; fallback to localStorage", e);
            __kv = null;
            DATA = _safeParseJSON(rawLS, {});
        }
    } else {
        // không có KVDB (thiếu sql.js) => dùng localStorage theo trường
        DATA = _safeParseJSON(rawLS, {});
    }

    ensureDataShape();

    // Đồng bộ dữ liệu giữa các bảng: nếu thiếu Môn/Lớp/GV thì các bảng phụ thuộc (phân công, tiết chuẩn...) phải trống.
    // Đồng thời tự dọn các liên kết mồ côi khi người dùng xoá / nạp lại dữ liệu.
    try{
        const changed = syncDerivedDataIntegrity();
        if (changed) saveStore();
    }catch(e){
        console.warn("syncDerivedDataIntegrity failed", e);
    }
    updateSchoolBadge();
    loadPage("khoi");
}

/* ============================================================
   KHỐI HỌC — EXTRACT NUMBER (Ví dụ: "Khối 6" → "6")
============================================================ */
function extractKhoiNumber(str){
    return (str+"").match(/\d+/)?.[0] || "";
}

/* ============================================================
   CHUẨN HOÁ TÊN LỚP — 6.1 / 6-1 / 6/1 → 6A1
============================================================ */
function normalizeClassName(name){
    if (!name) return "";
    name = name.trim();

    // Nếu đúng rồi → giữ nguyên
    if (/^\d+A\d+$/i.test(name)) return name.toUpperCase();

    // 6.1 → 6A1
    let m = name.match(/^(\d+)[\.\-_/ ]+(\d+)$/);
    if (m) return `${m[1]}A${m[2]}`.toUpperCase();

    // 6A01 → 6A1
    let m2 = name.match(/^(\d+)A0?(\d+)$/i);
    if (m2) return `${m2[1]}A${m2[2]}`.toUpperCase();

    return name.toUpperCase();
}


/* ============================================================
   PCCM / TKB: CHỈ DÙNG TÊN LỚP 2 (ten2) — và CHUẨN HOÁ để không bị 6/1 & 6A1
============================================================ */
function canonTen2FromLop(l){
    const t2 = (l && l.ten2 ? String(l.ten2) : "").trim();
    if (t2) return normalizeClassName(t2);   // chỉ lấy tên lớp 2
    // fallback nhẹ: nếu chưa có ten2 thì thử lấy từ "ten" để không bị mất lớp
    const t = (l && l.ten ? String(l.ten) : "").trim();
    return t ? normalizeClassName(t) : "";
}


/* ============================================================
   FORM CONFIG — ĐÚNG THEO EXCEL (Lớp + Môn)
============================================================ */
const FORM_CONFIG={
    lop:{
        label:"Lớp học",
        fields:[
            {k:"ten",label:"Tên lớp"},
            {k:"khoi",label:"Khối học"},
            {k:"buoi",label:"Buổi học"},
            {k:"diadiem",label:"Địa điểm"},
            {k:"dienthoai",label:"Điện thoại di động"},
            {k:"email",label:"Email"},
            {k:"zalo",label:"Zalo UID"},
            {k:"ghichu",label:"Ghi chú"}
        ]
    },

    monhoc:{
        label:"Môn học",
        fields:[
            {k:"ten",label:"Tên môn học"},
            {k:"ma",label:"Mã môn học"},
            {k:"giaoan",label:"Có giáo án"},
            {k:"ghichu",label:"Ghi chú"}
        ]
    },

    mon:{
        label:"Tiết chuẩn",
        fields:[
            {k:"khoi",label:"Khối học"},
            {k:"ten",label:"Môn học"},
            {k:"sotiet",label:"Số tiết/1 tuần"},
            {k:"gioihan",label:"Giới hạn số tiết/1 buổi"},
            {k:"ghichu",label:"Ghi chú"}
        ]
    },

    khoi:{label:"Khối học",fields:[
        {k:"ten",label:"Tên khối học"},
        {k:"makhoi",label:"Mã khối học"},
        {k:"ghichu",label:"Ghi chú"}
    ]},

    giaovien:{label:"Giáo viên",fields:[
        {k:"hodem",label:"Họ đệm"},
        {k:"ten",label:"Tên"},
        {k:"magv",label:"Mã GV"},
        {k:"email",label:"Email"},
        {k:"zalo",label:"Zalo"},
        {k:"ghichu",label:"Ghi chú"}
    ]},

    phong:{label:"Phòng học",fields:[
        {k:"ten",label:"Tên phòng"},
        {k:"tinhtrang",label:"Mã môn học"}
    ]}
};

/* ============================================================
   IMPORT EXCEL — CHUẨN (Lớp & Môn theo Excel)
============================================================ */
let IMPORT_SECTION="";
let IS_PCCM_IMPORT=false;

function triggerExcel(section){
    IMPORT_SECTION=section;
    IS_PCCM_IMPORT=false;
    document.getElementById("excelFile").click();
}

document.addEventListener("DOMContentLoaded",()=>{
    if (!document.getElementById("excelFile")){
        const inp=document.createElement("input");
        inp.type="file";
        inp.accept=".xlsx,.xls";
        inp.id="excelFile";
        inp.style.display="none";
        document.body.appendChild(inp);
    }
    document.getElementById("excelFile").addEventListener("change",readExcel);

    // Tiết chuẩn: hỗ trợ Ctrl/Shift chọn nhiều ô + Ctrl/Cmd+C, Ctrl/Cmd+V
    // (Chỉ kích hoạt khi đang ở trang Tiết chuẩn)
    document.addEventListener("keydown", tcGlobalKeyDown, true);
    document.addEventListener("paste", tcGlobalPaste, true);

    // Boot app (nạp DB theo trường) rồi render tab mặc định
    appBoot().catch(err=>{
        console.error(err);
        const c=document.getElementById('section-content');
        if (c) c.innerHTML = '<div style="padding:16px;color:#c00">Lỗi khởi tạo dữ liệu: '+(err && err.message ? err.message : err)+'</div>';
    });
});

function readExcel(e){
    const file = e.target.files && e.target.files[0];
    if (!file) return;

    if (!window.XLSX){
        alert("❌ Chưa tải thư viện Excel (XLSX). Hãy kiểm tra kết nối mạng hoặc thẻ <script src=...xlsx...> trong HTML.");
        e.target.value = "";
        return;
    }

    const reader = new FileReader();

    reader.onload = (evt)=>{
        try{
            const data = evt.target.result;

            // Ưu tiên ArrayBuffer (ổn định trên Chrome/Edge/Safari). Fallback binary nếu cần.
            let wb;
            if (data instanceof ArrayBuffer){
                wb = XLSX.read(data, { type: "array" });
            }else{
                wb = XLSX.read(data, { type: "binary" });
            }

            if (IS_PCCM_IMPORT) importPCCMFromExcel(wb);
            else importFromExcel(wb);
        }catch(err){
            console.error(err);
            alert("❌ Không đọc được file Excel. Vui lòng kiểm tra định dạng .xlsx/.xls hoặc thử lưu lại file rồi nhập lại.");
        }finally{
            // reset input để có thể chọn lại cùng 1 file
            e.target.value = "";
        }
    };

    reader.onerror = (err)=>{
        console.error(err);
        alert("❌ Lỗi đọc file. Vui lòng thử lại.");
        e.target.value = "";
    };

    // ArrayBuffer works best across browsers
    try{
        reader.readAsArrayBuffer(file);
    }catch(_){
        // Old fallback
        reader.readAsBinaryString(file);
    }
}

/* ============================================================
   IMPORT CHÍNH (Lớp / Môn / Mặc định)
============================================================ */
function importFromExcel(wb){
    const sheet=wb.Sheets[wb.SheetNames[0]];
    const rows=XLSX.utils.sheet_to_json(sheet,{defval:""});

    if (IMPORT_SECTION==="lop") return importExcel_Lop(rows);
    if (IMPORT_SECTION==="monhoc") return importExcel_MonHoc(rows);
    if (IMPORT_SECTION==="mon") return importExcel_Mon(rows);

    return importExcel_Default(rows);
}

/* ============================================================
   IMPORT LỚP (Chuẩn Excel)
============================================================ */
function importExcel_Lop(rows){
    rows.forEach(r=>{
        let obj={};

        obj.ten = r["Tên lớp"]?.trim() || "";
        obj.ten2 = r["Tên lớp 2"]?.trim() || "";
        obj.khoi = r["Khối học"]?.trim() || "";
        obj.buoi = r["Buổi học"]?.trim() || "";
        obj.diadiem = r["Địa điểm"]?.trim() || "";
        obj.dienthoai = r["Điện thoại di động"]?.trim() || "";
        obj.email = r["Email"]?.trim() || "";
        obj.zalo = r["Zalo UID"]?.trim() || "";
        obj.ghichu = r["Ghi chú"]?.trim() || "";

        // Chuẩn tên lớp
        let canon = normalizeClassName(obj.ten);
        if (!obj.ten2) obj.ten2 = canon;

        obj.id = autoID("lop");
        DATA.lop.push(obj);
    });

    saveStore();
    renderSectionInto("lop","section-content",document);
    alert("✔ Import lớp thành công!");
}

/* ============================================================
   IMPORT MÔN (Chuẩn Excel)
============================================================ */
function importExcel_Mon(rows){
    if (!Array.isArray(rows) || !rows.length){
        alert("⚠ File Excel trống hoặc không đọc được dữ liệu.");
        return;
    }

    if (!Array.isArray(DATA.mon)) DATA.mon = [];

    const _norm = (s)=> (s ?? "")
        .toString()
        .normalize('NFC')
        .trim()
        .replace(/\s+/g,' ')
        .toLowerCase();

    const _canonKhoi = (s)=>{
        const t = (s ?? "").toString().normalize('NFC').trim();
        const kn = extractKhoiNumber(t);
        return kn ? `Khối ${kn}` : t;
    };

    const _normKhoiKey = (s)=>{
        const t = (s ?? "").toString().normalize('NFC').trim();
        const m = t.match(/\d+/);
        if (m) return `khối ${m[0]}`;
        return _norm(t);
    };

    // ===== helper: tìm môn theo mọi alias (id/ten/ma/ma2) =====
    const _findMonhoc = (monValue)=>{
        const v = (monValue ?? "").toString().normalize('NFC').trim();
        if(!v) return null;
        const low = v.toLowerCase();
        return (DATA.monhoc || []).find(r=>{
            const id  = _normText(r.id).toLowerCase();
            const ten = _normText(r.ten).toLowerCase();
            const ma  = _normText(r.ma).toLowerCase();
            const ma2 = _normText(r.ma2).toLowerCase();
            return (id && id === low) || (ten && ten === low) || (ma && ma === low) || (ma2 && ma2 === low);
        }) || null;
    };

    const _canonMon = (monRaw)=>{
        const found = _findMonhoc(monRaw);
        if(!found) return _normText(monRaw);
        return _normText(found.ten) || _normText(monRaw);
    };

    // ===== 1) Dọn trùng trong dữ liệu cũ (Khối + Môn) để tránh dữ liệu phình / import lần 2 bị đúp =====
    // Đồng thời chuẩn hoá:
    // - Khối: về "Khối X"
    // - Môn: nếu đang lưu theo mã (KHTN...) thì map về tên chuẩn trong bảng Môn học (DATA.monhoc.ten)
    const _seen = new Map(); // key -> newIndex
    const _clean = [];
    let changed = false;

    (DATA.mon || []).forEach((m)=>{
        if (!m || typeof m !== "object"){
            _clean.push(m);
            return;
        }

        const khoi0Raw = (m?.khoi || "").toString().trim();
        const ten0Raw  = (m?.ten  || "").toString().trim();
        if (!khoi0Raw || !ten0Raw){
            _clean.push(m);
            return;
        }

        const khoi0 = _canonKhoi(khoi0Raw);
        const ten0  = _canonMon(ten0Raw);

        if (_normText(m.khoi) !== khoi0){ m.khoi = khoi0; changed = true; }
        if (_normText(m.ten)  !== ten0 ){ m.ten  = ten0;  changed = true; }

        const k0 = `${_normKhoiKey(khoi0)}|${_norm(ten0)}`;
        const exIdx = _seen.get(k0);

        if (exIdx === undefined){
            _seen.set(k0, _clean.length);
            _clean.push(m);
        }else{
            // merge vào dòng đã có (ưu tiên giá trị đang có; nếu trống/không hợp lệ thì lấy từ dòng trùng)
            const ex = _clean[exIdx];
            const st1 = String(ex.sotiet ?? "").trim();
            const gh1 = String(ex.gioihan ?? "").trim();
            const st2 = String(m.sotiet ?? "").trim();
            const gh2 = String(m.gioihan ?? "").trim();

            if ((st1 === "" || Number.isNaN(Number(st1))) && st2 !== "" && !Number.isNaN(Number(st2))){
                ex.sotiet = st2; changed = true;
            }
            if ((gh1 === "" || Number.isNaN(Number(gh1))) && gh2 !== "" && !Number.isNaN(Number(gh2))){
                ex.gioihan = gh2; changed = true;
            }

            const note1 = String(ex.ghichu ?? "").trim();
            const note2 = String(m.ghichu ?? "").trim();
            if (!note1 && note2){
                ex.ghichu = note2; changed = true;
            }

            changed = true; // bỏ dòng trùng
        }
    });

    if (_clean.length !== (DATA.mon||[]).length){
        DATA.mon = _clean;
    }

    // ===== 2) Chuẩn hoá dữ liệu từ Excel trước khi ghi =====
    const ops = [];
    let unknownMonCount = 0;

    rows.forEach(r=>{
        const khoiRaw = (r["Khối học"] ?? r["Khối"] ?? "").toString().trim() || "";
        const tenRaw  = (r["Môn học"] ?? r["Tên môn"] ?? r["Tên môn học"] ?? "").toString().trim() || "";
        if(!khoiRaw || !tenRaw) return;

        const khoi = _canonKhoi(khoiRaw);
        const found = _findMonhoc(tenRaw);
        const ten = found ? (_normText(found.ten) || _normText(tenRaw)) : _normText(tenRaw);
        if (!found) unknownMonCount++;

        const sotiet = (r["Số tiết/1 tuần"] ?? r["Số tiết"] ?? "").toString().trim() || "";
        const gioihan = (r["Giới hạn số tiết/1 buổi"] ?? r["Giới hạn"] ?? "").toString().trim() || "1";
        const ghichu = (r["Ghi chú"] ?? "").toString().trim() || "";

        const key = `${_normKhoiKey(khoi)}|${_norm(ten)}`;
        ops.push({ key, khoi, ten, sotiet, gioihan, ghichu, tenRaw, hasMonHoc: !!found });
    });

    if (!ops.length){
        alert("⚠ Không tìm thấy dòng hợp lệ (cần có cột Khối học + Môn học).");
        return;
    }

    // Gộp trùng trong file (cùng Khối+Môn) => lấy dòng CUỐI CÙNG
    const opMap = new Map();
    ops.forEach(op=>opMap.set(op.key, op));
    const uniqOps = Array.from(opMap.values());

    // ===== 3) Tính toán tác động để hỏi người dùng có muốn 'thêm mới' không =====
    let willUpdate = 0;
    let willAdd = 0;
    uniqOps.forEach(op=>{
        if (_seen.has(op.key)) willUpdate++;
        else willAdd++;
    });

    let mode = "1"; // 1=upsert, 2=update-only, 3=replace
    if (willAdd > 0 || unknownMonCount > 0){
        const msg =
`⚠ File Tiết chuẩn có thể làm phát sinh dòng mới (và dễ gây dữ liệu phình nếu Môn chưa khớp danh mục).

Tổng dòng (duy nhất theo Khối+Môn): ${uniqOps.length}
- Sẽ cập nhật: ${willUpdate}
- Sẽ thêm mới: ${willAdd}
${unknownMonCount ? ("- Dòng có Môn không khớp danh mục Môn học: " + unknownMonCount + "\n") : ""}

Chọn cách nhập:
1 = Cập nhật + Thêm mới
2 = Chỉ cập nhật (bỏ qua thêm mới)  ✅ (an toàn chống đúp)
3 = Thay thế toàn bộ (xóa Tiết chuẩn hiện tại rồi nhập lại)
0 = Huỷ`;

        const def = (willAdd > 0) ? "2" : "1";
        const ans = prompt(msg, def);
        if (ans === null) return;
        const v = String(ans).trim();
        if (v === "0") return;
        if (v === "1" || v === "2" || v === "3") mode = v;
        else mode = def;
    }

    // Replace: xoá sạch tiết chuẩn trước
    if (mode === "3"){
        DATA.mon = [];
        _seen.clear();
        changed = true;
    }

    // ===== 4) Apply =====
    let added = 0;
    let updated = 0;
    let skippedNew = 0;

    uniqOps.forEach(op=>{
        const idx = _seen.get(op.key);

        if (idx !== undefined){
            const obj = DATA.mon[idx];
            obj.khoi = op.khoi;
            obj.ten = op.ten;
            obj.sotiet = op.sotiet;
            obj.gioihan = op.gioihan || "1";
            obj.ghichu = op.ghichu;
            updated++;
        }else{
            if (mode === "2"){
                skippedNew++;
                return;
            }
            const obj = {
                id: autoID("mon"),
                khoi: op.khoi,
                ten: op.ten,
                sotiet: op.sotiet,
                gioihan: op.gioihan || "1",
                ghichu: op.ghichu
            };
            DATA.mon.push(obj);
            _seen.set(op.key, DATA.mon.length - 1);
            added++;
        }
    });

    // Đồng bộ thêm lần nữa để dọn các case "mã vs tên" vừa import (nếu có)
    try{ if(typeof ensureTietChuanSyncedFromMonHoc === "function") ensureTietChuanSyncedFromMonHoc(); }catch(_){}

    saveStore();
    renderSectionInto("mon","section-content",document);

    let tail = "";
    if (skippedNew) tail += `, Bỏ qua thêm mới: ${skippedNew}`;
    const warn = (unknownMonCount ? `\n⚠ Có ${unknownMonCount} dòng môn không khớp danh mục Môn học (kiểm tra lại mã/tên môn trong file hoặc bảng Môn học).` : "");

    alert(`✔ Import tiết chuẩn thành công! (Cập nhật: ${updated}, Thêm mới: ${added}${tail})${warn}`);
}



/* ============================================================
   IMPORT MÔN HỌC TỔNG HỢP (mon.xlsx)
   Cột: Tên môn học | Mã môn học | Mã môn học 2 | Có giáo án | Ghi chú
   Có giáo án: trống hoặc bất kỳ ký tự (x, v, 1, yes...) -> lưu "V"
============================================================ */
function importExcel_MonHoc(rows){
    rows.forEach(r=>{
        let obj={};

        obj.ten = (r["Tên môn học"] ?? r["Tên môn"] ?? r["Môn học"] ?? "").toString().trim();
        obj.ma  = (r["Mã môn học"] ?? r["Mã môn"] ?? "").toString().trim();
        obj.ma2 = (r["Mã môn học 2"] ?? r["Mã môn 2"] ?? "").toString().trim();

        const ga = (r["Có giáo án"] ?? "").toString().trim().toUpperCase();
        obj.giaoan = (ga && ga !== "0" && ga !== "NO" && ga !== "FALSE") ? "V" : "";

        obj.ghichu = (r["Ghi chú"] ?? "").toString().trim();

        obj.id = autoID("monhoc");
        DATA.monhoc.push(obj);
    });

    saveStore();
    renderSectionInto("monhoc","section-content",document);
    alert("✔ Import môn học (tổng hợp) thành công!");
}



/* ============================================================
   IMPORT MẶC ĐỊNH (khi không phải Lớp/Môn)
============================================================ */
function importExcel_Default(rows){
    const cfg=FORM_CONFIG[IMPORT_SECTION];

    rows.forEach(r=>{
        let obj={};

        cfg.fields.forEach(f=>{
            obj[f.k] = r[f.label] ?? r[f.k] ?? "";
        });

        obj.id = autoID(IMPORT_SECTION);
        DATA[IMPORT_SECTION].push(obj);
    });

    saveStore();
    renderSectionInto(IMPORT_SECTION,"section-content",document);
    alert("✔ Import thành công!");
}
/* ============================================================
   ==========  PART 2 / 4 – EXPORT + RENDER + MODAL ==========
============================================================ */


/* ============================================================
   EXPORT EXCEL — GIỮ NGUYÊN HEADER THEO EXCEL
============================================================ */
function exportExcel(section){
    const cfg = FORM_CONFIG[section];
    const rows = DATA[section].map((r,i)=>{
        let obj = { STT: i+1 };

        cfg.fields.forEach(f=>{
            let v = r[f.k] || "";
            // Tiết chuẩn: mặc định giới hạn = 1 để người dùng xuất Excel và chỉnh theo ý
            if (section === "mon" && f.k === "gioihan"){
                const s = (v ?? "").toString().trim();
                v = s === "" ? "1" : s;
            }
            obj[f.label] = v;
        });

        return obj;
    });

    const ws = XLSX.utils.json_to_sheet(rows);

    // set độ rộng cột cho Excel (tránh tình trạng các cột bằng nhau)
    try{
        const cols = (rows && rows.length) ? Object.keys(rows[0]) : ["STT", ...(cfg.fields||[]).map(f=>f.label)];
        const widths = cols.map((k, idx)=>{
            let maxLen = String(k||"").length;
            (rows||[]).forEach(r=>{
                const v = r?.[k];
                const s = (v==null) ? "" : String(v);
                if (s.length > maxLen) maxLen = s.length;
            });
            // STT thường ngắn
            if (idx === 0) maxLen = Math.max(4, Math.min(maxLen, 6));
            return { wch: Math.min(Math.max(8, maxLen + 2), 40) };
        });
        ws["!cols"] = widths;
    }catch(e){
        console.warn("exportExcel set column widths failed", e);
    }
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, section);

    XLSX.writeFile(wb, `${section}.xlsx`);
}


/* ============================================================
   HIỂN THỊ DANH SÁCH DỮ LIỆU TRONG TRANG TỔNG HỢP
============================================================ */
function renderSectionInto(section, containerId, doc=document){
    const cfg = FORM_CONFIG[section];
    const arr = DATA[section];
    const container = doc.getElementById(containerId);

    // selection count for UX
    const selSet = TABLE_SELECTION[section] || new Set();
    const selCount = selSet.size;
    const selIndex = (selCount === 1) ? Array.from(selSet.values())[0] : -1;

    const isEditing = (INLINE_EDIT.section === section && Number.isFinite(INLINE_EDIT.index) && INLINE_EDIT.index >= 0);
    const editingIndex = isEditing ? INLINE_EDIT.index : -1;

    let html = `
    <div class="action-bar" style="flex-wrap:wrap;align-items:center">
        <button class="btn primary" onclick="openModal('${section}')">Thêm mới</button>
        <button class="btn" onclick="triggerExcel('${section}')">Nhập Excel</button>
        <button class="btn" onclick="exportExcel('${section}')">Xuất Excel</button>

        ${isEditing
            ? `<button class="btn" onclick="tableCancelEdit()">Hủy</button>`
            : ``
        }

        <button class="btn danger" style="background:#c9302c;color:#fff" onclick="deleteSelectedRows('${section}')">
            🗑 Xóa đã chọn${selCount ? ` (${selCount})` : ""}
        </button>

        <!-- Xóa riêng mục -->
        <button class="btn danger"
                style="background:#c9302c;color:#fff"
                onclick="deleteSection('${section}')">
            🗑 Xóa mục này
        </button>

        <div style="margin-left:auto;color:#556;font-size:12px;white-space:nowrap">
            Mẹo: Click chọn/bỏ chọn • Ctrl/Shift chọn nhiều • Kích đúp để sửa • Enter để lưu
        </div>
    </div>

    <div class="table-wrap">
    <table>
        <tr>
            <th>STT</th>
            <th>ID</th>`;

    // Thêm các column header
    cfg.fields.forEach(f=>{
        // CÓ GIÁO ÁN: canh giữa header cho cột checkbox
        if (section === "monhoc" && f.k === "giaoan"){
            html += `<th style="text-align:center;width:120px">${f.label}</th>`;
            return;
        }
        html += `<th>${f.label}</th>`;
    });

    // (Đã bỏ cột "Đồng bộ" theo yêu cầu)
    html += `</tr>`;

    // Thêm các dòng dữ liệu
    arr.forEach((row,i)=>{
        const isSelected = (TABLE_SELECTION[section] && TABLE_SELECTION[section].has(i));
        const isEditing = (INLINE_EDIT.section === section && INLINE_EDIT.index === i);

        html += `<tr class="${isSelected?"row-selected":""} ${isEditing?"row-editing":""}"
            onclick="tableRowClick(event,'${section}',${i})"
            ondblclick="tableRowDblClick(event,'${section}',${i})">
            <td>${i+1}</td>
            <td>${row.id}</td>`;

        cfg.fields.forEach(f=>{
            if (isEditing){
                const val = row ? (row[f.k] || "") : "";

                // KHỐI => select
                if (f.k === "khoi"){
                    const opts = (DATA.khoi || []).map(k => (k.ten||"").trim()).filter(Boolean);
                    html += `<td>
                        <select class="inline-edit-select" id="edit_${section}_${i}_${f.k}" onkeydown="tableEditKeyDown(event,'${section}',${i})">
                            <option value="">-- Chọn khối --</option>
                            ${opts.map(k=>`<option value="${escapeHtml(k)}" ${k===val?"selected":""}>${escapeHtml(k)}</option>`).join("")}
                        </select>
                    </td>`;
                    return;
                }

                // BUỔI => select
                if (f.k === "buoi"){
                    const opts = ["Buổi sáng","Buổi chiều","Cả ngày"];
                    html += `<td>
                        <select class="inline-edit-select" id="edit_${section}_${i}_${f.k}" onkeydown="tableEditKeyDown(event,'${section}',${i})">
                            <option value="">-- Chọn buổi --</option>
                            ${opts.map(o=>`<option value="${escapeHtml(o)}" ${o===val?"selected":""}>${escapeHtml(o)}</option>`).join("")}
                        </select>
                    </td>`;
                    return;
                }

                // CÓ GIÁO ÁN => checkbox (lưu V/trống)
                if (f.k === "giaoan"){
                    const checked = (val || "").toString().trim().toUpperCase() === "V";
                    html += `<td style="text-align:center">
                        <input type="checkbox" id="edit_${section}_${i}_${f.k}" ${checked?"checked":""} onkeydown="tableEditKeyDown(event,'${section}',${i})">
                    </td>`;
                    return;
                }

                // default => input
                html += `<td>
                    <input class="inline-edit-input" id="edit_${section}_${i}_${f.k}" value="${escapeHtml(val)}" onkeydown="tableEditKeyDown(event,'${section}',${i})">
                </td>`;
                return;
            }

            // ===== VIEW MODE =====
            // Môn học (tổng hợp): cột "Có giáo án" hiển thị checkbox (trống/V)
            if (section === "monhoc" && f.k === "giaoan") {
                const checked = (row.giaoan || "").toString().trim().toUpperCase() === "V";
                html += `<td style="text-align:center">
                    <input type="checkbox" ${checked ? "checked" : ""} onchange="setMonhocGiaoan(${i}, this.checked)">
                </td>`;
            } else {
                html += `<td>${row[f.k] || ""}</td>`;
            }
        });

        // (Đã bỏ cột "Đồng bộ" theo yêu cầu)

        html += `
        </tr>`;
    });

    html += `
        </table>
    </div>`;

    
    container.innerHTML = html;
}

// Toggle checkbox "Có giáo án" ngay trên bảng Môn học (môn tổng hợp)
function setMonhocGiaoan(index, checked){
    if (!Array.isArray(DATA.monhoc) || index < 0 || index >= DATA.monhoc.length) return;
    DATA.monhoc[index].giaoan = checked ? "V" : "";
    saveStore();
    // Không render lại toàn bảng khi tick để tránh giật/scroll; chỉ lưu.
}



/* ============================================================
   MODAL (THÊM / SỬA)
============================================================ */
let CURRENT_SECTION = "";
let EDIT_INDEX = -1;

function openModal(section,index=-1){
    CURRENT_SECTION = section;
    EDIT_INDEX = index;

    const cfg = FORM_CONFIG[section];
    const row = index >= 0 ? DATA[section][index] : null;

    document.getElementById("modal-title").innerText =
        (index === -1 ? "Thêm " : "Sửa ") + cfg.label;

    // Build form
    let html = "";

    cfg.fields.forEach(f=>{
        const val = row ? (row[f.k] || "") : "";

        // ===== KHỐI =====
if (f.k === "khoi") {
    html += `
    <div class="form-row">
        <label>${f.label}</label>
        <select id="f_${f.k}">
            <option value="">-- Chọn khối --</option>
            ${DATA.khoi.map(k =>
                `<option value="${k.ten}" ${k.ten===val?"selected":""}>${k.ten}</option>`
            ).join("")}
        </select>
    </div>`;
    return;
}

// ===== BUỔI =====
if (f.k === "buoi") {
    const opts = ["Buổi sáng","Buổi chiều","Cả ngày"];
    html += `
    <div class="form-row">
        <label>${f.label}</label>
        <select id="f_${f.k}">
            <option value="">-- Chọn buổi --</option>
            ${opts.map(o =>
                `<option value="${o}" ${o===val?"selected":""}>${o}</option>`
            ).join("")}
        </select>
    </div>`;
    return;
}


// ===== CÓ GIÁO ÁN (checkbox) =====
if (f.k === "giaoan") {
    const checked = (val || "").toString().trim().toUpperCase() === "V";
    html += `
    <div class="form-row">
        <label>${f.label}</label>
        <input type="checkbox" id="f_${f.k}" ${checked ? "checked" : ""}>
    </div>`;
    return;
}

// ===== CÁC FIELD KHÁC =====
html += `
<div class="form-row">
    <label>${f.label}</label>
    <input id="f_${f.k}" value="${val}">
</div>`;

    });

    document.getElementById("modal-body").innerHTML = html;
    document.getElementById("modal").style.display = "flex";
}

function closeModal(){
    document.getElementById("modal").style.display = "none";
}


/* ============================================================
   SAVE DATA — CHUẨN HOÁ THEO EXCEL
============================================================ */
function saveData(){
    const section = CURRENT_SECTION;
    const cfg = FORM_CONFIG[section];

    let obj = {};

    cfg.fields.forEach(f=>{
        if (f.k === "giaoan") {
            obj[f.k] = document.getElementById("f_"+f.k).checked ? "V" : "";
        } else {
            obj[f.k] = document.getElementById("f_"+f.k).value.trim();
        }
    });

    // Chuẩn hoá tên lớp nếu là LỚP
    if (section === "lop"){
        const canon = normalizeClassName(obj.ten);
        obj.ten2 = canon; // luôn tự động chuẩn hoá tên lớp 2
    }

    if (EDIT_INDEX === -1){
        obj.id = autoID(section);
        DATA[section].push(obj);
    } else {
        obj.id = DATA[section][EDIT_INDEX].id;
        DATA[section][EDIT_INDEX] = obj;
    }

    saveStore();
    closeModal();
    renderSectionInto(section,"section-content",document);
}


/* ============================================================
   XÓA DỮ LIỆU
============================================================ */
function deleteRow(section,index){
    if (!confirm("Bạn có chắc muốn xóa?")) return;

    const removed = (DATA[section] || [])[index];
    DATA[section].splice(index,1);
    // Nếu xóa môn tổng hợp (monhoc) => hỏi có xóa luôn trong Tiết chuẩn (mon) không
    if (section === "monhoc" && removed){
        const ten = _normText(removed.ten);
        const ma  = _normText(removed.ma);
        const ma2 = _normText(removed.ma2);
        const candidates = Array.from(new Set([ma, ten, ma2].filter(Boolean)));

        if (candidates.length){
            const msg =
                "Bạn có muốn xóa môn này khỏi TIẾT CHUẨN (tất cả khối) không?\n" +
                "- Nếu chọn OK: sẽ xóa các dòng tiết chuẩn có Môn = " + candidates.join(" / ") + "\n" +
                "- Nếu chọn Cancel: chỉ xóa trong MÔN HỌC (môn tổng hợp).";
            if (confirm(msg)){
                const before = (DATA.mon || []).length;
                DATA.mon = (DATA.mon || []).filter(m => !candidates.includes(_normText(m.ten)));
                const after = (DATA.mon || []).length;

                // Không tự xóa PCCM để tránh mất phân công ngoài ý muốn
                if (before !== after){
                    // Nếu đang ở trang Tiết chuẩn (nếu có) thì nó sẽ tự render lại qua logic phía dưới
                }
            }
        }
    }


    // Nếu xóa tiết chuẩn (mon) => cũng dọn PCCM liên quan theo khối+môn
    if (section === "mon" && removed){
        const khoi = _normText(removed.khoi);
        const ten  = _normText(removed.ten);
        if (khoi && ten){
            const khoiNum = extractKhoiNumber(khoi);
            const shouldRemoveKey = (key)=>{
                if (!key.includes("|")) return false;
                const [lopCanon, monTen] = key.split("|");
                if (_normText(monTen) !== ten) return false;
                return extractKhoiNumber(lopCanon) === khoiNum;
            };
            for (const key in (DATA.pccmMatrix || {})){
                if (shouldRemoveKey(key)) delete DATA.pccmMatrix[key];
            }
            for (const key in (DATA.pccmRoomMatrix || {})){
                if (shouldRemoveKey(key)) delete DATA.pccmRoomMatrix[key];
            }
        }
    }

    // Đồng bộ liên kết sau khi xoá (tránh tình trạng xoá Môn/Lớp/GV nhưng phân công vẫn còn)
    try{ syncDerivedDataIntegrity(); }catch(e){ console.warn("syncDerivedDataIntegrity failed", e); }

    saveStore();

    // Nếu đang ở PCCM thì refresh PCCM; ngược lại render section bình thường
    const sc = document.getElementById("section-content");
    if (sc && typeof renderPCCM === "function" && sc.innerHTML && sc.innerHTML.includes("PCCM")) {
        sc.innerHTML = renderPCCM();
    } else {
        renderSectionInto(section,"section-content",document);
    }
}

/* ============================================================
   EXCEL-LIKE UX HELPERS (ROW SELECT + INLINE EDIT + BULK DELETE)
============================================================ */
function _getSelSet(section){
    if (!TABLE_SELECTION[section]) TABLE_SELECTION[section] = new Set();
    return TABLE_SELECTION[section];
}

function tableRowClick(evt, section, index){
    // ignore clicks on interactive elements
    const t = evt?.target;
    if (t && t.closest && t.closest("button, input, select, textarea, a")) return;

    // Tránh đụng với dblclick (trình duyệt sẽ bắn click 2 lần trước khi dblclick)
    if (evt && typeof evt.detail === "number" && evt.detail > 1) return;

    const set = _getSelSet(section);
    const last = TABLE_LAST_INDEX[section];
    const isCtrl = !!(evt && (evt.ctrlKey || evt.metaKey));
    const isShift = !!(evt && evt.shiftKey);

    if (isShift && typeof last === "number" && last >= 0){
        const a = Math.min(last, index);
        const b = Math.max(last, index);
        if (!isCtrl) set.clear();
        for (let i=a;i<=b;i++) set.add(i);
    } else if (isCtrl){
        if (set.has(index)) set.delete(index);
        else set.add(index);
    } else {
        // Click 1 lần: chọn; click lại cùng dòng: bỏ chọn
        if (set.size === 1 && set.has(index)) {
            set.clear();
        } else {
            set.clear();
            set.add(index);
        }
    }

    TABLE_LAST_INDEX[section] = index;
    // selection only; don't keep inline edit when user selects
    if (INLINE_EDIT.section === section && INLINE_EDIT.index !== index){
        INLINE_EDIT = { section:"", index:-1 };
    }

    renderSectionInto(section, "section-content", document);
}

function tableRowDblClick(evt, section, index){
    const t = evt?.target;
    if (t && t.closest && t.closest("button, input, select, textarea, a")) return;
    tableBeginEdit(section, index);
}

function tableBeginEdit(section, index){
    INLINE_EDIT = { section, index };
    // chọn đúng dòng đang sửa
    const set = _getSelSet(section);
    set.clear();
    set.add(index);
    TABLE_LAST_INDEX[section] = index;
    renderSectionInto(section, "section-content", document);
}

function tableCancelEdit(){
    const sec = INLINE_EDIT.section;
    INLINE_EDIT = { section:"", index:-1 };
    if (sec) renderSectionInto(sec, "section-content", document);
}

// Khi đang sửa inline: Enter = Lưu, Esc = Hủy
function tableEditKeyDown(evt, section, index){
    if (!evt) return;
    const key = evt.key || evt.code || "";
    if (key === "Enter"){
        evt.preventDefault();
        tableSaveEdit(section, index);
    }
    if (key === "Escape"){
        evt.preventDefault();
        tableCancelEdit();
    }
}

function tableSaveEdit(section, index){
    const cfg = FORM_CONFIG[section];
    const row = (DATA[section] || [])[index];
    if (!cfg || !row) return;

    const obj = {};
    cfg.fields.forEach(f=>{
        const id = `edit_${section}_${index}_${f.k}`;
        if (f.k === "giaoan"){
            const el = document.getElementById(id);
            obj[f.k] = (el && el.checked) ? "V" : "";
        } else {
            const el = document.getElementById(id);
            obj[f.k] = (el ? el.value : "").toString().trim();
        }
    });

    // Chuẩn hoá tên lớp nếu là LỚP
    if (section === "lop"){
        const canon = normalizeClassName(obj.ten);
        obj.ten2 = canon;
    }

    obj.id = row.id;
    DATA[section][index] = { ...row, ...obj, id: row.id };

    saveStore();
    INLINE_EDIT = { section:"", index:-1 };
    renderSectionInto(section, "section-content", document);
}

function deleteSelectedRows(section){
    const set = _getSelSet(section);
    const idxs = Array.from(set.values()).filter(i=>Number.isFinite(i)).sort((a,b)=>b-a);
    if (!idxs.length){
        alert("Chưa chọn dòng nào.");
        return;
    }

    if (!confirm(`Xóa ${idxs.length} dòng đã chọn?`)) return;

    const removed = [];
    idxs.forEach(i=>{
        if (i >= 0 && i < (DATA[section]||[]).length){
            removed.push(DATA[section][i]);
            DATA[section].splice(i, 1);
        }
    });

    // Special cleanup: monhoc => optional delete in Tiết chuẩn
    if (section === "monhoc" && removed.length){
        const candSet = new Set();
        removed.forEach(r=>{
            const ten = _normText(r?.ten);
            const ma  = _normText(r?.ma);
            const ma2 = _normText(r?.ma2);
            [ten, ma, ma2].filter(Boolean).forEach(x=>candSet.add(x));
        });
        const candidates = Array.from(candSet);
        if (candidates.length){
            const msg =
                "Bạn có muốn xóa các môn này khỏi TIẾT CHUẨN (tất cả khối) không?\n" +
                "- OK: xóa các dòng tiết chuẩn có Môn = " + candidates.join(" / ") + "\n" +
                "- Cancel: chỉ xóa trong MÔN HỌC (môn tổng hợp).";
            if (confirm(msg)){
                DATA.mon = (DATA.mon || []).filter(m => !candidates.includes(_normText(m.ten)));
            }
        }
    }

    // Special cleanup: mon (tiết chuẩn) => dọn PCCM theo khối+môn
    if (section === "mon" && removed.length){
        removed.forEach(r=>{
            const khoi = _normText(r?.khoi);
            const ten  = _normText(r?.ten);
            if (!khoi || !ten) return;
            const khoiNum = extractKhoiNumber(khoi);
            const shouldRemoveKey = (key)=>{
                if (!key.includes("|")) return false;
                const [lopCanon, monTen] = key.split("|");
                if (_normText(monTen) !== ten) return false;
                return extractKhoiNumber(lopCanon) === khoiNum;
            };
            for (const key in (DATA.pccmMatrix || {})){
                if (shouldRemoveKey(key)) delete DATA.pccmMatrix[key];
            }
            for (const key in (DATA.pccmRoomMatrix || {})){
                if (shouldRemoveKey(key)) delete DATA.pccmRoomMatrix[key];
            }
        });
    }

    saveStore();
    set.clear();
    TABLE_LAST_INDEX[section] = -1;
    INLINE_EDIT = { section:"", index:-1 };
    renderSectionInto(section, "section-content", document);
}
/* ============================================================
   ==========  PART 3 / 4 — PHÂN CÔNG CHUYÊN MÔN (PCCM)  ==========
   PCCM key = "LớpChuẩn|TênMôn"
   Ví dụ: "6A1|Toán"
============================================================ */
/* ============================================================
   ==========  PART 3 / 4 — PHÂN CÔNG CHUYÊN MÔN (PCCM)  ==========
   PCCM key = "LớpChuẩn|TênMôn"
   Ví dụ: "6A1|Toán"
============================================================ */

/* ============================================================
   RENDER PCCM TRONG TRANG TỔNG HỢP
============================================================ */
/* ============================================================
   ==========  PCCM HOÀN CHỈNH — AUTO FIX LỚP & MÔN  ==========
============================================================ */

/* KHÓA PCCM: "6A1|Toán" */

/* ============================================================
   RENDER PCCM TRÊN GIAO DIỆN
============================================================ */

/* ============================================================
   PCCM UI STATE (TAB + SELECTED)
============================================================ */
let PCCM_TAB = "lop";         // "lop" | "giaovien" | "monhoc" | "phong"
let PCCM_SELECTED_CLASS = ""; // vd "10A1"
let PCCM_SELECTED_GV = "";    // tên GV
let PCCM_SELECTED_ROOM = "";  // tên phòng (tab Phòng)

// Cache chỉnh sửa PCCM (tab Lớp) để nút Lưu có thể đọc DOM
let PCCM_CLASS_EDIT_CACHE = { cls:"", khoiName:"", mons:[] };

// Cache chỉnh sửa PCCM (tab Giáo viên)
let PCCM_TEACHER_EDIT_CACHE = { gv:"", rows:[] };
let PCCM_TEACHER_CLASS_FILTER = "Tất cả";

// Cache chỉnh sửa PCCM (tab Môn học)
let PCCM_SUBJECT_EDIT_CACHE = { monKey:"", rows:[] };

// Cache chỉnh sửa PCCM (tab Phòng)
let PCCM_ROOM_EDIT_CACHE = { room:"", rows:[] };

// Bộ lọc khối trong PCCM
let PCCM_KHOI = "Tất cả";     // "Tất cả" hoặc "Khối 6", "Khối 7", ...
let PCCM_ALLOWED_CLASS_SET = null;

// Bộ lọc môn trong tab Tiết chuẩn
let TC_MON = "Tất cả";

// Bộ lọc cho trang Tiết chuẩn (menu lớn)
let TC_KHOI = "Tất cả";

// Tiết chuẩn: chọn nhiều ô (Ctrl/Shift) + Copy/Paste
let TC_CELL_SELECTION = new Set(); // key "r,c"
let TC_CELL_ANCHOR = null;         // {r,c}


function renderPCCM() {
    // ===== Lớp (hiển thị tất cả, không lọc theo khối) =====
    const lopObjs = (DATA.lop || []).map(l=>{
        // chỉ lấy Tên lớp 2 (ten2), và CHUẨN HOÁ để không bị 6/1 & 6A1
        const ten2 = canonTen2FromLop(l);
        const khoi = (l.khoi || "").trim() || ("Khối " + extractKhoiNumber(ten2));
        return {ten2, khoi};
    }).filter(x=>x.ten2);

    const classNames = Array.from(new Set(lopObjs.map(x=>x.ten2))).sort((a,b)=>a.localeCompare(b,'vi'));

    if (!PCCM_SELECTED_CLASS && classNames.length) PCCM_SELECTED_CLASS = classNames[0];
    if (PCCM_SELECTED_CLASS && !classNames.includes(PCCM_SELECTED_CLASS) && classNames.length) {
        PCCM_SELECTED_CLASS = classNames[0];
    }

    // ===== Tiết chuẩn: CHỈ dùng để lookup số tiết/tuần & giới hạn theo khối =====
    // (Danh mục môn trong PCCM lấy từ Môn học tổng hợp: DATA.monhoc)
    const tcRowsAll = (DATA.mon || []).map(m=>({
        khoi: (m.khoi || "").trim(),
        ten: (m.ten || "").trim(),
        sotiet: (m.sotiet || "").toString().trim(),
        gioihan: (m.gioihan || "").toString().trim()
    })).filter(r=>r.khoi && r.ten);

    const pccmMonsAll = buildPCCMMonList(); // [{key, ten, ma, ma2}]

    // Danh mục môn (từ Môn học tổng hợp) nhưng chỉ HIỂN THỊ những môn có TIẾT CHUẨN (số tiết > 0)
    // ở ít nhất 1 khối. (Môn không có tiết sẽ không hiển thị trong Phân công.)
    const pccmMonsBase = pccmMonsAll.filter(m => monHasPositiveTietChuanAnyKhoi(m));

    // ===== Giáo viên: ưu tiên lấy từ bảng Giáo viên (mã GV), fallback từ PCCM nếu bảng GV trống =====
    let gvs = (DATA.giaovien || [])
        .map(g=>_normText(g?.magv))
        .filter(Boolean)
        .sort((a,b)=>a.localeCompare(b,'vi'));
    if (!gvs.length){
        const teachersSet = new Set();
        Object.values(DATA.pccmMatrix || {}).forEach(v=>{
            const s = (v||"").toString().trim();
            if (s) teachersSet.add(s);
        });
        gvs = Array.from(teachersSet).sort((a,b)=>a.localeCompare(b,'vi'));
    }
    if (!PCCM_SELECTED_GV && gvs.length) PCCM_SELECTED_GV = gvs[0];
    if (PCCM_SELECTED_GV && !gvs.includes(PCCM_SELECTED_GV) && gvs.length) PCCM_SELECTED_GV = gvs[0];

    const hasRooms = Array.isArray(DATA.phong) && DATA.phong.some(p=>_normText(p?.ten));

    // Tổng số tiết (tab Lớp): hiển thị cùng dòng nhập/xuất Excel
    let pccmTotalInfoHtml = "";
    if (PCCM_TAB === "lop" && PCCM_SELECTED_CLASS){
        const totals = pccmComputeTotalTietForClass(PCCM_SELECTED_CLASS, pccmMonsBase);
        if (totals){
            pccmTotalInfoHtml = `
                <span style="padding:6px 10px;background:#f0f5ff;border:1px solid #adc6ff;border-radius:999px;font-weight:700;color:#2f54eb">
                    Tổng tiết: ${totals.assigned}
                </span>`;
        }
    }

    // ===== UI: tabs =====
    const tabBtn = (key, text) => `
        <button class="btn ${PCCM_TAB===key?"primary":""}"
                onclick="setPCCMTab('${key}')">${text}</button>`;

    let html = `
    <div class="action-bar" style="gap:8px;flex-wrap:wrap;align-items:center">
        ${tabBtn("lop","Lớp")}
        ${tabBtn("giaovien","Giáo viên")}
        ${tabBtn("monhoc","Môn học")}
        ${tabBtn("phong","Phòng")}

        <div style="margin-left:auto;display:flex;gap:8px;flex-wrap:wrap;align-items:center">
            <button class="btn" onclick="triggerPCCMImport()">Nhập Excel</button>
            <button class="btn" onclick="exportPCCMExcel()">Xuất Excel</button>
            ${pccmTotalInfoHtml}
            <button class="btn primary" onclick="pccmSaveActive()">Lưu</button>
            <button class="btn danger" style="background:#c9302c;color:#fff" onclick="deleteAllPCCM()">🗑 Xóa Phân công</button>
        </div>
    </div>
    `;

    // ===== Content routing =====

    if (!classNames.length) {
        html += `
        <div style="padding:14px;background:#fff;border-radius:8px;border:1px solid #e3e8f3;margin-top:10px">
            ⚠ Chưa có <b>Lớp</b>.
        </div>`;
        return html;
    }

    
    if (!pccmMonsAll.length) {
        html += `
        <div style="padding:14px;background:#fff;border-radius:8px;border:1px solid #e3e8f3;margin-top:10px">
            ⚠ Chưa có <b>Môn học</b>. Vào menu <b>Môn học</b> để nhập file <b>mon.xlsx</b> trước.
        </div>`;
        return html;
    }

    if (!pccmMonsBase.length) {
        html += `
        <div style="padding:14px;background:#fff;border-radius:8px;border:1px solid #e3e8f3;margin-top:10px">
            ⚠ Chưa có <b>Tiết chuẩn</b> (Số tiết &gt; 0). Vào menu <b>Tiết chuẩn</b> để nhập/thiết lập số tiết trước, sau đó quay lại <b>Phân công</b>.
        </div>`;
        return html;
    }


    if (PCCM_TAB === "giaovien") {
        html += renderPCCM_ByTeacher(gvs, pccmMonsBase, classNames);
    } else if (PCCM_TAB === "monhoc") {
        html += renderPCCM_BySubject(classNames, pccmMonsBase);
    } else if (PCCM_TAB === "phong") {
        html += renderPCCM_ByRoom(classNames, pccmMonsBase);
    } else {
        // mặc định = tab Lớp
        html += renderPCCM_ByClass(classNames, pccmMonsBase);
    }

    return html;
}



/* ==============================
   TRANG TIẾT CHUẨN (MENU LỚN)
   ============================== */

function setTCKhoi(khoi){
    TC_KHOI = khoi || "Tất cả";
    const sc = document.getElementById("section-content");
    if (sc) sc.innerHTML = renderTietChuanPage();
}

// Giữ tên hàm cũ (được gọi từ HTML) nhưng thực chất set cho trang Tiết chuẩn
function setPCCMTCMon(mon){
    TC_MON = mon || "Tất cả";
    const sc = document.getElementById("section-content");
    if (sc) sc.innerHTML = renderTietChuanPage();
}

function renderTietChuanPage(){
    // Mỗi lần render lại bảng -> reset chọn nhiều ô để tránh lệch index
    try{
        TC_CELL_SELECTION = new Set();
        TC_CELL_ANCHOR = null;
    }catch(e){ /* ignore */ }

    // Đồng bộ Tiết chuẩn theo danh sách Môn học (nếu có dữ liệu)
    try{ ensureTietChuanSyncedFromMonHoc(); }catch(e){ /* ignore */ }

    // Nếu Môn học trống => không hiển thị dữ liệu Tiết chuẩn
    const monhocSet = new Set(
        (DATA.monhoc || [])
            .flatMap(r=>[r?.ten, r?.ma, r?.ma2, r?.id])
            .map(x=>_normText(x).toLowerCase())
            .filter(Boolean)
    );

    // lấy tất cả tiết chuẩn
    let monRowsAll = (DATA.mon || []).map(m=>({
        id: (m.id || "").toString().trim(),
        khoi: (m.khoi || "").trim(),
        ten: (m.ten || "").trim(),
        sotiet: (m.sotiet || "").toString().trim(),
        gioihan: (m.gioihan || "").toString().trim()
    })).filter(r=>r.khoi && r.ten);

    // lọc theo Môn học để tránh hiển thị dữ liệu cũ/khác
    if(monhocSet.size){
        monRowsAll = monRowsAll.filter(r=>monhocSet.has(_normText(r.ten).toLowerCase()));
    } else {
        monRowsAll = [];
    }

    // Danh sách khối ưu tiên lấy từ bảng Khối (nhanh hơn listbox)
    const khoiFromTable = (DATA.khoi || []).map(k => (k.ten || "").trim()).filter(Boolean);
    const khoiNames = Array.from(new Set(monRowsAll.map(r=>r.khoi))).filter(Boolean);

    const khoiUniq = Array.from(new Set((khoiFromTable.length ? khoiFromTable : khoiNames)));
    khoiUniq.sort((a,b)=>{
        const na = Number(extractKhoiNumber(a) || 0);
        const nb = Number(extractKhoiNumber(b) || 0);
        if (na !== nb) return na - nb;
        return a.localeCompare(b,'vi');
    });
    const khoiOptions = ["Tất cả", ...khoiUniq];

    if (!khoiOptions.includes(TC_KHOI)) TC_KHOI = "Tất cả";

    const rowsKhoi = (TC_KHOI === "Tất cả") ? monRowsAll : monRowsAll.filter(r => r.khoi === TC_KHOI);

    // Sort: ưu tiên theo Khối (từ bé -> lớn), sau đó theo thứ tự Môn trong Bảng Môn
    const __monOrder = new Map();
    (DATA.monhoc || []).forEach((r,i)=>{
        const keys = [r.id, r.ten, r.ma, r.ma2]
            .map(x => (x ?? '').toString().normalize('NFC').trim())
            .filter(Boolean);
        for (const k of keys){
            const low = k.toLowerCase();
            if (!__monOrder.has(low)) __monOrder.set(low, i);
        }
    });

    const filtered = rowsKhoi.slice().sort((a,b)=>{
        const na = Number(extractKhoiNumber(a.khoi) || 0);
        const nb = Number(extractKhoiNumber(b.khoi) || 0);
        if (na !== nb) return na - nb;
        const ka = (a.khoi||'').localeCompare((b.khoi||''),'vi');
        if (ka) return ka;

        const ia = __monOrder.get(((a.ten||'').toString().normalize('NFC').trim().toLowerCase())) ?? 1e9;
        const ib = __monOrder.get(((b.ten||'').toString().normalize('NFC').trim().toLowerCase())) ?? 1e9;
        if (ia !== ib) return ia - ib;

        return (a.ten||'').localeCompare((b.ten||''),'vi');
    });

    // ===== action bar (đồng bộ với các bảng khác) =====
    let html = `
    <div class="action-bar" style="flex-wrap:wrap;align-items:center;gap:10px">
        <div style="display:flex;gap:6px;flex-wrap:wrap">
            ${khoiOptions.map(k=>{
                const js = (k||"").toString().replace(/\\/g,"\\\\").replace(/'/g,"\\'");
                return `<button class="btn ${TC_KHOI===k?"primary":""}" onclick="setTCKhoi('${js}')">${escapeHtml(k)}</button>`;
            }).join("")}
        </div>

        <span style="flex:1"></span>

        <button class="btn" onclick="triggerExcel('mon')">Nhập Excel</button>
        <button class="btn" onclick="exportExcel('mon')">Xuất Excel</button>
        <button class="btn primary" onclick="tcSaveAllEdits()">Lưu</button>

        <div style="margin-left:10px;color:#556;font-size:12px;white-space:nowrap">
            Mẹo: Click chọn ô • Ctrl/Shift chọn nhiều • Ctrl/Cmd+C/V copy/paste • Kích đúp để sửa
        </div>
    </div>

    <div class="table-wrap">
    <table>
        <tr>
            <th style="width:60px">TT</th>
            <th style="width:140px">Khối học</th>
            <th>Môn học</th>
            <th style="width:160px;text-align:center">Số tiết/1 tuần</th>
            <th style="width:210px;text-align:center">Giới hạn số tiết/1 buổi</th>
        </tr>`;

    filtered.forEach((r,i)=>{
        const rid = (r.id || "").toString().trim();
        html += `
        <tr>
            <td style="text-align:center">${i+1}</td>
            <td>${escapeHtml(r.khoi)}</td>
            <td>${escapeHtml(r.ten)}</td>
            <td class="tc-cell" style="text-align:center;cursor:cell" 
                data-rowid="${escapeHtml(rid)}" data-field="sotiet" data-val="${escapeHtml(r.sotiet)}"
                data-r="${i}" data-c="0"
                onclick="tcCellClick(event,this)" ondblclick="tcBeginCellEdit(this)">${escapeHtml(r.sotiet)}</td>
            <td class="tc-cell" style="text-align:center;cursor:cell" 
                data-rowid="${escapeHtml(rid)}" data-field="gioihan" data-val="${escapeHtml(r.gioihan)}"
                data-r="${i}" data-c="1"
                onclick="tcCellClick(event,this)" ondblclick="tcBeginCellEdit(this)">${escapeHtml(r.gioihan)}</td>
        </tr>`;
    });

    if (!filtered.length){
        html += `<tr><td colspan="5" style="padding:14px;color:#666">Chưa có dữ liệu tiết chuẩn.</td></tr>`;
    }

    html += `</table></div>`;
    return html;
}
function setPCCMKhoi(khoi){
    PCCM_KHOI = khoi || "Tất cả";
    PCCM_SELECTED_CLASS = "";
    PCCM_SELECTED_GV = "";
    const sc = document.getElementById("section-content");
    if (sc) {
        const html = (sc.innerHTML || "");
        if (html.includes("PCCM") && typeof renderPCCM === "function") sc.innerHTML = renderPCCM();
        else if (typeof renderTietChuanPage === "function") sc.innerHTML = renderTietChuanPage();
    }
}

function setPCCMTietChuanMon(mon){
    TC_MON = mon || "Tất cả";
    const sc = document.getElementById("section-content");
    if (sc && typeof renderTietChuanPage === "function") sc.innerHTML = renderTietChuanPage();
}


function setPCCMTab(tab){
    PCCM_TAB = tab;
    document.getElementById("section-content").innerHTML = renderPCCM();
}
function setPCCMSelectedClass(cls){
    PCCM_SELECTED_CLASS = cls;
    document.getElementById("section-content").innerHTML = renderPCCM();
}
function setPCCMSelectedGV(name){
    PCCM_SELECTED_GV = name;
    document.getElementById("section-content").innerHTML = renderPCCM();
}

// Lưu theo đúng tab PCCM đang mở (nút Lưu ở action bar)
function pccmSaveActive(){
    switch (PCCM_TAB){
        case "giaovien":
            return (typeof pccmSaveTeacherEdits === "function") ? pccmSaveTeacherEdits() : void 0;
        case "monhoc":
            return (typeof pccmSaveSubjectEdits === "function") ? pccmSaveSubjectEdits() : void 0;
        case "phong":
            return (typeof pccmSaveRoomEdits === "function") ? pccmSaveRoomEdits() : void 0;
        case "lop":
        default:
            return (typeof pccmSaveClassEdits === "function") ? pccmSaveClassEdits() : void 0;
    }
}

function setPCCMSelectedRoom(room){
    PCCM_SELECTED_ROOM = room;
    const sc = document.getElementById("section-content");
    if (sc) sc.innerHTML = renderPCCM();
}

function setPCCMTeacherClassFilter(cls){
    PCCM_TEACHER_CLASS_FILTER = cls || "Tất cả";
    const sc = document.getElementById("section-content");
    if (sc) sc.innerHTML = renderPCCM();
}

function setPCCMTeacher(lopCanon, monTen, val){
    const key = `${lopCanon}|${monTen}`;
    val = (val||"").trim();
    if (val) DATA.pccmMatrix[key] = val;
    else delete DATA.pccmMatrix[key];
    saveStore();
}

function setPCCMRoom(lopCanon, monTen, val){
    const key = `${lopCanon}|${monTen}`;
    val = (val||"").trim();
    if (val) DATA.pccmRoomMatrix[key] = val;
    else delete DATA.pccmRoomMatrix[key];
    saveStore();
}

// ===== PCCM: Copy/Paste nhanh Giáo viên trong listbox (Ctrl/Cmd + C / V) =====
let PCCM_CLIPBOARD_GV = "";

function pccmHandleGVCopyPaste(ev, selectEl){
    try{
        const isCmd = !!(ev && (ev.ctrlKey || ev.metaKey));
        if (!isCmd) return;
        const key = (ev.key || "").toLowerCase();
        if (key === "c"){
            PCCM_CLIPBOARD_GV = (selectEl?.value ?? "").toString();
            ev.preventDefault();
            return;
        }
        if (key === "v"){
            const v = (PCCM_CLIPBOARD_GV ?? "").toString();
            if (!selectEl) return;

            // nếu value chưa có trong options thì thêm option tạm để tránh mất dữ liệu
            const exists = Array.from(selectEl.options || []).some(o => (o.value||"") === v);
            if (v && !exists){
                const opt = document.createElement("option");
                opt.value = v;
                opt.textContent = `(Đang lưu) ${v}`;
                selectEl.insertBefore(opt, selectEl.firstChild);
            }
            selectEl.value = v;
            ev.preventDefault();
        }
    }catch(e){
        // ignore
    }
}


// ===== PCCM: Double-click để mở listbox (Lớp/Môn), 1-click để chọn text copy =====
function pccmSelectText(el){
    try{
        if (!el) return;
        const tag = (el.tagName || "").toUpperCase();
        if (tag === "INPUT" || tag === "TEXTAREA"){
            // lưu lại giá trị trước khi người dùng paste/sửa
            try{ el.setAttribute('data-prev', (el.value ?? '').toString()); }catch(_){ }
            el.focus();
            el.select();
            return;
        }

        const range = document.createRange();
        range.selectNodeContents(el);
        const sel = window.getSelection();
        if (!sel) return;
        sel.removeAllRanges();
        sel.addRange(range);
    }catch(e){
        // ignore
    }
}


// ===== PCCM: Input (1-click) để copy/paste + Enter/Blur để commit sang select =====
function pccmDblInputKeyDown(ev, inputEl){
    try{
        if (!ev || !inputEl) return;
        if (ev.key === 'Enter'){
            ev.preventDefault();
            inputEl.blur();
            return;
        }
        if (ev.key === 'Escape'){
            ev.preventDefault();
            const prev = inputEl.getAttribute('data-prev');
            if (prev !== null && prev !== undefined) inputEl.value = prev;
            inputEl.blur();
            return;
        }
    }catch(e){
        // ignore
    }
}

function pccmDblInputCommit(inputEl){
    try{
        const el = (typeof inputEl === 'string') ? document.getElementById(inputEl) : inputEl;
        if (!el) return;
        // reset "click-armed" state so lần click sau lại trở về hành vi chọn-copy trước
        try{ el.setAttribute('data-openarmed','0'); }catch(_){ }
        const selectId = el.getAttribute('data-selectid');
        if (!selectId) return;
        const sel = document.getElementById(selectId);
        if (!sel) return;

        const kind = (el.getAttribute('data-kind') || sel.getAttribute('data-kind') || '').toString();

        // Nếu là Môn và select có data-class-select => refresh list theo lớp
        if (kind === 'mon'){
            const clsSelId = sel.getAttribute('data-class-select');
            if (clsSelId){
                const clsSel = document.getElementById(clsSelId);
                const clsVal = clsSel ? (clsSel.value || '') : '';
                pccmFillMonOptions(sel, clsVal);
            }
        }

        const raw = (el.value || '').toString();
        const txt = raw.normalize('NFC').trim();
        if (!txt){
            // không cho rỗng => revert
            const prev = el.getAttribute('data-prev');
            if (prev !== null && prev !== undefined) el.value = prev;
            return;
        }

        const low = txt.toLowerCase();
        let matched = null;

        const opts = Array.from(sel.options || []);
        for (const opt of opts){
            const v = (opt.value || '').toString().normalize('NFC').trim().toLowerCase();
            const t = (opt.textContent || '').toString().normalize('NFC').trim().toLowerCase();
            if (low === v || low === t){
                matched = opt.value;
                break;
            }
        }

        // fallback: match contains (duy nhất)
        if (matched === null){
            const cands = opts.filter(opt => {
                const t = (opt.textContent || '').toString().normalize('NFC').trim().toLowerCase();
                return t.includes(low);
            });
            if (cands.length === 1) matched = cands[0].value;
        }

        if (matched !== null){
            sel.value = matched;
            pccmDblSelSyncText(selectId);

            // Nếu commit là Lớp => refresh lại list Môn của dòng
            if (kind === 'cls'){
                pccmRowClassChanged(sel);
                pccmDblSelSyncText(selectId);
            }
        }else{
            const prev = el.getAttribute('data-prev');
            if (prev !== null && prev !== undefined) el.value = prev;
        }
    }catch(e){
        // ignore
    }
}


function pccmDblSelSyncText(selectId){
    try{
        const sel = document.getElementById(selectId);
        if (!sel) return;

        const inputId = sel.getAttribute('data-inputid');
        const textId  = sel.getAttribute('data-textid');
        const node = inputId ? document.getElementById(inputId) : (textId ? document.getElementById(textId) : null);
        if (!node) return;

        const opt = (sel.options && sel.selectedIndex >= 0) ? sel.options[sel.selectedIndex] : null;
        const display = opt ? (opt.textContent || '') : (sel.value || '');

        const tag = (node.tagName || '').toUpperCase();
        if (tag === 'INPUT' || tag === 'TEXTAREA'){
            node.value = display;
        }else{
            node.textContent = display;
        }
        try{ node.setAttribute('data-value', sel.value || ''); }catch(_){ }
    }catch(e){
        // ignore
    }
}


function pccmFillMonOptions(selectEl, clsVal){
    try{
        if (!selectEl) return;
        const curVal = (selectEl.value || "").toString();

        // base list: chỉ các môn có tiết chuẩn ở ít nhất 1 khối
        const base = buildPCCMMonList().filter(m => monHasPositiveTietChuanAnyKhoi(m));
        const allowed = clsVal ? pccmGetAllowedMonsForClass(clsVal, base) : base;

        const allowedKeys = new Set(allowed.map(m => (m.key || m.ten || "").toString()));

        // nếu current đang là 1 giá trị lạ (do dữ liệu cũ), giữ lại option tạm
        const extraOpt = (curVal && !allowedKeys.has(curVal))
            ? `<option value="${escapeHtml(curVal)}" selected>(Đang lưu) ${escapeHtml(curVal)}</option>`
            : "";

        const opts = allowed.map(m=>{
            const k = (m.key || m.ten || "").toString();
            const label = (m.ten || m.code || m.ma || m.key || '').toString();
            return `<option value="${escapeHtml(k)}" ${k===curVal?"selected":""}>${escapeHtml(label)}</option>`;
        }).join("");

        selectEl.innerHTML = `${extraOpt}${opts}`;

        // nếu cur không còn hợp lệ => chọn môn đầu tiên
        if (curVal && allowedKeys.has(curVal)){
            selectEl.value = curVal;
        }else if (allowed[0]){
            selectEl.value = (allowed[0].key || allowed[0].ten || "").toString();
        }
    }catch(e){
        // ignore
    }
}

function pccmRowClassChanged(clsSelectEl){
    try{
        if (!clsSelectEl) return;
        const monSelId = clsSelectEl.getAttribute("data-monselect");
        if (!monSelId) return;
        const monSel = document.getElementById(monSelId);
        if (!monSel) return;
        pccmFillMonOptions(monSel, (clsSelectEl.value || "").toString());
        pccmDblSelSyncText(monSelId);
    }catch(e){
        // ignore
    }
}


// PCCM listbox input:
// - Click lần 1: chọn text để copy
// - Click lần 2 (vẫn ở đúng ô đang focus): xổ listbox
// (Ctrl/Cmd+Click: mở listbox ngay)
function pccmInputClick(ev, inputEl, selId){
    try{
        if(!inputEl) return;

        // (NEW) Click 1 lần: mở listbox ngay.
        // Giữ Ctrl/Cmd khi click: chỉ chọn text để copy (không mở listbox).
        if(ev && (ev.ctrlKey || ev.metaKey)){
            pccmSelectText(inputEl);
            try{ inputEl.setAttribute('data-openarmed','0'); }catch(_){ }
            return;
        }

        try{ inputEl.setAttribute('data-openarmed','0'); }catch(_){ }
        pccmDblSelOpen(selId);
    }catch(e){
        // ignore
    }
}

function pccmDblSelOpen(selectId){
    try{
        const sel = document.getElementById(selectId);
        if (!sel) return;

        // Nếu là list Môn thì refresh option theo Lớp hiện tại (chỉ khi có data-class-select)
        const kind = sel.getAttribute('data-kind') || '';
        if (kind === 'mon'){
            const clsSelId = sel.getAttribute('data-class-select');
            if (clsSelId){
                const clsSel = document.getElementById(clsSelId);
                const clsVal = clsSel ? (clsSel.value || '') : '';
                pccmFillMonOptions(sel, clsVal);
            }
        }

        const inputId = sel.getAttribute('data-inputid');
        const textId  = sel.getAttribute('data-textid');
        const node = inputId ? document.getElementById(inputId) : (textId ? document.getElementById(textId) : null);
        if (node) node.style.display = 'none';

        // Hiển thị dạng listbox (size>1) khi mở
        const size = Number(sel.getAttribute('data-size') || 8);
        if (Number.isFinite(size) && size > 1) sel.size = size;

        sel.style.display = '';
        sel.disabled = false;
        sel.focus();

        // cố gắng mở dropdown ngay khi dblclick (tùy trình duyệt)
        try{ sel.click(); }catch(e){}
    }catch(e){
        // ignore
    }
}


function pccmDblSelClose(selectId, commit){
    try{
        const sel = document.getElementById(selectId);
        if (!sel) return;
        if (commit) pccmDblSelSyncText(selectId);

        const inputId = sel.getAttribute('data-inputid');
        const textId  = sel.getAttribute('data-textid');
        const node = inputId ? document.getElementById(inputId) : (textId ? document.getElementById(textId) : null);

        sel.style.display = 'none';
        sel.disabled = true;
        if (node) node.style.display = '';
    }catch(e){
        // ignore
    }
}



/* ============================================================
   PASTE FROM EXCEL (Tab-separated) — TIẾT CHUẨN
   - Cho phép copy bảng từ Excel rồi dán vào textarea
   - Nếu thiếu cột Khối học, sẽ dùng khối đang chọn ở PCCM
============================================================ */
function _normText(x){
    return (x ?? "")
        .toString()
        .normalize("NFC")
        .replace(/\r?\n/g, " ")
        .replace(/\s+/g, " ")
        .trim();
}

// ============================================================
// TIẾT CHUẨN: đồng bộ từ Môn học
// - Nếu Môn học có dữ liệu: tạo thiếu (Khối x Môn) với mặc định
//   Số tiết/tuần = 1, Giới hạn = 1.
// - Nếu Môn học trống: KHÔNG tự tạo. (render sẽ không hiển thị)
// ============================================================
function ensureTietChuanSyncedFromMonHoc(){
    const monhoc = Array.isArray(DATA.monhoc) ? DATA.monhoc : [];
    if(!monhoc.length) return;

    // Danh sách khối
    let khoiList = (DATA.khoi || []).map(k=>_normText(k?.ten)).filter(Boolean);
    if(!khoiList.length){
        // fallback: lấy từ danh sách lớp
        const set = new Set();
        (DATA.lop || []).forEach(l=>{
            const raw = _normText(l?.khoi) || _normText(l?.ten2) || _normText(l?.ten);
            const kn = extractKhoiNumber(raw);
            if(kn) set.add(`Khối ${kn}`);
        });
        khoiList = [...set];
    }
    if(!khoiList.length) return;

    // Danh sách môn (theo tên hiển thị)
    const monNames = [];
    const seen = new Set();
    monhoc.forEach(m=>{
        const ten = _normText(m?.ten);
        if(!ten) return;
        const k = ten.toLowerCase();
        if(seen.has(k)) return;
        seen.add(k);
        monNames.push(ten);
    });
    if(!monNames.length) return;

    if(!Array.isArray(DATA.mon)) DATA.mon = [];
    let changed = false;

    // Canon khoi: luôn về dạng "Khối X" nếu lấy được số
    const canonKhoi = (khoiRaw)=>{
        const t = _normText(khoiRaw);
        const n = extractKhoiNumber(t);
        return n ? `Khối ${n}` : t;
    };

    // Tìm môn theo mọi alias (id/ten/ma/ma2) để map code -> tên hiển thị
    const findMonhoc = (monValue)=>{
        const v = _normText(monValue);
        if(!v) return null;
        const low = v.toLowerCase();
        return (DATA.monhoc || []).find(r=>{
            const id  = _normText(r.id).toLowerCase();
            const ten = _normText(r.ten).toLowerCase();
            const ma  = _normText(r.ma).toLowerCase();
            const ma2 = _normText(r.ma2).toLowerCase();
            return (id && id === low) || (ten && ten === low) || (ma && ma === low) || (ma2 && ma2 === low);
        }) || null;
    };
    const canonMon = (monRaw)=>{
        const found = findMonhoc(monRaw);
        if(!found) return _normText(monRaw);
        return _normText(found.ten) || _normText(monRaw);
    };

    // Map existing (khoiCanon|tenCanonLower) + dọn trùng (merge) để tránh đúp
    const map = new Map();
    const cleaned = [];

    for(const r of (DATA.mon || [])){
        if(!r || typeof r !== "object"){
            cleaned.push(r);
            continue;
        }
        const khoi0 = canonKhoi(r.khoi);
        const ten0  = canonMon(r.ten);

        if(!khoi0 || !ten0){
            cleaned.push(r);
            continue;
        }

        const key = `${khoi0}|${ten0.toLowerCase()}`;

        // normalize dữ liệu đang lưu (để lần sau không lệch)
        if(_normText(r.khoi) !== khoi0){ r.khoi = khoi0; changed = true; }
        if(_normText(r.ten)  !== ten0 ){ r.ten  = ten0;  changed = true; }

        const ex = map.get(key);
        if(!ex){
            map.set(key, r);
            cleaned.push(r);
        }else{
            // merge dữ liệu (ưu tiên giá trị đang có; nếu trống/không hợp lệ thì lấy của dòng trùng)
            const st1 = String(ex.sotiet ?? "").trim();
            const gh1 = String(ex.gioihan ?? "").trim();
            const st2 = String(r.sotiet ?? "").trim();
            const gh2 = String(r.gioihan ?? "").trim();

            if((st1 === "" || Number.isNaN(Number(st1))) && st2 !== "" && !Number.isNaN(Number(st2))){
                ex.sotiet = st2;
                changed = true;
            }
            if((gh1 === "" || Number.isNaN(Number(gh1))) && gh2 !== "" && !Number.isNaN(Number(gh2))){
                ex.gioihan = gh2;
                changed = true;
            }

            const note1 = String(ex.ghichu ?? "").trim();
            const note2 = String(r.ghichu ?? "").trim();
            if(!note1 && note2){
                ex.ghichu = note2;
                changed = true;
            }

            // bỏ dòng trùng khỏi DATA.mon
            changed = true;
        }
    }

    if(cleaned.length !== (DATA.mon || []).length){
        DATA.mon = cleaned;
    }

    // Ensure all combos exist
    for(const khoiRaw of khoiList){
        const khoi = canonKhoi(khoiRaw);
        if(!khoi) continue;

        for(const tenRaw of monNames){
            const ten = canonMon(tenRaw) || _normText(tenRaw);
            if(!ten) continue;

            const key = `${khoi}|${ten.toLowerCase()}`;
            const ex = map.get(key);

            if(!ex){
                const obj = {
                    id: autoID("mon"),
                    khoi: khoi,
                    ten: ten,
                    sotiet: "1",
                    gioihan: "1"
                };
                DATA.mon.push(obj);
                map.set(key, obj);
                changed = true;
                continue;
            }

            // điền mặc định nếu thiếu/không hợp lệ
            const st = String(ex.sotiet ?? "").trim();
            const gh = String(ex.gioihan ?? "").trim();
            if(st === "" || Number.isNaN(Number(st))){
                ex.sotiet = "1";
                changed = true;
            }
            if(gh === "" || Number.isNaN(Number(gh))){
                ex.gioihan = "1";
                changed = true;
            }
        }
    }

    if(changed) saveStore();
}

// Room list helpers (hiển thị kèm Mã môn học trong ngoặc nếu có)
function getPhongOptionItems(){
    const items = (DATA.phong || [])
        .map(p=>({
            value: _normText(p?.ten),
            mon: _normText(p?.tinhtrang)
        }))
        .filter(x=>x.value)
        .sort((a,b)=>a.value.localeCompare(b.value,'vi'));
    return items.map(x=>({
        value: x.value,
        label: x.mon ? `${x.value} (${x.mon})` : x.value
    }));
}

/* ============================================================
   Helpers: hiển thị tên môn theo mã / id / tên
============================================================ */
function resolveMonDisplay(monValue){
    const v = _normText(monValue);
    if (!v) return "";
    const low = v.toLowerCase();
    const found = (DATA.monhoc || []).find(r=>{
        const id  = _normText(r.id).toLowerCase();
        const ten = _normText(r.ten).toLowerCase();
        const ma  = _normText(r.ma).toLowerCase();
        const ma2 = _normText(r.ma2).toLowerCase();
        return (id && id === low) || (ten && ten === low) || (ma && ma === low) || (ma2 && ma2 === low);
    });
    return found ? (_normText(found.ten) || v) : v;
}




function escapeHtml(s){
    s = (s===undefined || s===null) ? "" : String(s);
    return s.replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;")
            .replace(/"/g,"&quot;").replace(/'/g,"&#39;");
}

/* ===== PCCM helpers: danh mục môn (môn tổng hợp) + lookup tiết chuẩn ===== */
function buildPCCMMonList(){
    // HIỂN THỊ: TÊN MÔN (Âm nhạc, Chào cờ...)
    // LƯU/TRA PCCM: ưu tiên MÃ (SHDC, SHL...) để khớp dữ liệu PCCM đang lưu theo mã
    function looksLikeCode(s){
        s = _normText(s);
        if (!s) return false;
        const hasNonASCII = /[^\x00-\x7F]/.test(s);
        const hasSpace = /\s/.test(s);
        if (hasNonASCII || hasSpace) return false;
        return s.length <= 12;
    }

    const out = [];
    const seen = new Set();

    (DATA.monhoc || []).forEach(r=>{
        const ten = _normText(r.ten);
        const ma  = _normText(r.ma);
        const ma2 = _normText(r.ma2);

        const fields = [ten, ma, ma2].filter(Boolean);
        if (!fields.length) return;

        const displayName = _normText(fields.find(x=>!looksLikeCode(x)) || ten || ma || ma2);
        if (!displayName) return;

        const code = _normText(fields.find(x=>looksLikeCode(x)) || ma || ma2 || "");
        const storageKey = code || displayName;

        const normSeen = displayName.toLowerCase();
        if (seen.has(normSeen)) return;
        seen.add(normSeen);

        out.push({
            key: storageKey,      // key dùng để đọc/ghi PCCM (mã nếu có)
            ten: displayName,     // hiển thị
            code: code || "",     // hiển thị mã trên header
            ma: ma,
            ma2: ma2
        });
    });

    return out;
}

function pccmGetTeacher(lopCanon, monObj){
    if (!lopCanon || !monObj) return "";

    const clsRaw = _normText(lopCanon);
    const clsNorm = normalizeClassName(clsRaw);
    const classCandidates = Array.from(new Set([clsRaw, clsNorm].filter(Boolean)));

    const monCandidates = [];
    if (monObj.key) monCandidates.push(monObj.key);
    if (monObj.ten && monObj.ten !== monObj.key) monCandidates.push(monObj.ten);
    if (monObj.ma2) monCandidates.push(monObj.ma2);
    if (monObj.ma && monObj.ma !== monObj.key) monCandidates.push(monObj.ma);

    const matrix = DATA.pccmMatrix || {};
    for (const cls of classCandidates){
        for (const mk of monCandidates){
            const v = matrix[`${cls}|${mk}`];
            if (v) return String(v).trim();
        }
    }
    return "";
}

function pccmGetNumberFromMatrix(matrix, lopCanon, monObj){
    if (!matrix || !lopCanon || !monObj) return null;

    const clsRaw = _normText(lopCanon);
    const clsNorm = normalizeClassName(clsRaw);
    const classCandidates = Array.from(new Set([clsRaw, clsNorm].filter(Boolean)));

    const monCandidates = [];
    if (monObj.key) monCandidates.push(monObj.key);
    if (monObj.ten && monObj.ten !== monObj.key) monCandidates.push(monObj.ten);
    if (monObj.ma2) monCandidates.push(monObj.ma2);
    if (monObj.ma && monObj.ma !== monObj.key) monCandidates.push(monObj.ma);

    for (const cls of classCandidates){
        for (const mk of monCandidates){
            const raw = matrix[`${cls}|${mk}`];
            if (raw === undefined || raw === null) continue;
            const s = String(raw).trim();
            if (s === "") continue;
            const n = Number(s);
            if (!Number.isNaN(n)) return n;
        }
    }
    return null;
}

function pccmGetTiet(lopCanon, monObj){
    return pccmGetNumberFromMatrix(DATA.pccmTietMatrix || {}, lopCanon, monObj);
}
function pccmGetGioihan(lopCanon, monObj){
    return pccmGetNumberFromMatrix(DATA.pccmGioihanMatrix || {}, lopCanon, monObj);
}

// Tổng số tiết của 1 lớp (đã phân công giáo viên) / tổng tiết cần có
function pccmComputeTotalTietForClass(classDisplayName, monList){
    if (!classDisplayName) return null;
    const canon = normalizeClassName(classDisplayName);

    const lopObj = (DATA.lop || []).find(l=>{
        const t2 = canonTen2FromLop(l);
        return normalizeClassName(t2) === canon || normalizeClassName(l?.id) === canon;
    });
    if (!lopObj) return null;

    const rawKhoi = _normText(lopObj?.khoi) || _normText(lopObj?.ten2) || _normText(lopObj?.ten);
    const khoiNum = extractKhoiNumber(rawKhoi);
    const khoiName = khoiNum ? `Khối ${khoiNum}` : (_normText(lopObj?.khoi) || rawKhoi);

    let total = 0;
    let assigned = 0;

    (monList || []).forEach(m=>{
        // base theo Tiết chuẩn
        const tc = lookupTietChuan(khoiName, m?.ten);
        let base = 1;
        if (tc){
            const b = Number(tc.sotiet);
            if (!Number.isNaN(b)) base = b;
        }

        // override theo PCCM (nếu có)
        const ov = pccmGetTiet(canon, m);
        const tiet = (ov !== null) ? ov : base;
        if (!(Number.isFinite(tiet) && tiet > 0)) return;

        total += tiet;
        const gv = pccmGetTeacher(canon, m);
        if (gv) assigned += tiet;
    });

    return {total, assigned, khoiName, canon};
}

function pccmSetTeacher(lopCanon, monObj, val){
    if (!lopCanon || !monObj) return;
    // lưu theo key chính
    setPCCMTeacher(lopCanon, monObj.key || monObj.ten, val);

    // dọn key legacy để tránh "trùng" dữ liệu khi hiển thị
    const legacyKeys = [];
    if (monObj.ten && monObj.key && monObj.ten !== monObj.key) legacyKeys.push(`${lopCanon}|${monObj.ten}`);
    if (monObj.ma2 && monObj.key && monObj.ma2 !== monObj.key) legacyKeys.push(`${lopCanon}|${monObj.ma2}`);
    if (monObj.ma && monObj.key && monObj.ma !== monObj.key) legacyKeys.push(`${lopCanon}|${monObj.ma}`);

    legacyKeys.forEach(k=>{
        if (k === `${lopCanon}|${(monObj.key||"")}`) return;
        // nếu người dùng đang set theo key chính thì xóa legacy để khỏi hiển thị 2 cột
        if (DATA.pccmMatrix[k]) delete DATA.pccmMatrix[k];
    });
}

// ===== PCCM ROOM (Phòng học) =====
function pccmGetRoom(lopCanon, monObj){
    if (!lopCanon || !monObj) return "";
    const clsRaw = _normText(lopCanon);
    const clsNorm = normalizeClassName(clsRaw);
    const classCandidates = Array.from(new Set([clsRaw, clsNorm].filter(Boolean)));

    const monCandidates = [];
    if (monObj.key) monCandidates.push(monObj.key);
    if (monObj.ten && monObj.ten !== monObj.key) monCandidates.push(monObj.ten);
    if (monObj.ma2) monCandidates.push(monObj.ma2);
    if (monObj.ma && monObj.ma !== monObj.key) monCandidates.push(monObj.ma);

    for (const clsKey of classCandidates){
        for (const mk of monCandidates){
            const k = `${clsKey}|${mk}`;
            const v = (DATA.pccmRoomMatrix || {})[k];
            if (v) return v;
        }
    }
    return "";
}

function pccmSetRoom(lopCanon, monObj, val){
    if (!lopCanon || !monObj) return;
    setPCCMRoom(lopCanon, monObj.key || monObj.ten, val);

    // dọn key legacy để tránh lưu trùng (tương tự GV)
    const legacyKeys = [];
    if (monObj.ten && monObj.key && monObj.ten !== monObj.key) legacyKeys.push(`${lopCanon}|${monObj.ten}`);
    if (monObj.ma2 && monObj.key && monObj.ma2 !== monObj.key) legacyKeys.push(`${lopCanon}|${monObj.ma2}`);
    if (monObj.ma && monObj.key && monObj.ma !== monObj.key) legacyKeys.push(`${lopCanon}|${monObj.ma}`);
    legacyKeys.forEach(k=>{
        if (k === `${lopCanon}|${(monObj.key||"")}`) return;
        if (DATA.pccmRoomMatrix && DATA.pccmRoomMatrix[k]) delete DATA.pccmRoomMatrix[k];
    });
}

// ===== PCCM TIẾT / GIỚI HẠN theo Lớp|Môn (để màn "Sắp xếp TKB" lấy đúng tổng số tiết từ bảng phân công) =====
function pccmSetTietGioihanNoSave(lopCanon, monObj, sotietVal, gioihanVal){
    if (!lopCanon || !monObj) return;
    if (typeof DATA.pccmTietMatrix !== "object" || !DATA.pccmTietMatrix) DATA.pccmTietMatrix = {};
    if (typeof DATA.pccmGioihanMatrix !== "object" || !DATA.pccmGioihanMatrix) DATA.pccmGioihanMatrix = {};

    const monKey = (monObj.key || monObj.ten || "").toString().trim();
    if (!monKey) return;
    const primaryKey = `${lopCanon}|${monKey}`;

    const s1 = (sotietVal ?? "").toString().trim();
    const s2 = (gioihanVal ?? "").toString().trim();
    if (s1) DATA.pccmTietMatrix[primaryKey] = s1;
    else delete DATA.pccmTietMatrix[primaryKey];
    if (s2) DATA.pccmGioihanMatrix[primaryKey] = s2;
    else delete DATA.pccmGioihanMatrix[primaryKey];

    // dọn key legacy (tương tự GV/Phòng) để tránh lưu trùng theo tên/mã khác nhau
    const legacyKeys = [];
    if (monObj.ten && monObj.key && monObj.ten !== monObj.key) legacyKeys.push(`${lopCanon}|${monObj.ten}`);
    if (monObj.ma2 && monObj.key && monObj.ma2 !== monObj.key) legacyKeys.push(`${lopCanon}|${monObj.ma2}`);
    if (monObj.ma && monObj.key && monObj.ma !== monObj.key) legacyKeys.push(`${lopCanon}|${monObj.ma}`);
    legacyKeys.forEach(k=>{
        if (k === primaryKey) return;
        if (DATA.pccmTietMatrix && (k in DATA.pccmTietMatrix)) delete DATA.pccmTietMatrix[k];
        if (DATA.pccmGioihanMatrix && (k in DATA.pccmGioihanMatrix)) delete DATA.pccmGioihanMatrix[k];
    });
}

// Set nhanh theo monKey (dùng trong tab Môn học)
function pccmSetTeacherByMonKey(lopCanon, monKey, val){
    const mons = buildPCCMMonList();
    const monObj = mons.find(m => _normText(m.key) === _normText(monKey)) || {key: monKey, ten: monKey};
    pccmSetTeacher(lopCanon, monObj, val);
    saveStore();
}
function pccmSetRoomByMonKey(lopCanon, monKey, val){
    const mons = buildPCCMMonList();
    const monObj = mons.find(m => _normText(m.key) === _normText(monKey)) || {key: monKey, ten: monKey};
    pccmSetRoom(lopCanon, monObj, val);
    saveStore();
}

function lookupTietChuan(khoiName, monObj){
    const kNum = extractKhoiNumber(_normText(khoiName));
    if (!kNum || !monObj) return null;
    const m1 = (monObj.ten||"").toLowerCase();
    const m2 = (monObj.key||"").toLowerCase();
    const m3 = (monObj.ma2||"").toLowerCase();

    return (DATA.mon || []).find(r=>{
        const rk = extractKhoiNumber(_normText(r.khoi));
        if (rk !== kNum) return false;
        const rt = _normText(r.ten).toLowerCase();
        return (rt && (rt === m1 || rt === m2 || (m3 && rt === m3)));
    }) || null;
}

// ===== Helpers: kiểm tra "môn có tiết" =====
function _toPositiveNumberOrZero(x){
    const s = (x ?? "").toString().trim();
    if (s === "") return 0;
    const n = Number(s);
    if (!Number.isFinite(n)) return 0;
    return n;
}

function monHasPositiveTietChuanForKhoi(khoiName, monObj){
    const tc = lookupTietChuan(khoiName, monObj);
    if (!tc) return false;
    return _toPositiveNumberOrZero(tc.sotiet) > 0;
}

function monHasPositiveTietChuanAnyKhoi(monObj){
    const khoiNames = Array.from(new Set((DATA.mon || []).map(r=>_normText(r.khoi)).filter(Boolean)));
    for (const k of khoiNames){
        if (monHasPositiveTietChuanForKhoi(k, monObj)) return true;
    }
    return false;
}

function pccmSetTeacherFromInput(inp, lopCanon){
    const monObj = {
        key: inp.dataset.monkey || "",
        ten: inp.dataset.monten || "",
        ma: inp.dataset.monma || "",
        ma2: inp.dataset.monma2 || ""
    };
    pccmSetTeacher(lopCanon, monObj, inp.value);
    saveStore();
}

function buildTeacherDatalistHTML(){
    const items = (DATA.giaovien || [])
        .map(g=>{
            const code = _normText(g.magv);
            const name = `${(g.hodem||"").trim()} ${(g.ten||"").trim()}`.trim();
            return {code, name};
        })
        .filter(x=>x.code);

    const opts = items.map(x=>{
        // Chỉ hiển thị mã GV (tên tắt) theo yêu cầu
        const label = x.code;
        return `<option value="${escapeHtml(x.code)}" label="${escapeHtml(label)}"></option>`;
    }).join("");

    return `<datalist id="pccmTeacherCodes">${opts}</datalist>`;
}

function resolveTeacherCode(input){
    const raw = (input||"").trim();
    if (!raw) return "";
    const lower = raw.toLowerCase();

    // Ưu tiên match theo mã GV
    const byCode = (DATA.giaovien || []).find(g => _normText(g.magv).toLowerCase() === lower);
    if (byCode && byCode.magv) return _normText(byCode.magv);

    // Fallback: match theo tên đầy đủ
    const byName = (DATA.giaovien || []).find(g=>{
        const full = `${(g.hodem||"").trim()} ${(g.ten||"").trim()}`.trim().toLowerCase();
        return full && full === lower;
    });
    if (byName && byName.magv) return _normText(byName.magv);

    // Nếu không khớp, giữ nguyên (cho phép nhập tự do)
    return raw;
}

// Sửa lớp PCCM theo lớp+môn+GV (mã)
function pccmQuickEditApply(){
    const cls = document.getElementById("pccmQuickClass")?.value || "";
    const monKey = document.getElementById("pccmQuickMon")?.value || "";
    const gvInput = document.getElementById("pccmQuickGv")?.value || "";
    const roomInput = document.getElementById("pccmQuickRoom")?.value || "";
    const sotietInput = (document.getElementById("pccmQuickSoTiet")?.value || "").toString().trim();
    const gioihanInput = (document.getElementById("pccmQuickGioiHan")?.value || "").toString().trim();

    if (!cls || !monKey){
        alert("Chọn Lớp và Môn trước.");
        return;
    }

    const mons = buildPCCMMonList();
    const monObj = mons.find(m => (m.key||"") === monKey) || {key: monKey, ten: monKey};

    const gvCode = resolveTeacherCode(gvInput);
    pccmSetTeacher(cls, monObj, gvCode);

    // Phòng: lấy từ bảng Phòng (nếu có)
    if (typeof DATA.pccmRoomMatrix !== "object" || !DATA.pccmRoomMatrix) DATA.pccmRoomMatrix = {};
    pccmSetRoom(cls, monObj, roomInput);

    // Nếu người dùng nhập Số tiết/Giới hạn ở box thêm nhanh thì cập nhật TIẾT CHUẨN theo khối của lớp.
    // Nếu không nhập gì => giữ nguyên theo tiết chuẩn hiện có.
    function _validateNumber(v){
        const s = (v ?? "").toString().trim();
        if (s === "") return { ok:true, val:"" };
        const n = Number(s);
        if (!Number.isFinite(n)) return { ok:false };
        if (n <= 0) return { ok:false };
        return { ok:true, val: Number.isInteger(n) ? String(n) : String(n) };
    }

    const v1 = _validateNumber(sotietInput);
    const v2 = _validateNumber(gioihanInput);
    if (!v1.ok || !v2.ok){
        alert("⚠ Số tiết/Giới hạn không hợp lệ. Chỉ được để trống hoặc nhập số > 0.");
        return;
    }

    if (v1.val !== "" || v2.val !== ""){
        const khoiName = pccmGetKhoiNameForClass(cls);
        let tcRow = lookupTietChuan(khoiName, monObj);
        if (tcRow){
            if (v1.val !== "") tcRow.sotiet = v1.val;
            if (v2.val !== "") tcRow.gioihan = v2.val;
        } else {
            // phòng khi thiếu dữ liệu tiết chuẩn cho khối này
            DATA.mon = DATA.mon || [];
            DATA.mon.push({
                id: autoID("mon"),
                khoi: khoiName,
                ten: (monObj.ten || monObj.key || "").trim(),
                sotiet: v1.val,
                gioihan: v2.val,
                ghichu: ""
            });
        }
    }

    saveStore();

    // clear 2 ô nhập để tránh áp dụng nhầm lần sau
    try{
        const a = document.getElementById("pccmQuickSoTiet");
        const b = document.getElementById("pccmQuickGioiHan");
        if (a) a.value = "";
        if (b) b.value = "";
    }catch(e){}

    const sc = document.getElementById("section-content");
    if (sc && typeof renderPCCM === "function") sc.innerHTML = renderPCCM();
}

// ===== PCCM: lọc danh sách môn theo lớp (dựa theo tiết chuẩn: số tiết > 0) =====
function pccmGetKhoiNameForClass(cls){
    const canon = normalizeClassName(_normText(cls));
    const found = (DATA.lop || []).find(l => canonTen2FromLop(l) === canon);
    const khoi = _normText(found?.khoi);
    return khoi || ("Khối " + extractKhoiNumber(canon));
}

function pccmGetAllowedMonsForClass(cls, monList){
    const khoiName = pccmGetKhoiNameForClass(cls);
    const out = (monList || []).filter(m => monHasPositiveTietChuanForKhoi(khoiName, m));
    return out;
}

function pccmQuickOnClassChange(cls){
    const monSel = document.getElementById("pccmQuickMon");
    if (!monSel) return;

    // base list giống như màn Phân công: chỉ lấy môn có tiết chuẩn ở ít nhất 1 khối
    const base = buildPCCMMonList().filter(m => monHasPositiveTietChuanAnyKhoi(m));
    const allowed = pccmGetAllowedMonsForClass(cls, base);
    const cur = (monSel.value || "").toString();

    monSel.innerHTML = allowed.map(m=>{
        const k = (m.key||m.ten||"");
        const label = (m.ten || m.code || m.ma || m.key || '');
        return `<option value="${escapeHtml(k)}">${escapeHtml(label)}</option>`;
    }).join("");

    if (allowed.some(m => (m.key||m.ten||"") === cur)) monSel.value = cur;

    // reset 2 ô nhập nhanh số tiết/giới hạn (để mặc định theo Tiết chuẩn)
    try{
        const a = document.getElementById("pccmQuickSoTiet");
        const b = document.getElementById("pccmQuickGioiHan");
        if (a) a.value = "";
        if (b) b.value = "";
    }catch(e){}
}

// Box "Thêm nhanh" ở bên phải trong mục Phân công
function renderPCCMQuickBox(classNames, monList){
    const classes = (classNames || []).slice();
    const mons = (monList || []).slice();

    // Rooms (lấy từ bảng Phòng). Hiển thị kèm "(Mã môn học)" nếu có.
    const roomItems = getPhongOptionItems();

    // Teachers: chỉ hiển thị mã (tên tắt)
    const gvCodes = (DATA.giaovien || [])
        .map(g=>_normText(g?.magv))
        .filter(Boolean)
        .sort((a,b)=>a.localeCompare(b,'vi'));

    const defaultClass = classes.includes(PCCM_SELECTED_CLASS) ? PCCM_SELECTED_CLASS : (classes[0] || "");
    const allowedMons = pccmGetAllowedMonsForClass(defaultClass, mons);
    const defaultMonKey = (PCCM_TAB === "monhoc" && PCCM_SUBJ)
        ? PCCM_SUBJ
        : ((allowedMons[0] && allowedMons[0].key) ? allowedMons[0].key : "");
    const defaultGv = (PCCM_TAB === "giaovien" && PCCM_SELECTED_GV) ? PCCM_SELECTED_GV : "";

    return `
    <div style="width:280px">
        <div style="font-weight:900;margin:6px 0 8px">Thêm nhanh</div>
        <div style="background:#fff;border:1px solid #e3e8f3;border-radius:10px;padding:10px">
            <div style="font-weight:700;margin-bottom:6px">Lớp</div>
            <select id="pccmQuickClass" style="width:100%;margin-bottom:10px" onchange="pccmQuickOnClassChange(this.value)">
                ${classes.map(c=>`<option value="${escapeHtml(c)}" ${c===defaultClass?"selected":""}>${escapeHtml(c)}</option>`).join("")}
            </select>

            <div style="font-weight:700;margin-bottom:6px">Môn</div>
            <select id="pccmQuickMon" style="width:100%;margin-bottom:10px">
                ${allowedMons.map(m=>`<option value="${escapeHtml(m.key||m.ten||"")}" ${(m.key||m.ten||"")===defaultMonKey?"selected":""}>${escapeHtml(m.ten || m.code || m.ma || m.key || '')}</option>`).join("")}
            </select>

            <div style="font-weight:700;margin-bottom:6px">Giáo viên</div>
            <select id="pccmQuickGv" style="width:100%;margin-bottom:10px">
                <option value="" ${defaultGv===""?"selected":""}>(Chưa phân)</option>
                ${gvCodes.map(code=>`<option value="${escapeHtml(code)}" ${code===defaultGv?"selected":""}>${escapeHtml(code)}</option>`).join("")}
            </select>

            <div style="font-weight:700;margin-bottom:6px">Phòng học</div>
            <select id="pccmQuickRoom" style="width:100%;margin-bottom:10px">
                <option value="" selected>(Không chọn)</option>
                ${roomItems.map(it=>`<option value="${escapeHtml(it.value)}">${escapeHtml(it.label)}</option>`).join("")}
            </select>

            <div style="font-weight:700;margin-bottom:6px">Số tiết</div>
            <input id="pccmQuickSoTiet" class="inline-edit-input" type="number" min="0" step="1"
                   placeholder="(mặc định theo Tiết chuẩn)" style="width:100%;margin-bottom:10px;text-align:center">

            <div style="font-weight:700;margin-bottom:6px">Giới hạn</div>
            <input id="pccmQuickGioiHan" class="inline-edit-input" type="number" min="0" step="1"
                   placeholder="(mặc định theo Tiết chuẩn)" style="width:100%;margin-bottom:10px;text-align:center">

            <button class="btn primary" style="width:100%" onclick="pccmQuickEditApply()">Áp dụng</button>

        </div>
    </div>`;
}

// Lưu PCCM (tab Lớp): GV (listbox) + Tiết/Giới hạn (textbox) — chỉ lưu 1 lần
function pccmSaveClassEdits(){
    const cls = (PCCM_CLASS_EDIT_CACHE?.cls || "").trim();
    const khoiName = (PCCM_CLASS_EDIT_CACHE?.khoiName || "").trim();
    const mons = PCCM_CLASS_EDIT_CACHE?.mons || [];

    if (!cls || !khoiName || !Array.isArray(mons) || !mons.length){
        alert("⚠ Không có dữ liệu để lưu. Hãy chọn lớp trước.");
        return;
    }

    // validate: empty allowed; if not empty => >0
    function _validateNumber(v){
        const s = (v ?? "").toString().trim();
        if (s === "") return { ok:true, val:"" };
        const n = Number(s);
        if (!Number.isFinite(n)) return { ok:false };
        if (n <= 0) return { ok:false };
        return { ok:true, val: Number.isInteger(n) ? String(n) : String(n) };
    }

    // set teacher without saving each change
    function _setTeacherNoSave(lopCanon, monObj, val){
        if (!lopCanon || !monObj) return;
        val = (val || "").trim();
        const primaryKey = `${lopCanon}|${(monObj.key || monObj.ten || "").trim()}`;
        if (val) DATA.pccmMatrix[primaryKey] = val;
        else delete DATA.pccmMatrix[primaryKey];

        // xóa legacy keys để tránh trùng
        const legacyKeys = [];
        if (monObj.ten && monObj.key && monObj.ten !== monObj.key) legacyKeys.push(`${lopCanon}|${monObj.ten}`);
        if (monObj.ma2 && monObj.key && monObj.ma2 !== monObj.key) legacyKeys.push(`${lopCanon}|${monObj.ma2}`);
        if (monObj.ma && monObj.key && monObj.ma !== monObj.key) legacyKeys.push(`${lopCanon}|${monObj.ma}`);

        legacyKeys.forEach(k=>{
            if (k === primaryKey) return;
            if (DATA.pccmMatrix[k]) delete DATA.pccmMatrix[k];
        });
    }

    // set room without saving each change
    function _setRoomNoSave(lopCanon, monObj, val){
        if (!lopCanon || !monObj) return;
        if (typeof DATA.pccmRoomMatrix !== "object" || !DATA.pccmRoomMatrix) DATA.pccmRoomMatrix = {};
        val = (val || "").trim();
        const primaryKey = `${lopCanon}|${(monObj.key || monObj.ten || "").trim()}`;
        if (val) DATA.pccmRoomMatrix[primaryKey] = val;
        else delete DATA.pccmRoomMatrix[primaryKey];

        const legacyKeys = [];
        if (monObj.ten && monObj.key && monObj.ten !== monObj.key) legacyKeys.push(`${lopCanon}|${monObj.ten}`);
        if (monObj.ma2 && monObj.key && monObj.ma2 !== monObj.key) legacyKeys.push(`${lopCanon}|${monObj.ma2}`);
        if (monObj.ma && monObj.key && monObj.ma !== monObj.key) legacyKeys.push(`${lopCanon}|${monObj.ma}`);

        legacyKeys.forEach(k=>{
            if (k === primaryKey) return;
            if (DATA.pccmRoomMatrix[k]) delete DATA.pccmRoomMatrix[k];
        });
    }

    // pass 1: đọc DOM + check trùng khóa (Lớp|Môn)
    const allMons = buildPCCMMonList();
    const monMap = new Map(allMons.map(m=>[_normText(m.key||m.ten), m]));
    function _findMonObj(monKey){
        const k = _normText(monKey);
        return monMap.get(k)
            || allMons.find(m=>_normText(m.ten) === k)
            || { key: k, ten: k, code:"", ma:"", ma2:"" };
    }

    const ops = [];
    const seen = new Set();

    for (let idx=0; idx<mons.length; idx++){
        const oldMonObj = mons[idx];
        const newMonKey = (document.getElementById(`pccm_mon_${idx}`)?.value || (oldMonObj.key||oldMonObj.ten||""))
            .toString().trim();
        if (!newMonKey){
            alert(`⚠ Môn học không hợp lệ ở dòng ${idx+1}.`);
            return;
        }
        const newMonObj = _findMonObj(newMonKey);
        const key = `${cls}|${(newMonObj.key || newMonObj.ten || '').trim()}`;
        if (seen.has(key)){
            alert(`⚠ Trùng phân công (Lớp|Môn) ở dòng ${idx+1}: ${key}.`);
            return;
        }
        seen.add(key);

        ops.push({ idx, oldMonObj, newMonObj });
    }

    // clear key cũ nếu người dùng đổi Môn
    const cleared = new Set();
    for (const op of ops){
        const oldKey = `${cls}|${(op.oldMonObj?.key || op.oldMonObj?.ten || '').trim()}`;
        const newKey = `${cls}|${(op.newMonObj?.key || op.newMonObj?.ten || '').trim()}`;
        if (oldKey && newKey && oldKey !== newKey && !cleared.has(oldKey)){
            cleared.add(oldKey);
            _setTeacherNoSave(cls, op.oldMonObj, "");
            _setRoomNoSave(cls, op.oldMonObj, "");
            pccmSetTietGioihanNoSave(cls, op.oldMonObj, "", "");
        }
    }

    // apply all rows
    for (const op of ops){
        const idx = op.idx;
        const m = op.newMonObj;

        const gv = (document.getElementById(`pccm_gv_${idx}`)?.value || "").trim();
        const roomEl = document.getElementById(`pccm_room_${idx}`);
        const room = roomEl ? (roomEl.value || "").trim() : null;
        const sotietRaw = (document.getElementById(`pccm_sotiet_${idx}`)?.value || "").toString().trim();
        const gioihanRaw = (document.getElementById(`pccm_gioihan_${idx}`)?.value || "").toString().trim();

        // 1) GV (PCCM)
        _setTeacherNoSave(cls, m, gv);

        // 1b) Phòng (nếu có cột phòng)
        if (room !== null) _setRoomNoSave(cls, m, room);

        // 2) Tiết chuẩn theo khối
        const v1 = _validateNumber(sotietRaw);
        const v2 = _validateNumber(gioihanRaw);
        if (!v1.ok || !v2.ok){
            alert(`⚠ Dữ liệu không hợp lệ ở môn: ${(m.ten||m.key||"")} .
- Chỉ được để trống hoặc nhập số > 0.`);
            return;
        }

        // Lưu số tiết / giới hạn theo Lớp|Môn (nguồn cho màn "Sắp xếp TKB")
        pccmSetTietGioihanNoSave(cls, m, v1.val, v2.val);

        let tcRow = lookupTietChuan(khoiName, m);
        if (tcRow){
            tcRow.sotiet = v1.val;
            tcRow.gioihan = v2.val;
        } else {
            // tạo mới nếu người dùng có nhập gì đó
            if (v1.val !== "" || v2.val !== ""){
                DATA.mon = DATA.mon || [];
                DATA.mon.push({
                    id: autoID("mon"),
                    khoi: khoiName,
                    ten: (m.ten || m.key || "").trim(),
                    sotiet: v1.val,
                    gioihan: v2.val || "",
                    ghichu: ""
                });
            }
        }
    }

    saveStore();

    const sc = document.getElementById("section-content");
    if (sc && typeof renderPCCM === "function") sc.innerHTML = renderPCCM();
    alert("✔ Đã lưu PCCM.");
}

// Backward compatibility (nếu đâu đó còn gọi)
function pccmQuickAssignClass(){
    return pccmQuickEditApply();
}

// Lưu PCCM (tab Giáo viên): Phòng (listbox) + Tiết/Giới hạn (textbox) — lưu 1 lần
function pccmSaveTeacherEdits(){
    const rows = PCCM_TEACHER_EDIT_CACHE?.rows || [];
    const gv = (PCCM_TEACHER_EDIT_CACHE?.gv || "").toString().trim();
    if (!Array.isArray(rows) || !rows.length){
        alert("⚠ Không có dữ liệu để lưu.");
        return;
    }
    if (!gv){
        alert("⚠ Chưa chọn giáo viên.");
        return;
    }

    if (typeof DATA.pccmMatrix !== "object" || !DATA.pccmMatrix) DATA.pccmMatrix = {};
    if (typeof DATA.pccmRoomMatrix !== "object" || !DATA.pccmRoomMatrix) DATA.pccmRoomMatrix = {};

    // map lớp -> khối (để update tiết chuẩn đúng khối sau khi đổi lớp)
    const classToKhoi = new Map((DATA.lop||[]).map(l=>{
        const canon = canonTen2FromLop(l) || normalizeClassName(l.ten) || "";
        const khoi = (l.khoi || ("Khối " + extractKhoiNumber(canon))).trim();
        return [canon.trim(), khoi];
    }).filter(x=>x[0]));

    // map mônKey -> monObj
    const allMons = buildPCCMMonList();
    const monMap = new Map(allMons.map(m=>[_normText(m.key||m.ten), m]));
    function _findMonObj(monKey){
        const k = _normText(monKey);
        return monMap.get(k)
            || allMons.find(m=>_normText(m.ten) === k)
            || { key: k, ten: k, code:"", ma:"", ma2:"" };
    }

    function _validateNumber(v){
        const s = (v ?? "").toString().trim();
        if (s === "") return { ok:true, val:"" };
        const n = Number(s);
        if (!Number.isFinite(n)) return { ok:false };
        if (n <= 0) return { ok:false };
        return { ok:true, val: Number.isInteger(n) ? String(n) : String(n) };
    }

    function _setTeacherNoSave(lopCanon, monObj, val){
        if (!lopCanon || !monObj) return;
        val = (val || "").trim();
        const primaryKey = `${lopCanon}|${(monObj.key || monObj.ten || "").trim()}`;
        if (val) DATA.pccmMatrix[primaryKey] = val;
        else delete DATA.pccmMatrix[primaryKey];

        const legacyKeys = [];
        if (monObj.ten && monObj.key && monObj.ten !== monObj.key) legacyKeys.push(`${lopCanon}|${monObj.ten}`);
        if (monObj.ma2 && monObj.key && monObj.ma2 !== monObj.key) legacyKeys.push(`${lopCanon}|${monObj.ma2}`);
        if (monObj.ma && monObj.key && monObj.ma !== monObj.key) legacyKeys.push(`${lopCanon}|${monObj.ma}`);
        legacyKeys.forEach(k=>{ if (k !== primaryKey && DATA.pccmMatrix[k]) delete DATA.pccmMatrix[k]; });
    }

    function _setRoomNoSave(lopCanon, monObj, val){
        if (!lopCanon || !monObj) return;
        val = (val || "").trim();
        const primaryKey = `${lopCanon}|${(monObj.key || monObj.ten || "").trim()}`;
        if (val) DATA.pccmRoomMatrix[primaryKey] = val;
        else delete DATA.pccmRoomMatrix[primaryKey];

        const legacyKeys = [];
        if (monObj.ten && monObj.key && monObj.ten !== monObj.key) legacyKeys.push(`${lopCanon}|${monObj.ten}`);
        if (monObj.ma2 && monObj.key && monObj.ma2 !== monObj.key) legacyKeys.push(`${lopCanon}|${monObj.ma2}`);
        if (monObj.ma && monObj.key && monObj.ma !== monObj.key) legacyKeys.push(`${lopCanon}|${monObj.ma}`);
        legacyKeys.forEach(k=>{ if (k !== primaryKey && DATA.pccmRoomMatrix[k]) delete DATA.pccmRoomMatrix[k]; });
    }

    // pass 1: đọc DOM + check trùng khóa
    const ops = [];
    const seenKeys = new Set();
    for (let idx=0; idx<rows.length; idx++){
        const r = rows[idx];
        const newCls = (document.getElementById(`pccmT_cls_${idx}`)?.value || r.cls || "").toString().trim();
        const newMonKey = (document.getElementById(`pccmT_mon_${idx}`)?.value || (r.monObj?.key||r.monObj?.ten||""))
            .toString().trim();
        if (!newCls || !newMonKey){
            alert(`⚠ Dữ liệu Lớp/Môn không hợp lệ ở dòng ${idx+1}.`);
            return;
        }
        const newMonObj = _findMonObj(newMonKey);
        const newKey = `${newCls}|${(newMonObj.key || newMonObj.ten || "").trim()}`;
        if (seenKeys.has(newKey)){
            alert(`⚠ Trùng phân công (Lớp|Môn) ở dòng ${idx+1}: ${newKey}.\nHãy kiểm tra lại vì 1 Lớp + 1 Môn chỉ nên có 1 dòng.`);
            return;
        }
        seenKeys.add(newKey);

        ops.push({
            idx,
            r,
            oldCls: (r.cls||"").trim(),
            oldMonObj: r.monObj,
            newCls,
            newMonObj,
            room: (document.getElementById(`pccmT_room_${idx}`)?.value || "").toString().trim(),
            sotietRaw: (document.getElementById(`pccmT_sotiet_${idx}`)?.value || "").toString().trim(),
            gioihanRaw: (document.getElementById(`pccmT_gioihan_${idx}`)?.value || "").toString().trim()
        });
    }

    // clear các key cũ nếu người dùng đổi Lớp/Môn
    const cleared = new Set();
    ops.forEach(op=>{
        const oldKey = `${op.oldCls}|${(op.oldMonObj?.key || op.oldMonObj?.ten || "").toString().trim()}`;
        const newKey = `${op.newCls}|${(op.newMonObj?.key || op.newMonObj?.ten || "").toString().trim()}`;
        if (oldKey && newKey && oldKey !== newKey && !cleared.has(oldKey)){
            cleared.add(oldKey);
            _setTeacherNoSave(op.oldCls, op.oldMonObj, "");
            _setRoomNoSave(op.oldCls, op.oldMonObj, "");
        }
    });

    // apply
    for (const op of ops){
        const khoiName = classToKhoi.get(op.newCls) || ("Khối " + extractKhoiNumber(op.newCls));

        // 1) GV (giáo viên đang chọn) + Phòng
        _setTeacherNoSave(op.newCls, op.newMonObj, gv);
        _setRoomNoSave(op.newCls, op.newMonObj, op.room);

        // 2) Tiết/Giới hạn
        const v1 = _validateNumber(op.sotietRaw);
        const v2 = _validateNumber(op.gioihanRaw);
        if (!v1.ok || !v2.ok){
            alert(`⚠ Dữ liệu không hợp lệ ở: ${(op.newMonObj.ten||op.newMonObj.key||"")} (${khoiName}).\n- Chỉ được để trống hoặc nhập số > 0.`);
            return;
        }

        // Lưu số tiết / giới hạn theo Lớp|Môn (nguồn cho màn "Sắp xếp TKB")
        pccmSetTietGioihanNoSave(op.newCls, op.newMonObj, v1.val, v2.val);

        let tcRow = lookupTietChuan(khoiName, op.newMonObj);
        if (tcRow){
            tcRow.sotiet = v1.val;
            tcRow.gioihan = v2.val;
        } else {
            if (v1.val !== "" || v2.val !== ""){
                DATA.mon = DATA.mon || [];
                DATA.mon.push({
                    id: autoID("mon"),
                    khoi: khoiName,
                    ten: (op.newMonObj.ten || op.newMonObj.key || "").trim(),
                    sotiet: v1.val,
                    gioihan: v2.val || "",
                    ghichu: ""
                });
            }
        }
    }

    saveStore();
    const sc = document.getElementById("section-content");
    if (sc && typeof renderPCCM === "function") sc.innerHTML = renderPCCM();
    alert("✔ Đã lưu PCCM.");
}

// Lưu PCCM (tab Môn học): GV (listbox) + Phòng (listbox) + Tiết/Giới hạn (textbox) — lưu 1 lần
function pccmSaveSubjectEdits(){
    const rows = PCCM_SUBJECT_EDIT_CACHE?.rows || [];
    const monObj = PCCM_SUBJECT_EDIT_CACHE?.monObj || null;

    if (!monObj || !Array.isArray(rows) || !rows.length){
        alert("⚠ Không có dữ liệu để lưu.");
        return;
    }

    if (typeof DATA.pccmMatrix !== "object" || !DATA.pccmMatrix) DATA.pccmMatrix = {};
    if (typeof DATA.pccmRoomMatrix !== "object" || !DATA.pccmRoomMatrix) DATA.pccmRoomMatrix = {};

    // map lớp -> khối (để update tiết chuẩn đúng khối sau khi đổi lớp)
    const classToKhoi = new Map((DATA.lop||[]).map(l=>{
        const canon = canonTen2FromLop(l) || normalizeClassName(l.ten) || "";
        const khoi = (l.khoi || ("Khối " + extractKhoiNumber(canon))).trim();
        return [canon.trim(), khoi];
    }).filter(x=>x[0]));

    function _validateNumber(v){
        const s = (v ?? "").toString().trim();
        if (s === "") return { ok:true, val:"" };
        const n = Number(s);
        if (!Number.isFinite(n)) return { ok:false };
        if (n <= 0) return { ok:false };
        return { ok:true, val: Number.isInteger(n) ? String(n) : String(n) };
    }

    function _setTeacherNoSave(lopCanon, monObj, val){
        if (!lopCanon || !monObj) return;
        val = (val || "").trim();
        const primaryKey = `${lopCanon}|${(monObj.key || monObj.ten || "").trim()}`;
        if (val) DATA.pccmMatrix[primaryKey] = val;
        else delete DATA.pccmMatrix[primaryKey];

        const legacyKeys = [];
        if (monObj.ten && monObj.key && monObj.ten !== monObj.key) legacyKeys.push(`${lopCanon}|${monObj.ten}`);
        if (monObj.ma2 && monObj.key && monObj.ma2 !== monObj.key) legacyKeys.push(`${lopCanon}|${monObj.ma2}`);
        if (monObj.ma && monObj.key && monObj.ma !== monObj.key) legacyKeys.push(`${lopCanon}|${monObj.ma}`);
        legacyKeys.forEach(k=>{ if (k !== primaryKey && DATA.pccmMatrix[k]) delete DATA.pccmMatrix[k]; });
    }

    function _setRoomNoSave(lopCanon, monObj, val){
        if (!lopCanon || !monObj) return;
        val = (val || "").trim();
        const primaryKey = `${lopCanon}|${(monObj.key || monObj.ten || "").trim()}`;
        if (val) DATA.pccmRoomMatrix[primaryKey] = val;
        else delete DATA.pccmRoomMatrix[primaryKey];

        const legacyKeys = [];
        if (monObj.ten && monObj.key && monObj.ten !== monObj.key) legacyKeys.push(`${lopCanon}|${monObj.ten}`);
        if (monObj.ma2 && monObj.key && monObj.ma2 !== monObj.key) legacyKeys.push(`${lopCanon}|${monObj.ma2}`);
        if (monObj.ma && monObj.key && monObj.ma !== monObj.key) legacyKeys.push(`${lopCanon}|${monObj.ma}`);
        legacyKeys.forEach(k=>{ if (k !== primaryKey && DATA.pccmRoomMatrix[k]) delete DATA.pccmRoomMatrix[k]; });
    }

    // pass 1: đọc DOM + check trùng lớp
    const ops = [];
    const seenCls = new Set();
    for (let idx=0; idx<rows.length; idx++){
        const r = rows[idx];
        const newCls = (document.getElementById(`pccmS_cls_${idx}`)?.value || r.cls || "").toString().trim();
        if (!newCls){
            alert(`⚠ Dữ liệu Lớp không hợp lệ ở dòng ${idx+1}.`);
            return;
        }
        if (seenCls.has(newCls)){
            alert(`⚠ Trùng Lớp ở dòng ${idx+1}: ${newCls}.\nHãy kiểm tra lại vì 1 lớp chỉ nên xuất hiện 1 lần trong bảng này.`);
            return;
        }
        seenCls.add(newCls);

        ops.push({
            idx,
            r,
            oldCls: (r.cls||"").trim(),
            newCls,
            gv: (document.getElementById(`pccmS_gv_${idx}`)?.value || "").toString().trim(),
            room: (document.getElementById(`pccmS_room_${idx}`)?.value || "").toString().trim(),
            sotietRaw: (document.getElementById(`pccmS_sotiet_${idx}`)?.value || "").toString().trim(),
            gioihanRaw: (document.getElementById(`pccmS_gioihan_${idx}`)?.value || "").toString().trim()
        });
    }

    // clear key cũ nếu đổi lớp
    const cleared = new Set();
    ops.forEach(op=>{
        if (op.oldCls && op.newCls && op.oldCls !== op.newCls && !cleared.has(op.oldCls)){
            cleared.add(op.oldCls);
            _setTeacherNoSave(op.oldCls, monObj, "");
            _setRoomNoSave(op.oldCls, monObj, "");
        }
    });

    // apply
    for (const op of ops){
        const khoiName = classToKhoi.get(op.newCls) || ("Khối " + extractKhoiNumber(op.newCls));

        // 1) GV + Phòng
        _setTeacherNoSave(op.newCls, monObj, op.gv);
        _setRoomNoSave(op.newCls, monObj, op.room);

        // 2) Tiết chuẩn theo khối
        const v1 = _validateNumber(op.sotietRaw);
        const v2 = _validateNumber(op.gioihanRaw);
        if (!v1.ok || !v2.ok){
            alert(`⚠ Dữ liệu không hợp lệ ở lớp: ${op.newCls} (${khoiName}).\n- Chỉ được để trống hoặc nhập số > 0.`);
            return;
        }

        // Lưu số tiết / giới hạn theo Lớp|Môn (nguồn cho màn "Sắp xếp TKB")
        pccmSetTietGioihanNoSave(op.newCls, monObj, v1.val, v2.val);

        let tcRow = lookupTietChuan(khoiName, monObj);
        if (tcRow){
            tcRow.sotiet = v1.val;
            tcRow.gioihan = v2.val;
        } else {
            if (v1.val !== "" || v2.val !== ""){
                DATA.mon = DATA.mon || [];
                DATA.mon.push({
                    id: autoID("mon"),
                    khoi: khoiName,
                    ten: (monObj.ten || monObj.key || "").trim(),
                    sotiet: v1.val,
                    gioihan: v2.val || "",
                    ghichu: ""
                });
            }
        }
    }

    saveStore();
    const sc = document.getElementById("section-content");
    if (sc && typeof renderPCCM === "function") sc.innerHTML = renderPCCM();
    alert("✔ Đã lưu PCCM.");
}



/* ============================================================
   ĐỒNG BỘ MÔN (menu lớn "Môn học" = DATA.monhoc)
   - Ý tưởng: 1 môn tổng hợp (monhoc) được xem là "đồng bộ" nếu đã có đủ
     trong TIẾT CHUẨN (DATA.mon) của TẤT CẢ các khối.
   - Nếu chưa đủ, bấm checkbox sẽ tự bổ sung môn đó vào các khối còn thiếu.
============================================================ */
function getAllKhoiNamesForSync(){
    let khois = (DATA.khoi || []).map(k => _normText(k.ten)).filter(Boolean);

    // Nếu chưa có danh sách khối, thử lấy từ danh sách lớp
    if (!khois.length){
        khois = (DATA.lop || []).map(l=>{
            const canon = (l.ten2 || normalizeClassName(l.ten) || "").trim();
            const khoi = _normText(l.khoi || ("Khối " + extractKhoiNumber(canon)));
            return khoi;
        }).filter(Boolean);
    }

    // Nếu vẫn trống, thử lấy từ tiết chuẩn hiện có
    if (!khois.length){
        khois = (DATA.mon || []).map(m => _normText(m.khoi)).filter(Boolean);
    }

    return Array.from(new Set(khois)).sort((a,b)=>a.localeCompare(b,'vi'));
}

function _findMonhocByIdOrTen(x){
    x = _normText(x);
    if (!x) return null;
    return (DATA.monhoc || []).find(m => _normText(m.id) === x)
        || (DATA.monhoc || []).find(m => _normText(m.ten) === x)
        || null;
}
function _monhocCandidates(mh){
    if (!mh) return [];
    const ten = _normText(mh.ten);
    const ma  = _normText(mh.ma);
    const ma2 = _normText(mh.ma2);
    return Array.from(new Set([ma, ma2, ten].filter(Boolean)));
}
function _monhocCanonicalKey(mh){
    if (!mh) return "";
    return _normText(mh.ma) || _normText(mh.ma2) || _normText(mh.ten) || "";
}

/**
 * ĐỒNG BỘ MÔN: check đã có đủ trong TIẾT CHUẨN (DATA.mon) cho mọi khối hay chưa.
 * - Ưu tiên dùng MÃ MÔN (ma/ma2). Nếu không có mã thì dùng TÊN.
 * - Không tạo trùng kiểu "Nhạc" và "Âm nhạc": chỉ cần 1 trong các biến thể (mã/tên) là xem như đã có.
 */
function isMonSyncedAcrossKhoi(monhocIdOrTen){
    const mh = _findMonhocByIdOrTen(monhocIdOrTen);
    const candidates = mh ? _monhocCandidates(mh) : [_normText(monhocIdOrTen)].filter(Boolean);
    if (!candidates.length) return false;

    const khois = getAllKhoiNamesForSync();
    if (!khois.length) return false;

    return khois.every(khoi =>
        (DATA.mon || []).some(m =>
            _normText(m.khoi) === _normText(khoi) && candidates.includes(_normText(m.ten))
        )
    );
}

/**
 * Đồng bộ Môn học (DATA.monhoc) -> Tiết chuẩn (DATA.mon)
 * - Lưu vào tiết chuẩn bằng MÃ MÔN (ma/ma2) để tránh trùng (ví dụ: "Nhạc" vs "Âm nhạc")
 * - Khi hiển thị, UI dùng resolveMonDisplay(...) để hiện TÊN MÔN.
 * - Nếu trong 1 khối đã có môn theo TÊN nhưng chưa theo MÃ, hàm sẽ "chuẩn hoá" về MÃ thay vì thêm dòng mới.
 */
function syncMonAcrossKhoi(monhocIdOrTen){
    const mh = _findMonhocByIdOrTen(monhocIdOrTen);

    const displayName = mh ? (_normText(mh.ten) || _normText(monhocIdOrTen)) : _normText(monhocIdOrTen);
    const candidates = mh ? _monhocCandidates(mh) : [_normText(monhocIdOrTen)].filter(Boolean);
    const canonicalKey = mh ? _monhocCanonicalKey(mh) : _normText(monhocIdOrTen);

    if (!canonicalKey) return;

    const khois = getAllKhoiNamesForSync();
    if (!khois.length) {
        alert("⚠ Chưa có danh sách Khối. Hãy nhập 'Khối học' trước.");
        return;
    }

    // Lấy template số tiết/giới hạn từ 1 khối đã có (nếu có) để đỡ nhập lại
    const tpl = (DATA.mon || []).find(m => candidates.includes(_normText(m.ten)) && _normText(m.khoi));

    let added = 0;
    let normalized = 0;
    let removedDup = 0;

    khois.forEach(khoi=>{
        const khoiNorm = _normText(khoi);

        // Tất cả dòng "môn" của khối này khớp theo mã/tên
        const rows = (DATA.mon || []).filter(m =>
            _normText(m.khoi) === khoiNorm && candidates.includes(_normText(m.ten))
        );

        if (rows.length === 0){
            // chưa có -> thêm mới (ten = mã)
            DATA.mon.push({
                id: autoID("mon"),
                khoi: khoi,
                ten: canonicalKey,
                sotiet: tpl ? (tpl.sotiet || "") : "",
                gioihan: tpl ? (tpl.gioihan || "") : "",
                ghichu: ""
            });
            added++;
            return;
        }

        // đã có -> chuẩn hoá về canonicalKey để tránh "Nhạc" & "Âm nhạc" cùng tồn tại
        // Nếu có nhiều dòng, giữ 1 dòng "tốt nhất", xoá dòng còn lại.
        // Tiêu chí giữ: ưu tiên dòng có ten==canonicalKey; nếu không có thì chọn dòng có nhiều dữ liệu sotiet/gioihan hơn.
        let keep = rows.find(r => _normText(r.ten) === canonicalKey) || rows[0];
        for (const r of rows){
            const score = (x)=> (String(x.sotiet||"").trim()?1:0) + (String(x.gioihan||"").trim()?1:0);
            if (r !== keep && score(r) > score(keep)) keep = r;
        }

        // merge: nếu keep thiếu sotiet/gioihan mà tpl có, bổ sung nhẹ
        if (tpl){
            if (!String(keep.sotiet||"").trim() && String(tpl.sotiet||"").trim()) keep.sotiet = tpl.sotiet;
            if (!String(keep.gioihan||"").trim() && String(tpl.gioihan||"").trim()) keep.gioihan = tpl.gioihan;
        }

        if (_normText(keep.ten) !== canonicalKey){
            keep.ten = canonicalKey;
            normalized++;
        }

        // remove dups beyond keep
        if (rows.length > 1){
            const keepId = keep.id;
            DATA.mon = (DATA.mon || []).filter(m => {
                if (_normText(m.khoi) !== khoiNorm) return true;
                if (!candidates.includes(_normText(m.ten))) return true;
                // loại các dòng trùng, giữ lại đúng keepId
                return (m.id === keepId);
            });
            removedDup += (rows.length - 1);
        }
    });

    saveStore();

    // refresh UI: nếu đang ở PCCM thì renderPCCM, ngược lại render lại môn học
    const sc = document.getElementById("section-content");
    if (sc && typeof renderPCCM === "function" && sc.innerHTML && sc.innerHTML.includes("PCCM")) {
        sc.innerHTML = renderPCCM();
    } else {
        renderSectionInto("monhoc","section-content",document);
    }

    let msg = `✔ Đồng bộ "${displayName}" xong.`;
    if (added) msg += ` Thêm mới: ${added} khối.`;
    if (normalized) msg += ` Chuẩn hoá: ${normalized} khối.`;
    if (removedDup) msg += ` Xoá trùng: ${removedDup} dòng.`;
    alert(msg);
}



// (đã bỏ tính năng dán từ Excel)




function renderPCCM_ByClass(classNames, monList){
    // Left: danh sách lớp (đã lọc theo khối ở renderPCCM)
    const left = `
    <div style="width:220px">
        <div style="font-weight:800;margin:6px 0 8px">Lớp</div>
        <div style="border:1px solid #e3e8f3;border-radius:8px;overflow:auto;max-height:520px;background:#fff">
            ${(classNames||[]).map(c=>`
                <div onclick="setPCCMSelectedClass('${escapeHtml(c)}')"
                     style="padding:10px 12px;cursor:pointer;${c===PCCM_SELECTED_CLASS?"background:#eef2ff;font-weight:800":""}">
                    ${escapeHtml(c)}
                </div>
            `).join("")}
        </div>
    </div>`;

    if (!PCCM_SELECTED_CLASS && (classNames||[]).length) PCCM_SELECTED_CLASS = classNames[0];
    if (PCCM_SELECTED_CLASS && !(classNames||[]).includes(PCCM_SELECTED_CLASS) && (classNames||[]).length) PCCM_SELECTED_CLASS = classNames[0];

    const cls = PCCM_SELECTED_CLASS;

    const classToKhoi = new Map((DATA.lop||[]).map(l=>{
        const canon = (l.ten2 || normalizeClassName(l.ten) || "").trim();
        const khoi = (l.khoi || ("Khối " + extractKhoiNumber(canon))).trim();
        return [canon, khoi];
    }).filter(x=>x[0]));
    const khoiName = classToKhoi.get(cls) || ("Khối " + extractKhoiNumber(cls));

    // Danh sách môn: lấy từ môn tổng hợp, lookup sotiet/gioihan từ tiết chuẩn theo khối
    // Chỉ HIỂN THỊ môn có số tiết > 0 trong tiết chuẩn (môn không có tiết sẽ ẩn)
    const mons = (monList||[])
        .map(m=>{
            const tc = lookupTietChuan(khoiName, m);
            return {...m, sotiet: tc ? (tc.sotiet||"") : "", gioihan: tc ? (tc.gioihan||"") : ""};
        })
        .filter(m=> _toPositiveNumberOrZero(m.sotiet) > 0)
        // Chỉ hiển thị môn đã phân công giáo viên (môn chưa phân công sẽ ẩn)
        .filter(m=> {
            const gv = (pccmGetTeacher(cls, m) || "").trim();
            return !!gv;
        });

    // Giữ nguyên thứ tự môn theo Bảng Môn (không sort theo GV)

    // Map GV (mã -> tên) để hiển thị tooltip
    const gvMap = new Map((DATA.giaovien||[]).map(g=>{
        const code = _normText(g.magv);
        const name = `${(g.hodem||"").trim()} ${(g.ten||"").trim()}`.trim();
        return [code.toLowerCase(), {code, name}];
    }).filter(x=>x[0]));

    // Cache cho nút Lưu (đọc DOM)
    PCCM_CLASS_EDIT_CACHE = { cls, khoiName, mons };

    // Teacher options
    const gvItems = (DATA.giaovien || [])
        .map(g=>{
            const code = _normText(g.magv);
            const name = `${(g.hodem||"").trim()} ${(g.ten||"").trim()}`.trim();
            return {code, name};
        })
        .filter(x=>x.code)
        .sort((a,b)=>a.code.localeCompare(b.code,'vi'));
    const gvCodesSet = new Set(gvItems.map(x=>x.code));

    // Room options (từ bảng Phòng) — label: "Tên phòng (Mã môn học)"
    const roomItems = getPhongOptionItems();
    const roomNames = roomItems.map(x=>x.value);
    const roomSet = new Set(roomNames);


    let mid = `
    <div style="flex:1;min-width:560px">
        <div class="table-wrap"><table class="pccm-table">
            <colgroup>
                <col style="width:60px">
                <col>
                <col>
                <col>
                <col>
                <col>
            </colgroup>
            <tr>
                <th>TT</th>
                <th>Môn học</th>
                <th>Giáo viên</th>
                <th>Phòng học</th>
                <th>Số tiết</th>
                <th>Giới hạn</th>
            </tr>`;

    mons.forEach((m, idx)=>{
        const raw = (pccmGetTeacher(cls,m) || "").trim();
        const resolved = resolveTeacherCode(raw);
        const val = resolved || raw; // giữ nguyên nếu không match

        // nếu giá trị hiện tại không nằm trong list GV, thêm option tạm để không bị mất dữ liệu
        const extraOpt = (val && !gvCodesSet.has(val))
            ? `<option value="${escapeHtml(val)}" selected>(Đang lưu) ${escapeHtml(val)}</option>`
            : "";

        const teacherOpts = [
            `<option value="" ${val===""?"selected":""}>(Chưa phân)</option>`,
            ...gvItems.map(x=>{
                // chỉ hiển thị mã GV
                const label = x.code;
                return `<option value="${escapeHtml(x.code)}" ${x.code===val?"selected":""}>${escapeHtml(label)}</option>`;
            })
        ].join("");

        const roomVal = (pccmGetRoom(cls,m) || "").trim();
        const extraRoomOpt = (roomVal && !roomSet.has(roomVal))
            ? `<option value="${escapeHtml(roomVal)}" selected>(Đang lưu) ${escapeHtml(roomVal)}</option>`
            : "";
        const roomOpts = [
            `<option value="" ${roomVal===""?"selected":""}>(Không chọn)</option>`,
            ...roomItems.map(it=>`<option value="${escapeHtml(it.value)}" ${it.value===roomVal?"selected":""}>${escapeHtml(it.label)}</option>`)
        ].join("");

        // nếu đã lưu override (từ lần phân công trước) thì hiển thị đúng giá trị
        let sotietDisp = pccmGetTiet(cls, m);
        if(sotietDisp === null || sotietDisp === undefined){
            sotietDisp = (m.sotiet ?? "");
        }
        let gioihanDisp = pccmGetGioihan(cls, m);
        if(gioihanDisp === null || gioihanDisp === undefined){
            gioihanDisp = (m.gioihan ?? "");
        }
        // default 1 (như yêu cầu) nếu còn trống / không hợp lệ
        const _stNum = Number(String(sotietDisp).trim());
        if(String(sotietDisp).trim()==="" || Number.isNaN(_stNum)) sotietDisp = "1";
        const _ghNum = Number(String(gioihanDisp).trim());
        if(String(gioihanDisp).trim()==="" || Number.isNaN(_ghNum)) gioihanDisp = "1";

        mid += `
        <tr>
            <td style="text-align:center">${idx+1}</td>
            <td>
                <select id="pccm_mon_${idx}" class="inline-edit-select" style="width:100%" data-kind="mon">
                    ${(function(){
                        const curV = (m.key||m.ten||'').toString();
                        const curN = _normText(curV);
                        const opts = (mons||[]).map(mm=>{
                            const v = (mm.key||mm.ten||'').toString();
                            const vn = _normText(v);
                            const label = (mm.ten || mm.key || v).toString();
                            return `<option value="${escapeHtml(v)}" ${vn===curN?"selected":""}>${escapeHtml(label)}</option>`;
                        }).join('');
                        return opts;
                    })()}
                </select>
            </td>
            <td>
                <select id="pccm_gv_${idx}" class="inline-edit-select" style="width:100%" onkeydown="pccmHandleGVCopyPaste(event, this)">
                    ${extraOpt}
                    ${teacherOpts}
                </select>
            </td>
            <td>
                <select id="pccm_room_${idx}" class="inline-edit-select" style="width:100%">
                    ${extraRoomOpt}
                    ${roomOpts}
                </select>
            </td>
            <td style="text-align:center">
                <input id="pccm_sotiet_${idx}" class="inline-edit-input" type="number" min="0" step="1"
                       value="${escapeHtml(sotietDisp)}" placeholder="trống hoặc >0" style="text-align:center">
            </td>
            <td style="text-align:center">
                <input id="pccm_gioihan_${idx}" class="inline-edit-input" type="number" min="0" step="1"
                       value="${escapeHtml(gioihanDisp)}" placeholder="trống hoặc >0" style="text-align:center">
            </td>
        </tr>`;
    });

    mid += `</table></div></div>`;

    const right = renderPCCMQuickBox(classNames, monList);

    return `<div style="display:flex;gap:12px;align-items:flex-start">${left}${mid}${right}</div>`;
}




function renderPCCM_ByTeacher(gvs, monList, classNames){
    const left = `
    <div style="width:220px">
        <div style="font-weight:800;margin:6px 0 8px">Giáo viên</div>
        <div style="border:1px solid #e3e8f3;border-radius:8px;overflow:auto;max-height:520px;background:#fff">
            ${(gvs||[]).map(g=>`
                <div onclick="setPCCMSelectedGV('${escapeHtml(g)}')"
                     style="padding:10px 12px;cursor:pointer;${g===PCCM_SELECTED_GV?"background:#eef2ff;font-weight:800":""}">
                    ${escapeHtml(g)}
                </div>
            `).join("")}
        </div>
    </div>`;

    if (!PCCM_SELECTED_GV && (gvs||[]).length) PCCM_SELECTED_GV = gvs[0];
    if (PCCM_SELECTED_GV && !(gvs||[]).includes(PCCM_SELECTED_GV) && (gvs||[]).length) PCCM_SELECTED_GV = gvs[0];

    const gv = PCCM_SELECTED_GV;

    // class -> khối để lookup tiết chuẩn
    const classToKhoi = new Map((DATA.lop||[]).map(l=>{
        const canon = canonTen2FromLop(l) || normalizeClassName(l.ten) || "";
        const khoi = (l.khoi || ("Khối " + extractKhoiNumber(canon))).trim();
        return [canon, khoi];
    }).filter(x=>x[0]));

    // Rooms (từ bảng Phòng). Hiển thị kèm "(Mã môn học)" nếu có.
    const roomItems = getPhongOptionItems();
    const roomNames = roomItems.map(x=>x.value);
    const roomSet = new Set(roomNames);

    // lọc theo lớp (dropdown) — mặc định "Tất cả"
    const classFilterOptions = ["Tất cả", ...(classNames||[])];
    if (!PCCM_TEACHER_CLASS_FILTER) PCCM_TEACHER_CLASS_FILTER = "Tất cả";
    if (!classFilterOptions.includes(PCCM_TEACHER_CLASS_FILTER)) PCCM_TEACHER_CLASS_FILTER = "Tất cả";

    const rowsAll = [];
    (classNames||[]).forEach(cls=>{
        const khoiName = classToKhoi.get(cls) || ("Khối " + extractKhoiNumber(cls));
        (monList||[]).forEach(m=>{
            const val = (pccmGetTeacher(cls,m) || "").trim();
            if (!val) return;
            if (val !== (gv||"").trim()) return;

            // Ẩn môn không có tiết
            if (!monHasPositiveTietChuanForKhoi(khoiName, m)) return;

            const tc = lookupTietChuan(khoiName, m);
            rowsAll.push({
                cls,
                khoiName,
                monObj: m,
                room: (pccmGetRoom(cls, m) || ""),
                sotiet: tc ? (tc.sotiet||"") : "",
                gioihan: tc ? (tc.gioihan||"") : ""
            });
        });
    });

    let rows = rowsAll;
    if (PCCM_TEACHER_CLASS_FILTER !== "Tất cả") rows = rowsAll.filter(r=>r.cls === PCCM_TEACHER_CLASS_FILTER);

    // Sort: theo Lớp, sau đó theo thứ tự Môn trong Bảng Môn
    const __monOrder = new Map((monList||[]).map((m,i)=>[_normText(m.key||m.ten), i]));
    rows.sort((a,b)=>{
        const cl = a.cls.localeCompare(b.cls,'vi');
        if (cl) return cl;
        const ia = __monOrder.get(_normText(a.monObj?.key || a.monObj?.ten)) ?? 1e9;
        const ib = __monOrder.get(_normText(b.monObj?.key || b.monObj?.ten)) ?? 1e9;
        if (ia !== ib) return ia - ib;
        return (a.monObj.ten||a.monObj.key||'').localeCompare((b.monObj.ten||b.monObj.key||''),'vi');
    });

    PCCM_TEACHER_EDIT_CACHE = { gv, rows };

    let mid = `
    <div style="flex:1;min-width:560px">
        <div style="display:flex;align-items:center;gap:10px;margin:6px 0 10px;flex-wrap:wrap">
            <div style="font-weight:900;font-size:20px">Giáo viên ${escapeHtml(gv || "")}</div>

            <div style="margin-left:auto;display:flex;gap:8px;align-items:center;flex-wrap:wrap">
                <select onchange="setPCCMTeacherClassFilter(this.value)" title="Lọc theo lớp">
                    ${classFilterOptions.map(c=>`<option value="${escapeHtml(c)}" ${c===PCCM_TEACHER_CLASS_FILTER?"selected":""}>${escapeHtml(c)}</option>`).join("")}
                </select>
            </div>
        </div>

        <div class="table-wrap"><table class="pccm-table">
            <colgroup>
                <col style="width:60px">
                <col>
                <col>
                <col>
                <col>
                <col>
            </colgroup>
            <tr>
                <th>TT</th>
                <th>Môn học</th>
                <th>Lớp</th>
                <th>Phòng học</th>
                <th>Số tiết</th>
                <th>Giới hạn</th>
            </tr>`;

    rows.forEach((r,idx)=>{
        const clsVal = (r.cls || "").trim();
        const monKey = _normText(r.monObj?.key || r.monObj?.ten);

        // options Lớp
        const clsSet = new Set(classNames || []);
        const extraClsOpt = (clsVal && !clsSet.has(clsVal))
            ? `<option value="${escapeHtml(clsVal)}" selected>(Đang lưu) ${escapeHtml(clsVal)}</option>`
            : "";
        const clsOpts = (classNames||[])
            .map(c=>`<option value="${escapeHtml(c)}" ${c===clsVal?"selected":""}>${escapeHtml(c)}</option>`)
            .join("");

        // options Môn (lọc theo lớp)
        const allowedMons = pccmGetAllowedMonsForClass(clsVal, (monList||[]));
        const allowedKeys = new Set(allowedMons.map(m => (m.key||m.ten||"")));
        const extraMonOpt = (monKey && !allowedKeys.has(monKey))
            ? `<option value="${escapeHtml(monKey)}" selected>(Đang lưu) ${escapeHtml(monKey)}</option>`
            : "";
        const monOpts = allowedMons.map(m=>{
            const k = (m.key||m.ten||"").toString();
            const label = (m.ten || m.code || m.ma || m.key || '').toString();
            return `<option value="${escapeHtml(k)}" ${k===monKey?"selected":""}>${escapeHtml(label)}</option>`;
        }).join("");

        const roomVal = (r.room||"").trim();
        const extraRoomOpt = (roomVal && !roomSet.has(roomVal))
            ? `<option value="${escapeHtml(roomVal)}" selected>(Đang lưu) ${escapeHtml(roomVal)}</option>`
            : "";
        const roomOpts = [
            `<option value="" ${roomVal===""?"selected":""}>(Không chọn)</option>`,
            ...roomItems.map(it=>`<option value="${escapeHtml(it.value)}" ${it.value===roomVal?"selected":""}>${escapeHtml(it.label)}</option>`)
        ].join("");

        mid += `
        <tr>
            <td style="text-align:center">${idx+1}</td>
            <td>
                <select id="pccmT_mon_${idx}" class="inline-edit-select" style="width:100%" data-kind="mon">
                    ${extraMonOpt}${monOpts}
                </select>
            </td>
            <td>
                <select id="pccmT_cls_${idx}" class="inline-edit-select" style="width:100%" data-kind="cls" data-monselect="pccmT_mon_${idx}"
                        onchange="pccmRowClassChanged(this)">
                    ${extraClsOpt}${clsOpts}
                </select>
            </td>
            <td>
                <select id="pccmT_room_${idx}" class="inline-edit-select" style="width:100%">
                    ${extraRoomOpt}
                    ${roomOpts}
                </select>
            </td>
            <td style="text-align:center">
                <input id="pccmT_sotiet_${idx}" class="inline-edit-input" type="number" min="0" step="1"
                       value="${escapeHtml(r.sotiet||"")}" placeholder="trống hoặc >0" style="text-align:center">
            </td>
            <td style="text-align:center">
                <input id="pccmT_gioihan_${idx}" class="inline-edit-input" type="number" min="0" step="1"
                       value="${escapeHtml(r.gioihan||"")}" placeholder="trống hoặc >0" style="text-align:center">
            </td>
        </tr>`;
    });

    if (!rows.length){
        mid += `<tr><td colspan="6" style="padding:14px;color:#666">Không có dữ liệu phù hợp.</td></tr>`;
    }

    mid += `</table></div></div>`;

    const right = renderPCCMQuickBox(classNames, monList);
    return `<div style="display:flex;gap:12px;align-items:flex-start">${left}${mid}${right}</div>`;
}



let PCCM_SUBJ = ""; // môn (key) đang chọn trong tab Môn học

function setPCCMSelectedSubject(monKey){
    PCCM_SUBJ = monKey || "";
    const sc = document.getElementById("section-content");
    if (sc) sc.innerHTML = renderPCCM();
}

// Tab "Môn học": Lớp học, Giáo viên, Phòng học (nếu có), Số tiết, Giới hạn
function renderPCCM_BySubject(classNames, monList){
    const subjects = (monList || [])
        .map(m=>({
            key: _normText(m.key || m.ten),
            name: _normText(m.ten || m.code || m.ma || m.key)
        }))
        .filter(s=>s.key);

    if (!PCCM_SUBJ && subjects.length) PCCM_SUBJ = subjects[0].key;
    if (PCCM_SUBJ && !subjects.some(s=>s.key===PCCM_SUBJ) && subjects.length) PCCM_SUBJ = subjects[0].key;

    const monObj = (monList || []).find(m=>_normText(m.key||m.ten) === _normText(PCCM_SUBJ))
        || { key: PCCM_SUBJ, ten: PCCM_SUBJ };

    // Left: list môn
    const left = `
    <div style="width:220px">
        <div style="font-weight:800;margin:6px 0 8px">Môn học</div>
        <div style="border:1px solid #e3e8f3;border-radius:8px;overflow:auto;max-height:520px;background:#fff">
            ${subjects.map(s=>{
                const js = (s.key||"").replace(/\\/g,"\\\\").replace(/'/g,"\\'");
                return `
                <div onclick="setPCCMSelectedSubject('${js}')"
                     style="padding:10px 12px;cursor:pointer;${s.key===PCCM_SUBJ?"background:#eef2ff;font-weight:800":""}">
                    ${escapeHtml(s.name)}
                </div>`;
            }).join("")}
        </div>
    </div>`;

    // Teachers (mã GV)
    const gvCodes = (DATA.giaovien || [])
        .map(g=>_normText(g?.magv))
        .filter(Boolean)
        .sort((a,b)=>a.localeCompare(b,'vi'));
    const gvSet = new Set(gvCodes);

    // Rooms (từ bảng Phòng). Hiển thị kèm "(Mã môn học)" nếu có.
    const roomItems = getPhongOptionItems();
    const roomNames = roomItems.map(x=>x.value);
    const roomSet = new Set(roomNames);

    // class -> khối để lookup tiết chuẩn
    const classToKhoi = new Map((DATA.lop||[]).map(l=>{
        const canon = canonTen2FromLop(l) || normalizeClassName(l.ten) || "";
        const khoi = (l.khoi || ("Khối " + extractKhoiNumber(canon))).trim();
        return [canon, khoi];
    }).filter(x=>x[0]));

    // Chỉ hiển thị các lớp mà môn này có số tiết > 0 (tiết chuẩn)
    const rows = [];
    (classNames||[]).forEach(cls=>{
        const khoiName = classToKhoi.get(cls) || ("Khối " + extractKhoiNumber(cls));
        if (!monHasPositiveTietChuanForKhoi(khoiName, monObj)) return;
        const tc = lookupTietChuan(khoiName, monObj);
        rows.push({
            cls,
            khoiName,
            gv: _normText(pccmGetTeacher(cls, monObj)),
            room: _normText(pccmGetRoom(cls, monObj)),
            sotiet: tc ? (tc.sotiet||"") : "",
            gioihan: tc ? (tc.gioihan||"") : ""
        });
    });

    PCCM_SUBJECT_EDIT_CACHE = { monKey: _normText(monObj.key||monObj.ten), monObj, rows };

    let mid = `
    <div style="flex:1;min-width:720px">
        <div style="display:flex;align-items:center;gap:10px;margin:6px 0 10px;flex-wrap:wrap">
            <div style="font-weight:900;font-size:20px">Môn ${escapeHtml(monObj.ten || monObj.key || "")}</div>
        </div>

        <div class="table-wrap"><table class="pccm-table">
            <colgroup>
                <col style="width:60px">
                <col>
                <col>
                <col>
                <col>
                <col>
            </colgroup>
            <tr>
                <th>TT</th>
                <th>Lớp học</th>
                <th>Giáo viên</th>
                <th>Phòng học</th>
                <th>Số tiết</th>
                <th>Giới hạn</th>
            </tr>`;

    rows.forEach((r, idx)=>{
        const gvVal = r.gv;
        const extraGvOpt = (gvVal && !gvSet.has(gvVal))
            ? `<option value="${escapeHtml(gvVal)}" selected>(Đang lưu) ${escapeHtml(gvVal)}</option>`
            : "";
        const gvOpts = [
            `<option value="" ${gvVal===""?"selected":""}>(Chưa phân)</option>`,
            ...gvCodes.map(code=>`<option value="${escapeHtml(code)}" ${code===gvVal?"selected":""}>${escapeHtml(code)}</option>`)
        ].join("");

        const roomVal = r.room;
        const extraRoomOpt = (roomVal && !roomSet.has(roomVal))
            ? `<option value="${escapeHtml(roomVal)}" selected>(Đang lưu) ${escapeHtml(roomVal)}</option>`
            : "";
        const roomOpts = [
            `<option value="" ${roomVal===""?"selected":""}>(Không chọn)</option>`,
            ...roomItems.map(it=>`<option value="${escapeHtml(it.value)}" ${it.value===roomVal?"selected":""}>${escapeHtml(it.label)}</option>`)
        ].join("");

        mid += `
        <tr>
            <td style="text-align:center">${idx+1}</td>
            <td>
                <select id="pccmS_cls_${idx}" class="inline-edit-select" style="width:100%" data-kind="cls">
                    ${(classNames||[]).map(c=>`<option value="${escapeHtml(c)}" ${c===r.cls?"selected":""}>${escapeHtml(c)}</option>`).join("")}
                </select>
            </td>
            <td>
                <select id="pccmS_gv_${idx}" class="inline-edit-select" style="width:100%" onkeydown="pccmHandleGVCopyPaste(event, this)">
                    ${extraGvOpt}
                    ${gvOpts}
                </select>
            </td>
            <td>
                <select id="pccmS_room_${idx}" class="inline-edit-select" style="width:100%">
                    ${extraRoomOpt}
                    ${roomOpts}
                </select>
            </td>
            <td style="text-align:center">
                <input id="pccmS_sotiet_${idx}" class="inline-edit-input" type="number" min="0" step="1"
                       value="${escapeHtml(r.sotiet||"")}" placeholder="trống hoặc >0" style="text-align:center">
            </td>
            <td style="text-align:center">
                <input id="pccmS_gioihan_${idx}" class="inline-edit-input" type="number" min="0" step="1"
                       value="${escapeHtml(r.gioihan||"")}" placeholder="trống hoặc >0" style="text-align:center">
            </td>
        </tr>`;
    });

    if (!rows.length){
        mid += `<tr><td colspan="6" style="padding:14px;color:#666">Không có dữ liệu phù hợp.</td></tr>`;
    }

    mid += `</table></div></div>`;

    const right = renderPCCMQuickBox(classNames, monList);
    return `<div style="display:flex;gap:12px;align-items:flex-start">${left}${mid}${right}</div>`;
}


// Tab "Phòng": hiển thị theo phòng (lọc phòng) và cho phép chỉnh GV/Phòng/Tiết chuẩn
function renderPCCM_ByRoom(classNames, monList){
    const roomItems = getPhongOptionItems();

    // "" = chưa phân phòng
    const rooms = ["", ...roomItems.map(x=>x.value)];
    const roomLabelMap = new Map([
        ["", "(Chưa phân)"],
        ...roomItems.map(x=>[x.value, x.label])
    ]);

    if (PCCM_SELECTED_ROOM === undefined || PCCM_SELECTED_ROOM === null) PCCM_SELECTED_ROOM = "";
    if (!rooms.includes(PCCM_SELECTED_ROOM)) PCCM_SELECTED_ROOM = rooms[0] || "";

    // Left: list phòng
    const left = `
    <div style="width:220px">
        <div style="font-weight:800;margin:6px 0 8px">Phòng</div>
        <div style="border:1px solid #e3e8f3;border-radius:8px;overflow:auto;max-height:520px;background:#fff">
            ${rooms.map(r=>{
                const label = roomLabelMap.get(r) || (r||"");
                const js = (r||"").replace(/\\/g,"\\\\").replace(/'/g,"\\'");
                return `
                <div onclick="setPCCMSelectedRoom('${js}')"
                     style="padding:10px 12px;cursor:pointer;${r===PCCM_SELECTED_ROOM?"background:#eef2ff;font-weight:800":""}">
                    ${escapeHtml(label)}
                </div>`;
            }).join("")}
        </div>
    </div>`;

    // Teachers (mã GV)
    const gvCodes = (DATA.giaovien || [])
        .map(g=>_normText(g?.magv))
        .filter(Boolean)
        .sort((a,b)=>a.localeCompare(b,'vi'));
    const gvSet = new Set(gvCodes);

    // class -> khối để lookup tiết chuẩn
    const classToKhoi = new Map((DATA.lop||[]).map(l=>{
        const canon = canonTen2FromLop(l) || normalizeClassName(l.ten) || "";
        const khoi = (l.khoi || ("Khối " + extractKhoiNumber(canon))).trim();
        return [canon, khoi];
    }).filter(x=>x[0]));

    const selRoom = (PCCM_SELECTED_ROOM || "").trim();

    const rows = [];
    (classNames||[]).forEach(cls=>{
        const khoiName = classToKhoi.get(cls) || ("Khối " + extractKhoiNumber(cls));
        (monList||[]).forEach(m=>{
            if (!monHasPositiveTietChuanForKhoi(khoiName, m)) return;
            const roomVal = (pccmGetRoom(cls, m) || "").trim();
            if ((roomVal || "") !== selRoom) return;

            const tc = lookupTietChuan(khoiName, m);
            rows.push({
                cls,
                khoiName,
                monObj: m,
                gv: (pccmGetTeacher(cls, m) || "").trim(),
                room: roomVal,
                sotiet: tc ? (tc.sotiet||"") : "",
                gioihan: tc ? (tc.gioihan||"") : ""
            });
        });
    });

    // Sort: theo Lớp, sau đó theo thứ tự Môn trong Bảng Môn
    const __monOrderR = new Map((monList||[]).map((m,i)=>[_normText(m.key||m.ten), i]));
    rows.sort((a,b)=>{
        const cl = a.cls.localeCompare(b.cls,'vi');
        if (cl) return cl;
        const ia = __monOrderR.get(_normText(a.monObj?.key || a.monObj?.ten)) ?? 1e9;
        const ib = __monOrderR.get(_normText(b.monObj?.key || b.monObj?.ten)) ?? 1e9;
        if (ia !== ib) return ia - ib;
        const am = (a.monObj.ten || a.monObj.key || '');
        const bm = (b.monObj.ten || b.monObj.key || '');
        return am.localeCompare(bm,'vi');
    });

    PCCM_ROOM_EDIT_CACHE = { room: selRoom, rows };

    const roomTitle = roomLabelMap.get(selRoom) || (selRoom || "(Chưa phân)");

    let mid = `
    <div style="flex:1;min-width:760px">
        <div style="display:flex;align-items:center;gap:10px;margin:6px 0 10px;flex-wrap:wrap">
            <div style="font-weight:900;font-size:20px">Phòng ${escapeHtml(roomTitle)}</div>
        </div>

        <div class="table-wrap"><table class="pccm-table">
            <colgroup>
                <col style="width:60px">
                <col>
                <col>
                <col>
                <col>
                <col>
                <col>
            </colgroup>
            <tr>
                <th>TT</th>
                <th>Môn học</th>
                <th>Lớp học</th>
                <th>Giáo viên</th>
                <th>Phòng</th>
                <th>Số tiết</th>
                <th>Giới hạn</th>
            </tr>`;

    rows.forEach((r, idx)=>{
        const clsVal = (r.cls || "").trim();
        const monKey = _normText(r.monObj?.key || r.monObj?.ten);
        const monLabel = (r.monObj.code || r.monObj.ma || r.monObj.key || r.monObj.ten || "");

        const gvVal = (r.gv || "").trim();
        const extraGvOpt = (gvVal && !gvSet.has(gvVal))
            ? `<option value="${escapeHtml(gvVal)}" selected>(Đang lưu) ${escapeHtml(gvVal)}</option>`
            : "";
        const gvOpts = [
            `<option value="" ${gvVal===""?"selected":""}>(Chưa phân)</option>`,
            ...gvCodes.map(code=>`<option value="${escapeHtml(code)}" ${code===gvVal?"selected":""}>${escapeHtml(code)}</option>`)
        ].join("");

        const roomVal = (r.room || "").trim();
        const extraRoomOpt = (roomVal && !rooms.includes(roomVal))
            ? `<option value="${escapeHtml(roomVal)}" selected>(Đang lưu) ${escapeHtml(roomVal)}</option>`
            : "";
        const roomOpts = [
            `<option value="" ${roomVal===""?"selected":""}>(Không chọn)</option>`,
            ...roomItems.map(it=>`<option value="${escapeHtml(it.value)}" ${it.value===roomVal?"selected":""}>${escapeHtml(it.label)}</option>`)
        ].join("");

        // options Lớp
        const clsSet = new Set(classNames || []);
        const extraClsOpt = (clsVal && !clsSet.has(clsVal))
            ? `<option value="${escapeHtml(clsVal)}" selected>(Đang lưu) ${escapeHtml(clsVal)}</option>`
            : "";
        const clsOpts = (classNames||[])
            .map(c=>`<option value="${escapeHtml(c)}" ${c===clsVal?"selected":""}>${escapeHtml(c)}</option>`)
            .join("");

        // options Môn (lọc theo lớp)
        const allowedMons = pccmGetAllowedMonsForClass(clsVal, (monList||[]));
        const allowedKeys = new Set(allowedMons.map(m => (m.key||m.ten||"")));
        const extraMonOpt = (monKey && !allowedKeys.has(monKey))
            ? `<option value="${escapeHtml(monKey)}" selected>(Đang lưu) ${escapeHtml(monKey)}</option>`
            : "";
        const monOpts = allowedMons.map(m=>{
            const k = (m.key||m.ten||"").toString();
            const label = (m.ten || m.code || m.ma || m.key || '').toString();
            return `<option value="${escapeHtml(k)}" ${k===monKey?"selected":""}>${escapeHtml(label)}</option>`;
        }).join("");

        mid += `
        <tr>
            <td style="text-align:center">${idx+1}</td>
            <td>
                <input id="pccmR_mon_in_${idx}" class="pccm-dblsel-input" value="${escapeHtml(r.monObj.ten || r.monObj.key || monLabel || "")}"
                       data-selectid="pccmR_mon_${idx}" data-kind="mon"
                       onclick="pccmInputClick(event,this,'pccmR_mon_${idx}')"
                       onkeydown="pccmDblInputKeyDown(event,this)" onblur="pccmDblInputCommit(this)">
                <select id="pccmR_mon_${idx}" class="inline-edit-select pccm-dblsel-select" style="display:none" disabled
                        data-kind="mon" data-inputid="pccmR_mon_in_${idx}" data-class-select="pccmR_cls_${idx}" data-size="10"
                        onchange="pccmDblSelClose('pccmR_mon_${idx}', true)" onblur="pccmDblSelClose('pccmR_mon_${idx}', true)">
                    ${extraMonOpt}${monOpts}
                </select>
            </td>
            <td>
                <input id="pccmR_cls_in_${idx}" class="pccm-dblsel-input" value="${escapeHtml(clsVal)}"
                       data-selectid="pccmR_cls_${idx}" data-kind="cls"
                       onclick="pccmInputClick(event,this,'pccmR_cls_${idx}')"
                       onkeydown="pccmDblInputKeyDown(event,this)" onblur="pccmDblInputCommit(this)">
                <select id="pccmR_cls_${idx}" class="inline-edit-select pccm-dblsel-select" style="display:none" disabled
                        data-kind="cls" data-inputid="pccmR_cls_in_${idx}" data-monselect="pccmR_mon_${idx}" data-size="10"
                        onchange="pccmRowClassChanged(this); pccmDblSelClose('pccmR_cls_${idx}', true)" onblur="pccmDblSelClose('pccmR_cls_${idx}', true)">
                    ${extraClsOpt}${clsOpts}
                </select>
            </td>
            <td>
                <select id="pccmR_gv_${idx}" class="inline-edit-select" style="width:100%" onkeydown="pccmHandleGVCopyPaste(event, this)">
                    ${extraGvOpt}
                    ${gvOpts}
                </select>
            </td>
            <td>
                <select id="pccmR_room_${idx}" class="inline-edit-select" style="width:100%">
                    ${extraRoomOpt}
                    ${roomOpts}
                </select>
            </td>
            <td style="text-align:center">
                <input id="pccmR_sotiet_${idx}" class="inline-edit-input" type="number" min="0" step="1"
                       value="${escapeHtml(r.sotiet||"")}" placeholder="trống hoặc >0" style="text-align:center">
            </td>
            <td style="text-align:center">
                <input id="pccmR_gioihan_${idx}" class="inline-edit-input" type="number" min="0" step="1"
                       value="${escapeHtml(r.gioihan||"")}" placeholder="trống hoặc >0" style="text-align:center">
            </td>
        </tr>`;
    });

    if (!rows.length){
        mid += `<tr><td colspan="7" style="padding:14px;color:#666">Không có dữ liệu phù hợp.</td></tr>`;
    }

    mid += `</table></div></div>`;

    const right = renderPCCMQuickBox(classNames, monList);
    return `<div style="display:flex;gap:12px;align-items:flex-start">${left}${mid}${right}</div>`;
}


function pccmSaveRoomEdits(){
    const rows = PCCM_ROOM_EDIT_CACHE?.rows || [];
    if (!Array.isArray(rows) || !rows.length){
        alert("⚠ Không có dữ liệu để lưu.");
        return;
    }

    if (typeof DATA.pccmMatrix !== "object" || !DATA.pccmMatrix) DATA.pccmMatrix = {};
    if (typeof DATA.pccmRoomMatrix !== "object" || !DATA.pccmRoomMatrix) DATA.pccmRoomMatrix = {};

    // map lớp -> khối (để update tiết chuẩn đúng khối sau khi đổi lớp)
    const classToKhoi = new Map((DATA.lop||[]).map(l=>{
        const canon = canonTen2FromLop(l) || normalizeClassName(l.ten) || "";
        const khoi = (l.khoi || ("Khối " + extractKhoiNumber(canon))).trim();
        return [canon.trim(), khoi];
    }).filter(x=>x[0]));

    // map mônKey -> monObj
    const allMons = buildPCCMMonList();
    const monMap = new Map(allMons.map(m=>[_normText(m.key||m.ten), m]));
    function _findMonObj(monKey){
        const k = _normText(monKey);
        return monMap.get(k)
            || allMons.find(m=>_normText(m.ten) === k)
            || { key: k, ten: k, code:"", ma:"", ma2:"" };
    }

    function _validateNumber(v){
        const s = (v ?? "").toString().trim();
        if (s === "") return { ok:true, val:"" };
        const n = Number(s);
        if (!Number.isFinite(n)) return { ok:false };
        if (n <= 0) return { ok:false };
        return { ok:true, val: Number.isInteger(n) ? String(n) : String(n) };
    }

    function _setTeacherNoSave(lopCanon, monObj, val){
        if (!lopCanon || !monObj) return;
        val = (val || "").trim();
        const primaryKey = `${lopCanon}|${(monObj.key || monObj.ten || "").trim()}`;
        if (val) DATA.pccmMatrix[primaryKey] = val;
        else delete DATA.pccmMatrix[primaryKey];

        const legacyKeys = [];
        if (monObj.ten && monObj.key && monObj.ten !== monObj.key) legacyKeys.push(`${lopCanon}|${monObj.ten}`);
        if (monObj.ma2 && monObj.key && monObj.ma2 !== monObj.key) legacyKeys.push(`${lopCanon}|${monObj.ma2}`);
        if (monObj.ma && monObj.key && monObj.ma !== monObj.key) legacyKeys.push(`${lopCanon}|${monObj.ma}`);
        legacyKeys.forEach(k=>{ if (k !== primaryKey && DATA.pccmMatrix[k]) delete DATA.pccmMatrix[k]; });
    }

    function _setRoomNoSave(lopCanon, monObj, val){
        if (!lopCanon || !monObj) return;
        val = (val || "").trim();
        const primaryKey = `${lopCanon}|${(monObj.key || monObj.ten || "").trim()}`;
        if (val) DATA.pccmRoomMatrix[primaryKey] = val;
        else delete DATA.pccmRoomMatrix[primaryKey];

        const legacyKeys = [];
        if (monObj.ten && monObj.key && monObj.ten !== monObj.key) legacyKeys.push(`${lopCanon}|${monObj.ten}`);
        if (monObj.ma2 && monObj.key && monObj.ma2 !== monObj.key) legacyKeys.push(`${lopCanon}|${monObj.ma2}`);
        if (monObj.ma && monObj.key && monObj.ma !== monObj.key) legacyKeys.push(`${lopCanon}|${monObj.ma}`);
        legacyKeys.forEach(k=>{ if (k !== primaryKey && DATA.pccmRoomMatrix[k]) delete DATA.pccmRoomMatrix[k]; });
    }

    // pass 1: đọc DOM + check trùng khóa
    const ops = [];
    const seenKeys = new Set();
    for (let idx=0; idx<rows.length; idx++){
        const r = rows[idx];

        const newCls = (document.getElementById(`pccmR_cls_${idx}`)?.value || r.cls || "").toString().trim();
        const newMonKey = (document.getElementById(`pccmR_mon_${idx}`)?.value || (r.monObj?.key||r.monObj?.ten||""))
            .toString().trim();
        if (!newCls || !newMonKey){
            alert(`⚠ Dữ liệu Lớp/Môn không hợp lệ ở dòng ${idx+1}.`);
            return;
        }
        const newMonObj = _findMonObj(newMonKey);
        const newKey = `${newCls}|${(newMonObj.key || newMonObj.ten || "").trim()}`;
        if (seenKeys.has(newKey)){
            alert(`⚠ Trùng phân công (Lớp|Môn) ở dòng ${idx+1}: ${newKey}.\nHãy kiểm tra lại vì 1 Lớp + 1 Môn chỉ nên có 1 dòng.`);
            return;
        }
        seenKeys.add(newKey);

        ops.push({
            idx,
            r,
            oldCls: (r.cls||"").trim(),
            oldMonObj: r.monObj,
            newCls,
            newMonObj,
            gv: (document.getElementById(`pccmR_gv_${idx}`)?.value || "").toString().trim(),
            room: (document.getElementById(`pccmR_room_${idx}`)?.value || "").toString().trim(),
            sotietRaw: (document.getElementById(`pccmR_sotiet_${idx}`)?.value || "").toString().trim(),
            gioihanRaw: (document.getElementById(`pccmR_gioihan_${idx}`)?.value || "").toString().trim()
        });
    }

    // clear các key cũ nếu người dùng đổi Lớp/Môn
    const cleared = new Set();
    ops.forEach(op=>{
        const oldKey = `${op.oldCls}|${(op.oldMonObj?.key || op.oldMonObj?.ten || "").toString().trim()}`;
        const newKey = `${op.newCls}|${(op.newMonObj?.key || op.newMonObj?.ten || "").toString().trim()}`;
        if (oldKey && newKey && oldKey !== newKey && !cleared.has(oldKey)){
            cleared.add(oldKey);
            _setTeacherNoSave(op.oldCls, op.oldMonObj, "");
            _setRoomNoSave(op.oldCls, op.oldMonObj, "");
        }
    });

    // apply
    for (const op of ops){
        const khoiName = classToKhoi.get(op.newCls) || ("Khối " + extractKhoiNumber(op.newCls));

        _setTeacherNoSave(op.newCls, op.newMonObj, op.gv);
        _setRoomNoSave(op.newCls, op.newMonObj, op.room);

        const v1 = _validateNumber(op.sotietRaw);
        const v2 = _validateNumber(op.gioihanRaw);
        if (!v1.ok || !v2.ok){
            alert(`⚠ Dữ liệu không hợp lệ ở: ${(op.newMonObj.ten||op.newMonObj.key||"")} (${khoiName}).\n- Chỉ được để trống hoặc nhập số > 0.`);
            return;
        }

        // Lưu số tiết / giới hạn theo Lớp|Môn (nguồn cho màn "Sắp xếp TKB")
        pccmSetTietGioihanNoSave(op.newCls, op.newMonObj, v1.val, v2.val);

        let tcRow = lookupTietChuan(khoiName, op.newMonObj);
        if (tcRow){
            tcRow.sotiet = v1.val;
            tcRow.gioihan = v2.val;
        } else {
            if (v1.val !== "" || v2.val !== ""){
                DATA.mon = DATA.mon || [];
                DATA.mon.push({
                    id: autoID("mon"),
                    khoi: khoiName,
                    ten: (op.newMonObj.ten || op.newMonObj.key || "").trim(),
                    sotiet: v1.val,
                    gioihan: v2.val || "",
                    ghichu: ""
                });
            }
        }
    }

    saveStore();
    const sc = document.getElementById("section-content");
    if (sc && typeof renderPCCM === "function") sc.innerHTML = renderPCCM();
    alert("✔ Đã lưu Phân công.");
}





function renderPCCM_TongHop(lops, mons){
    // mons: danh mục môn tổng hợp [{key,ten,ma,ma2}]
    const monsSorted = (mons || []).slice().sort((a,b)=>{
        const aHas = (lops||[]).some(l=>!!pccmGetTeacher(l,a));
        const bHas = (lops||[]).some(l=>!!pccmGetTeacher(l,b));
        if (aHas !== bHas) return aHas ? -1 : 1;
        return (a.ten||a.key||"").localeCompare((b.ten||b.key||""),'vi');
    });

    let html = `
    <div class="table-wrap">
    <table class="pccm-table">
        <tr><th>Lớp / Môn</th>`;

    monsSorted.forEach(m => html += `<th>${renderPCCMMonHeader(m)}</th>`);
    html += `</tr>`;

    (lops||[]).forEach(l => {
        html += `<tr><td><b>${escapeHtml(l)}</b></td>`;
        monsSorted.forEach(m => {
            const val = pccmGetTeacher(l, m);
            html += `
            <td>
                <input value="${escapeHtml(val)}"
                       data-monkey="${escapeHtml(m.key||"")}"
                       data-monten="${escapeHtml(m.ten||"")}"
                       data-monma="${escapeHtml(m.ma||"")}"
                       data-monma2="${escapeHtml(m.ma2||"")}"
                       oninput="pccmSetTeacherFromInput(this,'${escapeHtml(l)}')">
            </td>`;
        });
        html += `</tr>`;
    });

    html += `</table></div>`;
    return html;
}


/* ============================================================
   XÓA TOÀN BỘ PCCM
============================================================ */


function renderPCCM_TietChuan(monRows){
    const rows = (monRows || []).map(m=>({
        id: (m.id || "").toString().trim(),
        khoi: (m.khoi || "").trim(),
        ten: (m.ten || "").trim(),
        sotiet: (m.sotiet || "").toString().trim(),
        gioihan: (m.gioihan || "").toString().trim()
    })).filter(r => r.ten && r.khoi);

    const monNames = Array.from(new Set(rows.map(r=>r.ten))).sort((a,b)=>a.localeCompare(b,'vi'));
    if (!TC_MON) TC_MON = "Tất cả";
    if (TC_MON !== "Tất cả" && !monNames.includes(TC_MON)) TC_MON = "Tất cả";

    const filtered = (TC_MON === "Tất cả") ? rows : rows.filter(r => r.ten === TC_MON);

    let mid = `
    <div style="flex:1;min-width:760px">
        <div style="display:flex;align-items:center;gap:10px;margin:6px 0 10px;flex-wrap:wrap">
            <div style="font-weight:900;font-size:20px">Tiết chuẩn</div>

            <div style="margin-left:auto;display:flex;align-items:center;gap:8px;flex-wrap:wrap">
                <span style="font-weight:700">Môn</span>
                <select onchange="setPCCMTCMon(this.value)">
                    <option value="Tất cả"${TC_MON==="Tất cả"?" selected":""}>(Chọn tất cả)</option>
                    ${monNames.map(n=>`<option value="${escapeHtml(n)}"${n===TC_MON?" selected":""}>${escapeHtml(n)}</option>`).join("")}
                </select>

                <button class="btn" onclick="triggerExcel('mon')">Nhập Excel</button>
                <button class="btn" onclick="exportExcel('mon')">Xuất Excel</button>
                <button class="btn primary" onclick="tcSaveAllEdits()">Lưu</button>
            </div>
        </div>

        <div style="font-size:12px;color:#667;margin:0 0 10px;line-height:1.35">
            * Sửa trực tiếp <b>Số tiết</b> và <b>Giới hạn</b> trên bảng rồi bấm <b>Lưu</b> (giống Excel).<br>
            * Ô trống được phép. Nếu nhập số thì phải <b>&gt; 0</b> (không âm).
        </div>

        <div class="table-wrap">
        <table>
            <tr>
                <th style="width:60px">TT</th>
                <th style="width:140px">Khối học</th>
                <th>Môn học</th>
                <th style="width:160px">Số tiết/1 tuần</th>
                <th style="width:210px">Giới hạn số tiết/1 buổi</th>
            </tr>`;

    filtered.forEach((r,i)=>{
        const rid = r.id || `${r.khoi}__${r.ten}`;
        mid += `
        <tr>
            <td style="text-align:center">${i+1}</td>
            <td>${escapeHtml(r.khoi)}</td>
            <td>${escapeHtml(r.ten)}</td>
            <td style="text-align:center">
                <input class="inline-edit-input tc-edit" type="number" min="0" step="1"
                       data-rowid="${escapeHtml(rid)}" data-field="sotiet" value="${escapeHtml(r.sotiet)}"
                       placeholder="(trống hoặc >0)" style="text-align:center">
            </td>
            <td style="text-align:center">
                <input class="inline-edit-input tc-edit" type="number" min="0" step="1"
                       data-rowid="${escapeHtml(rid)}" data-field="gioihan" value="${escapeHtml(r.gioihan)}"
                       placeholder="(trống hoặc >0)" style="text-align:center">
            </td>
        </tr>`;
    });

    if (!filtered.length){
        mid += `<tr><td colspan="5" style="padding:14px;color:#666">Chưa có dữ liệu tiết chuẩn.</td></tr>`;
    }

    mid += `</table></div></div>`;

    return `<div style="display:flex;gap:12px">${mid}</div>`;
}

// Lưu toàn bộ chỉnh sửa Tiết chuẩn (inline edit)
function tcSaveAllEdits(){
    // Hỗ trợ 2 kiểu UI:
    // 1) input.tc-edit (cũ)
    // 2) ô text .tc-cell (mới) + input chỉ xuất hiện khi click sửa
    const byId = {};

    const cells = Array.from(document.querySelectorAll(".tc-cell[data-rowid][data-field]"));
    if (cells.length){
        cells.forEach(td=>{
            const id = (td.dataset.rowid || "").toString();
            const field = (td.dataset.field || "").toString();
            if (!id || !field) return;
            if (!byId[id]) byId[id] = {};

            // Nếu đang sửa (có input) thì lấy value từ input, ngược lại lấy từ data-val
            const inp = td.querySelector("input");
            const val = inp ? (inp.value ?? "") : (td.dataset.val ?? "");
            byId[id][field] = (val ?? "").toString().trim();
        });
    } else {
        const inputs = Array.from(document.querySelectorAll(".tc-edit"));
        if (!inputs.length) return;
        inputs.forEach(inp=>{
            const id = (inp.dataset.rowid || "").toString();
            const field = (inp.dataset.field || "").toString();
            if (!id || !field) return;
            if (!byId[id]) byId[id] = {};
            byId[id][field] = (inp.value ?? "").toString().trim();
        });
    }

    // validate helper: empty allowed; if not empty => >0
    function _validateNumber(v){
        const s = (v ?? "").toString().trim();
        if (s === "") return { ok:true, val:"" };
        const n = Number(s);
        if (!Number.isFinite(n)) return { ok:false };
        if (n <= 0) return { ok:false };
        // lưu dạng integer nếu là số nguyên
        const out = Number.isInteger(n) ? String(n) : String(n);
        return { ok:true, val: out };
    }

    // map id -> row in DATA.mon
    const mapIdToRow = new Map();
    (DATA.mon||[]).forEach(r=>{
        const id = (r.id || "").toString().trim();
        if (id) mapIdToRow.set(id, r);
    });

    // apply
    for (const [rid, upd] of Object.entries(byId)){
        const row = mapIdToRow.get(rid);
        if (!row) continue;

        const v1 = _validateNumber(upd.sotiet);
        const v2 = _validateNumber(upd.gioihan);
        if (!v1.ok || !v2.ok){
            alert(`⚠ Dữ liệu không hợp lệ ở: ${row.ten} (${row.khoi}).\n- Chỉ được để trống hoặc nhập số > 0.`);
            return;
        }
        row.sotiet = v1.val;
        row.gioihan = v2.val;
    }

    saveStore();
    // refresh đúng màn đang mở
    const sc = document.getElementById("section-content");
    if (sc) sc.innerHTML = renderTietChuanPage();
    alert("✔ Đã lưu tiết chuẩn.");
}

/* =======================
   TIẾT CHUẨN: CHỌN NHIỀU Ô + COPY/PASTE
   - Click: chọn 1 ô
   - Ctrl/Cmd + Click: thêm/bớt ô
   - Shift + Click: chọn theo vùng (rect)
   - Ctrl/Cmd + C: copy (TSV)
   - Ctrl/Cmd + V: paste (TSV) vào ô đang chọn
======================= */

function tcIsActive(){
    return !!document.querySelector(".tc-cell[data-r][data-c]");
}

function tcKey(r,c){
    return `${Number(r)},${Number(c)}`;
}

function tcGetCell(r,c){
    return document.querySelector(`.tc-cell[data-r="${Number(r)}"][data-c="${Number(c)}"]`);
}

function tcGetCellValue(td){
    if(!td) return "";
    const inp = td.querySelector("input");
    if(inp) return (inp.value ?? "").toString();
    // ưu tiên data-val để đồng bộ với cơ chế Lưu
    return (td.dataset.val ?? td.textContent ?? "").toString();
}

function tcSetCellValue(td, v){
    if(!td) return;
    const val = (v ?? "").toString().trim();
    const inp = td.querySelector("input");
    if(inp) inp.value = val;
    td.dataset.val = val;
    if(!inp) td.textContent = val;
}

function tcUpdateSelectionUI(){
    const cells = Array.from(document.querySelectorAll(".tc-cell[data-r][data-c]"));
    if(!cells.length) return;
    cells.forEach(td=>{
        const r = Number(td.dataset.r);
        const c = Number(td.dataset.c);
        const k = tcKey(r,c);
        td.classList.toggle("tc-selected", TC_CELL_SELECTION.has(k));
    });
}

function tcClearSelection(){
    TC_CELL_SELECTION = new Set();
    TC_CELL_ANCHOR = null;
    tcUpdateSelectionUI();
}

function tcSetSingleSelection(r,c){
    TC_CELL_SELECTION = new Set([tcKey(r,c)]);
    TC_CELL_ANCHOR = {r:Number(r), c:Number(c)};
    tcUpdateSelectionUI();
}

function tcToggleSelection(r,c){
    const k = tcKey(r,c);
    if(TC_CELL_SELECTION.has(k)) TC_CELL_SELECTION.delete(k);
    else TC_CELL_SELECTION.add(k);
    TC_CELL_ANCHOR = {r:Number(r), c:Number(c)};
    tcUpdateSelectionUI();
}

function tcSelectRange(a, b){
    const r1 = Math.min(Number(a?.r), Number(b?.r));
    const r2 = Math.max(Number(a?.r), Number(b?.r));
    const c1 = Math.min(Number(a?.c), Number(b?.c));
    const c2 = Math.max(Number(a?.c), Number(b?.c));
    const next = new Set();
    for(let r=r1; r<=r2; r++){
        for(let c=c1; c<=c2; c++){
            if(tcGetCell(r,c)) next.add(tcKey(r,c));
        }
    }
    TC_CELL_SELECTION = next;
    TC_CELL_ANCHOR = {r:Number(b?.r), c:Number(b?.c)};
    tcUpdateSelectionUI();
}

// Click vào ô (không tự bật edit) để chọn giống Excel
function tcCellClick(ev, td){
    try{
        if(!ev || !td) return;
        // Nếu đang sửa (có input) thì để người dùng thao tác trong input
        if(td.querySelector("input")) return;

        const r = Number(td.dataset.r);
        const c = Number(td.dataset.c);
        if(!Number.isFinite(r) || !Number.isFinite(c)) return;

        const isCmd = !!(ev.ctrlKey || ev.metaKey);
        const isShift = !!ev.shiftKey;

        if(isShift && TC_CELL_ANCHOR){
            tcSelectRange(TC_CELL_ANCHOR, {r,c});
        } else if(isCmd){
            tcToggleSelection(r,c);
        } else {
            tcSetSingleSelection(r,c);
        }

        // tránh bôi đen chữ khi shift-click
        ev.preventDefault();
    }catch(e){
        // ignore
    }
}

function tcGetSelectionRect(){
    if(!TC_CELL_SELECTION || TC_CELL_SELECTION.size === 0) return null;
    let minR = 1e9, maxR = -1e9, minC = 1e9, maxC = -1e9;
    for(const k of TC_CELL_SELECTION){
        const [r,c] = (k||"").split(",").map(Number);
        if(!Number.isFinite(r) || !Number.isFinite(c)) continue;
        if(r < minR) minR = r;
        if(r > maxR) maxR = r;
        if(c < minC) minC = c;
        if(c > maxC) maxC = c;
    }
    if(minR > maxR || minC > maxC) return null;
    return {minR, maxR, minC, maxC};
}

function tcBuildCopyText(){
    const rect = tcGetSelectionRect();
    if(!rect) return "";
    const lines = [];
    for(let r=rect.minR; r<=rect.maxR; r++){
        const row = [];
        for(let c=rect.minC; c<=rect.maxC; c++){
            const k = tcKey(r,c);
            const td = tcGetCell(r,c);
            const v = (td && TC_CELL_SELECTION.has(k)) ? tcGetCellValue(td) : "";
            row.push((v ?? "").toString());
        }
        lines.push(row.join("\t"));
    }
    return lines.join("\n");
}

function tcFallbackCopy(text){
    const ta = document.createElement("textarea");
    ta.value = text;
    ta.setAttribute("readonly","readonly");
    ta.style.position = "fixed";
    ta.style.opacity = "0";
    ta.style.left = "-9999px";
    document.body.appendChild(ta);
    ta.select();
    try{ document.execCommand("copy"); }catch(e){ /* ignore */ }
    ta.remove();
}

function tcCopySelectionToClipboard(){
    try{
        const text = tcBuildCopyText();
        if(!text) return;
        if(navigator.clipboard && navigator.clipboard.writeText){
            navigator.clipboard.writeText(text).catch(()=>tcFallbackCopy(text));
        } else {
            tcFallbackCopy(text);
        }
    }catch(e){
        // ignore
    }
}

function tcParseClipboard(text){
    let t = (text ?? "").toString();
    t = t.replace(/\r\n/g,"\n").replace(/\r/g,"\n");
    // bỏ dòng trống cuối (Excel hay thêm)
    while(t.endsWith("\n")) t = t.slice(0,-1);
    if(!t) return [];
    const rows = t.split("\n").map(line=> line.split("\t"));
    return rows;
}

function tcPasteMatrix(matrix){
    const rect = tcGetSelectionRect();
    if(!rect) return;
    if(!Array.isArray(matrix) || !matrix.length) return;

    // Nếu clipboard chỉ có 1 giá trị -> dán vào tất cả ô đang chọn
    if(matrix.length === 1 && (matrix[0]||[]).length === 1 && TC_CELL_SELECTION.size > 1){
        const v = (matrix[0][0] ?? "").toString().trim();
        for(const k of TC_CELL_SELECTION){
            const [r,c] = (k||"").split(",").map(Number);
            const td = tcGetCell(r,c);
            if(td) tcSetCellValue(td, v);
        }
        return;
    }

    const startR = rect.minR;
    const startC = rect.minC;
    for(let i=0;i<matrix.length;i++){
        const row = matrix[i] || [];
        for(let j=0;j<row.length;j++){
            const td = tcGetCell(startR+i, startC+j);
            if(!td) continue;
            tcSetCellValue(td, (row[j] ?? "").toString());
        }
    }
}

function tcClearSelectedCells(){
    for(const k of TC_CELL_SELECTION){
        const [r,c] = (k||"").split(",").map(Number);
        const td = tcGetCell(r,c);
        if(td) tcSetCellValue(td, "");
    }
}

// Lắng nghe Ctrl/Cmd+C và Delete/Backspace (toàn trang)
function tcGlobalKeyDown(ev){
    try{
        if(!tcIsActive()) return;
        if(!TC_CELL_SELECTION || TC_CELL_SELECTION.size === 0) return;

        const tag = (ev?.target?.tagName || "").toUpperCase();
        const isTyping = (tag === "INPUT" || tag === "TEXTAREA" || tag === "SELECT");

        const isCmd = !!(ev.ctrlKey || ev.metaKey);
        const key = (ev.key || "").toLowerCase();

        if(isCmd && key === "c"){
            // nếu đang gõ trong input -> để copy bình thường
            if(isTyping) return;
            ev.preventDefault();
            tcCopySelectionToClipboard();
            return;
        }

        if(key === "escape"){
            if(isTyping) return;
            tcClearSelection();
            return;
        }

        if((key === "delete" || key === "backspace") && !isTyping){
            ev.preventDefault();
            tcClearSelectedCells();
            tcUpdateSelectionUI();
            return;
        }
    }catch(e){
        // ignore
    }
}

// Lắng nghe paste (Ctrl/Cmd+V) để dán vào bảng Tiết chuẩn
function tcGlobalPaste(ev){
    try{
        if(!tcIsActive()) return;
        if(!TC_CELL_SELECTION || TC_CELL_SELECTION.size === 0) return;

        const tag = (ev?.target?.tagName || "").toUpperCase();
        const isTyping = (tag === "INPUT" || tag === "TEXTAREA" || tag === "SELECT");
        if(isTyping) return; // để paste vào input bình thường

        const text = (ev.clipboardData || window.clipboardData)?.getData("text");
        if(typeof text !== "string") return;

        const matrix = tcParseClipboard(text);
        if(!matrix.length) return;

        ev.preventDefault();
        tcPasteMatrix(matrix);
        tcUpdateSelectionUI();
    }catch(e){
        // ignore
    }
}

// Tiết chuẩn: click vào ô để sửa (input chỉ xuất hiện khi cần)
function tcBeginCellEdit(td){
    try{
        if (!td) return;
        if (td.querySelector("input")) return; // đang sửa

        const rowid = (td.dataset.rowid || "").toString();
        const field = (td.dataset.field || "").toString();
        const cur = (td.dataset.val ?? td.textContent ?? "").toString().trim();
        if (!rowid || !field) return;

        td.innerHTML = `<input class="inline-edit-input tc-edit" type="number" min="0" step="1"
                            data-rowid="${escapeHtml(rowid)}" data-field="${escapeHtml(field)}"
                            value="${escapeHtml(cur)}" style="text-align:center">`;
        const inp = td.querySelector("input");
        if (!inp) return;
        inp.focus();
        inp.select();

        inp.onkeydown = (ev)=>{
            const k = (ev.key || "");
            if (k === "Enter"){
                ev.preventDefault();
                inp.blur();
            }
            if (k === "Escape"){
                ev.preventDefault();
                inp.value = cur;
                inp.blur();
            }
        };

        inp.onblur = ()=>{
            const v = (inp.value ?? "").toString().trim();
            td.dataset.val = v;
            td.innerHTML = escapeHtml(v);
        };
    }catch(e){
        // ignore
    }
}

function tcQuickUpsert(){
    const khoi = _normText(document.getElementById("tc_khoi")?.value || "");
    const ten = _normText(document.getElementById("tc_mon")?.value || "");
    const sotiet = _normText(document.getElementById("tc_sotiet")?.value || "");
    const gioihan = _normText(document.getElementById("tc_gioihan")?.value || "");
    if (!khoi || !ten || !sotiet) return alert("⚠ Cần có Khối + Môn + Số tiết/tuần.");

    const existing = (DATA.mon||[]).find(m => _normText(m.khoi)===khoi && _normText(m.ten)===ten);
    if (existing){
        existing.sotiet = sotiet;
        existing.gioihan = gioihan;
    } else {
        DATA.mon.push({ id:autoID("mon"), khoi, ten, sotiet, gioihan, ghichu:"" });
    }
    saveStore();
    document.getElementById("section-content").innerHTML = renderPCCM();
}

function tcDelete(khoi, ten){
    khoi = _normText(khoi);
    ten = _normText(ten);
    if (!khoi || !ten) return;

    if (!confirm(`Xóa tiết chuẩn: ${ten} (${khoi}) ?`)) return;

    // 1) Xóa khỏi DATA.mon (tiết chuẩn)
    DATA.mon = (DATA.mon || []).filter(m => !(_normText(m.khoi)===khoi && _normText(m.ten)===ten));

    // 2) Xóa phân công PCCM liên quan: chỉ trong các lớp thuộc khối đó
    const khoiNum = extractKhoiNumber(khoi);
    const shouldRemoveKey = (key)=>{
        if (!key.includes("|")) return false;
        const [lopCanon, monTen] = key.split("|");
        if (_normText(monTen) !== ten) return false;
        // lopCanon dạng 6A1/10A1..., khối lấy số đầu
        return extractKhoiNumber(lopCanon) === khoiNum;
    };

    for (const key in (DATA.pccmMatrix || {})){
        if (shouldRemoveKey(key)) delete DATA.pccmMatrix[key];
    }
    for (const key in (DATA.pccmRoomMatrix || {})){
        if (shouldRemoveKey(key)) delete DATA.pccmRoomMatrix[key];
    }

    saveStore();

    // refresh đúng màn đang mở (PCCM hoặc Tiết chuẩn)
    const sc = document.getElementById("section-content");
    if (sc) {
        const html = (sc.innerHTML || "");
        if (html.includes("PCCM") && typeof renderPCCM === "function") {
            sc.innerHTML = renderPCCM();
        } else if (typeof renderTietChuanPage === "function") {
            sc.innerHTML = renderTietChuanPage();
        }
    }
}



function deleteAllPCCM(){
    if (!confirm("⚠ Bạn có chắc muốn XÓA TOÀN BỘ PCCM?")) return;

    DATA.pccmMatrix = {};
    DATA.pccmRoomMatrix = {};
    DATA.pccmTietMatrix = {};
    DATA.pccmGioihanMatrix = {};
    saveStore();

    document.getElementById("section-content").innerHTML = renderPCCM();
    alert("✔ Đã xóa sạch PCCM");
}

/* ============================================================
   UPDATE Ô PCCM
============================================================ */
function updatePCCMCell(el){
    // giữ tương thích với UI cũ (ma trận tổng hợp)
    let lop = el.dataset.lop;
    let mon = el.dataset.mon;
    setPCCMTeacher(lop, mon, el.value);
}

/* ============================================================
   TRIGGER IMPORT PCCM
============================================================ */

function mapPCCMMonNameByCode(monRaw){
    const m = _normText(monRaw);
    if (!m) return "";
    const low = m.toLowerCase();
    const found = (DATA.monhoc || []).find(r=>{
        const ten = _normText(r.ten).toLowerCase();
        const ma  = _normText(r.ma).toLowerCase();
        const ma2 = _normText(r.ma2).toLowerCase();
        return (ten && ten === low) || (ma && ma === low) || (ma2 && ma2 === low);
    });
    if (!found) return m;

    const ten = _normText(found.ten);
    const ma  = _normText(found.ma);
    const ma2 = _normText(found.ma2);
    const fields = [ten, ma, ma2].filter(Boolean);

    function looksLikeCode(s){
        s = _normText(s);
        if (!s) return false;
        const hasNonASCII = /[^\x00-\x7F]/.test(s);
        const hasSpace = /\s/.test(s);
        if (hasNonASCII || hasSpace) return false;
        return s.length <= 12;
    }

    const displayName = _normText(fields.find(x=>!looksLikeCode(x)) || ten || ma || ma2 || m);
    return displayName || m;
}

function triggerPCCMImport(){
    IS_PCCM_IMPORT = true;
    document.getElementById("excelFile").click();
}

/* ============================================================
   IMPORT PCCM EXCEL — AUTO FIX LỖI + BỔ SUNG LỚP/MÔN
============================================================ */

function importPCCMFromExcel(wb){
    // Hỗ trợ 2 dạng:
    // (A) Dạng ma trận (như file pccm.xlsx): hàng 1 là mã/tên môn, cột A là lớp, các ô là GV
    // (B) Dạng 3 cột: Lớp | Môn học | Giáo viên
    const preferSheet = (wb.SheetNames||[]).includes("M3") ? "M3" : (wb.SheetNames||[])[0];
    const sheet = wb.Sheets[preferSheet];
    if (!sheet){
        alert("❌ Không tìm thấy sheet để nhập PCCM.");
        return;
    }

    const rows = XLSX.utils.sheet_to_json(sheet,{header:1,defval:""});
    if (!rows || !rows.length){
        alert("❌ Sheet rỗng.");
        return;
    }

    // helpers
    function looksLikeCode(s){
        s = _normText(s);
        if (!s) return false;
        const hasNonASCII = /[^\x00-\x7F]/.test(s);
        const hasSpace = /\s/.test(s);
        if (hasNonASCII || hasSpace) return false;
        return s.length <= 12;
    }

    function resolveMon(monHeader){
        const raw = _normText(monHeader);
        if (!raw) return null;
        const low = raw.toLowerCase();

        let found = (DATA.monhoc || []).find(r=>{
            const ten = _normText(r.ten).toLowerCase();
            const ma  = _normText(r.ma).toLowerCase();
            const ma2 = _normText(r.ma2).toLowerCase();
            return (ten && ten === low) || (ma && ma === low) || (ma2 && ma2 === low);
        });

        if (!found){
            // không có trong danh mục môn => dùng raw làm cả tên lẫn key
            return { key: raw, ten: raw, code: looksLikeCode(raw) ? raw : "" };
        }

        const ten = _normText(found.ten);
        const ma  = _normText(found.ma);
        const ma2 = _normText(found.ma2);
        const fields = [ten, ma, ma2].filter(Boolean);

        const displayName = _normText(fields.find(x=>!looksLikeCode(x)) || ten || ma || ma2 || raw) || raw;
        const code = _normText(fields.find(x=>looksLikeCode(x)) || ma || ma2 || "") || (looksLikeCode(raw) ? raw : "");
        const key = code || displayName; // KEY lưu PCCM ưu tiên mã
        return { key, ten: displayName, code, ma, ma2 };
    }

    function canonLop(x){
        const raw = String(x||"").trim();
        if (!raw) return "";
        // normalizeClassName có thể đổi 6/1 => 6A1
        try{
            const n = normalizeClassName(raw);
            return (n && String(n).trim()) ? String(n).trim() : raw;
        }catch(_){
            return raw;
        }
    }

    // ---- Detect long format (3 columns) ----
    // Nếu hàng đầu có chứa "Lớp" và "Môn" => long format
    const header0 = rows[0].map(x=>String(x||"").trim().toLowerCase());
    const hasLop = header0.some(x=>x==="lớp" || x==="lop" || x==="tên lớp" || x==="ten lop");
    const hasMon = header0.some(x=>x==="môn học" || x==="mon hoc" || x==="môn" || x==="mon");
    const hasGV  = header0.some(x=>x==="giáo viên" || x==="giao vien" || x==="gv");

    if (hasLop && hasMon && hasGV){
        const objs = XLSX.utils.sheet_to_json(sheet,{defval:""});
        objs.forEach(r=>{
            const lop = canonLop(r["Lớp"] || r["lop"] || r["Tên lớp"] || r["ten lop"] || "");
            const mon = resolveMon(r["Môn học"] || r["Mon hoc"] || r["Môn"] || r["Mon"] || "");
            const gv  = _normText(r["Giáo viên"] || r["Giao vien"] || r["GV"] || r["gv"] || "");
            if (!lop || !mon) return;
            if (gv) pccmSetTeacher(lop, mon, gv);
        });
        saveStore();
        alert("✔ Đã nhập PCCM từ Excel (dạng 3 cột).");
        const sc = document.getElementById("section-content");
        if (sc && typeof renderPCCM === "function") sc.innerHTML = renderPCCM();
        return;
    }

    // ---- Matrix format ----
    // tìm dòng header: dòng đầu tiên có ít nhất 3 ô có dữ liệu
    let headerRow = 0;
    for (let i=0;i<Math.min(rows.length,10);i++){
        const nonEmpty = rows[i].filter(x=>String(x||"").trim()!=="").length;
        if (nonEmpty >= 3){
            headerRow = i; break;
        }
    }

    const header = rows[headerRow].map(x=>String(x||"").trim());
    const monHeaders = header.slice(1); // bỏ cột lớp (* / Lớp)

    const monObjs = monHeaders.map(h=>resolveMon(h));

    for (let i=headerRow+1;i<rows.length;i++){
        const line = rows[i];
        if (!line || !line.length) continue;

        const lop = canonLop(line[0]);
        if (!lop) continue;

        for (let j=0;j<monObjs.length;j++){
            const mon = monObjs[j];
            if (!mon) continue;
            const gv = _normText(line[j+1] || "");
            if (gv){
                pccmSetTeacher(lop, mon, gv);
            }
        }
    }

    saveStore();
    alert("✔ Đã nhập PCCM từ Excel (dạng ma trận).");
    const sc = document.getElementById("section-content");
    if (sc && typeof renderPCCM === "function") sc.innerHTML = renderPCCM();
}

/* ============================================================
   EXPORT PCCM
============================================================ */

function exportPCCMExcel(){
    // Xuất theo khối đang chọn; nếu "Tất cả" thì xuất toàn bộ
    const lopObjs = (DATA.lop || []).map(l=>{
        const canon = (l.ten2 || normalizeClassName(l.ten) || "").trim();
        const khoi = (l.khoi || ("Khối " + extractKhoiNumber(canon))).trim();
        return { canon, khoi };
    }).filter(x=>x.canon);

    const lops = (PCCM_KHOI === "Tất cả")
        ? lopObjs.map(x=>x.canon)
        : lopObjs.filter(x=>extractKhoiNumber(x.khoi)===extractKhoiNumber(PCCM_KHOI)).map(x=>x.canon);

    const monList = buildPCCMMonList();
    const headerMons = monList.map(m=>m.ten || m.key);

    let rows = [];
    rows.push(["Lớp / Môn", ...headerMons]);

    lops.forEach(l=>{
        let line = [l];
        monList.forEach(m=>{
            line.push(pccmGetTeacher(l,m) || "");
        });
        rows.push(line);
    });

    const ws = XLSX.utils.aoa_to_sheet(rows);

    // set độ rộng cột cho file PCCM (tránh các cột bằng nhau)
    try{
        const colCount = rows.reduce((m, r)=>Math.max(m, (r||[]).length), 0);
        const widths = Array.from({length: colCount}).map((_, idx)=>{
            let maxLen = 0;
            (rows||[]).forEach(r=>{
                const v = (r && idx < r.length) ? r[idx] : "";
                const s = (v==null) ? "" : String(v);
                if (s.length > maxLen) maxLen = s.length;
            });
            if (idx === 0) maxLen = Math.max(10, Math.min(maxLen, 24)); // cột "Lớp / Môn"
            return { wch: Math.min(Math.max(8, maxLen + 2), 40) };
        });
        ws["!cols"] = widths;
    }catch(e){
        console.warn("exportPCCMExcel set column widths failed", e);
    }
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "PCCM");
    XLSX.writeFile(wb, "PCCM.xlsx");
}



/* ============================================================
   ==========  PART 4 / 4 — TIỆN ÍCH & HOÀN TẤT FILE ==========
============================================================ */

/* ============================================================
   XÓA SẠCH TOÀN BỘ DỮ LIỆU
============================================================ */
function clearAllData(){
    if (!confirm("⚠ Bạn có chắc muốn XÓA SẠCH toàn bộ dữ liệu?")) return;

    const sid = CTX.schoolId || getSchoolId();

    // Xóa backup localStorage của trường hiện tại
    try{ localStorage.removeItem(_lsKey(sid)); }catch(e){}

    // Xóa dữ liệu trong KVDB của trường hiện tại (nếu có)
    try{
        if (__kv) __kv.set("DATA_JSON", "{}");
    }catch(e){}

    location.reload();
}


/* ============================================================
   TẠO ID TỰ ĐỘNG (K001 / L001 / M001…)
============================================================ */
function autoID(section){
    const map = { khoi:"K", lop:"L", giaovien:"GV", monhoc:"MH", mon:"M", phong:"P" };
    let prefix = map[section] || "X";

    let max = 0;
    DATA[section].forEach(item=>{
        const m = (item.id || "").match(/(\d+)$/);
        if (m) max = Math.max(max, parseInt(m[1],10));
    });

    return prefix + String(max+1).padStart(3,"0");
}


/* ============================================================
   REFRESH VIEW SAU KHI LƯU / IMPORT
============================================================ */
function refreshView(section){
    renderSectionInto(section, "section-content", document);
}

function afterDataChanged(section){
    saveStore();
    refreshView(section);
}


/* ============================================================
   ĐIỀU HƯỚNG VỀ TRANG CHÍNH
============================================================ */
function backToMain(){
    window.location.href = "../index.html";
}


/* ============================================================
   HIỂN THỊ GỢI Ý BẢNG TRỐNG
============================================================ */
function renderEmptyMessage(text){
    return `
        <div style="padding:20px; text-align:center; color:#777;">
            ${text}
        </div>
    `;
}


/* ============================================================
   UTILS
============================================================ */
function isEmptyObject(obj){
    return Object.keys(obj).length === 0;
}

function debug(...args){
    // Bật lên khi cần:
    // console.log("[DEBUG]", ...args);
}
/* ============================================================
   XÓA TOÀN BỘ DỮ LIỆU RIÊNG CỦA MỤC (Khối / Lớp / Môn ...)
============================================================ */
function deleteSection(section){
    if (!confirm("⚠ Bạn có chắc muốn XÓA toàn bộ dữ liệu mục: " + section.toUpperCase() + " ?"))
        return;

    DATA[section] = [];

    // Nếu xóa lớp → phải xóa TKB tương ứng
    if (section === "lop") DATA.tkb = {};

    // Đồng bộ liên kết sau khi xoá mục dữ liệu (tránh tình trạng xoá Môn/Lớp/GV nhưng phân công/tiết chuẩn còn)
    try{ syncDerivedDataIntegrity(); }catch(e){ console.warn("syncDerivedDataIntegrity failed", e); }

    saveStore();

    const sc = document.getElementById("section-content");
    if (sc && typeof renderPCCM === "function" && sc.innerHTML && sc.innerHTML.includes("PCCM")) {
        sc.innerHTML = renderPCCM();
    } else {
        renderSectionInto(section, "section-content", document);
    }

    alert("✔ Đã xóa sạch dữ liệu mục: " + section.toUpperCase());
}



/* ============================================================
   THÔNG BÁO HOÀN TẤT
============================================================ */
debug("✔ app.js đã tải thành công (PART 1 → PART 4)");

/* ============================================================
   SCHOOL SWITCHER (Facebook-like)
============================================================ */
function getSchoolList(){
  return JSON.parse(localStorage.getItem("TKB_SCHOOL_LIST") || "[]");
}
function addSchoolToList(sid){
  let list = getSchoolList();
  if(!list.includes(sid)){
    list.push(sid);
    localStorage.setItem("TKB_SCHOOL_LIST", JSON.stringify(list));
  }
}
function showSchoolSwitcher(){
  closeSchoolSwitcher();
  const schools = getSchoolList();
  const current = getSchoolId();

  let html = `<div class="modal" id="schoolModal" style="display:flex">
    <div class="modal-content" style="width:320px">
      <h3>Chọn trường</h3>
      <div style="margin-top:10px">`;

  schools.forEach(s=>{
    const sid = _sanitizeSchoolId(s);
    const js = (sid||"").replace(/\\/g,"\\\\").replace(/'/g,"\\'");
    const name = _prettySchoolLabel(_getSchoolName(sid) || sid);
    html += `
      <div class="school-item ${sid===current?'active':''}" onclick="switchSchool('${js}')">
        <div class="school-name" title="${escapeHtml(name)}">${escapeHtml(name)}</div>
        <div style="display:flex;gap:6px;align-items:center">
          <button class="school-edit" onclick="event.stopPropagation();renameSchool('${js}')">Sửa</button>
          <button class="school-del" onclick="event.stopPropagation();deleteSchool('${js}')">Xóa</button>
        </div>
      </div>`;
  });

  html += `</div>
    <button class="btn primary" style="width:100%;margin-top:10px" onclick="addNewSchool()">➕ Thêm trường</button>
    <div style="text-align:right;margin-top:10px">
      <button class="btn" onclick="closeSchoolSwitcher()">Đóng</button>
    </div>
  </div></div>`;

  document.body.insertAdjacentHTML("beforeend", html);
}
function closeSchoolSwitcher(){
  const m = document.getElementById("schoolModal");
  if(m) m.remove();
}
function switchSchool(sid){
  const u = new URL(window.location.href);
  u.searchParams.set("school", sid);
  window.location.href = u.toString();
}

// Sửa tên hiển thị (có dấu) của trường hiện tại trong danh sách
function renameSchool(sid){
  sid = _sanitizeSchoolId(sid);
  if(!sid) return;

  const cur = _getSchoolName(sid) || "";
  const next = prompt("Nhập tên hiển thị của trường (có dấu):", cur);
  if(next === null) return;
  const name = String(next||"").trim();
  if(!name) return;

  _setSchoolName(sid, name);
  // nếu đang là trường hiện tại -> cập nhật badge
  if((CTX.schoolId||getSchoolId()) === sid){
    CTX.schoolLabel = _prettySchoolLabel(name) || name;
    try{ localStorage.setItem("TKB_LAST_SCHOOL_LABEL", CTX.schoolLabel); }catch(e){ /* ignore */ }
    updateSchoolBadge();
  }
  showSchoolSwitcher();
}

// Xóa trường khỏi danh sách + xóa dữ liệu lưu trên máy (localStorage + IndexedDB/sql.js)
async function deleteSchool(sid){
  sid = _sanitizeSchoolId(sid);
  if(!sid) return;

  const msg = `Xóa trường "${sid}"?\n\n- Sẽ xóa dữ liệu của trường này khỏi máy (localStorage/IndexedDB).\n- Thao tác này không thể hoàn tác.`;
  if(!confirm(msg)) return;

  // 1) Xóa khỏi danh sách trường
  try{
    const list = getSchoolList().filter(x=>_sanitizeSchoolId(x) !== sid);
    localStorage.setItem("TKB_SCHOOL_LIST", JSON.stringify(list));
  }catch(e){ /* ignore */ }

  // 2) Xóa backup localStorage
  try{ localStorage.removeItem(_lsKey(sid)); }catch(e){ /* ignore */ }

  // 2b) Xóa tên hiển thị
  try{ _deleteSchoolName(sid); }catch(e){ /* ignore */ }

  // 3) Xóa DB trong IndexedDB (KVDB/sql.js lưu ở objectStore "files")
  try{ await _kvdbDeleteDbByName(`TKB::SCHOOL::${sid}`); }catch(e){ /* ignore */ }

  // 4) Nếu đang ở trường bị xóa -> chuyển sang trường khác / default
  const cur = getSchoolId();
  if (cur === sid){
    const nextList = getSchoolList();
    const next = _sanitizeSchoolId(nextList[0] || "default");
    const u = new URL(window.location.href);
    u.searchParams.set("school", next);
    window.location.href = u.toString();
    return;
  }

  // Refresh modal
  showSchoolSwitcher();
}

function _kvdbDeleteDbByName(dbName){
  return new Promise((resolve)=>{
    try{
      const req = indexedDB.open("TKB_SQLJS_DB", 1);
      req.onupgradeneeded = ()=>{
        try{
          const db = req.result;
          if (db && !db.objectStoreNames.contains("files")) db.createObjectStore("files");
        }catch(e){ /* ignore */ }
      };
      req.onsuccess = ()=>{
        const db = req.result;
        try{
          const tx = db.transaction("files","readwrite");
          const st = tx.objectStore("files");
          st.delete(dbName);
          tx.oncomplete = ()=>{ try{ db.close(); }catch(_){ } resolve(true); };
          tx.onerror = ()=>{ try{ db.close(); }catch(_){ } resolve(false); };
        }catch(e){
          try{ db.close(); }catch(_){ }
          resolve(false);
        }
      };
      req.onerror = ()=> resolve(false);
    }catch(e){
      resolve(false);
    }
  });
}
function addNewSchool(){
  // Modal tạo trường mới + chọn cấp (TH/THCS/THPT)
  const old = document.getElementById("addSchoolModal");
  if (old) old.remove();

  const html = `
    <div class="modal" id="addSchoolModal" style="display:flex">
      <div class="modal-content" style="width:360px">
        <h3>Thêm trường</h3>
        <div style="margin-top:12px">
          <div style="font-weight:700;margin-bottom:6px">Mã trường</div>
          <input id="newSchoolId" placeholder="VD: THCS_NguyenTrai" style="width:100%;margin-bottom:12px">

          <div style="font-weight:700;margin-bottom:6px">Cấp trường</div>
          <select id="newSchoolLevel" style="width:100%;margin-bottom:12px">
            <option value="TH">TH (Tiểu học)</option>
            <option value="THCS" selected>THCS</option>
            <option value="THPT">THPT</option>
          </select>

          <div style="display:flex;gap:10px;justify-content:flex-end">
            <button class="btn" onclick="closeAddSchoolModal()">Hủy</button>
            <button class="btn primary" onclick="createNewSchoolFromModal()">Tạo & mở</button>
          </div>
          <div style="margin-top:10px;font-size:12px;color:#667;line-height:1.35">
            * Khi tạo trường mới, hệ thống sẽ tự tạo dữ liệu <b>Khối</b> theo cấp bạn chọn.
          </div>
        </div>
      </div>
    </div>
  `;
  document.body.insertAdjacentHTML("beforeend", html);
}

function closeAddSchoolModal(){
  document.getElementById("addSchoolModal")?.remove();
}

function _seedKhoiByLevel(level){
  const lv = (level||"").toString().trim().toUpperCase();
  let from = 6, to = 9;
  if (lv === "TH") { from = 1; to = 5; }
  if (lv === "THPT") { from = 10; to = 12; }

  const out = [];
  let idx = 1;
  for (let k=from;k<=to;k++){
    out.push({
      id: "K" + String(idx).padStart(3,"0"),
      ten: `Khối ${k}`,
      makhoi: `K${k}`,
      ghichu: ""
    });
    idx++;
  }
  return out;
}

async function createNewSchoolFromModal(){
  const raw = document.getElementById("newSchoolId")?.value || "";
  const level = document.getElementById("newSchoolLevel")?.value || "THCS";
  if(!raw.trim()) return alert("Nhập mã trường.");

  const sid = _sanitizeSchoolId(raw);
  addSchoolToList(sid);
  // lưu tên hiển thị theo input (có thể có dấu)
  try{ _setSchoolName(sid, String(raw).trim()); }catch(e){ /* ignore */ }

  // Nếu trường chưa có dữ liệu, seed khối theo cấp
  const key = _lsKey(sid);
  const existing = _safeParseJSON(localStorage.getItem(key), null);
  const hasAny = existing && typeof existing === "object" && (
    Array.isArray(existing.khoi) && existing.khoi.length
  );

  if (!hasAny){
    const seeded = {
      khoi: _seedKhoiByLevel(level),
      lop: [],
      giaovien: [],
      monhoc: [],
      mon: [],
      phong: [],
      pccmMatrix: {},
      pccmRoomMatrix: {},
      tkb: {},
      tkbConfig: (existing && existing.tkbConfig) ? existing.tkbConfig : undefined
    };

    try{ localStorage.setItem(key, JSON.stringify(seeded)); }catch(e){}

    // seed KVDB nếu có (không bắt buộc)
    try{
      if (window.KVDB){
        const kv = await KVDB.open(`TKB::SCHOOL::${sid}`);
        await kv.set("DATA_JSON", JSON.stringify(seeded));
      }
    }catch(e){
      console.warn("Seed KVDB failed", e);
    }
  }

  closeAddSchoolModal();
  switchSchool(sid);
}

function renderPCCMMonHeader(monObj){
    const name = escapeHtml((monObj && (monObj.ten || monObj.key)) || "");
    const code = escapeHtml((monObj && (monObj.code || "")) || "");
    if (!code) return `<span class="pccm-th-inline"><span class="pccm-th-name" title="${name}">${name}</span></span>`;
    // inline: Tên + (Mã)
    const label = `${name} (${code})`;
    return `<span class="pccm-th-inline" title="${label}"><span class="pccm-th-name">${name}</span><span class="pccm-th-code">(${code})</span></span>`;
}
