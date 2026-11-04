(function () {
"use strict";
/* ============================================================
   XUẤT EXCEL (.xlsx) THUẦN JAVASCRIPT — không cần thư viện ngoài
   Ghi file ZIP chuẩn OOXML để Excel / Google Sheets mở được.
   ============================================================ */
const CRC_TABLE = (() => {
  const t = new Uint32Array(256);
  for (let n = 0; n < 256; n++) {
    let c = n;
    for (let k = 0; k < 8; k++) c = (c & 1) ? (0xEDB88320 ^ (c >>> 1)) : (c >>> 1);
    t[n] = c >>> 0;
  }
  return t;
})();

function crc32(u8) {
  let c = 0xFFFFFFFF;
  for (let i = 0; i < u8.length; i++) c = CRC_TABLE[(c ^ u8[i]) & 0xFF] ^ (c >>> 8);
  return (c ^ 0xFFFFFFFF) >>> 0;
}

function xlsxCol(n) {
  let s = '';
  while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); }
  return s;
}

function xmlEsc(v) {
  return String(v == null ? '' : v)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;').replace(/'/g, '&apos;')
    .replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, '');
}

function zipStore(files) {
  const enc = new TextEncoder();
  const chunks = [], central = [];
  let offset = 0;
  const DOS_TIME = 0, DOS_DATE = 22562; // 2024-01-01
  for (const f of files) {
    const nameb = enc.encode(f.name);
    const data = f.data;
    const crc = crc32(data);
    const lh = new Uint8Array(30 + nameb.length);
    const dv = new DataView(lh.buffer);
    dv.setUint32(0, 0x04034b50, true);
    dv.setUint16(4, 20, true);
    dv.setUint16(6, 0x0800, true); // UTF-8
    dv.setUint16(8, 0, true);      // store
    dv.setUint16(10, DOS_TIME, true);
    dv.setUint16(12, DOS_DATE, true);
    dv.setUint32(14, crc, true);
    dv.setUint32(18, data.length, true);
    dv.setUint32(22, data.length, true);
    dv.setUint16(26, nameb.length, true);
    dv.setUint16(28, 0, true);
    lh.set(nameb, 30);
    chunks.push(lh, data);
    central.push({ nameb, crc, size: data.length, offset });
    offset += lh.length + data.length;
  }
  const cdChunks = [];
  let cdSize = 0;
  for (const c of central) {
    const cd = new Uint8Array(46 + c.nameb.length);
    const dv = new DataView(cd.buffer);
    dv.setUint32(0, 0x02014b50, true);
    dv.setUint16(4, 20, true);
    dv.setUint16(6, 20, true);
    dv.setUint16(8, 0x0800, true);
    dv.setUint16(10, 0, true);
    dv.setUint16(12, DOS_TIME, true);
    dv.setUint16(14, DOS_DATE, true);
    dv.setUint32(16, c.crc, true);
    dv.setUint32(20, c.size, true);
    dv.setUint32(24, c.size, true);
    dv.setUint16(28, c.nameb.length, true);
    dv.setUint16(30, 0, true);
    dv.setUint16(32, 0, true);
    dv.setUint16(34, 0, true);
    dv.setUint16(36, 0, true);
    dv.setUint32(38, 0, true);
    dv.setUint32(42, c.offset, true);
    cd.set(c.nameb, 46);
    cdChunks.push(cd);
    cdSize += cd.length;
  }
  const eocd = new Uint8Array(22);
  const dv = new DataView(eocd.buffer);
  dv.setUint32(0, 0x06054b50, true);
  dv.setUint16(4, 0, true);
  dv.setUint16(6, 0, true);
  dv.setUint16(8, central.length, true);
  dv.setUint16(10, central.length, true);
  dv.setUint32(12, cdSize, true);
  dv.setUint32(16, offset, true);
  dv.setUint16(20, 0, true);
  return new Blob([...chunks, ...cdChunks, eocd], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
}

const XLSX_STYLES = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
  '<fonts count="3"><font><sz val="10"/><name val="Calibri"/></font>' +
  '<font><b/><sz val="10"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font>' +
  '<font><b/><sz val="10"/><name val="Calibri"/></font></fonts>' +
  '<fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill>' +
  '<fill><patternFill patternType="solid"><fgColor rgb="FF0B5CAB"/><bgColor indexed="64"/></patternFill></fill></fills>' +
  '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>' +
  '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
  '<cellXfs count="3"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>' +
  '<xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>' +
  '<xf numFmtId="0" fontId="2" fillId="0" borderId="0" xfId="0" applyFont="1"/></cellXfs>' +
  '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>' +
  '</styleSheet>';

function xlsxSheetXml(sheet) {
  const rows = sheet.rows || [];
  const nCols = rows.reduce((m, r) => Math.max(m, r.length), 1);
  const widths = [];
  for (let c = 0; c < nCols; c++) {
    let w = 8;
    for (let i = 0; i < Math.min(rows.length, 400); i++) {
      const v = rows[i][c];
      const len = v == null ? 0 : String(v).length;
      if (len > w) w = Math.min(55, len + 2);
    }
    widths.push(w);
  }
  const lastCol = xlsxCol(nCols), lastRow = Math.max(rows.length, 1);
  let cols = '<cols>';
  widths.forEach((w, i) => { cols += `<col min="${i + 1}" max="${i + 1}" width="${w}" customWidth="1"/>`; });
  cols += '</cols>';
  let data = '<sheetData>';
  rows.forEach((row, ri) => {
    const r = ri + 1;
    data += `<row r="${r}"${ri === 0 ? ' ht="26" customHeight="1"' : ''}>`;
    row.forEach((v, ci) => {
      const ref = xlsxCol(ci + 1) + r;
      const style = ri === 0 ? ' s="1"' : '';
      if (v === null || v === undefined || v === '') {
        if (ri === 0) data += `<c r="${ref}"${style}/>`;
        return;
      }
      if (typeof v === 'number' && isFinite(v)) {
        data += `<c r="${ref}"${style}><v>${v}</v></c>`;
      } else {
        data += `<c r="${ref}"${style} t="inlineStr"><is><t xml:space="preserve">${xmlEsc(v)}</t></is></c>`;
      }
    });
    data += '</row>';
  });
  data += '</sheetData>';
  return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
    '<dimension ref="A1:' + lastCol + lastRow + '"/>' +
    '<sheetViews><sheetView workbookViewId="0" tabSelected="1">' +
    '<pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/>' +
    '<selection pane="bottomLeft" activeCell="A2" sqref="A2"/>' +
    '</sheetView></sheetViews>' +
    '<sheetFormatPr defaultRowHeight="15"/>' +
    cols + data +
    `<autoFilter ref="A1:${lastCol}${lastRow}"/>` +
    '<pageMargins left="0.35" right="0.35" top="0.5" bottom="0.5" header="0.3" footer="0.3"/>' +
    '</worksheet>';
}

function buildXlsxBlob(sheets) {
  const enc = new TextEncoder();
  const safe = sheets.map((s, i) => {
    let name = String(s.name || ('Sheet' + (i + 1))).replace(/[\[\]\*\?\/\\:]/g, ' ').slice(0, 31) || ('Sheet' + (i + 1));
    return { name, rows: s.rows || [] };
  });
  const files = [];
  files.push({ name: '[Content_Types].xml', data: enc.encode('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
    '<Default Extension="xml" ContentType="application/xml"/>' +
    '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
    safe.map((_, i) => `<Override PartName="/xl/worksheets/sheet${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`).join('') +
    '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>' +
    '</Types>') });
  files.push({ name: '_rels/.rels', data: enc.encode('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>' +
    '</Relationships>') });
  files.push({ name: 'xl/workbook.xml', data: enc.encode('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
    '<bookViews><workbookView xWindow="0" yWindow="0" windowWidth="22000" windowHeight="12000"/></bookViews><sheets>' +
    safe.map((s, i) => `<sheet name="${xmlEsc(s.name)}" sheetId="${i + 1}" r:id="rId${i + 1}"/>`).join('') +
    '</sheets></workbook>') });
  files.push({ name: 'xl/_rels/workbook.xml.rels', data: enc.encode('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    safe.map((_, i) => `<Relationship Id="rId${i + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${i + 1}.xml"/>`).join('') +
    `<Relationship Id="rId${safe.length + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>` +
    '</Relationships>') });
  files.push({ name: 'xl/styles.xml', data: enc.encode(XLSX_STYLES) });
  safe.forEach((s, i) => {
    files.push({ name: `xl/worksheets/sheet${i + 1}.xml`, data: enc.encode(xlsxSheetXml(s)) });
  });
  return zipStore(files);
}

function saveBlob(blob, filename) {
  try {
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = filename;
    a.rel = 'noopener';
    a.style.display = 'none';
    document.body.appendChild(a);
    a.click();
    setTimeout(() => { URL.revokeObjectURL(url); a.remove(); }, 4000);
    return true;
  } catch (e) {
    console.error('download failed', e);
    return false;
  }
}

function downloadXlsx(filename, sheets) {
  const blob = buildXlsxBlob(sheets);
  const ok = saveBlob(blob, filename);
  if (!ok) toast('Trình duyệt chặn tải file. Hãy dùng Chrome/Edge và mở web app ở tab mới.', 'err');
  return ok;
}

function stamp() {
  const d = new Date();
  const p = n => String(n).padStart(2, '0');
  return `${d.getFullYear()}${p(d.getMonth() + 1)}${p(d.getDate())}_${p(d.getHours())}${p(d.getMinutes())}`;
}


/* ============================================================
   LỚP GỌI API + PHIÊN ĐĂNG NHẬP
   - Chạy trong Google Apps Script  → google.script.run
   - Chạy trên GitHub Pages         → JSONP tới Apps Script /exec
   ============================================================ */
const GAS_RUN = (typeof google !== 'undefined' && google.script && google.script.run) ? google.script.run : null;

const LS_TOKEN = 'ctt_token_v1';

/* API CỐ ĐỊNH — đọc duy nhất từ config.js (window.APP_CONFIG) */
function fixedApiConfig() {
  const cfg = (typeof window !== 'undefined' && window.APP_CONFIG) ? window.APP_CONFIG : {};
  return { url: String(cfg.API_URL || '').trim(), code: String(cfg.ACCESS_CODE || '') };
}

function isValidApiUrl(u) {
  try {
    const x = new URL(u);
    if (x.protocol === 'https:') return true;
    if (x.protocol === 'http:' && (x.hostname === 'localhost' || x.hostname === '127.0.0.1')) return true;
    return false;
  } catch (e) { return false; }
}

function apiUid() {
  try {
    if (window.crypto && crypto.randomUUID) return crypto.randomUUID();
  } catch (e) {}
  return 'c' + Date.now().toString(36) + Math.random().toString(36).slice(2, 10);
}

let JSONP_SEQ = 0;
function jsonpCall(action, payload, timeoutMs) {
  return new Promise((resolve, reject) => {
    if (!API.url) return reject(new Error('NO_API'));
    const cb = '__ctt_cb_' + (++JSONP_SEQ) + '_' + Date.now();
    const params = new URLSearchParams();
    params.set('action', action);
    params.set('callback', cb);
    params.set('_t', String(Date.now()));
    if (API.code) params.set('code', API.code);
    if (API.token) params.set('token', API.token);
    if (payload) params.set('payload', JSON.stringify(payload));

    const script = document.createElement('script');
    let done = false;
    const timer = setTimeout(() => finish(new Error('API không phản hồi sau ' + Math.round((timeoutMs || 120000) / 1000) + ' giây. Kiểm tra URL /exec và kết nối mạng.')), timeoutMs || 120000);

    function finish(err, data) {
      if (done) return;
      done = true;
      clearTimeout(timer);
      try { delete window[cb]; } catch (e) { window[cb] = undefined; }
      script.remove();
      if (err) reject(err); else resolve(data);
    }
    window[cb] = data => finish(null, data);
    script.onerror = () => finish(new Error('Không tải được API. Kiểm tra URL /exec, quyền truy cập web app và mạng.'));
    script.src = API.url + (API.url.indexOf('?') === -1 ? '?' : '&') + params.toString();
    document.head.appendChild(script);
  });
}

const API = {
  mode: GAS_RUN ? 'gas' : 'http',
  url: '',
  code: '',
  token: '',
  user: null,
  available: false,

  init() {
    if (GAS_RUN) {
      this.mode = 'gas';
      this.available = true;
      try { this.code = sessionStorage.getItem(SS_CODE) || ''; } catch (e) {}
    } else {
      this.mode = 'http';
      const c = fixedApiConfig();
      this.url = c.url;
      this.code = c.code;
      this.available = !!this.url && isValidApiUrl(this.url);
    }
    try { this.token = localStorage.getItem(LS_TOKEN) || ''; } catch (e) { this.token = ''; }
  },

  setToken(t) {
    this.token = String(t || '');
    try {
      if (this.token) localStorage.setItem(LS_TOKEN, this.token);
      else localStorage.removeItem(LS_TOKEN);
    } catch (e) {}
  },

  clearToken() { this.setToken(''); this.user = null; },

  host() {
    try { return new URL(this.url).host; } catch (e) { return this.url || ''; }
  },

  _err(res) {
    const e = new Error((res && res.error) || 'Lỗi từ API');
    e.code = (res && res.code) || '';
    return e;
  },

  _gas(name, args) {
    return new Promise((resolve, reject) => {
      if (!GAS_RUN) return reject(new Error('NO_GAS'));
      let done = false;
      const timer = setTimeout(() => {
        if (done) return; done = true;
        reject(new Error('Máy chủ phản hồi quá lâu (quá 120 giây). Kiểm tra kết nối mạng rồi thử lại.'));
      }, 120000);
      try {
        GAS_RUN
          .withSuccessHandler(res => {
            if (done) return; done = true; clearTimeout(timer);
            if (res && res.ok === false) reject(this._err(res));
            else resolve(res);
          })
          .withFailureHandler(err => {
            if (done) return; done = true; clearTimeout(timer);
            const m = (err && err.message) || String(err);
            const e = new Error(/authorization|permission|required/i.test(m)
              ? 'Chưa cấp quyền truy cập Google Sheet. Hãy mở lại web app và bấm "Cho phép".'
              : m);
            if (/authorization|permission|required/i.test(m)) e.code = 'AUTH';
            reject(e);
          })[name](args);
      } catch (e) {
        if (!done) { done = true; clearTimeout(timer); reject(e); }
      }
    });
  },

  async _http(action, payload, timeout) {
    const res = await jsonpCall(action, payload, timeout);
    if (res && res.ok === false) throw this._err(res);
    return res;
  },

  async login(username, password) {
    const payload = { username: username, password: password, code: this.code };
    const res = this.mode === 'gas'
      ? await this._gas('login', payload)
      : await this._http('login', payload, 45000);
    if (!res || !res.token) throw new Error('Đăng nhập không trả về phiên làm việc.');
    this.setToken(res.token);
    this.user = res.user || null;
    return res;
  },

  async me() {
    if (!this.token) throw Object.assign(new Error('Chưa đăng nhập.'), { code: 'AUTH' });
    const res = this.mode === 'gas'
      ? await this._gas('me', { token: this.token })
      : await this._http('me', null, 30000);
    this.user = res.user || null;
    return res.user;
  },

  async logout() {
    try {
      if (this.token) {
        if (this.mode === 'gas') await this._gas('logout', { token: this.token });
        else await this._http('logout', null, 15000);
      }
    } catch (e) { /* bỏ qua */ }
    this.clearToken();
  },

  async changePassword(oldPassword, newPassword) {
    const payload = { old_password: oldPassword, new_password: newPassword, code: this.code, token: this.token };
    const res = this.mode === 'gas'
      ? await this._gas('changePassword', payload)
      : await this._http('changePassword', payload, 45000);
    if (res && res.user) this.user = res.user;
    return res;
  },

  bootstrap() {
    if (this.mode === 'gas') return this._gas('getBootstrap', { token: this.token, code: this.code });
    return this._http('bootstrap', null, 90000);
  },

  saveTouch(payload) {
    if (this.mode === 'gas') return this._gas('saveTouch', Object.assign({ token: this.token, code: this.code }, payload));
    const p = Object.assign({}, payload);
    p.client_id = p.client_id || apiUid();
    return this._http('saveTouch', p, 60000);
  },

  deleteTouch(id) {
    if (this.mode === 'gas') return this._gas('deleteTouch', { id: id, token: this.token, code: this.code });
    return this._http('deleteTouch', { id: id }, 45000);
  },

  saveAppointment(payload) {
    if (this.mode === 'gas') return this._gas('saveAppointment', Object.assign({ token: this.token, code: this.code }, payload));
    return this._http('saveAppointment', payload, 60000);
  },

  updateAppointment(payload) {
    if (this.mode === 'gas') return this._gas('updateAppointment', Object.assign({ token: this.token, code: this.code }, payload));
    return this._http('updateAppointment', payload, 45000);
  },

  ping() {
    if (this.mode === 'gas') return Promise.resolve({ ok: true, rows: ALL.length });
    return this._http('ping', null, 25000);
  }
};


/* ============================================================
   LÕI ỨNG DỤNG — dữ liệu đọc trực tiếp từ Google Sheet
   ============================================================ */
const SVC = {
  tt:   { key: 'tt',   label: 'Chuyển trả trước', short: 'Trả trước', icon: '💳', note: 'Đang trả sau → chuyển trả trước' },
  mytv: { key: 'mytv', label: 'MyTV',             short: 'MyTV',      icon: '📺', note: 'Truyền hình + kho phim' },
  mesh: { key: 'mesh', label: 'Mesh Wi-Fi',       short: 'Mesh',      icon: '📶', note: 'Phủ sóng toàn nhà' },
  cam:  { key: 'cam',  label: 'Camera an ninh',   short: 'Cam',       icon: '📷', note: 'Giám sát từ điện thoại' }
};
const SVC_KEYS = ['tt', 'mytv', 'mesh', 'cam'];
const UNIT_LABEL = { VNPT: 'VNPT KV' };
const unitLabel = u => UNIT_LABEL[u] || u || '';

const RESULT = {
  QUAN_TAM:       { label: 'Quan tâm',            cls: 'res-quan_tam' },
  DA_CHOT:        { label: 'Đã chốt',             cls: 'res-da_chot' },
  HEN_LAI:        { label: 'Hẹn lại',             cls: 'res-hen_lai' },
  TU_CHOI:        { label: 'Từ chối',             cls: 'res-tu_choi' },
  KHONG_LIEN_HE:  { label: 'Không liên hệ được',  cls: 'res-khong_lien_he' }
};

const SCRIPT_TXT = {
  tt: 'Anh/chị đang dùng trả sau. Chuyển sang trả trước giúp chủ động kiểm soát cước hằng tháng, không lo phát sinh ngoài ý muốn, vẫn giữ nguyên số điện thoại và các dịch vụ đang dùng. Anh/chị đồng ý để em hướng dẫn chuyển đổi ngay tại nhà không ạ?',
  mytv: 'Nhà mình chưa lắp MyTV. Nhân dịp thay modem, em lắp thêm MyTV trên cùng đường truyền — không tốn thêm dây, nhiều kênh và kho phim cho cả nhà. Anh/chị có muốn em kiểm tra hạ tầng và báo giá ưu đãi không ạ?',
  mesh: 'Nhà mình diện tích rộng, nhiều phòng nên sóng Wi-Fi có thể chưa phủ hết. Em bổ sung thêm Mesh để đi đến đâu cũng có Wi-Fi mạnh, thiết bị tự chuyển vùng, không phải đổi mật khẩu. Anh/chị có nhu cầu em tư vấn gói phù hợp nhé?',
  cam: 'Nhân dịp thay modem, em giới thiệu gói camera an ninh VNPT: xem trực tiếp trên điện thoại, lưu trữ đám mây, lắp đặt nhanh. Anh/chị muốn em khảo sát vị trí lắp đặt không ạ?'
};

/* ---------- tiện ích ---------- */
const $  = (s, r) => (r || document).querySelector(s);
const $$ = (s, r) => Array.prototype.slice.call((r || document).querySelectorAll(s));
const esc = s => String(s == null ? '' : s).replace(/[&<>"']/g, m => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[m]));
const nf = n => new Intl.NumberFormat('vi-VN').format(Number(n) || 0);
const p2 = n => String(n).padStart(2, '0');
const norm = s => String(s == null ? '' : s).normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/đ/g, 'd').replace(/Đ/g, 'D').toLowerCase().trim();
const todayISO = () => { const d = new Date(); return `${d.getFullYear()}-${p2(d.getMonth() + 1)}-${p2(d.getDate())}`; };
const nowHHMM = () => { const d = new Date(); return `${p2(d.getHours())}:${p2(d.getMinutes())}`; };
const fmtDate = iso => { if (!iso) return ''; const a = String(iso).slice(0, 10).split('-'); return a.length === 3 ? `${a[2]}/${a[1]}/${a[0]}` : String(iso); };
const shortTeam = t => String(t || '').normalize('NFC').replace(/^Tổ Kỹ thuật Địa bàn\s+/i, 'Tổ ');
const empName = nv => { const s = String(nv || ''); const i = s.indexOf(' - '); return i > -1 ? s.slice(i + 3) : s; };
const empCode = nv => { const s = String(nv || ''); const i = s.indexOf(' - '); return i > -1 ? s.slice(0, i) : ''; };
const uniq = arr => Array.from(new Set(arr.filter(v => v !== '' && v != null)));
const cleanText = v => { const s = String(v == null ? '' : v).trim(); return /^#(N\/A|VALUE!|REF!|DIV\/0!|NAME\?|NUM!)$/i.test(s) ? '' : s; };
const displayName = r => cleanText(r.ten) || 'Khách chưa có tên';
const cleanPhone = v => cleanText(v).replace(/[^0-9+]/g, '');
const displayPhone = r => cleanText(r.sdt);
const pctNum = (n, d) => d > 0 ? (n / d * 100).toLocaleString('vi-VN', { minimumFractionDigits: 1, maximumFractionDigits: 1 }) : '0,0';
const pct = (n, d) => pctNum(n, d) + '%';
const byName = (a, b) => String(a).localeCompare(String(b), 'vi');

function toast(msg, type, ms) {
  const w = $('#toasts');
  if (!w) return;
  const el = document.createElement('div');
  el.className = 'toast ' + (type || '');
  el.innerHTML = `<span>${type === 'ok' ? '✓' : type === 'err' ? '✕' : type === 'warn' ? '!' : 'ℹ'}</span><span>${esc(msg)}</span>`;
  w.appendChild(el);
  setTimeout(() => { el.style.opacity = '0'; el.style.transform = 'translateY(8px)'; el.style.transition = '.25s'; setTimeout(() => el.remove(), 300); }, ms || 3400);
}

let _confirmResolve = null;
function confirmDialog(opts) {
  return new Promise(resolve => {
    _confirmResolve = resolve;
    $('#confirm-title').textContent = opts.title || 'Xác nhận';
    $('#confirm-msg').innerHTML = opts.msg || '';
    $('#confirm-yes').textContent = opts.yes || 'Đồng ý';
    $('#confirm-yes').className = 'btn ' + (opts.danger === false ? 'primary' : 'danger');
    $('#confirm-modal').classList.add('show');
    $('#confirm-overlay').classList.add('show');
  });
}
function closeConfirm(v) {
  $('#confirm-modal').classList.remove('show');
  $('#confirm-overlay').classList.remove('show');
  if (_confirmResolve) { const r = _confirmResolve; _confirmResolve = null; r(v); }
}

/* ---------- trạng thái ---------- */
const LS_CONFIG = 'ctt_config_v1';

let ALL = [];
let BY_MA = new Map();
let TOUCHES = [];
let LATEST = new Map();
let TOUCH_BY_MA = new Map();
let AGREED = new Map();
let APPOINTMENTS = [];
let APPT_BY_MA = new Map();
let APPT_ACTIVE = new Set();
let APPT_LATEST = new Map();
let META = {};
let CONFIG = { target: 15 };
let USER = null;
let CURRENT_MA = null;
let FORM_CLIENT_ID = null;
let FORM_MODE = 'touch';
let FORM_SVC = { tu_van: new Set(), dong_y: new Set() };
let LOADED = false;

const UI = {
  tab: 'list',
  filters: { q: '', to: '', kv: '', nv: '', xa: '', ap: '', duong: '', thay: 'chua', cham: 'all', hen: 'all', need: [], sort: 'uu_tien', topN: 0 },
  page: 1,
  pageSize: 50,
  unit: {
    TTVT: { q: '', from: '', to: '', emp: '' },
    VNPT: { q: '', from: '', to: '', emp: '' }
  },
  appt: { q: '', from: '', to: '', status: 'CHO_THAY', emp: '', kv: '' },
  report: { q: '', from: '', to: '', donvi: '', team: '', emp: '', xa: '', kq: '', svc: '' },
  stats: { from: '', to: '', unit: '' }
};

function loadLocal() {
  try {
    const c = JSON.parse(localStorage.getItem(LS_CONFIG) || 'null');
    if (c && c.target) CONFIG.target = c.target;
  } catch (e) {}
}
function saveLocal() {
  try { localStorage.setItem(LS_CONFIG, JSON.stringify(CONFIG)); } catch (e) {}
}

function recomputeTouches() {
  LATEST = new Map();
  TOUCH_BY_MA = new Map();
  AGREED = new Map();
  const sorted = TOUCHES.slice().sort((a, b) => (a.ts || '').localeCompare(b.ts || ''));
  for (const t of sorted) {
    if (!TOUCH_BY_MA.has(t.ma)) TOUCH_BY_MA.set(t.ma, []);
    TOUCH_BY_MA.get(t.ma).push(t);
    LATEST.set(t.ma, t);
    if (t.dong_y && t.dong_y.length) {
      if (!AGREED.has(t.ma)) AGREED.set(t.ma, new Set());
      t.dong_y.forEach(k => AGREED.get(t.ma).add(k));
    }
  }
}

function recomputeAppointments() {
  APPT_BY_MA = new Map();
  APPT_ACTIVE = new Set();
  APPT_LATEST = new Map();
  const sorted = APPOINTMENTS.slice().sort((a, b) => String(a.cap_nhat || a.ngay_tao || '').localeCompare(String(b.cap_nhat || b.ngay_tao || '')));
  for (const a of sorted) {
    APPT_LATEST.set(a.ma, a);
    if (a.trang_thai === 'CHO_THAY') {
      APPT_BY_MA.set(a.ma, a);
      APPT_ACTIVE.add(a.ma);
    }
  }
  APPOINTMENTS.sort((a, b) => String(a.ngay_hen || '').localeCompare(String(b.ngay_hen || '')) ||
    String(a.gio_hen || '').localeCompare(String(b.gio_hen || '')));
}

/* ---------- tải dữ liệu ---------- */
function showLoading(title, sub) {
  $('#loading').classList.remove('hidden');
  $('#load-spin').classList.remove('hidden');
  $('#load-err').classList.add('hidden');
  $('#load-title').textContent = title || 'Đang đọc dữ liệu từ Google Sheet…';
  $('#load-sub').textContent = sub || 'Lần đầu tải có thể mất vài giây.';
}
function hideLoading() { $('#loading').classList.add('hidden'); }

function showLoadError(err) {
  const box = $('#load-err');
  const m = String((err && err.message) || err || '');
  let html = '';
  if (m === 'NO_GAS' || m === 'NO_API') {
    html = `<h4>Chưa cấu hình API</h4>
      <p>Mở file <code>config.js</code> của website và điền <code>API_URL</code> (URL <code>/exec</code> đã triển khai). Chỉ quản trị sửa một lần; nhân viên không phải nhập.</p>`;
  } else if (/quyền|authoriz|permission|cấp quyền/i.test(m)) {
    html = `<h4>Chưa cấp quyền truy cập</h4><p>${esc(m)}</p><p>Mở lại link web app API và bấm <b>Cho phép</b> khi Google hỏi quyền truy cập Sheet.</p>`;
  } else if (/mã truy cập/i.test(m)) {
    html = `<h4>Sai hoặc thiếu mã truy cập</h4><p>${esc(m)}</p>`;
  } else if (API.mode === 'http') {
    html = `<h4>Không gọi được API</h4><p>${esc(m)}</p>
      <p>Kiểm tra: (1) URL <code>/exec</code> đúng, (2) web app triển khai quyền <b>Bất kỳ ai</b>, (3) mạng ổn định.</p>`;
  } else {
    html = `<h4>Không đọc được dữ liệu</h4><p>${esc(m)}</p>
      <p>Kiểm tra: (1) file Sheet còn tồn tại, (2) tài khoản triển khai còn quyền sửa Sheet, (3) mạng ổn định.</p>`;
  }
  $('#load-spin').classList.add('hidden');
  $('#load-title').textContent = 'Có lỗi khi tải dữ liệu';
  $('#load-sub').textContent = 'Chi tiết bên dưới.';
  box.innerHTML = html + `<div style="margin-top:12px;display:flex;gap:8px"><button type="button" class="btn primary" data-act="retry-load">Thử lại</button></div>`;
  box.classList.remove('hidden');
}

async function loadData() {
  showLoading();
  try {
    const res = await API.bootstrap();
    if (!res || !res.rows) throw new Error('Dữ liệu trả về không hợp lệ.');
    META = res.meta || {};
    USER = res.user || API.user || USER;
    API.user = USER;
    ALL = res.rows.map((a, i) => ({
      id: i,
      ma: a[0], ten: a[1], sdt: a[2], to: a[3], nv: a[4], kv: a[5],
      traSau: Number(a[6]) === 1 || a[6] === true,
      dc: a[7], sn: a[8], duong: a[9], ap: a[10], xa: a[11],
      modem: a[12], daThay: Number(a[13]) === 1 || a[13] === true,
      mesh: a[14] || '', cam: a[15] || '', mytv: a[16] || ''
    }));
    BY_MA = new Map(ALL.map(r => [r.ma, r]));
    TOUCHES = res.touches || [];
    APPOINTMENTS = res.appointments || [];
    if (res.config && res.config.target) CONFIG.target = res.config.target;
    recomputeTouches();
    recomputeAppointments();
    buildOptions();
    renderUserUI();
    LOADED = true;
    renderAll();
    hideLoading();
    hideLogin();
    const conn = $('#conn');
    conn.className = 'conn ok';
    $('#conn-text').textContent = `Đã kết nối · ${nf(ALL.length)} hộ`;
    if (res.warning) toast(res.warning, 'warn', 6000);
  } catch (e) {
    const conn = $('#conn');
    conn.className = 'conn err';
    $('#conn-text').textContent = 'Chưa kết nối';
    hideLoading();
    if (e.code === 'AUTH') {
      API.clearToken();
      showLogin(e.message);
      return;
    }
    showLoadError(e);
  }
}

function showCodePrompt(msg) {
  const box = $('#load-err');
  $('#load-spin').classList.add('hidden');
  $('#load-title').textContent = 'Nhập mã truy cập';
  $('#load-sub').textContent = msg || 'Trang này được bảo vệ bằng mã truy cập do quản trị viên cung cấp.';
  box.innerHTML = `<div class="form-field" style="margin-bottom:10px"><label>Mã truy cập</label>
    <input id="access-code" type="password" placeholder="Nhập mã…" autocomplete="off"></div>
    <div style="display:flex;gap:8px"><button type="button" class="btn primary" data-act="submit-code">Vào trang</button></div>`;
  box.classList.remove('hidden');
  $('#access-code').focus();
  $('#access-code').addEventListener('keydown', e => { if (e.key === 'Enter') submitCode(); });
}
function submitCode() {
  const v = ($('#access-code') && $('#access-code').value || '').trim();
  if (!v) return toast('Nhập mã truy cập', 'err');
  try { sessionStorage.setItem(SS_CODE, v); } catch (e) {}
  loadData(v);
}

/* ---------- danh mục bộ lọc ---------- */
function buildOptions() {
  const tos = uniq(ALL.map(r => r.to)).sort(byName);
  const xas = uniq(ALL.map(r => r.xa)).sort(byName);
  const kvs = uniq(ALL.map(r => r.kv)).sort(byName);
  fillSelect('#f-to', tos.map(v => [v, shortTeam(v)]), 'Tất cả tổ');
  fillSelect('#f-kv', kvs.map(v => [v, v]), 'Tất cả KV');
  fillSelect('#f-xa', xas.map(v => [v, v]), 'Tất cả xã');
  rebuildNvSelect();
  rebuildApSelect();
  rebuildDuongDatalist();
  // đếm nhu cầu toàn danh sách
  const cnt = { tt: 0, mytv: 0, mesh: 0, cam: 0 };
  for (const r of ALL) gapKeys(r).forEach(k => cnt[k]++);
  SVC_KEYS.forEach(k => { const el = $('#need-n-' + k); if (el) el.textContent = nf(cnt[k]); });
  buildQuickCounts();
  renderQuickFilters();
}

function fillSelect(sel, pairs, firstLabel) {
  const el = $(sel);
  if (!el) return;
  const cur = el.value;
  el.innerHTML = `<option value="">${esc(firstLabel)}</option>` + pairs.map(p => `<option value="${esc(p[0])}">${esc(p[1])}</option>`).join('');
  if (cur && pairs.some(p => p[0] === cur)) el.value = cur;
}

function employeesForTeam() {
  const f = UI.filters;
  return uniq(ALL.filter(r => !f.to || r.to === f.to).map(r => r.nv))
    .sort((a, b) => byName(empName(a), empName(b)));
}

function rebuildNvSelect() {
  const f = UI.filters;
  if (f.nv && employeesForTeam().indexOf(f.nv) === -1) f.nv = '';
}

function rebuildApSelect() {
  const f = UI.filters;
  const list = ALL.filter(r => !f.xa || r.xa === f.xa).map(r => r.ap);
  const pairs = uniq(list).sort(byName).map(v => [v, v]);
  fillSelect('#f-ap', pairs, 'Tất cả ấp');
  if (f.ap && !pairs.some(p => p[0] === f.ap)) { f.ap = ''; $('#f-ap').value = ''; }
}

function rebuildDuongDatalist() {
  const f = UI.filters;
  const list = uniq(ALL.filter(r => (!f.xa || r.xa === f.xa) && (!f.ap || r.ap === f.ap)).map(r => r.duong)).sort(byName);
  const dl = $('#dl-duong');
  if (dl) dl.innerHTML = list.map(v => `<option value="${esc(v)}"></option>`).join('');
}

/* ---------- KPI ---------- */
function renderKPIs() {
  const total = ALL.length;
  const daThay = ALL.filter(r => r.daThay).length;
  const touched = new Set(TOUCHES.map(t => t.ma));
  const today = todayISO();
  const todayTouches = TOUCHES.filter(t => t.ngay === today);
  const usersToday = new Set(todayTouches.map(t => t.nguoi || '—')).size;
  const targetTotal = Math.max(1, usersToday) * (CONFIG.target || 15);
  const agreed = new Set(TOUCHES.filter(t => t.dong_y && t.dong_y.length).map(t => t.ma));
  const svcSets = { tt: new Set(), mytv: new Set(), mesh: new Set(), cam: new Set() };
  TOUCHES.forEach(t => (t.dong_y || []).forEach(k => svcSets[k] && svcSets[k].add(t.ma)));
  const totalSvc = SVC_KEYS.reduce((s, k) => s + svcSets[k].size, 0);

  $('#kpi-total').textContent = nf(total);
  $('#kpi-total-sub').textContent = `Chưa thay ${nf(total - daThay)} · Đã thay ${nf(daThay)}`;
  $('#kpi-covered').textContent = nf(touched.size);
  const rate = total ? touched.size / total * 100 : 0;
  $('#kpi-covered-rate').textContent = ` · ${pct(touched.size, total)}`;
  $('#kpi-covered-sub').textContent = `Còn ${nf(Math.max(0, total - touched.size))} hộ chưa chạm`;
  $('#kpi-covered-bar').style.width = Math.min(100, rate) + '%';

  $('#kpi-today').textContent = nf(todayTouches.length);
  $('#kpi-today-sub').textContent = `Mục tiêu ${CONFIG.target || 15} lượt/người · ${usersToday || 0} người hôm nay`;
  $('#kpi-today-bar').style.width = Math.min(100, todayTouches.length / targetTotal * 100) + '%';

  $('#kpi-agreed').textContent = nf(agreed.size);
  $('#kpi-agreed-sub').textContent = totalSvc ? `${nf(totalSvc)} dịch vụ đã được khách đồng ý` : 'Chưa có dịch vụ nào được chốt';

  $('#kpi-svc').innerHTML = SVC_KEYS.map(k =>
    `<div><span><i class="dot" style="background:${svcColor(k)}"></i>${SVC[k].short}</span><b>${nf(svcSets[k].size)}</b></div>`
  ).join('');
}

function svcColor(k) {
  return { tt: '#e07b00', mytv: '#6b5bd2', mesh: '#0e8f8a', cam: '#178a5c' }[k];
}

function renderBadges() {
  const need = ALL.filter(r => !r.daThay && !LATEST.has(r.ma)).length;
  $('#badge-list').textContent = nf(need);
  const t = TOUCHES.filter(x => x.donvi === 'TTVT').length;
  const v = TOUCHES.filter(x => x.donvi === 'VNPT').length;
  $('#badge-ttvt').textContent = nf(t);
  $('#badge-vnpt').textContent = nf(v);
  $('#badge-report').textContent = nf(TOUCHES.length);
  const ba = $('#badge-appt');
  if (ba) ba.textContent = nf(APPOINTMENTS.filter(a => a.trang_thai === 'CHO_THAY').length);
}

/* ---------- danh sách ---------- */
function gapKeys(r) {
  const g = [];
  if (r.traSau) g.push('tt');
  if (!r.mytv) g.push('mytv');
  if (!r.mesh) g.push('mesh');
  if (!r.cam) g.push('cam');
  return g;
}
function score(r) {
  let s = gapKeys(r).length * 10;
  if (!r.daThay) s += 12;
  if (r.traSau) s += 3;
  return s;
}

function filterRowMatches(r, f, q, dq, ignoreNv) {
  if (f.to && r.to !== f.to) return false;
  if (!ignoreNv && f.nv && r.nv !== f.nv) return false;
  if (f.kv && r.kv !== f.kv) return false;
  if (f.xa && r.xa !== f.xa) return false;
  if (f.ap && r.ap !== f.ap) return false;
  if (dq && norm(r.duong).indexOf(dq) === -1) return false;
  if (f.thay === 'chua' && r.daThay) return false;
  if (f.thay === 'da' && !r.daThay) return false;
  const latest = LATEST.get(r.ma);
  if (f.cham === 'chua' && latest) return false;
  if (f.cham === 'da' && !latest) return false;
  if (f.hen === 'co' && !APPT_ACTIVE.has(r.ma)) return false;
  if (f.hen === 'chua' && APPT_ACTIVE.has(r.ma)) return false;
  if (f.need.length) { const g = gapKeys(r); if (!f.need.every(k => g.indexOf(k) > -1)) return false; }
  if (q) {
    const hay = norm(r.ten) + ' ' + norm(r.ma) + ' ' + norm(r.sdt) + ' ' + norm(r.dc) + ' ' + norm(r.duong) + ' ' + norm(r.ap) + ' ' + norm(r.xa) + ' ' + norm(r.kv);
    if (hay.indexOf(q) === -1) return false;
  }
  return true;
}

function currentList() {
  const f = UI.filters;
  const q = norm(f.q);
  const dq = norm(f.duong);
  let arr = ALL.filter(r => filterRowMatches(r, f, q, dq, false));
  const t = (ma) => LATEST.get(ma);
  if (f.sort === 'uu_tien') arr.sort((a, b) => score(b) - score(a) || byName(a.ten, b.ten));
  else if (f.sort === 'ten') arr.sort((a, b) => byName(a.ten, b.ten));
  else if (f.sort === 'duong') arr.sort((a, b) => byName(a.duong, b.duong) || String(a.sn).localeCompare(String(b.sn), 'vi', { numeric: true }));
  else if (f.sort === 'moi_nhat') arr.sort((a, b) => {
    const ta = t(a.ma), tb = t(b.ma);
    if (!ta && !tb) return score(b) - score(a);
    if (!ta) return 1; if (!tb) return -1;
    return String(tb.ts).localeCompare(String(ta.ts));
  });
  else if (f.sort === 'chua_cham') arr.sort((a, b) => {
    const ta = t(a.ma), tb = t(b.ma);
    if (!ta && !tb) return score(b) - score(a);
    if (!ta) return -1; if (!tb) return 1;
    return String(ta.ts).localeCompare(String(tb.ts));
  });
  if (f.topN) arr = arr.slice(0, f.topN);
  return arr;
}

function buildQuickCounts() {
  const total = ALL.length;
  const chua = ALL.filter(r => !r.daThay).length;
  const a = $('#qf-n-chua'), b = $('#qf-n-da'), c = $('#qf-n-all');
  if (a) a.textContent = nf(chua);
  if (b) b.textContent = nf(total - chua);
  if (c) c.textContent = nf(total);
}

function renderQuickFilters() {
  $$('#qf-thay .qf-pill').forEach(b => b.classList.toggle('on', b.dataset.v === UI.filters.thay));
  const btn = $('#qf-emp-btn');
  if (!btn) return;
  const label = $('#qf-emp-label');
  if (UI.filters.nv) {
    btn.classList.add('on');
    label.innerHTML = `<b>${esc(empName(UI.filters.nv))}</b> <span class="muted">${esc(empCode(UI.filters.nv))}</span>`;
    btn.title = UI.filters.nv;
  } else {
    btn.classList.remove('on');
    label.textContent = 'Tất cả nhân viên';
    btn.title = 'Chọn nhân viên kỹ thuật';
  }
}

function openEmpPanel() {
  const pop = $('#qf-emp-pop');
  if (!pop) return;
  if (pop.classList.contains('show')) return closeEmpPanel();
  $('#qf-emp-search').value = '';
  renderEmpList('');
  pop.classList.add('show');
  $('#qf-emp-btn').setAttribute('aria-expanded', 'true');
  setTimeout(() => { const el = $('#qf-emp-search'); if (el) el.focus(); }, 60);
}

function closeEmpPanel() {
  const pop = $('#qf-emp-pop');
  if (!pop || !pop.classList.contains('show')) return;
  pop.classList.remove('show');
  const btn = $('#qf-emp-btn');
  if (btn) btn.setAttribute('aria-expanded', 'false');
}

function renderEmpList(term) {
  const box = $('#qf-emp-list');
  if (!box) return;
  const f = UI.filters;
  const q = norm(f.q), dq = norm(f.duong);
  const base = ALL.filter(r => filterRowMatches(r, f, q, dq, true));
  const cnt = {};
  base.forEach(r => { cnt[r.nv] = (cnt[r.nv] || 0) + 1; });
  const t = norm(term || '');
  const list = employeesForTeam().filter(nv => !t || norm(nv).indexOf(t) > -1);
  let html = `<button type="button" class="qf-emp-item ${!f.nv ? 'on' : ''}" data-nv="">
      <span class="qei-name">Tất cả nhân viên</span>
      <span class="qei-count">${nf(base.length)}</span>
    </button>`;
  html += list.map(nv => {
    return `<button type="button" class="qf-emp-item ${f.nv === nv ? 'on' : ''}" data-nv="${esc(nv)}">
      <span class="qei-name">${esc(empName(nv))}<small>${esc(empCode(nv) || 'Không mã')}</small></span>
      <span class="qei-count">${nf(cnt[nv] || 0)}</span>
    </button>`;
  }).join('');
  box.innerHTML = list.length ? html : '<div class="qf-empty">Không tìm thấy nhân viên phù hợp</div>';
}

function setQuickNv(nv) {
  UI.filters.nv = nv || '';
  UI.page = 1;
  closeEmpPanel();
  renderList();
}

function setQuickThay(v) {
  UI.filters.thay = v;
  UI.page = 1;
  renderList();
}

function renderFiltersUI() {
  $$('#need-chips .need-chip').forEach(b => b.classList.toggle('on', UI.filters.need.indexOf(b.dataset.svc) > -1));
  const fc = $('#filter-count');
  if (fc) {
    const n = [UI.filters.q, UI.filters.to, UI.filters.kv, UI.filters.nv, UI.filters.xa, UI.filters.ap, UI.filters.duong,
      UI.filters.thay !== 'chua' ? 'x' : '', UI.filters.cham !== 'all' ? 'x' : '', UI.filters.hen !== 'all' ? 'x' : ''].filter(Boolean).length + UI.filters.need.length;
    fc.textContent = n; fc.style.display = n ? 'inline-block' : 'none';
  }
  $('#f-cham').value = UI.filters.cham;
  $('#f-kv').value = UI.filters.kv;
  $('#f-hen').value = UI.filters.hen;
  $('#f-sort').value = UI.filters.sort;
  $('#top15-note').classList.toggle('hidden', !UI.filters.topN);
  $('#btn-top15').classList.toggle('primary', !!UI.filters.topN);
}

function renderList() {
  if (!LOADED) return;
  renderQuickFilters();
  const arr = currentList();
  const total = arr.length;
  const pages = Math.max(1, Math.ceil(total / UI.pageSize));
  if (UI.page > pages) UI.page = pages;
  if (UI.page < 1) UI.page = 1;
  const start = (UI.page - 1) * UI.pageSize;
  const slice = arr.slice(start, start + UI.pageSize);

  $('#result-count').textContent = nf(total);
  const touched = arr.filter(r => LATEST.has(r.ma)).length;
  $('#result-extra').textContent = total ? `· đã chạm ${nf(touched)} · chưa chạm ${nf(total - touched)}` : '';
  $('#badge-list').textContent = nf(ALL.filter(r => !r.daThay && !LATEST.has(r.ma)).length);

  const body = $('#list-body');
  if (!total) {
    body.innerHTML = '';
    $('#list-empty').classList.remove('hidden');
  } else {
    $('#list-empty').classList.add('hidden');
    body.innerHTML = slice.map((r, i) => {
      const t = LATEST.get(r.ma);
      const addrMain = [r.sn, r.duong].filter(Boolean).join(' ') || r.dc;
      const addrSub = [r.ap, r.xa].filter(Boolean).join(', ');
      const ap = APPT_BY_MA.get(r.ma);
      return `<tr class="${t ? 'touched' : ''} ${ap ? 'has-appt' : ''}" data-act="open" data-ma="${esc(r.ma)}">
        <td class="col-stt hide-sm">${start + i + 1}</td>
        <td>
          <div class="cust-name">${esc(displayName(r))}</div>
          <div class="cust-meta"><span class="mono">${esc(r.ma)}</span>${phoneLink(r)}</div>
        </td>
        <td>
          <div class="addr"><div class="addr-main">${esc(addrMain)}</div><div class="addr-sub">${esc(addrSub)}</div></div>
        </td>
        <td class="hide-sm">
          <div class="emp-name">${esc(empName(r.nv))}</div>
          <div class="emp-team">${esc(shortTeam(r.to))} · ${esc(r.kv)}</div>
        </td>
        <td><div class="chips">${gapChips(r)}</div></td>
        <td>${statusCell(t, r)}</td>
        <td><div class="row-actions">
          ${phoneIcon(r)}
          <button type="button" class="icon-btn" data-act="open" data-ma="${esc(r.ma)}" title="Nhập phiếu chạm">✏️</button>
        </div></td>
      </tr>`;
    }).join('');
  }
  renderPager(total, pages, start, slice.length);
  renderMobileCards(slice);
}

function phoneLink(r) {
  const p = displayPhone(r);
  if (!p) return '<span class="muted" style="font-size:11.5px">Chưa có SĐT</span>';
  const tel = cleanPhone(p);
  return `<a class="phone-link" href="tel:${esc(tel)}" title="Gọi ${esc(p)}">${esc(p)}</a>`;
}
function phoneIcon(r) {
  const tel = cleanPhone(r.sdt);
  if (!tel || tel.length < 8) return '';
  return `<a class="icon-btn" href="tel:${esc(tel)}" title="Gọi ${esc(displayPhone(r))}">📞</a>`;
}
function renderMobileCards(slice) {
  const box = $('#list-cards');
  if (!box) return;
  box.innerHTML = slice.map(r => {
    const t = LATEST.get(r.ma);
    const addrMain = [r.sn, r.duong].filter(Boolean).join(' ') || r.dc;
    const addrSub = [r.ap, r.xa].filter(Boolean).join(', ');
    return `<div class="mob-card ${t ? 'touched' : ''}" data-act="open" data-ma="${esc(r.ma)}">
      <div class="mc-top">
        <div style="min-width:0">
          <div class="mc-name">${esc(displayName(r))}</div>
          <div class="mc-meta"><span class="mono">${esc(r.ma)}</span>${phoneLink(r)}</div>
        </div>
        ${t ? `<span class="badge ${(RESULT[t.ket_qua] || {}).cls || ''}">${esc((RESULT[t.ket_qua] || {}).label || t.ket_qua)}</span>` : '<span class="badge pending">Chưa chạm</span>'}
      </div>
      <div class="mc-addr">${esc(addrMain)}<span>${esc(addrSub)}</span></div>
      <div class="mc-bottom">
        <div class="chips">${gapChips(r)}</div>
        <button type="button" class="btn primary sm" data-act="open" data-ma="${esc(r.ma)}">Chạm →</button>
      </div>
    </div>`;
  }).join('');
}

function gapChips(r) {
  const g = gapKeys(r);
  if (!g.length) return '<span class="chip none">Đủ 4 dịch vụ</span>';
  return g.map(k => `<span class="chip ${k}">${k === 'tt' ? 'Trả sau' : 'Thiếu ' + SVC[k].short}</span>`).join('');
}

function apptMini(a) {
  if (!a) return '';
  const today = todayISO();
  const cls = a.trang_thai === 'DA_THAY' ? 'hen-da' : a.trang_thai === 'HUY' ? 'hen-huy' : 'hen-cho';
  const over = a.trang_thai === 'CHO_THAY' && a.ngay_hen < today;
  return `<div class="touch-when" style="margin-top:4px"><span class="badge ${over ? 'res-tu_choi' : cls}" style="padding:1px 6px;font-size:10px">📅 ${over ? 'Quá hẹn ' : ''}${fmtDate(a.ngay_hen)}${a.gio_hen ? ' ' + esc(a.gio_hen) : ''}</span></div>`;
}

function statusCell(t, r) {
  const ap = r ? APPT_BY_MA.get(r.ma) : null;
  if (!t) return '<span class="badge pending">Chưa chạm</span>' + apptMini(ap);
  const res = RESULT[t.ket_qua] || { label: t.ket_qua, cls: '' };
  return `<span class="badge ${res.cls}">${esc(res.label)}</span>
    <div class="touch-when"><span class="badge unit-${t.donvi === 'TTVT' ? 'ttvt' : 'vnpt'}" style="padding:1px 6px;font-size:10px">${esc(unitLabel(t.donvi))}</span> ${fmtDate(t.ngay)}</div>` + apptMini(ap);
}

function renderPager(total, pages, start, count) {
  const el = $('#pager-list');
  if (!el) return;
  const nums = [];
  const push = n => { if (nums.indexOf(n) === -1) nums.push(n); };
  push(1);
  for (let i = UI.page - 2; i <= UI.page + 2; i++) if (i > 1 && i < pages) push(i);
  if (pages > 1) push(pages);
  nums.sort((a, b) => a - b);
  let parts = [];
  let prev = 0;
  nums.forEach(n => {
    if (n - prev > 1) parts.push('<span class="muted" style="padding:0 2px">…</span>');
    parts.push(`<button type="button" class="pg ${n === UI.page ? 'active' : ''}" data-act="page" data-p="${n}">${n}</button>`);
    prev = n;
  });
  el.innerHTML = `<div class="pager-info">Hiển thị <b>${total ? start + 1 : 0}–${start + count}</b> / ${nf(total)} hộ
      <select data-act="pagesize" style="margin-left:8px">
        ${[25, 50, 100, 200].map(n => `<option value="${n}" ${n === UI.pageSize ? 'selected' : ''}>${n}/trang</option>`).join('')}
      </select></div>
    <div class="pager-btns">
      <button type="button" class="pg" data-act="page" data-p="${UI.page - 1}" ${UI.page <= 1 ? 'disabled' : ''}>‹</button>
      ${parts.join('')}
      <button type="button" class="pg" data-act="page" data-p="${UI.page + 1}" ${UI.page >= pages ? 'disabled' : ''}>›</button>
    </div>`;
}

/* ---------- phiếu chạm / hẹn thay ---------- */
function initials(name) {
  const parts = String(name || '').trim().split(/\s+/).filter(Boolean);
  if (!parts.length) return '?';
  return parts.slice(-2).map(w => w[0]).join('').toUpperCase();
}

function userLine() {
  if (!USER) return '';
  return `<div class="user-line">
    <span class="avatar sm">${esc(initials(USER.name))}</span>
    <div class="ul-meta"><b>${esc(USER.name)}</b><span>${esc(USER.role)} · ${esc(unitLabel(USER.unit))}</span></div>
    <span class="ul-auto">Tự động ghi vào sheet</span>
  </div>`;
}

function openDrawer(ma, mode) {
  const r = BY_MA.get(ma);
  if (!r) return;
  CURRENT_MA = ma;
  FORM_MODE = mode === 'appt' ? 'appt' : 'touch';
  FORM_CLIENT_ID = apiUid();
  FORM_SVC = { tu_van: new Set(), dong_y: new Set() };
  $('#dr-title').textContent = FORM_MODE === 'appt' ? 'Hẹn thay modem · ' + displayName(r) : 'Phiếu chạm · ' + displayName(r);
  $('#dr-sub').textContent = `${r.ma} · ${shortTeam(r.to)} · ${r.kv}`;
  $('#dr-body').innerHTML = drawerHtml(r, gapKeys(r));
  updateDrawerFooter();
  $('#overlay').classList.add('show');
  $('#drawer').classList.add('open');
  $('#drawer').setAttribute('aria-hidden', 'false');
  document.body.style.overflow = 'hidden';
  setTimeout(() => {
    const el = FORM_MODE === 'appt' ? $('#a-ngay') : $('#f-kq');
    if (el) el.focus();
  }, 120);
}

function closeDrawer() {
  $('#overlay').classList.remove('show');
  $('#drawer').classList.remove('open');
  $('#drawer').setAttribute('aria-hidden', 'true');
  document.body.style.overflow = '';
  CURRENT_MA = null;
}

function updateDrawerFooter() {
  const btn = $('#btn-save');
  if (!btn) return;
  btn.disabled = false;
  btn.textContent = FORM_MODE === 'appt' ? '📅 Lưu hẹn thay' : '💾 Lưu phiếu chạm';
  $('#dr-hint').textContent = FORM_MODE === 'appt'
    ? 'Không tạo phiếu chạm — khách chuyển sang tab KH hẹn thay'
    : 'Ghi trực tiếp vào Google Sheet';
}

function setFormMode(mode) {
  if (!CURRENT_MA || !BY_MA.get(CURRENT_MA)) return;
  if (FORM_MODE === mode) return;
  FORM_MODE = mode === 'appt' ? 'appt' : 'touch';
  FORM_SVC = { tu_van: new Set(), dong_y: new Set() };
  const r = BY_MA.get(CURRENT_MA);
  $('#dr-title').textContent = FORM_MODE === 'appt' ? 'Hẹn thay modem · ' + displayName(r) : 'Phiếu chạm · ' + displayName(r);
  $('#dr-body').innerHTML = drawerHtml(r, gapKeys(r));
  updateDrawerFooter();
}

function drawerHtml(r, g) {
  const t = LATEST.get(r.ma);
  const active = APPT_BY_MA.get(r.ma);
  const latestAppt = APPT_LATEST.get(r.ma);
  const hist = (TOUCH_BY_MA.get(r.ma) || []).slice().sort((a, b) => String(b.ts).localeCompare(String(a.ts)));
  const apptHist = APPOINTMENTS.filter(a => a.ma === r.ma).slice().sort((a, b) => String(b.ngay_tao || '').localeCompare(String(a.ngay_tao || '')));
  const chips = [
    `<span class="svc-status ${r.traSau ? 'miss' : 'have'}">💳 ${r.traSau ? 'Đang trả sau' : 'Đang trả trước'}</span>`,
    `<span class="svc-status ${r.mytv ? 'have' : 'miss'}">📺 ${r.mytv ? 'Đã có MyTV' : 'Chưa có MyTV'}</span>`,
    `<span class="svc-status ${r.mesh ? 'have' : 'miss'}">📶 ${r.mesh ? 'Đã có Mesh' : 'Chưa có Mesh'}</span>`,
    `<span class="svc-status ${r.cam ? 'have' : 'miss'}">📷 ${r.cam ? 'Đã có Cam' : 'Chưa có Cam'}</span>`,
    `<span class="svc-status neutral">${r.daThay ? '✓ Đã thay modem' : '⏳ Chưa thay modem'}</span>`,
    active ? `<span class="svc-status" style="background:#eeecfd;color:#4c3fb3">📅 Hẹn ${fmtDate(active.ngay_hen)} ${esc(active.gio_hen || '')}</span>` : ''
  ].join('');

  const touchRows = SVC_KEYS.map(k => {
    const have = g.indexOf(k) === -1;
    const note = k === 'tt'
      ? (have ? 'Đang trả trước — không cần chuyển' : 'Đang trả sau — có thể chuyển trả trước')
      : (have ? 'Khách đã có thiết bị/dịch vụ' : 'Chưa có — cần tiếp thị');
    return `<div class="svc-row ${have ? 'have' : 'miss'}" data-svc="${k}">
      <div class="svc-ico">${SVC[k].icon}</div>
      <div class="svc-main"><div class="svc-name">${SVC[k].label}</div><div class="svc-note">${note}</div></div>
      ${have
        ? `<span class="svc-ok">✓ ${k === 'tt' ? 'Đang trả trước' : 'Đã có'}</span>`
        : `<div class="svc-checks">
            <label class="chk"><input type="checkbox" data-svc-chk="tu_van" data-svc="${k}"> Tư vấn</label>
            <label class="chk agree"><input type="checkbox" data-svc-chk="dong_y" data-svc="${k}"> Đồng ý</label>
           </div>`}
    </div>`;
  }).join('');

  const apptRows = SVC_KEYS.map(k => {
    const have = g.indexOf(k) === -1;
    const selected = active ? (active.dich_vu || []).indexOf(k) > -1 : false;
    return `<label class="svc-row appt-svc-row ${have ? 'have' : 'miss'} ${selected ? 'picked' : ''}">
      <input type="checkbox" data-appt-svc="${k}" ${have ? 'disabled' : ''} ${selected ? 'checked' : ''}>
      <div class="svc-ico">${SVC[k].icon}</div>
      <div class="svc-main"><div class="svc-name">${SVC[k].label}</div>
        <div class="svc-note">${have ? (k === 'tt' ? 'Đang trả trước' : 'Đã có') : 'Khách chưa có — có thể tiếp thị'}</div></div>
      <span class="svc-ok">${have ? '—' : (selected ? 'Đã chọn' : 'Chọn')}</span>
    </label>`;
  }).join('');

  const scripts = g.length ? `<details class="script-item" open>
      <summary>🗣 Kịch bản tiếp thị gợi ý (${g.length} dịch vụ)</summary>
      <div class="script-txt" style="padding-left:11px">
        ${g.map(k => `<div style="margin-bottom:9px"><b style="font-size:12px">${SVC[k].icon} ${SVC[k].label}</b>
          <div>${esc(SCRIPT_TXT[k])}</div>
          <button type="button" class="btn sm copy" data-act="copy-script" data-key="${k}">📋 Sao chép</button></div>`).join('')}
      </div>
    </details>` : '';

  const history = hist.length ? `<div class="sec">
      <div class="sec-title">Lịch sử chạm (${hist.length})</div>
      <div class="panel" style="padding:4px 12px">
        ${hist.slice(0, 6).map(h => `<div class="history-item">
          <span class="history-dot" style="background:${h.donvi === 'TTVT' ? '#e07b00' : '#0b5cab'}"></span>
          <div style="flex:1;min-width:0">
            <div><b>${esc((RESULT[h.ket_qua] || {}).label || h.ket_qua)}</b> · ${esc(unitLabel(h.donvi))} · ${esc(h.nguoi)}</div>
            <div class="history-meta">${fmtDate(h.ngay)} ${esc(h.gio || '')}${(h.dong_y && h.dong_y.length) ? ' · Đồng ý: ' + h.dong_y.map(k => SVC[k] ? SVC[k].short : k).join(', ') : ''}${(h.tu_van && h.tu_van.length) ? ' · Tư vấn: ' + h.tu_van.map(k => SVC[k] ? SVC[k].short : k).join(', ') : ''}</div>
            ${h.ghi_chu ? `<div class="history-note">“${esc(h.ghi_chu)}”</div>` : ''}
          </div>
        </div>`).join('')}
      </div>
    </div>` : '';

  const apptHistory = apptHist.length ? `<div class="sec">
      <div class="sec-title">Lịch sử hẹn thay (${apptHist.length})</div>
      <div class="panel" style="padding:4px 12px">
        ${apptHist.slice(0, 6).map(a => `<div class="history-item">
          <span class="history-dot" style="background:${a.trang_thai === 'DA_THAY' ? '#178a5c' : a.trang_thai === 'HUY' ? '#9aa8ba' : '#6b5bd2'}"></span>
          <div style="flex:1;min-width:0">
            <div><b>${fmtDate(a.ngay_hen)} ${esc(a.gio_hen || '')}</b> · ${esc(apptStatusLabel(a.trang_thai))}</div>
            <div class="history-meta">${esc(a.nguoi)} · ${esc(unitLabel(a.donvi))}${(a.dich_vu || []).length ? ' · ' + a.dich_vu.map(k => SVC[k] ? SVC[k].short : k).join(', ') : ''}</div>
            ${a.ghi_chu ? `<div class="history-note">“${esc(a.ghi_chu)}”</div>` : ''}
          </div>
        </div>`).join('')}
      </div>
    </div>` : '';

  const modeTabs = `<div class="drawer-modes">
      <button type="button" data-mode="touch" class="${FORM_MODE === 'touch' ? 'on' : ''}">📞 Chạm tiếp thị</button>
      <button type="button" data-mode="appt" class="${FORM_MODE === 'appt' ? 'on' : ''}">📅 Gọi hẹn thay</button>
    </div>`;

  const customerSec = `
  <div class="sec">
    <div class="sec-title">Khách hàng &amp; hiện trạng</div>
    <div class="panel">
      <div class="info-grid">
        <div class="info-item full"><div class="k">Khách hàng</div><div class="v">${esc(displayName(r))} <span class="muted mono">· ${esc(r.ma)}</span></div></div>
        <div class="info-item"><div class="k">Số điện thoại</div><div class="v">${phoneLink(r)}</div></div>
        <div class="info-item"><div class="k">Mã khu vực</div><div class="v mono">${esc(r.kv)}</div></div>
        <div class="info-item full"><div class="k">Địa chỉ</div><div class="v">${esc(r.dc)}</div></div>
        <div class="info-item"><div class="k">Tổ phụ trách</div><div class="v">${esc(shortTeam(r.to))}</div></div>
        <div class="info-item"><div class="k">NV quản lý</div><div class="v">${esc(empName(r.nv))}</div></div>
        <div class="info-item"><div class="k">Modem hiện tại</div><div class="v mono">${esc(r.modem || '—')}</div></div>
        <div class="info-item"><div class="k">Lần chạm gần nhất</div><div class="v">${t ? fmtDate(t.ngay) + ' · ' + esc(unitLabel(t.donvi)) : 'Chưa chạm'}</div></div>
      </div>
      <div class="status-chips">${chips}</div>
    </div>
  </div>`;

  if (FORM_MODE === 'appt') {
    const apDate = active ? active.ngay_hen : todayISO();
    const apGio = active ? (active.gio_hen || '') : '';
    const apNote = active ? active.ghi_chu : '';
    const banner = active ? `<div class="appt-existing">
        <span>📅 Đã hẹn <b>${fmtDate(active.ngay_hen)} ${esc(active.gio_hen || '')}</b> · ${esc(apptStatusLabel(active.trang_thai))}</span>
        <span class="muted">${esc(active.nguoi)} · ${esc(unitLabel(active.donvi))}</span>
      </div>` : '';
    return `${modeTabs}${customerSec}
      <div class="sec">
        <div class="sec-title">${active ? 'Cập nhật hẹn thay' : 'Gọi hẹn thay modem'}</div>
        <div class="panel">
          ${banner}
          ${userLine()}
          <div class="callout" style="margin:10px 0 12px;padding:8px 11px;font-size:12px">
            Lưu hẹn thay <b>không tạo phiếu chạm</b>. Sau khi lưu, khách sẽ nằm trong tab <b>KH hẹn thay</b> và được sắp xếp theo ngày hẹn gần nhất.
          </div>
          <div class="form-grid">
            <div class="form-field">
              <label>Ngày hẹn thay <span class="req">*</span></label>
              <input id="a-ngay" type="date" value="${esc(apDate)}" max="2100-12-31">
            </div>
            <div class="form-field">
              <label>Giờ hẹn</label>
              <input id="a-gio" type="time" value="${esc(apGio)}">
            </div>
            <div class="form-field full">
              <label>Ghi chú hẹn <span class="req">*</span></label>
              <textarea id="a-ghichu" maxlength="400" placeholder="Bắt buộc: nội dung trao đổi, hẹn ai, lưu ý lắp đặt…">${esc(apNote)}</textarea>
            </div>
          </div>
        </div>
      </div>
      <div class="sec">
        <div class="sec-title">Dịch vụ tiếp thị kèm theo (nếu có)</div>
        <div class="panel">
          <div class="svc-rows appt-svc-rows">${apptRows}</div>
          <div class="mini-note">Có thể chọn nhiều dịch vụ. Không bắt buộc nếu khách chỉ hẹn thay modem.</div>
        </div>
      </div>
      ${apptHistory}`;
  }

  return `${modeTabs}${customerSec}
  <div class="sec">
    <div class="sec-title">Người chạm &amp; kết quả</div>
    <div class="panel">
      ${userLine()}
      <div class="form-grid" style="margin-top:10px">
        <div class="form-field">
          <label>Ngày chạm <span class="req">*</span></label>
          <input id="f-ngay" type="date" value="${todayISO()}" max="2100-12-31">
        </div>
        <div class="form-field">
          <label>Giờ chạm</label>
          <input id="f-gio" type="time" value="${nowHHMM()}">
        </div>
        <div class="form-field full">
          <label>Kết quả chạm <span class="req">*</span></label>
          <select id="f-kq">
            <option value="QUAN_TAM" selected>Quan tâm — cần theo tiếp</option>
            <option value="DA_CHOT">Đã chốt — khách đồng ý dịch vụ</option>
            <option value="HEN_LAI">Hẹn lại — khách bận, chưa quyết</option>
            <option value="TU_CHOI">Từ chối — không có nhu cầu</option>
            <option value="KHONG_LIEN_HE">Không liên hệ được</option>
          </select>
        </div>
        <div class="form-field full">
          <label>Ghi chú</label>
          <textarea id="f-ghichu" placeholder="Ví dụ: khách hẹn 17h mai gọi lại, chỉ gắn 2 mesh…" maxlength="300"></textarea>
        </div>
      </div>
    </div>
  </div>

  <div class="sec">
    <div class="sec-title">Dịch vụ cần tiếp thị</div>
    <div class="panel">
      <div class="svc-rows">${touchRows}</div>
      <div class="mini-note">Chọn <b>Tư vấn</b> khi đã giới thiệu; chọn <b>Đồng ý</b> khi khách chấp thuận để đưa vào thống kê chốt dịch vụ.</div>
    </div>
  </div>

  <div class="sec">${scripts}</div>
  ${history}`;
}

function syncServiceChecks() {
  $$('#dr-body [data-svc-chk]').forEach(inp => {
    const kind = inp.dataset.svcChk, k = inp.dataset.svc;
    const on = kind === 'tu_van' ? FORM_SVC.tu_van.has(k) : FORM_SVC.dong_y.has(k);
    inp.checked = on;
    const label = inp.closest('.chk');
    if (label) label.classList.toggle('on', on);
  });
}

async function saveTouch() {
  const r = BY_MA.get(CURRENT_MA);
  if (!r) return;
  const ngay = $('#f-ngay').value;
  const gio = $('#f-gio').value;
  const ketQua = $('#f-kq').value;
  const ghiChu = ($('#f-ghichu').value || '').trim();

  if (!ngay) { toast('Chọn ngày chạm', 'err'); $('#f-ngay').focus(); return; }
  if (ketQua === 'DA_CHOT' && FORM_SVC.dong_y.size === 0) {
    toast('Kết quả "Đã chốt" cần chọn ít nhất 1 dịch vụ khách đồng ý', 'err');
    return;
  }

  const payload = {
    ma: r.ma,
    ngay: ngay,
    gio: gio,
    ket_qua: ketQua,
    tu_van: Array.from(FORM_SVC.tu_van),
    dong_y: Array.from(FORM_SVC.dong_y),
    ghi_chu: ghiChu,
    client_id: FORM_CLIENT_ID
  };
  const btn = $('#btn-save');
  btn.disabled = true;
  btn.textContent = '⏳ Đang ghi vào sheet…';
  try {
    const res = await API.saveTouch(payload);
    if (!res || !res.touch) throw new Error('Không nhận được dữ liệu xác nhận từ sheet');
    TOUCHES.push(res.touch);
    recomputeTouches();
    FORM_CLIENT_ID = null;
    closeDrawer();
    renderAll();
    toast('Đã lưu phiếu chạm — người nhập và đơn vị tự động ghi từ tài khoản', 'ok');
    if (res.warning) toast(res.warning, 'warn', 5000);
  } catch (e) {
    btn.disabled = false;
    btn.textContent = '💾 Lưu phiếu chạm';
    if (e.code === 'AUTH') { closeDrawer(); API.clearToken(); showLogin(e.message); return; }
    toast('Lỗi khi ghi vào sheet: ' + e.message, 'err', 7000);
  }
}

async function saveAppointment() {
  const r = BY_MA.get(CURRENT_MA);
  if (!r) return;
  const ngay = $('#a-ngay').value;
  const gio = $('#a-gio').value;
  const ghiChu = ($('#a-ghichu').value || '').trim();
  const dichVu = $$('#dr-body [data-appt-svc]').filter(x => x.checked && !x.disabled).map(x => x.dataset.apptSvc);
  if (!ngay) { toast('Chọn ngày hẹn thay', 'err'); $('#a-ngay').focus(); return; }
  if (!ghiChu) { toast('Ghi chú hẹn thay là bắt buộc', 'err'); $('#a-ghichu').focus(); return; }

  const btn = $('#btn-save');
  btn.disabled = true;
  btn.textContent = '⏳ Đang lưu hẹn…';
  try {
    const res = await API.saveAppointment({ ma: r.ma, ngay_hen: ngay, gio_hen: gio, dich_vu: dichVu, ghi_chu: ghiChu });
    if (!res || !res.appointment) throw new Error('Không nhận được xác nhận hẹn thay');
    const i = APPOINTMENTS.findIndex(a => a.id === res.appointment.id);
    if (i > -1) APPOINTMENTS[i] = res.appointment; else APPOINTMENTS.push(res.appointment);
    recomputeAppointments();
    closeDrawer();
    renderAll();
    showTab('appt');
    toast(res.updated ? 'Đã cập nhật hẹn thay' : 'Đã lưu hẹn thay — khách chuyển vào mục KH hẹn thay', 'ok');
  } catch (e) {
    btn.disabled = false;
    btn.textContent = '📅 Lưu hẹn thay';
    if (e.code === 'AUTH') { closeDrawer(); API.clearToken(); showLogin(e.message); return; }
    toast('Lỗi khi lưu hẹn: ' + e.message, 'err', 7000);
  }
}

function apptStatusLabel(s) {
  return s === 'DA_THAY' ? 'Đã thay' : s === 'HUY' ? 'Đã hủy' : 'Chờ thay';
}

function renderAll() {
  renderKPIs();
  renderBadges();
  renderFiltersUI();
  renderList();
  renderUserUI();
  if (UI.tab === 'appt') renderAppointments();
  if (UI.tab === 'ttvt') renderUnit('TTVT');
  if (UI.tab === 'vnpt') renderUnit('VNPT');
  if (UI.tab === 'report') renderReport();
  if (UI.tab === 'stats') renderStats();
}

function renderUserUI() {
  const u = USER;
  if (!u) return;
  const nm = $('#user-name'), role = $('#user-role'), av = $('#user-avatar');
  if (nm) nm.textContent = u.name || u.username;
  if (role) role.textContent = (u.role || '') + ' · ' + unitLabel(u.unit);
  if (av) av.textContent = initials(u.name);
  const umName = $('#um-name'), umSub = $('#um-sub'), umAreas = $('#um-areas');
  if (umName) umName.textContent = u.name || u.username;
  if (umSub) umSub.textContent = (u.role || '') + ' · ' + unitLabel(u.unit) + ' · ' + (u.username || '');
  if (umAreas) umAreas.textContent = u.full ? 'Toàn quyền: xem tất cả khu vực' : ('Khu vực: ' + (u.areas || []).join(', '));
  const must = $('#must-change');
  if (must) must.classList.toggle('hidden', !u.must_change);
  const pwUser = $('#pw-user');
  if (pwUser) pwUser.textContent = u.username || '';
}


/* ============================================================
   CÁC MÀN HÌNH: TTVT / VNPT, BÁO CÁO, THỐNG KÊ, HƯỚNG DẪN
   ============================================================ */

/* ---------- lọc lượt chạm ---------- */
function filterTouches(base, f) {
  const q = norm(f.q || '');
  return base.filter(t => {
    const r = BY_MA.get(t.ma) || {};
    if (f.from && t.ngay < f.from) return false;
    if (f.to && t.ngay > f.to) return false;
    if (f.donvi && t.donvi !== f.donvi) return false;
    if (f.emp && t.nguoi !== f.emp) return false;
    if (f.kq && t.ket_qua !== f.kq) return false;
    if (f.svc && ((t.tu_van || []).indexOf(f.svc) === -1) && ((t.dong_y || []).indexOf(f.svc) === -1)) return false;
    if (f.team && r.to !== f.team) return false;
    if (f.xa && r.xa !== f.xa) return false;
    if (q) {
      const hay = norm(t.ma) + ' ' + norm(t.ten || r.ten) + ' ' + norm(t.nguoi) + ' ' + norm(t.sdt || r.sdt) + ' ' + norm(t.dc || r.dc) + ' ' + norm(r.duong) + ' ' + norm(r.ap) + ' ' + norm(r.xa);
      if (hay.indexOf(q) === -1) return false;
    }
    return true;
  }).sort((a, b) => String(b.ts).localeCompare(String(a.ts)));
}

function touchCustomer(t) {
  const r = BY_MA.get(t.ma) || {};
  const pick = (a, b) => cleanText(a) || cleanText(b);
  return {
    ten: pick(t.ten, r.ten), sdt: pick(t.sdt, r.sdt), to: pick(t.to, r.to),
    nv: pick(t.nv, r.nv), xa: pick(t.xa, r.xa), ap: pick(t.ap, r.ap),
    duong: pick(t.duong, r.duong), sn: pick(t.sn, r.sn), dc: pick(t.dc, r.dc),
    daThay: r.daThay, kv: r.kv
  };
}

function svcChips(list, kind) {
  if (!list || !list.length) return '<span class="chip none">—</span>';
  return list.map(k => `<span class="chip ${k === 'tt' ? 'tt' : k}" ${kind === 'tu_van' ? 'style="opacity:.85"' : ''}>${esc(SVC[k] ? SVC[k].short : k)}</span>`).join('');
}

function touchRowHtml(t, opts) {
  opts = opts || {};
  const c = touchCustomer(t);
  const res = RESULT[t.ket_qua] || { label: t.ket_qua, cls: '' };
  const addr = [c.sn, c.duong].filter(Boolean).join(' ') || c.dc;
  const cnt = (t.dong_y || []).length;
  return `<tr data-act="open" data-ma="${esc(t.ma)}">
    <td class="col-stt hide-sm">${opts.stt || ''}</td>
    <td><div style="font-weight:700">${fmtDate(t.ngay)}</div><div class="touch-when">${esc(t.gio || '')}</div></td>
    <td>
      <div class="cust-name">${esc(displayName(c))}</div>
      <div class="cust-meta"><span class="mono">${esc(t.ma)}</span>${/^0\d{8,10}$/.test(cleanPhone(c.sdt)) ? `<a class="phone-link" href="tel:${esc(cleanPhone(c.sdt))}">${esc(c.sdt)}</a>` : '<span class="muted" style="font-size:11.5px">Chưa có SĐT</span>'}</div>
    </td>
    <td><div class="addr"><div class="addr-main">${esc(addr)}</div><div class="addr-sub">${esc([c.ap, c.xa].filter(Boolean).join(', '))}</div></div></td>
    <td class="hide-sm"><div class="emp-name">${esc(empName(c.nv))}</div><div class="emp-team">${esc(shortTeam(c.to))}</div></td>
    ${opts.showUnit ? `<td><span class="badge unit-${t.donvi === 'TTVT' ? 'ttvt' : 'vnpt'}">${esc(unitLabel(t.donvi))}</span></td>` : ''}
    <td><div class="emp-name" style="font-weight:600">${esc(t.nguoi)}</div></td>
    <td><span class="badge ${res.cls}">${esc(res.label)}</span></td>
    <td><div class="chips" style="max-width:190px">${svcChips(t.dong_y, 'dong_y')}</div>${(t.tu_van || []).length ? `<div class="touch-when">Tư vấn: ${(t.tu_van || []).map(k => esc(SVC[k] ? SVC[k].short : k)).join(', ')}</div>` : ''}</td>
    <td class="hide-sm"><div style="max-width:170px;font-size:12px;color:#4a5b70">${esc(t.ghi_chu || '')}</div></td>
    <td><button type="button" class="icon-btn" data-act="del-touch" data-id="${esc(t.id)}" title="Xóa lượt chạm" style="color:#c44">🗑</button></td>
  </tr>`;
}

/* ---------- TTVT / VNPT ---------- */
function unitTouches(unit, f) {
  const base = TOUCHES.filter(t => t.donvi === unit);
  return filterTouches(base, f);
}

function renderUnit(unit) {
  const key = unit === 'TTVT' ? 'ttvt' : 'vnpt';
  const root = $('#unit-' + key);
  if (!root || !LOADED) return;
  const st = UI.unit[unit];
  const base = TOUCHES.filter(t => t.donvi === unit);
  const ho = new Set(base.map(t => t.ma));
  const today = base.filter(t => t.ngay === todayISO()).length;
  const chot = new Set(base.filter(t => (t.dong_y || []).length).map(t => t.ma));
  const svc = { tt: new Set(), mytv: new Set(), mesh: new Set(), cam: new Set() };
  base.forEach(t => (t.dong_y || []).forEach(k => svc[k] && svc[k].add(t.ma)));
  const emps = uniq(base.map(t => t.nguoi)).sort(byName);

  root.innerHTML = `
  <div class="unit-hero ${key}">
    <div>
      <h2>${unit === 'TTVT' ? '🟠 TTVT đã chạm' : '🔵 VNPT KV đã chạm'}</h2>
      <p>Toàn bộ lượt chạm do nhân viên <b>${unitLabel(unit)}</b> nhập, ghi trực tiếp vào sheet <code style="color:#fff;background:rgba(255,255,255,.18);padding:1px 5px;border-radius:4px">CHAM_TIEP_THI</code> của Google Sheet.</p>
    </div>
    <div class="unit-kpis">
      <div class="ukpi"><b>${nf(base.length)}</b><span>Lượt chạm</span></div>
      <div class="ukpi"><b>${nf(ho.size)}</b><span>Hộ đã chạm</span></div>
      <div class="ukpi"><b>${nf(today)}</b><span>Hôm nay</span></div>
      <div class="ukpi"><b>${nf(chot.size)}</b><span>Hộ đồng ý</span></div>
    </div>
  </div>

  <div class="card">
    <div class="toolbar">
      <div class="fg grow">
        <label>Tìm khách, SĐT, mã TB, người nhập</label>
        <input id="u-q-${unit}" value="${esc(st.q)}" placeholder="Gõ để tìm…" autocomplete="off">
      </div>
      <div class="fg"><label>Từ ngày</label><input type="date" id="u-from-${unit}" value="${esc(st.from)}"></div>
      <div class="fg"><label>Đến ngày</label><input type="date" id="u-to-${unit}" value="${esc(st.to)}"></div>
      <div class="fg"><label>Người nhập</label>
        <select id="u-emp-${unit}"><option value="">Tất cả</option>${emps.map(e => `<option value="${esc(e)}" ${e === st.emp ? 'selected' : ''}>${esc(e)}</option>`).join('')}</select>
      </div>
      <button type="button" class="btn sm" data-act="reset-unit" data-unit="${unit}">↺ Xóa lọc</button>
      <div class="spacer"></div>
      <button type="button" class="btn primary sm" data-act="export-unit" data-unit="${unit}">⬇ Tải Excel</button>
    </div>
    <div style="padding:10px 16px;border-bottom:1px solid var(--line);display:flex;gap:14px;flex-wrap:wrap;font-size:12px;color:var(--gray)" id="u-sum-${unit}"></div>
    <div class="table-wrap" style="max-height:none">
      <table>
        <thead><tr>
          <th class="col-stt hide-sm">#</th><th>Ngày chạm</th><th>Khách hàng</th><th>Địa chỉ</th>
          <th class="hide-sm">Tổ / NV</th><th>Người nhập</th><th>Kết quả</th><th>Dịch vụ đồng ý</th>
          <th class="hide-sm">Ghi chú</th><th></th>
        </tr></thead>
        <tbody id="u-body-${unit}"></tbody>
      </table>
    </div>
    <div id="u-empty-${unit}" class="empty hidden"><div class="big">📭</div><div>Chưa có lượt chạm nào của ${unitLabel(unit)} khớp bộ lọc.</div></div>
    <div class="pager"><div class="pager-info" id="u-count-${unit}"></div></div>
  </div>`;

  ['q', 'from', 'to', 'emp'].forEach(k => {
    const el = $('#u-' + k + '-' + unit);
    if (!el) return;
    const ev = k === 'q' ? 'input' : 'change';
    el.addEventListener(ev, debounce(() => {
      UI.unit[unit][k] = el.value;
      refreshUnitTable(unit);
    }, k === 'q' ? 220 : 0));
  });
  refreshUnitTable(unit);
}

function refreshUnitTable(unit) {
  const st = UI.unit[unit];
  const list = unitTouches(unit, st);
  const body = $('#u-body-' + unit);
  if (!body) return;
  const shown = list.slice(0, 400);
  body.innerHTML = shown.map((t, i) => touchRowHtml(t, { stt: i + 1 })).join('');
  $('#u-empty-' + unit).classList.toggle('hidden', list.length > 0);
  const ho = new Set(list.map(t => t.ma));
  const chot = new Set(list.filter(t => (t.dong_y || []).length).map(t => t.ma));
  $('#u-sum-' + unit).innerHTML = `
    <span><b style="color:var(--ink)">${nf(list.length)}</b> lượt</span>
    <span><b style="color:var(--ink)">${nf(ho.size)}</b> hộ</span>
    <span><b style="color:var(--green)">${nf(chot.size)}</b> hộ đồng ý</span>
    <span>Đơn vị: <b style="color:var(--ink)">${unitLabel(unit)}</b></span>`;
  $('#u-count-' + unit).innerHTML = list.length
    ? `Hiển thị ${nf(shown.length)} / ${nf(list.length)} lượt chạm${list.length > 400 ? ' — tải Excel để xem đầy đủ' : ''}`
    : 'Chưa có dữ liệu';
}

/* ---------- BÁO CÁO CHI TIẾT ---------- */
function renderReport() {
  const root = $('#report-root');
  if (!root || !LOADED) return;
  const f = UI.report;
  const teams = uniq(ALL.map(r => r.to)).sort(byName);
  const xas = uniq(ALL.map(r => r.xa)).sort(byName);
  const emps = uniq(TOUCHES.map(t => t.nguoi)).sort(byName);
  const svcOpts = SVC_KEYS.map(k => `<option value="${k}" ${f.svc === k ? 'selected' : ''}>${SVC[k].label}</option>`).join('');
  const kqOpts = Object.keys(RESULT).map(k => `<option value="${k}" ${f.kq === k ? 'selected' : ''}>${RESULT[k].label}</option>`).join('');

  root.innerHTML = `
  <div class="card" style="margin-bottom:14px">
    <div class="card-h no-print">
      <div><div class="card-title">📑 Báo cáo chi tiết lượt chạm <span class="sub">có cột ngày chạm</span></div>
        <div class="sub" style="font-size:11.5px;color:var(--gray);margin-top:2px">Dữ liệu lấy trực tiếp từ sheet CHAM_TIEP_THI — lọc, xem và tải Excel</div>
      </div>
      <div style="display:flex;gap:8px">
        <button type="button" class="btn sm" data-act="print-report">🖨 In</button>
        <button type="button" class="btn primary sm" data-act="export-report">⬇ Tải Excel</button>
      </div>
    </div>
    <div class="toolbar no-print">
      <div class="fg grow"><label>Tìm kiếm</label><input id="rp-q" value="${esc(f.q)}" placeholder="Tên khách, SĐT, mã TB, người nhập…" autocomplete="off"></div>
      <div class="fg"><label>Từ ngày</label><input type="date" id="rp-from" value="${esc(f.from)}"></div>
      <div class="fg"><label>Đến ngày</label><input type="date" id="rp-to" value="${esc(f.to)}"></div>
      <div class="fg"><label>Đơn vị</label><select id="rp-donvi"><option value="">Tất cả</option><option value="TTVT" ${f.donvi === 'TTVT' ? 'selected' : ''}>TTVT</option><option value="VNPT" ${f.donvi === 'VNPT' ? 'selected' : ''}>VNPT KV</option></select></div>
      <div class="fg"><label>Tổ</label><select id="rp-team"><option value="">Tất cả tổ</option>${teams.map(t => `<option value="${esc(t)}" ${f.team === t ? 'selected' : ''}>${esc(shortTeam(t))}</option>`).join('')}</select></div>
      <div class="fg"><label>Nhân viên nhập</label><select id="rp-emp"><option value="">Tất cả</option>${emps.map(e => `<option value="${esc(e)}" ${f.emp === e ? 'selected' : ''}>${esc(e)}</option>`).join('')}</select></div>
      <div class="fg"><label>Xã</label><select id="rp-xa"><option value="">Tất cả xã</option>${xas.map(x => `<option value="${esc(x)}" ${f.xa === x ? 'selected' : ''}>${esc(x)}</option>`).join('')}</select></div>
      <div class="fg"><label>Kết quả</label><select id="rp-kq"><option value="">Tất cả</option>${kqOpts}</select></div>
      <div class="fg"><label>Dịch vụ</label><select id="rp-svc"><option value="">Tất cả</option>${svcOpts}</select></div>
      <button type="button" class="btn sm" data-act="reset-report">↺ Xóa lọc</button>
    </div>
    <div style="padding:10px 16px;border-bottom:1px solid var(--line)" id="rp-sum"></div>
    <div class="table-wrap" style="max-height:none">
      <table>
        <thead><tr>
          <th class="col-stt hide-sm">#</th><th>Ngày chạm</th><th>Khách hàng</th><th>Địa chỉ</th>
          <th class="hide-sm">Tổ / NV</th><th>Đơn vị</th><th>Người nhập</th><th>Kết quả</th>
          <th>Dịch vụ</th><th class="hide-sm">Ghi chú</th><th class="no-print"></th>
        </tr></thead>
        <tbody id="rp-body"></tbody>
      </table>
    </div>
    <div id="rp-empty" class="empty hidden"><div class="big">📄</div><div>Không có lượt chạm nào khớp bộ lọc.</div></div>
    <div class="pager"><div class="pager-info" id="rp-count"></div></div>
  </div>`;

  ['q', 'from', 'to', 'donvi', 'team', 'emp', 'xa', 'kq', 'svc'].forEach(k => {
    const el = $('#rp-' + k);
    if (!el) return;
    el.addEventListener(k === 'q' ? 'input' : 'change', debounce(() => {
      UI.report[k] = el.value;
      refreshReportTable();
    }, k === 'q' ? 220 : 0));
  });
  refreshReportTable();
}

function reportList() {
  return filterTouches(TOUCHES, UI.report);
}

function refreshReportTable() {
  const list = reportList();
  const shown = list.slice(0, 500);
  const body = $('#rp-body');
  if (!body) return;
  body.innerHTML = shown.map((t, i) => touchRowHtml(t, { stt: i + 1, showUnit: true })).join('');
  $('#rp-empty').classList.toggle('hidden', list.length > 0);
  const ho = new Set(list.map(t => t.ma));
  const chot = new Set(list.filter(t => (t.dong_y || []).length).map(t => t.ma));
  const units = { TTVT: 0, VNPT: 0 };
  list.forEach(t => units[t.donvi] = (units[t.donvi] || 0) + 1);
  $('#rp-sum').innerHTML = `<div style="display:flex;gap:18px;flex-wrap:wrap;font-size:12.5px">
      <span><b style="font-size:16px">${nf(list.length)}</b> <span class="muted">lượt chạm</span></span>
      <span><b style="font-size:16px">${nf(ho.size)}</b> <span class="muted">hộ</span></span>
      <span><b style="font-size:16px;color:var(--green)">${nf(chot.size)}</b> <span class="muted">hộ đồng ý</span></span>
      <span><b style="color:#a35c00">${nf(units.TTVT)}</b> <span class="muted">TTVT</span></span>
      <span><b style="color:var(--blue)">${nf(units.VNPT)}</b> <span class="muted">VNPT KV</span></span>
    </div>`;
  $('#rp-count').innerHTML = list.length ? `Hiển thị ${nf(shown.length)} / ${nf(list.length)} lượt${list.length > 500 ? ' — tải Excel để xem đầy đủ' : ''}` : 'Chưa có dữ liệu';
}

/* ---------- THỐNG KÊ ---------- */
function renderStats() {
  const root = $('#stats-root');
  if (!root || !LOADED) return;
  const f = UI.stats;
  root.innerHTML = `
  <div class="stats-toolbar no-print">
    <div class="fg"><label>Từ ngày</label><input type="date" id="st-from" value="${esc(f.from)}"></div>
    <div class="fg"><label>Đến ngày</label><input type="date" id="st-to" value="${esc(f.to)}"></div>
    <div class="fg"><label>Đơn vị</label><select id="st-unit"><option value="">Cả hai đơn vị</option><option value="TTVT" ${f.unit === 'TTVT' ? 'selected' : ''}>TTVT</option><option value="VNPT" ${f.unit === 'VNPT' ? 'selected' : ''}>VNPT KV</option></select></div>
    <button type="button" class="btn sm" data-act="stats-this-month">Tháng này</button>
    <button type="button" class="btn sm" data-act="reset-stats">Toàn bộ thời gian</button>
    <div class="spacer" style="flex:1"></div>
    <button type="button" class="btn primary sm" data-act="export-all-stats">⬇ Tải Excel toàn bộ báo cáo</button>
  </div>
  <div id="stats-body"></div>`;
  ['from', 'to', 'unit'].forEach(k => {
    const el = $('#st-' + k);
    el.addEventListener('change', () => { UI.stats[k] = el.value; renderStatsBody(); });
  });
  renderStatsBody();
}

function statRange() {
  return filterTouches(TOUCHES, UI.stats);
}

function statTableCard(title, sub, headers, rows, key) {
  STATS_CACHE[key] = { name: title.replace(/[\\\/\?\*\[\]:]/g, ' ').slice(0, 31), rows: [headers].concat(rows) };
  return `<div class="card">
    <div class="card-h"><div><div class="card-title">${title}</div>${sub ? `<div class="sub" style="font-size:11.5px;color:var(--gray);margin-top:2px">${sub}</div>` : ''}</div>
    <button type="button" class="btn sm" data-act="export-stats" data-key="${key}">⬇ Excel</button></div>
    <div class="table-wrap" style="max-height:52vh">
      <table><thead><tr>${headers.map((h, i) => `<th${i > 0 ? ' style="text-align:right"' : ''}>${esc(h)}</th>`).join('')}</tr></thead>
      <tbody>${rows.length ? rows.map(r => `<tr style="cursor:default">${r.map((c, i) => `<td${i > 0 ? ' style="text-align:right;font-variant-numeric:tabular-nums"' : ''}>${typeof c === 'number' ? nf(c) : esc(c)}</td>`).join('')}</tr>`).join('') : `<tr><td colspan="${headers.length}" class="muted" style="text-align:center;padding:20px">Chưa có dữ liệu trong kỳ</td></tr>`}</tbody></table>
    </div>
  </div>`;
}

let STATS_CACHE = {};

function renderStatsBody() {
  const body = $('#stats-body');
  if (!body) return;
  const ts = statRange();
  const ho = new Set(ts.map(t => t.ma));
  const chotHo = new Set(ts.filter(t => (t.dong_y || []).length).map(t => t.ma));
  const daChot = new Set(ts.filter(t => t.ket_qua === 'DA_CHOT').map(t => t.ma));
  const totalHo = ALL.length;
  const coveredAll = new Set(TOUCHES.map(t => t.ma));
  const covRate = totalHo ? coveredAll.size / totalHo * 100 : 0;
  const quanTam = new Set(ts.filter(t => t.ket_qua === 'QUAN_TAM' || t.ket_qua === 'DA_CHOT').map(t => t.ma));
  const svcAgree = { tt: new Set(), mytv: new Set(), mesh: new Set(), cam: new Set() };
  const svcTuVan = { tt: new Set(), mytv: new Set(), mesh: new Set(), cam: new Set() };
  ts.forEach(t => {
    (t.tu_van || []).forEach(k => svcTuVan[k] && svcTuVan[k].add(t.ma));
    (t.dong_y || []).forEach(k => svcAgree[k] && svcAgree[k].add(t.ma));
  });
  const need = { tt: 0, mytv: 0, mesh: 0, cam: 0 };
  ALL.forEach(r => gapKeys(r).forEach(k => need[k]++));
  const homNay = ts.filter(t => t.ngay === todayISO()).length;

  /* 1. tổng quan */
  const funnel = [
    { label: 'Hộ trong danh sách', val: totalHo, color: '#0b5cab' },
    { label: 'Chưa thay modem', val: ALL.filter(r => !r.daThay).length, color: '#3b9fd4' },
    { label: 'Đã chạm', val: coveredAll.size, color: '#0e8f8a' },
    { label: 'Khách quan tâm', val: quanTam.size, color: '#6b5bd2' },
    { label: 'Đồng ý ≥ 1 dịch vụ', val: chotHo.size, color: '#178a5c' }
  ];
  const funnelHtml = funnel.map(x => {
    const p = totalHo ? x.val / totalHo * 100 : 0;
    return `<div class="funnel-row">
      <div class="funnel-lbl">${esc(x.label)}</div>
      <div class="funnel-track"><div class="funnel-fill" style="width:${Math.max(2, p)}%;background:${x.color}">${pctNum(x.val, totalHo)}%</div></div>
      <div class="funnel-val">${nf(x.val)}</div>
    </div>`;
  }).join('');

  /* 2. theo đơn vị */
  const byUnit = {};
  ts.forEach(t => {
    if (!byUnit[t.donvi]) byUnit[t.donvi] = { luot: 0, ho: new Set(), chot: new Set(), svc: { tt: new Set(), mytv: new Set(), mesh: new Set(), cam: new Set() } };
    const u = byUnit[t.donvi];
    u.luot++; u.ho.add(t.ma);
    if ((t.dong_y || []).length) u.chot.add(t.ma);
    (t.dong_y || []).forEach(k => u.svc[k] && u.svc[k].add(t.ma));
  });
  const unitRows = ['TTVT', 'VNPT'].filter(u => byUnit[u]).map(u => {
    const x = byUnit[u];
    return [unitLabel(u), x.luot, x.ho.size, x.chot.size, x.svc.tt.size, x.svc.mytv.size, x.svc.mesh.size, x.svc.cam.size, pct(x.chot.size, x.ho.size)];
  });
  const unitTotal = ['TTVT', 'VNPT'].reduce((a, u) => {
    const x = byUnit[u];
    if (!x) return a;
    a[1] += x.luot; a[2] += x.ho.size;
    return a;
  }, [0, 0, 0]);
  const unitCard = statTableCard('Hiệu suất theo đơn vị', `Trong kỳ đã chọn · ${nf(ts.length)} lượt chạm`,
    ['Đơn vị', 'Lượt chạm', 'Hộ đã chạm', 'Hộ đồng ý', 'Chốt trả trước', 'Chốt MyTV', 'Chốt Mesh', 'Chốt Cam', 'Tỉ lệ chốt/hộ'],
    unitRows, 'byUnit');

  /* 3. theo ngày */
  const days = {};
  ts.forEach(t => {
    if (!days[t.ngay]) days[t.ngay] = { luot: 0, ho: new Set(), TTVT: 0, VNPT: 0, chot: new Set() };
    const d = days[t.ngay];
    d.luot++; d.ho.add(t.ma); d[t.donvi] = (d[t.donvi] || 0) + 1;
    if ((t.dong_y || []).length) d.chot.add(t.ma);
  });
  const dayKeys = Object.keys(days).sort();
  const chartDays = dayKeys.slice(-21);
  const maxDay = Math.max(1, ...chartDays.map(k => days[k].luot));
  const chartHtml = chartDays.length ? `<div class="day-chart">${chartDays.map(k => {
    const d = days[k];
    const h = Math.round(d.luot / maxDay * 88);
    const hv = Math.round(d.VNPT / maxDay * 88);
    return `<div class="day-col" title="${fmtDate(k)}: ${d.luot} lượt">
      <div class="bar-n">${d.luot}</div>
      <div class="day-bars">
        <div class="bar-v ttvt" style="height:${Math.max(2, h - hv)}px"></div>
        <div class="bar-v vnpt" style="height:${Math.max(0, hv)}px"></div>
      </div>
      <div class="day-lbl">${k.slice(8, 10)}/${k.slice(5, 7)}</div>
    </div>`;
  }).join('')}</div>
  <div class="mini-note"><span style="display:inline-block;width:9px;height:9px;background:#d98200;border-radius:2px"></span> TTVT &nbsp; <span style="display:inline-block;width:9px;height:9px;background:#0b5cab;border-radius:2px"></span> VNPT KV · 21 ngày gần nhất có phát sinh</div>` : '<div class="empty" style="padding:20px">Chưa có dữ liệu</div>';

  const dayRows = dayKeys.slice().reverse().map(k => {
    const d = days[k];
    return [fmtDate(k), d.luot, d.ho.size, d.TTVT, d.VNPT, d.chot.size];
  });
  const dayCard = statTableCard('Lượt chạm theo ngày', 'Mỗi dòng là một ngày có phát sinh chạm',
    ['Ngày chạm', 'Lượt chạm', 'Hộ', 'TTVT', 'VNPT KV', 'Hộ đồng ý'], dayRows, 'byDay');

  /* 4. theo nhân viên */
  const empMap = {};
  ts.forEach(t => {
    const k = t.nguoi || '—';
    if (!empMap[k]) empMap[k] = { luot: 0, ho: new Set(), chot: new Set(), donvi: t.donvi, svc: { tt: new Set(), mytv: new Set(), mesh: new Set(), cam: new Set() } };
    const e = empMap[k];
    e.luot++; e.ho.add(t.ma); e.donvi = t.donvi;
    if ((t.dong_y || []).length) e.chot.add(t.ma);
    (t.dong_y || []).forEach(x => e.svc[x] && e.svc[x].add(t.ma));
  });
  const empList = Object.keys(empMap).map(k => ({ name: k, ...empMap[k] })).sort((a, b) => b.ho.size - a.ho.size || b.luot - a.luot);
  const empRows = empList.map(e => [e.name, e.donvi, e.luot, e.ho.size, e.chot.size, e.svc.tt.size, e.svc.mytv.size, e.svc.mesh.size, e.svc.cam.size, pct(e.chot.size, e.ho.size)]);
  const empCard = statTableCard('Hiệu suất theo nhân viên nhập', `Xếp theo số hộ đã chạm · ${nf(empList.length)} nhân viên`,
    ['Nhân viên', 'Đơn vị', 'Lượt', 'Hộ', 'Hộ đồng ý', 'TT', 'MyTV', 'Mesh', 'Cam', 'Tỉ lệ chốt'],
    empRows, 'byEmp');

  const leader = empList.slice(0, 8).map((e, i) => `<div class="rank-item">
      <div class="rank-no">${i + 1}</div>
      <div class="rank-name">${esc(e.name)} <span class="rank-sub">${esc(unitLabel(e.donvi))}</span></div>
      <div class="rank-val">${nf(e.ho.size)} <span class="rank-sub">hộ</span></div>
    </div>`).join('') || '<div class="muted" style="font-size:12px">Chưa có dữ liệu</div>';

  /* 5. theo tổ (lũy kế) */
  const teamMap = {};
  ALL.forEach(r => {
    if (!teamMap[r.to]) teamMap[r.to] = { total: 0, touched: 0, chot: new Set(), miss: 0 };
    const t = teamMap[r.to];
    t.total++;
    if (LATEST.has(r.ma)) t.touched++;
    if (!r.daThay) t.miss++;
    if (AGREED.has(r.ma)) t.chot.add(r.ma);
  });
  const teamRows = Object.keys(teamMap).sort(byName).map(k => {
    const t = teamMap[k];
    return [shortTeam(k), t.total, t.touched, t.total - t.touched, t.chot.size, pct(t.touched, t.total)];
  });
  const teamCard = statTableCard('Độ phủ theo tổ kỹ thuật', 'Lũy kế đến hiện tại (không phụ thuộc bộ lọc ngày)',
    ['Tổ', 'Tổng hộ', 'Đã chạm', 'Chưa chạm', 'Hộ đồng ý', 'Tỉ lệ chạm'], teamRows, 'byTeam');

  /* 6. theo xã */
  const xaMap = {};
  ALL.forEach(r => {
    if (!xaMap[r.xa]) xaMap[r.xa] = { total: 0, touched: 0, chot: new Set(), miss: 0 };
    const t = xaMap[r.xa];
    t.total++;
    if (LATEST.has(r.ma)) t.touched++;
    if (!r.daThay) t.miss++;
    if (AGREED.has(r.ma)) t.chot.add(r.ma);
  });
  const xaRows = Object.keys(xaMap).sort(byName).map(k => {
    const t = xaMap[k];
    return [k, t.total, t.touched, t.total - t.touched, t.chot.size, pct(t.touched, t.total)];
  });
  const xaCard = statTableCard('Độ phủ theo xã', 'Lũy kế đến hiện tại',
    ['Xã', 'Tổng hộ', 'Đã chạm', 'Chưa chạm', 'Hộ đồng ý', 'Tỉ lệ chạm'], xaRows, 'byXa');

  /* 7. theo dịch vụ */
  const svcRows = SVC_KEYS.map(k => [
    SVC[k].label + (k === 'tt' ? ' (đang trả sau)' : ''),
    need[k], svcTuVan[k].size, svcAgree[k].size,
    pct(svcAgree[k].size, need[k])
  ]);
  const svcCard = statTableCard('Hiệu quả tiếp thị theo dịch vụ', 'Nhu cầu = số hộ còn thiếu dịch vụ trong toàn danh sách',
    ['Dịch vụ', 'Nhu cầu (hộ)', 'Đã tư vấn (hộ)', 'Đồng ý (hộ)', 'Tỉ lệ chốt/nhu cầu'], svcRows, 'byService');

  /* 8. kết quả chạm */
  const resMap = {};
  ts.forEach(t => { resMap[t.ket_qua] = (resMap[t.ket_qua] || 0) + 1; });
  const resRows = Object.keys(RESULT).map(k => [RESULT[k].label, resMap[k] || 0, pct(resMap[k] || 0, ts.length)]);
  const resCard = statTableCard('Phân bố kết quả chạm', `Tổng ${nf(ts.length)} lượt trong kỳ`,
    ['Kết quả', 'Số lượt', 'Tỉ lệ'], resRows, 'byResult');

  /* 9. địa bàn còn tồn */
  const areaMap = {};
  ALL.forEach(r => {
    const k = r.xa + ' › ' + (r.ap || '(không rõ ấp)');
    if (!areaMap[k]) areaMap[k] = { total: 0, touched: 0 };
    areaMap[k].total++;
    if (LATEST.has(r.ma)) areaMap[k].touched++;
  });
  const areaRows = Object.keys(areaMap).map(k => {
    const a = areaMap[k];
    return [k, a.total, a.touched, a.total - a.touched, pct(a.touched, a.total)];
  }).sort((a, b) => b[3] - a[3]).slice(0, 20);
  const areaCard = statTableCard('Địa bàn còn nhiều hộ chưa chạm', 'Top 20 xã/ấp theo số hộ chưa chạm (lũy kế)',
    ['Xã / Ấp', 'Tổng hộ', 'Đã chạm', 'Chưa chạm', 'Tỉ lệ chạm'], areaRows, 'byArea');

  body.innerHTML = `
  <div class="stats-grid">
    <div class="stat-card">
      <div class="stat-h"><h3>Độ phủ tiếp thị</h3><span class="sub">toàn danh sách</span></div>
      <div class="donut-wrap">
        <div class="donut" style="--p:${covRate.toFixed(1)}"><div><b>${pct(coveredAll.size, totalHo)}</b><span>đã chạm</span></div></div>
        <div class="legend">
          <div><i class="dot" style="background:#0b5cab"></i>Đã chạm <b>${nf(coveredAll.size)}</b></div>
          <div><i class="dot" style="background:#e8eef6"></i>Chưa chạm <b>${nf(totalHo - coveredAll.size)}</b></div>
          <div><i class="dot" style="background:#178a5c"></i>Hộ đồng ý <b>${nf(new Set(TOUCHES.filter(t => (t.dong_y || []).length).map(t => t.ma)).size)}</b></div>
          <div><i class="dot" style="background:#e07b00"></i>Chạm hôm nay <b>${nf(homNay)}</b></div>
        </div>
      </div>
    </div>
    <div class="stat-card">
      <div class="stat-h"><h3>Phễu tiếp thị</h3><span class="sub">lũy kế</span></div>
      <div class="funnel">${funnelHtml}</div>
    </div>
    <div class="stat-card">
      <div class="stat-h"><h3>Bảng xếp hạng nhân viên</h3><span class="sub">theo hộ đã chạm</span></div>
      <div class="rank">${leader}</div>
      <div class="mini-note">Tính trong kỳ đã chọn.</div>
    </div>
  </div>

  <div class="card">
    <div class="card-h"><div><div class="card-title">Nhịp độ chạm theo ngày</div><div class="sub" style="font-size:11.5px;color:var(--gray)">Cột xếp chồng TTVT + VNPT KV</div></div></div>
    <div class="card-b day-chart-wrap">${chartHtml}</div>
  </div>

  ${unitCard}
  ${empCard}
  ${svcCard}
  ${dayCard}
  ${resCard}
  ${teamCard}
  ${xaCard}
  ${areaCard}
  <div class="card"><div class="card-b" style="font-size:12px;color:var(--gray)">
    💡 Bảng “Độ phủ”, “Theo tổ”, “Theo xã” và “Địa bàn còn tồn” luôn tính lũy kế trên toàn bộ dữ liệu để phản ánh tiến độ thực tế; các bảng còn lại theo bộ lọc ngày/đơn vị ở trên.
  </div></div>`;
}

/* ---------- HƯỚNG DẪN / CẤU TRÚC SHEET ---------- */
const COLUMNS = [
  ['MA_TB', 'Mã thiết bị / mã thuê bao', 'Duy nhất mỗi hộ — dùng để đối chiếu khi ghi ngược vào sheet'],
  ['TEN_KH', 'Tên khách hàng', 'Hiển thị trên danh sách và phiếu chạm'],
  ['SO_DT', 'Số điện thoại', 'Bấm để gọi ngay trên điện thoại'],
  ['TO_VT', 'Tổ kỹ thuật địa bàn', '5 tổ: An Nhơn Tây, Nhuận Đức, Thái Mỹ, Tân An Hội 1, Tân An Hội 2'],
  ['Nhân viên kỹ thuật', 'Nhân viên / CTV phụ trách', 'Dạng “MÃ - Tên”, dùng để lọc theo nhân viên'],
  ['MA_KV', 'Mã khu vực', '32 mã TAH001…TAH032'],
  ['DOKIEM_OBSS_TT_TRATRUOC', 'Loại thanh toán hiện tại', 'TRẢ TRƯỚC / TRẢ SAU — nguồn xác định nhu cầu chuyển trả trước'],
  ['DOKIEM_DIACHI_LD', 'Địa chỉ đầy đủ', 'Đã tách sẵn từ địa chỉ gốc'],
  ['Số nhà', 'Số nhà / số hẻm', 'Có thể trống'],
  ['Đường', 'Tên đường', 'Dùng cho bộ lọc theo đường'],
  ['Ấp', 'Ấp / khu phố', 'Có thể trống'],
  ['Xã', 'Xã / thị trấn', '13 xã, thị trấn'],
  ['TYPE_HIENTAI', 'Model modem hiện tại', 'GW040, GW020, 045A, GX40E5…'],
  ['TRANGTHAI_THAYTHE', 'Trạng thái thay modem', 'CHUA_THAY_THE / DA_THAY_THE'],
  ['TB_MESH', 'Mã thiết bị Mesh', 'Trống = chưa có Mesh (nhiều mã cách nhau dấu “;”)'],
  ['TB_CAM', 'Mã camera', 'Trống = chưa có camera'],
  ['TB_MYTV', 'Mã MyTV', 'Trống = chưa có MyTV']
];

const STATUS_COLS = [
  ['TRANGTHAI_TIEP_THI', 'CHUA_CHAM / DA_CHAM / DA_CHOT'],
  ['NGAY_CHAM_CUOI', 'Ngày của lượt chạm gần nhất'],
  ['GIO_CHAM_CUOI', 'Giờ của lượt chạm gần nhất'],
  ['NGUOI_CHAM_CUOI', 'Tên người nhập lượt chạm gần nhất'],
  ['DON_VI_CHAM', 'TTVT hoặc VNPT (giao diện hiển thị VNPT KV)'],
  ['KET_QUA_CHAM', 'Quan tâm / Đã chốt / Hẹn lại / Từ chối / Không liên hệ được'],
  ['DICH_VU_TU_VAN', 'Các dịch vụ đã tư vấn (lũy kế)'],
  ['DICH_VU_DONG_Y', 'Các dịch vụ khách đã đồng ý (lũy kế)'],
  ['SO_LAN_CHAM', 'Tổng số lần đã chạm hộ này'],
  ['GHI_CHU_CHAM', 'Ghi chú của lượt chạm gần nhất'],
  ['NGAY_HEN_THAY', 'Ngày khách hẹn thay modem'],
  ['GIO_HEN_THAY', 'Giờ hẹn thay'],
  ['TRANGTHAI_HEN', 'CHO_THAY / DA_THAY / HUY'],
  ['DICH_VU_HEN', 'Dịch vụ tiếp thị kèm theo lúc hẹn'],
  ['GHI_CHU_HEN', 'Ghi chú hẹn thay'],
  ['NGUOI_HEN', 'Nhân viên tạo hẹn']
];

const APPT_COLS = [
  'ID', 'THOI_GIAN_TAO', 'NGAY_HEN', 'GIO_HEN', 'MA_TB', 'TEN_KH', 'SO_DT',
  'TO_VT', 'NHAN_VIEN', 'MA_KV', 'XA', 'AP', 'DUONG', 'SO_NHA', 'DIA_CHI',
  'DICH_VU', 'GHI_CHU', 'NGUOI_HEN', 'DON_VI', 'TRANG_THAI', 'CAP_NHAT'
];

const HISTORY_COLS = [
  'ID', 'THOI_GIAN', 'NGAY_CHAM', 'GIO_CHAM', 'MA_TB', 'TEN_KH', 'SO_DT',
  'TO_VT', 'NHAN_VIEN', 'MA_KV', 'XA', 'AP', 'DUONG', 'SO_NHA', 'DIA_CHI',
  'NGUOI_NHAP', 'DON_VI', 'KET_QUA', 'DICH_VU_TU_VAN', 'DICH_VU_DONG_Y',
  'GHI_CHU', 'TRANGTHAI_THAY_MODEM'
];

function renderGuide() {
  const root = $('#guide-root');
  if (!root || !LOADED) return;
  const spread = META.spreadsheetUrl || '';
  const hist = META.historySheet || 'CHAM_TIEP_THI';
  const isHttp = API.mode === 'http';
  const apiBanner = isHttp ? `<div class="notice" style="align-items:center;flex-wrap:wrap">
      <span>🌐 Trang tĩnh trên <b>GitHub Pages</b> · API cố định: <b>${esc(API.host() || 'chưa cấu hình')}</b></span>
      <span style="margin-left:auto;color:var(--gray);font-size:12px">Quản trị sửa 1 lần trong <code>config.js</code>; nhân viên không phải nhập.</span>
    </div>` : '';
  const steps = isHttp ? `
      <ol>
        <li>Đưa thư mục web lên GitHub → <b>Settings → Pages</b> → chọn nhánh có <code>index.html</code>.</li>
        <li>Triển khai <code>Code.gs</code> thành <b>Ứng dụng web</b> (Thực thi: Tôi, Truy cập: Bất kỳ ai) → lấy URL <code>/exec</code>.</li>
        <li>Quản trị dán URL <code>/exec</code> vào <code>config.js</code> (một lần). Tất cả nhân viên dùng chung, không phải nhập API.</li>
        <li>Đăng nhập bằng tài khoản được phân (mật khẩu mặc định <b>123456</b>) và đổi mật khẩu ngay.</li>
        <li>Chạm khách → chọn dịch vụ đã tư vấn / khách đồng ý → <b>Lưu</b>. Hoặc chọn <b>Gọi hẹn thay</b> để đưa khách sang tab KH hẹn thay.</li>
      </ol>` : `
      <ol>
        <li>Mở Google Sheet → <b>Tiện ích mở rộng → Apps Script</b>, dán mã đã cung cấp.</li>
        <li><b>Triển khai → Ứng dụng web</b> → quyền <i>Chạy với tư cách: Tôi</i>, truy cập <i>Bất kỳ ai</i>.</li>
        <li>Gửi link và tài khoản cho nhân viên TTVT/VNPT KV; mỗi người đăng nhập và chỉ thấy khu vực được phân.</li>
        <li>Chạm khách → chọn dịch vụ / khách đồng ý → <b>Lưu</b>; hoặc <b>Gọi hẹn thay</b> để chuyển sang tab KH hẹn thay.</li>
        <li>Xem lại tại tab <b>TTVT đã chạm</b>, <b>VNPT KV đã chạm</b>, <b>Báo cáo chi tiết</b>; tải Excel ở từng bảng thống kê.</li>
      </ol>`;
  root.innerHTML = `
  <div class="file-info">
    <div class="file-chip"><b>File nguồn</b><span>${esc(META.fileName || 'cham thay modem - TACH DIA CHI.xlsx')}</span></div>
    <div class="file-chip"><b>Sheet dữ liệu</b><span>${esc(META.sheetName || 'Sheet1')}</span></div>
    <div class="file-chip"><b>Số hộ</b><span>${nf(ALL.length)}</span></div>
    <div class="file-chip"><b>Sheet lịch sử</b><span>${esc(hist)}</span></div>
    <div class="file-chip"><b>Lượt chạm đã ghi</b><span>${nf(TOUCHES.length)}</span></div>
    <div class="file-chip"><b>Tài khoản</b><span>${USER ? esc(USER.name) + ' · ' + esc(USER.role) : '—'}</span></div>
    <div class="file-chip"><b>Khu vực</b><span>${USER ? (USER.full ? 'Toàn quyền' : esc(USER.areas.join(', '))) : '—'}</span></div>
    ${spread ? `<a class="btn" href="${esc(spread)}" target="_blank" rel="noopener" style="align-self:center">↗ Mở Google Sheet</a>` : ''}
  </div>
  ${apiBanner}
  ${isHttp ? `<div class="callout warn" style="margin-bottom:16px">🔒 API đã cố định trong <code>config.js</code>. Nếu website công khai, nên đặt <code>ACCESS_CODE</code> trong <code>Code.gs</code> và <code>config.js</code>; nhân viên vẫn chỉ cần đăng nhập.</div>` : ''}

  <div class="two-col">
    <div class="card"><div class="card-h"><div class="card-title">📌 Cách hoạt động</div></div><div class="card-b guide">
      <div class="callout ok"><b>Đăng nhập theo tài khoản được phân ô.</b> Mỗi người chỉ thấy hộ thuộc MA_KV của mình; Giám đốc, Phó Giám đốc và HTKT thấy toàn bộ. Khi lưu chạm hoặc hẹn thay, <b>tên người nhập và đơn vị được tự động ghi</b> từ tài khoản — không phải nhập tay.</div>
      ${steps}
      <div class="callout warn">Không mở khóa cột tiêu đề của Sheet1 và không đổi tên 17 cột gốc — ứng dụng dựa vào đúng thứ tự cột để đọc dữ liệu. Các cột trạng thái mới (R→AG) do ứng dụng tự thêm, có thể ẩn đi nếu không dùng.</div>
      <p><b>Ghi chú vận hành:</b> nhiều người nhập cùng lúc vẫn an toàn nhờ khóa ghi (LockService). Nếu hai người cùng lưu một hộ, cả hai lượt đều được ghi vào lịch sử; cột trạng thái lấy lượt mới nhất.</p>
    </div></div>

    <div class="card"><div class="card-h"><div class="card-title">🧭 Quy tắc xác định “còn thiếu dịch vụ”</div></div><div class="card-b guide">
      <table class="struct-table">
        <thead><tr><th style="width:32%">Nhu cầu</th><th>Điều kiện lấy từ sheet</th></tr></thead>
        <tbody>
          <tr><td style="text-align:left"><b>Chuyển trả trước</b></td><td><code>DOKIEM_OBSS_TT_TRATRUOC = "TRẢ SAU"</code></td></tr>
          <tr><td style="text-align:left"><b>Chưa có MyTV</b></td><td><code>TB_MYTV</code> trống</td></tr>
          <tr><td style="text-align:left"><b>Chưa có Mesh</b></td><td><code>TB_MESH</code> trống</td></tr>
          <tr><td style="text-align:left"><b>Chưa có Camera</b></td><td><code>TB_CAM</code> trống</td></tr>
          <tr><td style="text-align:left"><b>Chưa thay modem</b></td><td><code>TRANGTHAI_THAYTHE = "CHUA_THAY_THE"</code></td></tr>
        </tbody>
      </table>
      <p style="margin-top:10px">Mỗi hộ được tính <b>điểm ưu tiên</b> = số dịch vụ còn thiếu × 10 + 12 nếu chưa thay modem + 3 nếu đang trả sau. Bấm <b>⚡ Gợi ý 15 hộ ưu tiên</b> để lấy danh sách chạm gọn trong ngày.</p>
      <div class="callout">Ô <code>TB_MESH</code> có thể chứa nhiều mã cách nhau dấu <code>;</code> (ví dụ <code>mesh0048130;mesh0048131</code>) — ứng dụng chỉ cần biết ô có dữ liệu hay không.</div>
    </div></div>
  </div>

  <div class="two-col">
    <div class="card"><div class="card-h"><div class="card-title">👤 Tài khoản &amp; phân quyền</div></div><div class="card-b guide">
      <p><b>${USER ? esc(USER.name) : '—'}</b> · ${USER ? esc(USER.role) + ' · ' + esc(unitLabel(USER.unit)) : ''} · <code>${USER ? esc(USER.username) : ''}</code></p>
      <p>Khu vực được phân: <b>${USER ? (USER.full ? 'Tất cả (toàn quyền)' : esc(USER.areas.join(', '))) : '—'}</b></p>
      <ul>
        <li><b>GIÁM ĐỐC, PHÓ GIÁM ĐỐC, HTKT:</b> xem và nhập cho tất cả khu vực.</li>
        <li><b>TỔ TRƯỞNG, NHÂN VIÊN:</b> chỉ thấy hộ có MA_KV nằm trong phân ô của tài khoản.</li>
        <li>Tài khoản quản lý tại sheet <code>TAI_KHOAN</code>; mật khẩu mặc định <b>123456</b>.</li>
        <li>Đổi mật khẩu bất cứ lúc nào ở menu tên tài khoản (góc trên bên phải).</li>
      </ul>
      <button type="button" class="btn sm" data-act="change-password">🔑 Đổi mật khẩu</button>
    </div></div>

    <div class="card"><div class="card-h"><div class="card-title">📅 KH hẹn thay modem</div></div><div class="card-b guide">
      <p>Trong phiếu chạm, chọn tab <b>Gọi hẹn thay</b> khi khách hẹn ngày thay modem. Phiếu hẹn <b>không tạo phiếu chạm</b>.</p>
      <ul>
        <li>Nhập <b>ngày hẹn</b>, giờ hẹn (không bắt buộc), dịch vụ tiếp thị kèm theo (nếu có) và <b>ghi chú bắt buộc</b>.</li>
        <li>Khách chuyển sang tab <b>KH hẹn thay</b>, sắp xếp theo ngày hẹn gần nhất.</li>
        <li>Có bộ lọc ngày hẹn, trạng thái, người hẹn, khu vực; nút <b>Đã thay</b>, <b>Dời</b>, <b>Hủy</b>.</li>
        <li>Đánh dấu <b>Đã thay</b> sẽ cập nhật trạng thái hộ sang đã thay modem.</li>
      </ul>
      <p>Sheet lịch sử hẹn: <code>HEN_THAY_MODEM</code> (${APPT_COLS.length} cột).</p>
      <div class="chips" style="max-width:none">${APPT_COLS.slice(0, 12).map(c => `<span class="chip ok" style="font-family:ui-monospace,Consolas,monospace;font-size:10px">${esc(c)}</span>`).join('')}</div>
    </div></div>
  </div>

  <div class="card">
    <div class="card-h"><div><div class="card-title">🗂 Cấu trúc sheet dữ liệu gốc</div><div class="sub" style="font-size:11.5px;color:var(--gray)">${esc(META.sheetName || 'Sheet1')} · ${nf(ALL.length)} dòng dữ liệu · 17 cột</div></div>
      <button type="button" class="btn sm" data-act="export-structure">⬇ Tải file mẫu cấu trúc</button></div>
    <div class="table-wrap" style="max-height:none">
      <table class="struct-table">
        <thead><tr><th>#</th><th>Tên cột</th><th>Ý nghĩa</th><th>Ghi chú sử dụng trong app</th></tr></thead>
        <tbody>${COLUMNS.map((c, i) => `<tr><td>${i + 1}</td><td class="col-name">${esc(c[0])}</td><td>${esc(c[1])}</td><td>${esc(c[2])}</td></tr>`).join('')}</tbody>
      </table>
    </div>
  </div>

  <div class="two-col">
    <div class="card"><div class="card-h"><div class="card-title">➕ Cột trạng thái ứng dụng tự thêm vào Sheet1</div></div><div class="card-b guide">
      <p>Bắt đầu từ cột <b>R</b>, ngay sau 17 cột gốc. Đây là nơi bạn có thể lọc/báo cáo trực tiếp trên Sheet mà không cần mở web.</p>
      <table class="struct-table"><thead><tr><th>Tên cột</th><th>Giá trị</th></tr></thead>
        <tbody>${STATUS_COLS.map(c => `<tr><td class="col-name">${esc(c[0])}</td><td>${esc(c[1])}</td></tr>`).join('')}</tbody></table>
    </div></div>

    <div class="card"><div class="card-h"><div class="card-title">🧾 Sheet lịch sử <code>${esc(hist)}</code></div></div><div class="card-b guide">
      <p>Mỗi lần lưu là <b>một dòng mới</b>, không bao giờ ghi đè — giữ đầy đủ lịch sử nhiều lần chạm, nhiều nhân viên, nhiều đơn vị.</p>
      <div class="chips" style="max-width:none">${HISTORY_COLS.map(c => `<span class="chip ok" style="font-family:ui-monospace,Consolas,monospace;font-size:10px">${esc(c)}</span>`).join('')}</div>
      <p style="margin-top:12px"><b>Xóa lượt chạm:</b> dùng nút 🗑 ở tab TTVT/VNPT KV hoặc Báo cáo chi tiết — ứng dụng xóa đúng dòng trong sheet lịch sử và tính lại trạng thái của hộ đó. Không nên xóa trực tiếp trên sheet vì cột trạng thái sẽ không tự cập nhật.</p>
      <div class="callout ok">Dữ liệu nằm hoàn toàn trên Google Sheet. Sao lưu bằng <b>File → Tạo bản sao</b> trong Google Sheet.</div>
    </div></div>
  </div>

  <div class="two-col">
    <div class="card"><div class="card-h"><div class="card-title">⚙️ Cấu hình</div></div><div class="card-b guide">
      <div class="form-field" style="max-width:280px">
        <label>Mục tiêu lượt chạm / người / ngày</label>
        <input type="number" id="cfg-target" min="1" max="200" value="${CONFIG.target || 15}">
      </div>
      <p style="margin-top:8px">Dùng để tính thanh tiến độ “Lượt chạm hôm nay”. Lưu trên trình duyệt của từng người.</p>
      <button type="button" class="btn sm" data-act="save-target">Lưu mục tiêu</button>
      <p style="margin-top:14px" class="muted">Phiên bản 1.0 · Dữ liệu cập nhật lúc ${esc(META.updatedAt || '—')}</p>
    </div></div>

    <div class="card"><div class="card-h"><div class="card-title">❓ Xử lý sự cố</div></div><div class="card-b guide">
      <ul>
        <li><b>Không tải được dữ liệu:</b> mở lại link web app, bấm “Cho phép” khi Google hỏi quyền; kiểm tra tài khoản triển khai còn quyền sửa Sheet.</li>
        <li><b>Lưu báo lỗi:</b> thường do mất mạng hoặc Sheet bị khóa — thử lại sau vài giây, dữ liệu biểu mẫu vẫn còn trên form.</li>
        <li><b>Sai tên cột:</b> khôi phục đúng 17 cột gốc, không chèn/xóa cột ở giữa.</li>
        <li><b>Tải Excel bị chặn:</b> dùng Chrome/Edge, mở web app ở tab mới (không nhúng trong iframe của trang khác).</li>
        <li><b>Dữ liệu quá lớn:</b> nếu lịch sử vượt ~20.000 dòng, nên lưu trữ bớt các đợt cũ sang file khác để trang tải nhanh.</li>
      </ul>
    </div></div>
  </div>`;

  $('#cfg-target').addEventListener('change', e => {
    const v = Math.max(1, Math.min(200, parseInt(e.target.value, 10) || 15));
    CONFIG.target = v; e.target.value = v; saveLocal(); renderKPIs();
  });
}

/* ---------- XUẤT EXCEL ---------- */
const EXPORT_HEADERS = ['STT', 'Ngày chạm', 'Giờ', 'Mã TB', 'Tên khách hàng', 'Số điện thoại', 'Tổ', 'Nhân viên quản lý', 'Xã', 'Ấp', 'Đường', 'Số nhà', 'Địa chỉ', 'Đơn vị chạm', 'Người nhập', 'Kết quả', 'Dịch vụ đã tư vấn', 'Dịch vụ khách đồng ý', 'Ghi chú', 'Trạng thái thay modem'];

function reportRows(list) {
  return [EXPORT_HEADERS].concat(list.map((t, i) => {
    const c = touchCustomer(t);
    return [i + 1, fmtDate(t.ngay), t.gio || '', t.ma, c.ten, c.sdt, shortTeam(c.to), c.nv, c.xa, c.ap, c.duong, c.sn, c.dc,
      unitLabel(t.donvi), t.nguoi, (RESULT[t.ket_qua] || {}).label || t.ket_qua,
      (t.tu_van || []).map(k => SVC[k] ? SVC[k].label : k).join(', '),
      (t.dong_y || []).map(k => SVC[k] ? SVC[k].label : k).join(', '),
      t.ghi_chu || '', c.daThay === false ? 'CHUA_THAY_THE' : 'DA_THAY_THE'];
  }));
}

function exportReport() {
  const list = reportList();
  if (!list.length) return toast('Không có dữ liệu để xuất', 'warn');
  downloadXlsx(`Bao_cao_cham_tiep_thi_${stamp()}.xlsx`, [{ name: 'Bao cao chi tiet', rows: reportRows(list) }]);
  toast(`Đã xuất ${nf(list.length)} lượt chạm ra Excel`, 'ok');
}

function exportUnit(unit) {
  const list = unitTouches(unit, UI.unit[unit]);
  if (!list.length) return toast('Không có dữ liệu để xuất', 'warn');
  downloadXlsx(`${unit}_da_cham_${stamp()}.xlsx`, [{ name: `${unit} da cham`, rows: reportRows(list) }]);
  toast(`Đã xuất ${nf(list.length)} lượt chạm ${unitLabel(unit)}`, 'ok');
}

function exportList() {
  const list = currentList();
  if (!list.length) return toast('Không có hộ nào để xuất', 'warn');
  const headers = ['STT', 'Mã TB', 'Tên khách hàng', 'Số điện thoại', 'Tổ', 'Nhân viên quản lý', 'Mã khu vực', 'Xã', 'Ấp', 'Đường', 'Số nhà', 'Địa chỉ', 'Modem hiện tại', 'Trạng thái thay modem', 'Thanh toán', 'MyTV', 'Mesh', 'Camera', 'Dịch vụ còn thiếu', 'Điểm ưu tiên', 'Trạng thái chạm', 'Ngày chạm cuối', 'Người chạm cuối', 'Đơn vị', 'Kết quả', 'Tư vấn', 'Đồng ý', 'Ghi chú', 'Số lần chạm'];
  const rows = list.map((r, i) => {
    const t = LATEST.get(r.ma);
    return [i + 1, r.ma, r.ten, r.sdt, r.to, r.nv, r.kv, r.xa, r.ap, r.duong, r.sn, r.dc, r.modem,
      r.daThay ? 'DA_THAY_THE' : 'CHUA_THAY_THE', r.traSau ? 'TRẢ SAU' : 'TRẢ TRƯỚC',
      r.mytv ? 'CÓ' : 'CHƯA CÓ', r.mesh ? 'CÓ' : 'CHƯA CÓ', r.cam ? 'CÓ' : 'CHƯA CÓ',
      gapKeys(r).map(k => SVC[k].label).join(', '), score(r),
      t ? 'DA_CHAM' : 'CHUA_CHAM', t ? fmtDate(t.ngay) : '', t ? t.nguoi : '', t ? unitLabel(t.donvi) : '',
      t ? ((RESULT[t.ket_qua] || {}).label || t.ket_qua) : '',
      t ? (t.tu_van || []).map(k => SVC[k] ? SVC[k].label : k).join(', ') : '',
      t ? (t.dong_y || []).map(k => SVC[k] ? SVC[k].label : k).join(', ') : '',
      t ? (t.ghi_chu || '') : '', t ? 1 : 0];
  });
  downloadXlsx(`Danh_sach_cham_${stamp()}.xlsx`, [{ name: 'Danh sach cham', rows: [headers].concat(rows) }]);
  toast(`Đã xuất ${nf(list.length)} hộ ra Excel`, 'ok');
}

function exportStats(key) {
  const c = STATS_CACHE[key];
  if (!c) return toast('Chưa có dữ liệu', 'warn');
  downloadXlsx(`Thong_ke_${key}_${stamp()}.xlsx`, [{ name: c.name, rows: c.rows }]);
  toast('Đã tải báo cáo Excel', 'ok');
}

function exportAllStats() {
  const sheets = Object.keys(STATS_CACHE).map(k => ({ name: STATS_CACHE[k].name, rows: STATS_CACHE[k].rows }));
  if (!sheets.length) return toast('Chưa có dữ liệu', 'warn');
  // sheet tổng quan
  const sum = [['CHỈ SỐ', 'GIÁ TRỊ'],
    ['Tổng hộ trong danh sách', ALL.length],
    ['Hộ đã chạm', new Set(TOUCHES.map(t => t.ma)).size],
    ['Tổng lượt chạm', TOUCHES.length],
    ['Hộ khách đồng ý', new Set(TOUCHES.filter(t => (t.dong_y || []).length).map(t => t.ma)).size],
    ['Ngày xuất', new Date().toLocaleString('vi-VN')]];
  sheets.unshift({ name: 'Tong quan', rows: sum });
  downloadXlsx(`Thong_ke_cham_tiep_thi_${stamp()}.xlsx`, sheets);
  toast(`Đã xuất ${sheets.length} bảng thống kê ra Excel`, 'ok');
}

function exportStructure() {
  const rows = [['STT', 'TÊN CỘT', 'Ý NGHĨA', 'GHI CHÚ']].concat(COLUMNS.map((c, i) => [i + 1, c[0], c[1], c[2]]));
  const rows2 = [['TÊN CỘT', 'GIÁ TRỊ']].concat(STATUS_COLS);
  const rows3 = [['TÊN CỘT LỊCH SỬ']].concat(HISTORY_COLS.map(c => [c]));
  downloadXlsx(`Cau_truc_sheet_${stamp()}.xlsx`, [
    { name: 'Cau truc 17 cot', rows },
    { name: 'Cot trang thai them', rows: rows2 },
    { name: 'Cot lich su', rows: rows3 }
  ]);
  toast('Đã tải file cấu trúc sheet', 'ok');
}

/* ---------- copy script ---------- */
function copyText(txt) {
  const done = () => toast('Đã sao chép kịch bản', 'ok');
  if (navigator.clipboard && navigator.clipboard.writeText) {
    navigator.clipboard.writeText(txt).then(done).catch(() => fallbackCopy(txt, done));
  } else fallbackCopy(txt, done);
}
function fallbackCopy(txt, done) {
  const ta = document.createElement('textarea');
  ta.value = txt; ta.style.position = 'fixed'; ta.style.opacity = '0';
  document.body.appendChild(ta); ta.select();
  try { document.execCommand('copy'); done(); } catch (e) { toast('Không sao chép được', 'err'); }
  ta.remove();
}

/* ============================================================
   KH HẸN THAY MODEM
   ============================================================ */
function dayDiff(iso) {
  const p = String(iso || '').split('-').map(Number);
  if (p.length !== 3 || !p[0]) return 0;
  const a = new Date(p[0], p[1] - 1, p[2]);
  const n = new Date();
  const b = new Date(n.getFullYear(), n.getMonth(), n.getDate());
  return Math.round((a - b) / 86400000);
}

function apptDays(a) {
  if (a.trang_thai === 'DA_THAY') return '<span class="badge hen-da">Đã thay</span>';
  if (a.trang_thai === 'HUY') return '<span class="badge hen-huy">Đã hủy</span>';
  const d = dayDiff(a.ngay_hen);
  if (d < 0) return `<span class="appt-days over">Quá ${-d} ngày</span>`;
  if (d === 0) return '<span class="appt-days today">Hôm nay</span>';
  if (d === 1) return '<span class="appt-days">Ngày mai</span>';
  return `<span class="appt-days">Còn ${d} ngày</span>`;
}

function apptStatusBadge(s) {
  if (s === 'DA_THAY') return '<span class="badge hen-da">Đã thay</span>';
  if (s === 'HUY') return '<span class="badge hen-huy">Đã hủy</span>';
  return '<span class="badge hen-cho">Chờ thay</span>';
}

function apptList() {
  const f = UI.appt;
  const q = norm(f.q || '');
  return APPOINTMENTS.filter(a => {
    if (f.status && f.status !== 'all' && a.trang_thai !== f.status) return false;
    if (f.from && a.ngay_hen < f.from) return false;
    if (f.to && a.ngay_hen > f.to) return false;
    if (f.emp && a.nguoi !== f.emp) return false;
    if (f.kv && a.kv !== f.kv) return false;
    if (q) {
      const hay = norm(a.ma) + ' ' + norm(a.ten) + ' ' + norm(a.sdt) + ' ' + norm(a.dc) + ' ' + norm(a.nguoi) + ' ' + norm(a.ghi_chu) + ' ' + norm(a.kv);
      if (hay.indexOf(q) === -1) return false;
    }
    return true;
  }).sort((x, y) => String(x.ngay_hen).localeCompare(String(y.ngay_hen)) ||
    String(x.gio_hen || '').localeCompare(String(y.gio_hen || '')) ||
    String(x.ten).localeCompare(String(y.ten), 'vi'));
}

function apptRowHtml(a, idx) {
  const diff = dayDiff(a.ngay_hen);
  const rowCls = a.trang_thai === 'CHO_THAY' ? (diff < 0 ? 'appt-overdue' : diff === 0 ? 'appt-today' : '') : '';
  const addr = [a.sn, a.duong].filter(Boolean).join(' ') || a.dc || '';
  const phone = cleanText(a.sdt);
  const svc = (a.dich_vu || []).length
    ? `<div class="svc-inline">${a.dich_vu.map(k => `<span class="chip ${k}">${esc(SVC[k] ? SVC[k].short : k)}</span>`).join('')}</div>`
    : '<span class="chip none">Không kèm</span>';
  const actions = a.trang_thai === 'CHO_THAY'
    ? `<div class="appt-actions">
        <button type="button" class="btn sm green" data-act="appt-done" data-id="${esc(a.id)}">✓ Đã thay</button>
        <button type="button" class="btn sm" data-act="appt-reschedule" data-id="${esc(a.id)}">Dời</button>
        <button type="button" class="btn sm danger" data-act="appt-cancel" data-id="${esc(a.id)}">Hủy</button>
        <button type="button" class="icon-btn" data-act="open" data-ma="${esc(a.ma)}" title="Mở hộ">↗</button>
      </div>`
    : `<div class="appt-actions"><button type="button" class="icon-btn" data-act="open" data-ma="${esc(a.ma)}" title="Mở hộ">↗</button></div>`;
  return `<tr class="${rowCls}">
    <td class="col-stt hide-sm">${idx}</td>
    <td>
      <div class="appt-date">${fmtDate(a.ngay_hen)}</div>
      <div class="touch-when">${esc(a.gio_hen || '')}</div>
      ${apptDays(a)}
    </td>
    <td>
      <div class="cust-name">${esc(displayName(a))}</div>
      <div class="cust-meta"><span class="mono">${esc(a.ma)}</span>${/^0\d{8,10}$/.test(cleanPhone(phone)) ? `<a class="phone-link" href="tel:${esc(cleanPhone(phone))}">${esc(phone)}</a>` : '<span class="muted" style="font-size:11px">Chưa có SĐT</span>'}</div>
    </td>
    <td><div class="addr"><div class="addr-main">${esc(addr)}</div><div class="addr-sub">${esc([a.ap, a.xa].filter(Boolean).join(', '))}</div></div></td>
    <td class="hide-sm"><div class="mono" style="font-weight:700">${esc(a.kv)}</div><div class="addr-sub">${esc(a.xa || '')}</div></td>
    <td>${svc}</td>
    <td><div class="emp-name" style="font-weight:600">${esc(a.nguoi || '—')}</div><div class="touch-when">${esc(unitLabel(a.donvi))}</div></td>
    <td class="hide-sm"><div style="max-width:190px;font-size:12px;color:#4a5b70">${esc(a.ghi_chu || '')}</div></td>
    <td>${apptStatusBadge(a.trang_thai)}</td>
    <td>${actions}</td>
  </tr>`;
}

function renderAppointments() {
  const root = $('#appt-root');
  if (!root || !LOADED) return;
  const f = UI.appt;
  const emps = uniq(APPOINTMENTS.map(a => a.nguoi)).sort(byName);
  const kvs = uniq(ALL.map(r => r.kv)).sort(byName);
  root.innerHTML = `
  <div class="unit-hero appt">
    <div>
      <h2>📅 KH hẹn thay modem</h2>
      <p>Khách đã hẹn ngày thay modem — sắp xếp theo ngày hẹn gần nhất. Lưu hẹn không tạo phiếu chạm.</p>
    </div>
    <div class="unit-kpis">
      <div class="ukpi"><b id="ap-kpi-active">0</b><span>Chờ thay</span></div>
      <div class="ukpi"><b id="ap-kpi-today">0</b><span>Hôm nay</span></div>
      <div class="ukpi"><b id="ap-kpi-over">0</b><span>Quá hẹn</span></div>
      <div class="ukpi"><b id="ap-kpi-done">0</b><span>Đã thay</span></div>
    </div>
  </div>

  <div class="card">
    <div class="toolbar">
      <div class="fg grow"><label>Tìm khách, SĐT, mã TB, ghi chú</label><input id="ap-q" value="${esc(f.q)}" placeholder="Gõ để tìm…" autocomplete="off"></div>
      <div class="fg"><label>Từ ngày hẹn</label><input type="date" id="ap-from" value="${esc(f.from)}"></div>
      <div class="fg"><label>Đến ngày hẹn</label><input type="date" id="ap-to" value="${esc(f.to)}"></div>
      <div class="fg"><label>Trạng thái</label>
        <select id="ap-status">
          <option value="CHO_THAY" ${f.status === 'CHO_THAY' ? 'selected' : ''}>Chờ thay</option>
          <option value="DA_THAY" ${f.status === 'DA_THAY' ? 'selected' : ''}>Đã thay</option>
          <option value="HUY" ${f.status === 'HUY' ? 'selected' : ''}>Đã hủy</option>
          <option value="all" ${f.status === 'all' ? 'selected' : ''}>Tất cả</option>
        </select>
      </div>
      <div class="fg"><label>Người hẹn</label>
        <select id="ap-emp"><option value="">Tất cả</option>${emps.map(e => `<option value="${esc(e)}" ${f.emp === e ? 'selected' : ''}>${esc(e)}</option>`).join('')}</select>
      </div>
      <div class="fg"><label>Khu vực</label>
        <select id="ap-kv"><option value="">Tất cả KV</option>${kvs.map(k => `<option value="${esc(k)}" ${f.kv === k ? 'selected' : ''}>${esc(k)}</option>`).join('')}</select>
      </div>
      <button type="button" class="btn sm" data-act="reset-appt">↺ Xóa lọc</button>
      <div class="spacer"></div>
      <button type="button" class="btn primary sm" data-act="export-appt">⬇ Tải Excel</button>
    </div>
    <div style="padding:10px 16px;border-bottom:1px solid var(--line);display:flex;gap:18px;flex-wrap:wrap;font-size:12px;color:var(--gray)" id="ap-sum"></div>
    <div class="table-wrap" style="max-height:none">
      <table>
        <thead><tr>
          <th class="col-stt hide-sm">#</th><th>Ngày hẹn</th><th>Khách hàng</th><th>Địa chỉ</th>
          <th class="hide-sm">Khu vực</th><th>Dịch vụ</th><th>Người hẹn</th>
          <th class="hide-sm">Ghi chú</th><th>Trạng thái</th><th style="width:235px"></th>
        </tr></thead>
        <tbody id="ap-body"></tbody>
      </table>
    </div>
    <div id="ap-empty" class="empty hidden"><div class="big">📅</div><div>Không có khách hẹn thay nào khớp bộ lọc.</div></div>
    <div class="pager"><div class="pager-info" id="ap-count"></div></div>
  </div>`;

  ['q', 'from', 'to', 'status', 'emp', 'kv'].forEach(k => {
    const el = $('#ap-' + k);
    if (!el) return;
    el.addEventListener(k === 'q' ? 'input' : 'change', debounce(() => {
      UI.appt[k] = el.value;
      refreshApptTable();
    }, k === 'q' ? 220 : 0));
  });
  refreshApptTable();
}

function refreshApptTable() {
  const list = apptList();
  const shown = list.slice(0, 400);
  const body = $('#ap-body');
  if (!body) return;
  body.innerHTML = shown.map((a, i) => apptRowHtml(a, i + 1)).join('');
  $('#ap-empty').classList.toggle('hidden', list.length > 0);
  const active = APPOINTMENTS.filter(a => a.trang_thai === 'CHO_THAY');
  const today = todayISO();
  const td = active.filter(a => a.ngay_hen === today).length;
  const over = active.filter(a => a.ngay_hen < today).length;
  const done = APPOINTMENTS.filter(a => a.trang_thai === 'DA_THAY').length;
  $('#ap-kpi-active').textContent = nf(active.length);
  $('#ap-kpi-today').textContent = nf(td);
  $('#ap-kpi-over').textContent = nf(over);
  $('#ap-kpi-done').textContent = nf(done);
  const ho = new Set(list.map(a => a.ma));
  $('#ap-sum').innerHTML = `
    <span><b style="color:var(--ink)">${nf(list.length)}</b> hẹn</span>
    <span><b style="color:var(--ink)">${nf(ho.size)}</b> hộ</span>
    <span style="color:var(--red)"><b>${nf(list.filter(a => a.trang_thai === 'CHO_THAY' && a.ngay_hen < today).length)}</b> quá hẹn</span>
    <span style="color:var(--green)"><b>${nf(list.filter(a => a.trang_thai === 'DA_THAY').length)}</b> đã thay</span>
    <span class="muted">Sắp xếp theo ngày hẹn gần nhất</span>`;
  $('#ap-count').innerHTML = list.length
    ? `Hiển thị ${nf(shown.length)} / ${nf(list.length)} hẹn${list.length > 400 ? ' — tải Excel để xem đầy đủ' : ''}`
    : 'Chưa có hẹn thay nào';
}

function exportAppointments() {
  const list = apptList();
  if (!list.length) return toast('Không có hẹn thay để xuất', 'warn');
  const headers = ['STT', 'Ngày hẹn', 'Giờ hẹn', 'Mã TB', 'Tên khách hàng', 'Số điện thoại', 'Địa chỉ',
    'Khu vực', 'Xã', 'Ấp', 'Đường', 'Số nhà', 'Dịch vụ tiếp thị', 'Ghi chú', 'Người hẹn', 'Đơn vị',
    'Trạng thái', 'Ngày tạo', 'Cập nhật'];
  const rows = list.map((a, i) => [
    i + 1, fmtDate(a.ngay_hen), a.gio_hen || '', a.ma, a.ten, a.sdt, a.dc,
    a.kv, a.xa, a.ap, a.duong, a.sn,
    (a.dich_vu || []).map(k => SVC[k] ? SVC[k].label : k).join(', '),
    a.ghi_chu, a.nguoi, unitLabel(a.donvi),
    apptStatusLabel(a.trang_thai), a.ngay_tao || '', a.cap_nhat || ''
  ]);
  downloadXlsx(`KH_hen_thay_${stamp()}.xlsx`, [{ name: 'KH hen thay', rows: [headers].concat(rows) }]);
  toast(`Đã xuất ${nf(list.length)} hẹn thay ra Excel`, 'ok');
}


/* ============================================================
   ĐIỀU KHIỂN SỰ KIỆN & KHỞI ĐỘNG
   ============================================================ */
function debounce(fn, wait) {
  let t = null;
  return function () {
    const args = arguments, ctx = this;
    clearTimeout(t);
    t = setTimeout(() => fn.apply(ctx, args), wait || 200);
  };
}

function showTab(tab) {
  closeEmpPanel();
  closeUserMenu();
  UI.tab = tab;
  $$('.tab-btn').forEach(b => b.classList.toggle('active', b.dataset.tab === tab));
  $$('.tab-pane').forEach(p => p.classList.toggle('active', p.id === 'tab-' + tab));
  if (tab === 'appt') renderAppointments();
  if (tab === 'ttvt') renderUnit('TTVT');
  if (tab === 'vnpt') renderUnit('VNPT');
  if (tab === 'report') renderReport();
  if (tab === 'stats') renderStats();
  if (tab === 'guide') renderGuide();
  if (tab === 'list') renderList();
  try { history.replaceState(null, '', '#' + tab); } catch (e) {}
  window.scrollTo({ top: 0, behavior: 'smooth' });
}

async function deleteTouch(id) {
  const t = TOUCHES.filter(x => String(x.id) === String(id))[0];
  if (!t) return;
  const c = touchCustomer(t);
  const ok = await confirmDialog({
    title: 'Xóa lượt chạm này?',
    msg: `Lượt chạm <b>${esc(t.donvi)} · ${esc(t.nguoi)} · ${fmtDate(t.ngay)} ${esc(t.gio || '')}</b> của hộ <b>${esc(c.ten)}</b> sẽ bị xóa khỏi sheet <code>CHAM_TIEP_THI</code>. Cột trạng thái của hộ sẽ được tính lại từ các lượt còn lại.`,
    yes: 'Xóa lượt chạm'
  });
  if (!ok) return;
  try {
    await API.deleteTouch(id);
    TOUCHES = TOUCHES.filter(x => String(x.id) !== String(id));
    recomputeTouches();
    renderAll();
    toast('Đã xóa lượt chạm và cập nhật lại sheet', 'ok');
  } catch (e) {
    if (e.code === 'AUTH') { API.clearToken(); showLogin(e.message); return; }
    toast('Không xóa được: ' + e.message, 'err', 6000);
  }
}

function resetFilters() {
  closeEmpPanel();
  UI.filters = { q: '', to: '', kv: '', nv: '', xa: '', ap: '', duong: '', thay: 'chua', cham: 'all', hen: 'all', need: [], sort: 'uu_tien', topN: 0 };
  UI.page = 1;
  $('#f-q').value = ''; $('#f-duong').value = '';
  $('#f-to').value = ''; $('#f-kv').value = ''; $('#f-xa').value = ''; $('#f-ap').value = '';
  rebuildNvSelect(); rebuildApSelect(); rebuildDuongDatalist();
  renderFiltersUI(); renderList();
}

/* ---------- đăng nhập / người dùng ---------- */
function showLogin(msg) {
  hideLoading();
  const el = $('#login-screen');
  if (!el) return;
  el.classList.remove('hidden');
  const err = $('#login-error');
  if (err) {
    err.textContent = msg || '';
    err.classList.toggle('hidden', !msg);
  }
  const btn = $('#login-btn');
  if (btn) { btn.disabled = false; btn.textContent = 'Đăng nhập'; }
  setTimeout(() => {
    const u = $('#login-user');
    if (u && !u.value) u.focus();
  }, 100);
}

function hideLogin() {
  const el = $('#login-screen');
  if (el) el.classList.add('hidden');
}

async function handleLogin() {
  const u = ($('#login-user').value || '').trim();
  const p = $('#login-pass').value || '';
  if (!u || !p) { showLogin('Nhập tài khoản và mật khẩu.'); return; }
  const btn = $('#login-btn');
  btn.disabled = true;
  btn.textContent = '⏳ Đang đăng nhập…';
  try {
    const res = await API.login(u, p);
    USER = res.user; API.user = USER;
    $('#login-pass').value = '';
    hideLogin();
    await loadData();
  } catch (e) {
    btn.disabled = false;
    btn.textContent = 'Đăng nhập';
    showLogin(e.message);
  }
}

function openUserMenu() {
  const m = $('#user-menu');
  if (!m) return;
  m.classList.add('show');
  $('#user-menu-btn').setAttribute('aria-expanded', 'true');
}
function closeUserMenu() {
  const m = $('#user-menu');
  if (!m) return;
  m.classList.remove('show');
  const b = $('#user-menu-btn');
  if (b) b.setAttribute('aria-expanded', 'false');
}

function openPasswordModal() {
  closeUserMenu();
  $('#pw-old').value = ''; $('#pw-new').value = ''; $('#pw-confirm').value = '';
  const err = $('#pw-error');
  err.textContent = '';
  err.className = 'api-status';
  $('#pw-modal').classList.add('show');
  $('#pw-overlay').classList.add('show');
  setTimeout(() => $('#pw-old').focus(), 100);
}
function closePasswordModal() {
  $('#pw-modal').classList.remove('show');
  $('#pw-overlay').classList.remove('show');
}
async function savePassword() {
  const oldPw = $('#pw-old').value || '';
  const newPw = $('#pw-new').value || '';
  const conf = $('#pw-confirm').value || '';
  const err = $('#pw-error');
  const fail = m => { err.className = 'api-status err'; err.textContent = m; };
  if (!oldPw) return fail('Nhập mật khẩu hiện tại.');
  if (newPw.length < 6) return fail('Mật khẩu mới phải có ít nhất 6 ký tự.');
  if (newPw !== conf) return fail('Nhập lại mật khẩu mới không khớp.');
  if (newPw === oldPw) return fail('Mật khẩu mới phải khác mật khẩu hiện tại.');
  const btn = $('[data-act="save-password"]');
  const oldText = btn.textContent;
  btn.disabled = true; btn.textContent = 'Đang lưu…';
  err.className = 'api-status'; err.textContent = 'Đang cập nhật…';
  try {
    const res = await API.changePassword(oldPw, newPw);
    USER = res.user || USER; API.user = USER;
    renderUserUI();
    closePasswordModal();
    toast('Đã đổi mật khẩu thành công', 'ok');
  } catch (e) {
    btn.disabled = false; btn.textContent = oldText;
    if (e.code === 'AUTH') { closePasswordModal(); API.clearToken(); showLogin(e.message); return; }
    fail(e.message);
  }
}

async function doLogout() {
  closeUserMenu();
  const ok = await confirmDialog({ title: 'Đăng xuất?', msg: 'Bạn sẽ cần đăng nhập lại để tiếp tục nhập liệu.', yes: 'Đăng xuất', danger: false });
  if (!ok) return;
  await API.logout();
  location.reload();
}

/* ---------- thao tác hẹn thay ---------- */
async function apptSetStatus(id, status) {
  const a = APPOINTMENTS.filter(x => String(x.id) === String(id))[0];
  if (!a) return;
  const done = status === 'DA_THAY';
  const c = { ten: a.ten, ngay: a.ngay_hen, gio: a.gio_hen };
  const ok = await confirmDialog({
    title: done ? 'Xác nhận đã thay modem?' : 'Hủy hẹn thay?',
    msg: done
      ? `Hộ <b>${esc(c.ten)}</b> — hẹn ngày <b>${fmtDate(c.ngay)} ${esc(c.gio || '')}</b>. Xác nhận modem đã được thay? Trạng thái hộ sẽ chuyển sang <b>Đã thay</b>.`
      : `Hủy hẹn của hộ <b>${esc(c.ten)}</b> ngày <b>${fmtDate(c.ngay)}</b>? Hộ sẽ trở lại danh sách chưa hẹn.`,
    yes: done ? 'Đã thay modem' : 'Hủy hẹn'
  });
  if (!ok) return;
  try {
    await API.updateAppointment({ id: id, status: status });
    a.trang_thai = status;
    a.cap_nhat = nowHHMM();
    if (done) {
      const r = BY_MA.get(a.ma);
      if (r) r.daThay = true;
    }
    recomputeAppointments();
    renderAll();
    toast(done ? 'Đã xác nhận thay modem' : 'Đã hủy hẹn', 'ok');
  } catch (e) {
    if (e.code === 'AUTH') { API.clearToken(); showLogin(e.message); return; }
    toast('Không cập nhật được hẹn: ' + e.message, 'err', 6000);
  }
}

/* ---------- sự kiện ---------- */
function bindEvents() {
  /* đăng nhập */
  $('#login-btn').addEventListener('click', handleLogin);
  $('#login-pass').addEventListener('keydown', e => { if (e.key === 'Enter') handleLogin(); });
  $('#login-user').addEventListener('keydown', e => { if (e.key === 'Enter') $('#login-pass').focus(); });
  $('#login-eye').addEventListener('click', () => {
    const inp = $('#login-pass');
    inp.type = inp.type === 'password' ? 'text' : 'password';
    $('#login-eye').style.opacity = inp.type === 'text' ? '1' : '.65';
  });

  /* menu người dùng */
  $('#user-menu-btn').addEventListener('click', e => {
    e.stopPropagation();
    const m = $('#user-menu');
    if (m.classList.contains('show')) closeUserMenu(); else openUserMenu();
  });

  $('#btn-refresh').addEventListener('click', async () => {
    if (!API.token) return showLogin('');
    $('#conn').className = 'conn';
    $('#conn-text').textContent = 'Đang làm mới…';
    try {
      const res = await API.bootstrap();
      if (!res || !res.rows) throw new Error('Dữ liệu trả về không hợp lệ');
      META = res.meta || META;
      USER = res.user || USER; API.user = USER;
      ALL = res.rows.map((a, i) => ({
        id: i, ma: a[0], ten: a[1], sdt: a[2], to: a[3], nv: a[4], kv: a[5],
        traSau: Number(a[6]) === 1 || a[6] === true, dc: a[7], sn: a[8], duong: a[9], ap: a[10], xa: a[11],
        modem: a[12], daThay: Number(a[13]) === 1 || a[13] === true,
        mesh: a[14] || '', cam: a[15] || '', mytv: a[16] || ''
      }));
      BY_MA = new Map(ALL.map(r => [r.ma, r]));
      TOUCHES = res.touches || [];
      APPOINTMENTS = res.appointments || [];
      recomputeTouches(); recomputeAppointments(); buildOptions(); renderUserUI(); renderAll();
      $('#conn').className = 'conn ok';
      $('#conn-text').textContent = `Đã kết nối · ${nf(ALL.length)} hộ`;
      toast('Đã đọc lại dữ liệu mới nhất từ Google Sheet', 'ok');
    } catch (e) {
      $('#conn').className = 'conn err';
      $('#conn-text').textContent = 'Mất kết nối';
      if (e.code === 'AUTH') { API.clearToken(); showLogin(e.message); return; }
      toast('Không làm mới được: ' + e.message, 'err', 6000);
    }
  });

  /* bộ lọc danh sách */
  $('#f-q').addEventListener('input', debounce(e => { UI.filters.q = e.target.value; UI.page = 1; renderList(); }, 250));
  const bindSel = (sel, fn) => $(sel).addEventListener('change', fn);
  bindSel('#f-to', e => { closeEmpPanel(); UI.filters.to = e.target.value; UI.filters.nv = ''; UI.page = 1; rebuildNvSelect(); renderList(); });
  bindSel('#f-kv', e => { UI.filters.kv = e.target.value; UI.page = 1; renderList(); });
  bindSel('#f-xa', e => {
    UI.filters.xa = e.target.value; UI.filters.ap = ''; UI.filters.duong = '';
    UI.page = 1; rebuildApSelect(); rebuildDuongDatalist();
    $('#f-duong').value = ''; renderList();
  });
  bindSel('#f-ap', e => {
    UI.filters.ap = e.target.value; UI.filters.duong = '';
    UI.page = 1; rebuildDuongDatalist(); $('#f-duong').value = ''; renderList();
  });
  $('#f-duong').addEventListener('input', debounce(e => { UI.filters.duong = e.target.value; UI.page = 1; renderList(); }, 250));
  bindSel('#f-cham', e => { UI.filters.cham = e.target.value; UI.page = 1; renderList(); });
  bindSel('#f-hen', e => { UI.filters.hen = e.target.value; UI.page = 1; renderList(); });
  bindSel('#f-sort', e => { UI.filters.sort = e.target.value; UI.page = 1; renderList(); });

  $$('#need-chips .need-chip').forEach(b => b.addEventListener('click', () => {
    const k = b.dataset.svc, i = UI.filters.need.indexOf(k);
    if (i > -1) UI.filters.need.splice(i, 1); else UI.filters.need.push(k);
    UI.page = 1; renderFiltersUI(); renderList();
  }));

  $('#btn-reset-filters').addEventListener('click', resetFilters);

  /* lọc nhanh */
  $('#qf-thay').addEventListener('click', e => {
    const b = e.target.closest('[data-v]');
    if (!b) return;
    setQuickThay(b.dataset.v);
  });
  $('#qf-emp-btn').addEventListener('click', () => openEmpPanel());
  $('#qf-emp-close').addEventListener('click', e => { e.stopPropagation(); closeEmpPanel(); });
  $('#qf-emp-search').addEventListener('input', debounce(e => renderEmpList(e.target.value), 150));
  $('#qf-emp-list').addEventListener('click', e => {
    const b = e.target.closest('[data-nv]');
    if (!b) return;
    setQuickNv(b.dataset.nv || '');
  });

  $('#btn-filter-toggle').addEventListener('click', () => {
    const body = $('#filter-body'), btn = $('#btn-filter-toggle');
    const open = body.classList.toggle('open');
    btn.setAttribute('aria-expanded', open ? 'true' : 'false');
    $('#filter-chevron').textContent = open ? '▴' : '▾';
  });
  $('#btn-top15').addEventListener('click', () => {
    UI.filters.cham = 'chua'; UI.filters.thay = 'chua'; UI.filters.topN = 15;
    if (UI.filters.sort !== 'uu_tien') UI.filters.sort = 'uu_tien';
    UI.page = 1;
    renderFiltersUI(); renderList();
    toast('Đã lấy 15 hộ ưu tiên nhất theo bộ lọc hiện tại', 'ok');
    $('#list-body') && $('#list-body').closest('.table-card').scrollIntoView({ behavior: 'smooth', block: 'start' });
  });
  $('#btn-clear-top15').addEventListener('click', () => {
    UI.filters.topN = 0; UI.page = 1; renderFiltersUI(); renderList();
  });

  /* click toàn cục */
  document.addEventListener('click', e => {
    if (e.target.closest('a[href^="tel:"]')) return;
    if (!e.target.closest('#user-box')) closeUserMenu();
    const modeBtn = e.target.closest('[data-mode]');
    if (modeBtn && $('#drawer').classList.contains('open')) { setFormMode(modeBtn.dataset.mode); return; }
    const tab = e.target.closest('.tab-btn');
    if (tab) { showTab(tab.dataset.tab); return; }
    const el = e.target.closest('[data-act]');
    if (!el) {
      if (e.target.id === 'overlay') closeDrawer();
      if (e.target.id === 'confirm-overlay') closeConfirm(false);
      if (e.target.id === 'pw-overlay') closePasswordModal();
      if (!e.target.closest('#qf-emp')) closeEmpPanel();
      return;
    }
    if (!e.target.closest('#qf-emp')) closeEmpPanel();
    const act = el.dataset.act;
    switch (act) {
      case 'open': e.preventDefault(); openDrawer(el.dataset.ma, el.dataset.mode || 'touch'); break;
      case 'close-drawer': closeDrawer(); break;
      case 'save-touch': if (FORM_MODE === 'appt') saveAppointment(); else saveTouch(); break;
      case 'page': UI.page = parseInt(el.dataset.p, 10) || 1; renderList(); break;
      case 'export-list': exportList(); break;
      case 'retry-load': loadData(); break;
      case 'confirm-yes': closeConfirm(true); break;
      case 'confirm-no': closeConfirm(false); break;
      case 'copy-script': copyText(SCRIPT_TXT[el.dataset.key] || ''); break;
      case 'del-touch': e.stopPropagation(); deleteTouch(el.dataset.id); break;
      case 'export-unit': exportUnit(el.dataset.unit); break;
      case 'reset-unit': UI.unit[el.dataset.unit] = { q: '', from: '', to: '', emp: '' }; renderUnit(el.dataset.unit); break;
      case 'export-report': exportReport(); break;
      case 'print-report': window.print(); break;
      case 'reset-report': UI.report = { q: '', from: '', to: '', donvi: '', team: '', emp: '', xa: '', kq: '', svc: '' }; renderReport(); break;
      case 'export-stats': exportStats(el.dataset.key); break;
      case 'export-all-stats': exportAllStats(); break;
      case 'export-structure': exportStructure(); break;
      case 'export-appt': exportAppointments(); break;
      case 'reset-appt': UI.appt = { q: '', from: '', to: '', status: 'CHO_THAY', emp: '', kv: '' }; renderAppointments(); break;
      case 'appt-done': apptSetStatus(el.dataset.id, 'DA_THAY'); break;
      case 'appt-cancel': apptSetStatus(el.dataset.id, 'HUY'); break;
      case 'appt-reschedule': { const a = APPOINTMENTS.filter(x => String(x.id) === String(el.dataset.id))[0]; if (a) openDrawer(a.ma, 'appt'); break; }
      case 'stats-this-month': {
        const d = new Date();
        UI.stats.from = `${d.getFullYear()}-${p2(d.getMonth() + 1)}-01`;
        UI.stats.to = todayISO();
        renderStats(); break;
      }
      case 'reset-stats': UI.stats = { from: '', to: '', unit: '' }; renderStats(); break;
      case 'change-password': openPasswordModal(); break;
      case 'logout': doLogout(); break;
      case 'pw-cancel': closePasswordModal(); break;
      case 'save-password': savePassword(); break;
      case 'save-target': {
        const v = Math.max(1, Math.min(200, parseInt($('#cfg-target').value, 10) || 15));
        CONFIG.target = v; saveLocal(); renderKPIs(); toast('Đã lưu mục tiêu ' + v + ' lượt/người/ngày', 'ok'); break;
      }
    }
  });

  /* đổi trang / checkbox dịch vụ */
  document.addEventListener('change', e => {
    if (e.target.matches('[data-act="pagesize"]')) {
      UI.pageSize = parseInt(e.target.value, 10) || 50; UI.page = 1; renderList(); return;
    }
    const apptChk = e.target.closest('[data-appt-svc]');
    if (apptChk) {
      const row = apptChk.closest('.appt-svc-row');
      if (row) {
        row.classList.toggle('picked', apptChk.checked);
        const ok = row.querySelector('.svc-ok');
        if (ok) ok.textContent = apptChk.disabled ? '—' : (apptChk.checked ? 'Đã chọn' : 'Chọn');
      }
      return;
    }
    const chk = e.target.closest('[data-svc-chk]');
    if (chk) {
      const kind = chk.dataset.svcChk, k = chk.dataset.svc;
      if (kind === 'tu_van') { if (chk.checked) FORM_SVC.tu_van.add(k); else { FORM_SVC.tu_van.delete(k); FORM_SVC.dong_y.delete(k); } }
      if (kind === 'dong_y') { if (chk.checked) { FORM_SVC.dong_y.add(k); FORM_SVC.tu_van.add(k); } else FORM_SVC.dong_y.delete(k); }
      syncServiceChecks();
      return;
    }
  });

  /* phím tắt */
  document.addEventListener('keydown', e => {
    if (e.key === 'Escape') {
      if ($('#user-menu').classList.contains('show')) { closeUserMenu(); return; }
      if ($('#pw-modal').classList.contains('show')) { closePasswordModal(); return; }
      if ($('#qf-emp-pop').classList.contains('show')) { closeEmpPanel(); return; }
      if ($('#confirm-modal').classList.contains('show')) closeConfirm(false);
      else if ($('#drawer').classList.contains('open')) closeDrawer();
    }
    if (e.key === 'Enter' && (e.ctrlKey || e.metaKey) && $('#drawer').classList.contains('open')) {
      e.preventDefault();
      if (FORM_MODE === 'appt') saveAppointment(); else saveTouch();
    }
    if (e.key === '/' && !/INPUT|TEXTAREA|SELECT/.test(document.activeElement.tagName) && UI.tab === 'list') {
      e.preventDefault(); $('#f-q').focus();
    }
  });
}

/* ---------- khởi động ---------- */
async function boot() {
  loadLocal();
  API.init();
  bindEvents();

  if (API.mode === 'http' && !API.available) {
    hideLogin();
    showLoading();
    showLoadError({ message: 'NO_API' });
    return;
  }
  if (!API.token) {
    hideLoading();
    showLogin();
    return;
  }
  showLoading('Đang kiểm tra phiên đăng nhập…');
  try {
    const u = await API.me();
    USER = u; API.user = u;
    hideLogin();
    renderUserUI();
    await loadData();
  } catch (e) {
    hideLoading();
    API.clearToken();
    showLogin(e.code === 'AUTH' ? 'Phiên đăng nhập đã hết hạn. Vui lòng đăng nhập lại.' : ('Không kết nối được API: ' + e.message));
  }
}

if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', boot);
else boot();

})();
