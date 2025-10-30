// =======================
// Konstanta & Variabel Global
// =======================
const ROLL_PCT = { SLOT: 0.003, CASINO: 0.006, POKER: 0.003 };
const CB_SLOT_PCT = 0.05;

let specialCases = {};
let allData = [];
let finishing = [];
let summaryData = {};
let finalData = [];
let tiketData = []; // hanya 1 deklarasi untuk tiket

// small global lock for processing (prevents double click across heavy operations)
let isProcessing = false;

// =======================
// Inject spinner & modal CSS
// =======================
(function injectSpinnerStyle() {
  if (document.getElementById("script-global-spinner")) return;
  const css = `
    @keyframes spin { from{ transform: rotate(0deg); } to{ transform: rotate(360deg); } }
    .btn-loading { position: relative; opacity: 0.9; pointer-events: none; }
    .btn-loading::before {
      content: "";
      position: absolute;
      left: 10px;
      top: 50%;
      width: 14px;
      height: 14px;
      margin-top: -7px;
      border: 2px solid rgba(255,255,255,0.7);
      border-top-color: #ffb200;
      border-radius: 50%;
      animation: spin 0.7s linear infinite;
    }
    .special-modal-overlay {
      position: fixed;
      inset: 0;
      background: rgba(0,0,0,0.45);
      display: flex;
      align-items: center;
      justify-content: center;
      z-index: 9999;
      padding: 20px;
    }
    .special-modal {
      width: 760px;
      max-width: 100%;
      background: #fff;
      border-radius: 10px;
      box-shadow: 0 10px 30px rgba(0,0,0,0.25);
      padding: 18px;
      font-family: system-ui, -apple-system, "Segoe UI", Roboto, "Helvetica Neue", Arial;
      position: relative;
    }
    .special-modal h3 { margin: 0 0 8px 0; font-size: 16px; }
    .special-modal .row { display:flex; gap:10px; margin-bottom:8px; align-items:center; }
    .special-modal .col { flex: 1; }
    .special-modal input[type="text"] { width:100%; padding:8px 10px; border-radius:6px; border:1px solid #ddd; }
    .special-modal .actions { display:flex; gap:8px; justify-content:flex-end; margin-top:12px; }
    .pill { display:inline-flex; align-items:center; gap:8px; background:#eef2ff; padding:6px 10px; border-radius:999px; margin:4px; }
    .pill button { margin-left:6px; background:transparent; border:none; cursor:pointer; font-weight:bold; color:#a00; }
    .special-modal .close-btn { position:absolute; right:18px; top:12px; background:transparent; border:none; font-size:18px; cursor:pointer; }
    .special-input-group { display:grid; grid-template-columns: repeat(5, 1fr); gap:8px; margin-top:6px; }
    .muted { color:#6b7280; font-size:14px; }

    /* ensure primary buttons have consistent height so spinner centers nicely */
    button.primary { display: inline-flex; align-items: center; gap:8px; justify-content:center; min-height:40px; padding:10px 14px; }
  `;
  const s = document.createElement("style");
  s.id = "script-global-spinner";
  s.textContent = css;
  document.head.appendChild(s);
})();

// =======================
// Ensure Special Modal DOM exists (create if missing)
// =======================
(function ensureSpecialModalDOM() {
  if (document.getElementById("specialModal")) return;

  const overlay = document.createElement("div");
  overlay.id = "specialModal";
  overlay.className = "special-modal-overlay";
  overlay.style.display = "none";

  overlay.innerHTML = `
    <div class="special-modal" role="dialog" aria-modal="true">
      <button class="close-btn" id="closeSpecialModal" title="Close">&times;</button>
      <h3>Case Spesial — Override Rate</h3>
      <p class="muted">Masukkan Account ID (maks 5 per group). ID akan dibersihkan otomatis.</p>

      <div style="margin-top:10px;">
        <label><b>Slot — rate 0.5%</b></label>
        <div class="special-input-group" id="slot05_inputs">
          <input id="slot05_1" type="text" placeholder="ID 1" />
          <input id="slot05_2" type="text" placeholder="ID 2" />
          <input id="slot05_3" type="text" placeholder="ID 3" />
          <input id="slot05_4" type="text" placeholder="ID 4" />
          <input id="slot05_5" type="text" placeholder="ID 5" />
        </div>
        <div style="text-align:right; margin-top:8px;">
          <button id="confirmSlot05" class="btn primary">Tambah Slot 0.5%</button>
        </div>
      </div>

      <hr style="margin:12px 0;" />

      <div>
        <label><b>Casino — rate 0.8%</b></label>
        <div class="special-input-group" id="casino08_inputs">
          <input id="casino08_1" type="text" placeholder="ID 1" />
          <input id="casino08_2" type="text" placeholder="ID 2" />
          <input id="casino08_3" type="text" placeholder="ID 3" />
          <input id="casino08_4" type="text" placeholder="ID 4" />
          <input id="casino08_5" type="text" placeholder="ID 5" />
        </div>
        <div style="text-align:right; margin-top:8px;">
          <button id="confirmCasino08" class="btn primary">Tambah Casino 0.8%</button>
        </div>
      </div>

      <hr style="margin:12px 0;" />

      <div>
        <label><b>Casino — rate 1.0%</b></label>
        <div class="special-input-group" id="casino10_inputs">
          <input id="casino10_1" type="text" placeholder="ID 1" />
          <input id="casino10_2" type="text" placeholder="ID 2" />
          <input id="casino10_3" type="text" placeholder="ID 3" />
          <input id="casino10_4" type="text" placeholder="ID 4" />
          <input id="casino10_5" type="text" placeholder="ID 5" />
        </div>
        <div style="text-align:right; margin-top:8px;">
          <button id="confirmCasino10" class="btn primary">Tambah Casino 1.0%</button>
        </div>
      </div>

      <div style="margin-top:12px;">
        <h4>Active Cases</h4>
        <div id="activeCases" style="min-height:36px"></div>
      </div>

      <div class="actions" style="margin-top:8px;">
        <button id="closeSpecialModalFooter" class="btn">Tutup</button>
      </div>
    </div>
  `;

  document.body.appendChild(overlay);

  if (!document.getElementById("openSpecialBtn")) {
    const openBtn = document.createElement("button");
    openBtn.id = "openSpecialBtn";
    openBtn.className = "btn primary";
    openBtn.style.position = "fixed";
    openBtn.style.right = "16px";
    openBtn.style.bottom = "16px";
    openBtn.style.zIndex = "9999";
    openBtn.textContent = "Case Spesial";
    openBtn.disabled = true; // enable after processing files
    document.body.appendChild(openBtn);
  }
})();

// =======================
// Modal Handler wiring
// =======================
const modal = document.getElementById("specialModal");
const openBtn = document.getElementById("openSpecialBtn");
const closeBtn = document.getElementById("closeSpecialModal");
const closeFooterBtn = document.getElementById("closeSpecialModalFooter");

if (openBtn && modal && closeBtn) {
  openBtn.addEventListener("click", () => (modal.style.display = "flex"));
  closeBtn.addEventListener("click", () => (modal.style.display = "none"));
  closeFooterBtn?.addEventListener("click", () => (modal.style.display = "none"));
  modal.addEventListener("click", (e) => {
    if (e.target === modal) modal.style.display = "none";
  });
}

// =======================
// Helper Functions
// =======================
function parseNumber(val) {
  if (val === undefined || val === null || val === "") return 0;
  if (typeof val === "string") val = val.replace(/\s+/g, "").replace(/,/g, "");
  const n = Number(val);
  return isNaN(n) ? 0 : n;
}

function cleanId(val) {
  let id = (val || "").toString().trim();
  id = id.replace(/[\u200B-\u200D\uFEFF]/g, "");
  const selectedPrefix = (document.getElementById("prefixOption")?.value || "").toLowerCase();
  if (selectedPrefix && id.toLowerCase().startsWith(selectedPrefix)) {
    id = id.slice(selectedPrefix.length);
  }
  id = id.replace(/\s+/g, "").toLowerCase();
  id = id.replace(/[^a-z0-9]/g, "");
  return id;
}

function dropTotalish(arr) {
  return arr.filter((r) => {
    let id = (r.ID || "").toString().toLowerCase().replace(/[^a-z0-9]/g, "");
    return !(id.startsWith("total") || ["total", "grandtotal", "subtotal"].includes(id));
  });
}

function getOverrideRate(id, type, defaultRate) {
  const key = `${id}|${type}`;
  return specialCases[key] ?? defaultRate;
}

function getBrandName(prefix) {
  const map = {
    as7: "ASIASLOT777",
    gcr: "GACOR88",
    jt7: "JITU77",
    k88: "HOKIGACOR77",
    ses: "DEWAGACOR77",
    dst: "DEWASLOTO",
    slh: "SLOTTHAILAND",
  };
  return map[prefix?.toLowerCase()] || "NAMA-WEB";
}

function formatDateRange(finishDateStr) {
  try {
    const d = new Date(finishDateStr);
    const start = new Date(d);
    start.setDate(d.getDate() - 6);
    const startDay = start.getDate();
    const endDay = d.getDate();
    const monthName = new Intl.DateTimeFormat("id-ID", { month: "long" }).format(d);
    const year = d.getFullYear();
    return `${startDay}–${endDay} ${monthName} ${year}`;
  } catch {
    return finishDateStr || "";
  }
}

// universal loading control for buttons
function setLoading(buttonId, isLoading, textWhenFalse = null) {
  const btn = document.getElementById(buttonId);
  if (!btn) return;
  if (isLoading) {
    btn.classList.add("btn-loading");
    btn.disabled = true;
  } else {
    btn.classList.remove("btn-loading");
    btn.disabled = false;
    if (textWhenFalse !== null) btn.innerHTML = textWhenFalse;
  }
}

// =======================
// Case Spesial (Override Rate)
// =======================
function renderActiveCases() {
  const wrap = document.getElementById("activeCases");
  if (!wrap) return;

  const entries = Object.entries(specialCases);
  if (!entries.length) {
    wrap.innerHTML = "<span class='muted'>Belum ada case spesial aktif.</span>";
    return;
  }

  // Group by type
  const grouped = entries.reduce((acc, [k, v]) => {
    const [id, type] = k.split("|");
    if (!acc[type]) acc[type] = [];
    acc[type].push({ id, rate: v });
    return acc;
  }, {});

  const typeColor = { SLOT: "#2563eb", CASINO: "#7c3aed", POKER: "#059669" };

  let html = "";
  for (const [type, arr] of Object.entries(grouped)) {
    html += `
      <div style="margin-bottom:12px;">
        <div style="font-weight:600;color:${typeColor[type] || '#374151'};margin-bottom:4px;">
          ${type} — rate ${((arr[0]?.rate || 0) * 100).toFixed(2)}%
        </div>
        <div style="display:flex;flex-wrap:wrap;gap:6px;margin-top:4px;">
          ${arr.map(x => `
            <span style="
              background:#f9fafb;
              border:1px solid #e5e7eb;
              border-radius:6px;
              padding:3px 8px;
              font-size:0.85rem;
              display:flex;
              align-items:center;
              gap:8px;
            ">
              ${x.id}
              <button
                title="Hapus"
                onclick="removeSpecialCase('${x.id}','${type}')"
                style="border:none;background:none;color:#dc2626;cursor:pointer;font-weight:600;font-size:1rem;line-height:1;"
              >×</button>
            </span>
          `).join('')}
        </div>
      </div>
    `;
  }

  wrap.innerHTML = html;
}

function removeSpecialCase(id, type) {
  const key = `${id}|${type}`;
  if (specialCases.hasOwnProperty(key)) {
    delete specialCases[key];
    renderActiveCases();
    renderAll();
  }
}

function addSpecialCasesFromInputs(prefix, type, rate) {
  const ids = [];
  for (let i = 1; i <= 5; i++) {
    const el = document.getElementById(`${prefix}_${i}`);
    if (el && String(el.value || "").trim() !== "") {
      const cleaned = cleanId(el.value);
      if (cleaned) {
        specialCases[`${cleaned}|${type.toUpperCase()}`] = rate;
        ids.push(cleaned);
      }
      el.value = "";
    }
  }
  if (!ids.length) return alert("⚠️ Masukkan minimal satu Account ID!");
  renderActiveCases();
  renderAll();
  alert(`✅ ${ids.length} Account ID berhasil ditambahkan untuk ${type.toUpperCase()} (${(rate * 100).toFixed(1)}%)`);
}

document.getElementById("confirmSlot05")?.addEventListener("click", () => addSpecialCasesFromInputs("slot05", "SLOT", 0.005));
document.getElementById("confirmCasino08")?.addEventListener("click", () => addSpecialCasesFromInputs("casino08", "CASINO", 0.008));
document.getElementById("confirmCasino10")?.addEventListener("click", () => addSpecialCasesFromInputs("casino10", "CASINO", 0.01));

// =======================
// Proses File Input (Main Processing)
// =======================
document.getElementById("processBtn")?.addEventListener("click", async () => {
  if (isProcessing) return;
  const finishDate = (document.getElementById("finishDate")?.value || "").trim();
  if (!finishDate) return alert("⚠️ Pilih tanggal finishing dulu!");

  const btn = document.getElementById("processBtn");
  const originalText = btn?.innerHTML || "Proses";
  try {
    isProcessing = true;
    if (btn) {
      btn.disabled = true;
      btn.innerHTML = `<span style="display:inline-block;width:14px;height:14px;border:2px solid #fff;border-top:2px solid transparent;border-radius:50%;margin-right:6px;vertical-align:middle;animation:spin 0.8s linear infinite;"></span> Memproses...`;
    }

    // reset old data
    allData.length = 0;
    finishing = [];
    finalData = [];

    const files = [
      { id: "slot1", gt: "SLOT" },
      { id: "casino1", gt: "CASINO" },
      { id: "poker1", gt: "POKER" },
    ];

    async function processFiles(fileInput) {
      const input = document.getElementById(fileInput.id);
      if (!input || input.files.length === 0) return;
      const file = input.files[0];

      try {
        if (file.name.toLowerCase().endsWith(".csv")) {
          const text = await file.text();
          const parsed = Papa.parse(text, { header: true, skipEmptyLines: true });
          parsed.data.forEach((row) => {
            const id = cleanId(row["Account ID"] || row["User ID"] || row["ID"] || row["Username"] || row["User Name"] || row["Player"] || "");
            const net_to = parseNumber(row["Net Turnover"] || row["NetTO"] || row["Turnover"] || row["Net To"] || row["Net Turn Over"] || 0);
            const memberWin = parseNumber(row["Member Win"] || row["Win"] || row["MemberWin"] || 0);
            if (id) allData.push({ ID: id, GAME_TYPE: fileInput.gt, NET_TO: net_to, MEMBER_WIN: memberWin });
          });
        } else {
          const data = await file.arrayBuffer();
          const workbook = XLSX.read(data, { type: "array" });
          const sheet = workbook.Sheets[workbook.SheetNames[0]];
          const rows = XLSX.utils.sheet_to_json(sheet, { defval: "", range: 2 });
          rows.forEach((row) => {
            const id = cleanId(row["Account ID"] || row["User ID"] || row["ID"] || row["Username"] || row["User Name"] || row["Player"] || "");
            const net_to = parseNumber(row["Net Turnover"] || row["NetTO"] || row["Turnover"] || row["Net To"] || row["Net Turn Over"] || 0);
            const memberWin = parseNumber(row["Member Win"] || row["Win"] || row["MemberWin"] || 0);
            if (id) allData.push({ ID: id, GAME_TYPE: fileInput.gt, NET_TO: net_to, MEMBER_WIN: memberWin });
          });
        }
      } catch (err) {
        console.error("⚠️ Gagal membaca file:", err);
        alert("⚠️ Gagal membaca file. Pastikan formatnya benar.");
      }
    }

    for (const f of files) await processFiles(f);

    if (!allData.length) {
      alert("❌ Tidak ada data terbaca. Cek file/kepala kolom.");
      return;
    }

    allData = dropTotalish(allData);

    renderAll();

    // enable tiket process button if present
    const pTEl = document.getElementById("processTiket");
    if (pTEl) pTEl.disabled = false;

    if (openBtn) openBtn.disabled = false;

    // Tampilkan Ringkasan Agent setelah proses selesai
    const ringkasanAgent = document.getElementById("ringkasan-agent");
    if (ringkasanAgent) {
      ringkasanAgent.style.display = "block"; // Menampilkan Ringkasan Agent
    }

  } catch (err) {
    console.error(err);
    alert("❌ Terjadi kesalahan saat memproses file:\n" + (err && err.message ? err.message : err));
  } finally {
    isProcessing = false;
    if (btn) {
      btn.disabled = false;
      btn.innerHTML = originalText;
    }
  }
});

// =======================
// Render Tabel & Ringkasan
// =======================
function renderAll() {
  // group by ID|GAME_TYPE
  const grouped = {};
  allData.forEach((r) => {
    const key = `${r.ID}|${r.GAME_TYPE}`;
    if (!grouped[key]) grouped[key] = { NET_TO: 0, MEMBER_WIN: 0 };
    grouped[key].NET_TO += parseNumber(r.NET_TO);
    grouped[key].MEMBER_WIN += parseNumber(r.MEMBER_WIN || 0);
  });

  // pivot to per-ID
  const pivot = {};
  Object.entries(grouped).forEach(([key, val]) => {
    const [id, gt] = key.split("|");
    if (!pivot[id]) pivot[id] = { SLOT: { NET_TO: 0, MEMBER_WIN: 0 }, CASINO: { NET_TO: 0 }, POKER: { NET_TO: 0 } };
    pivot[id][gt].NET_TO = val.NET_TO;
    if (gt === "SLOT") pivot[id][gt].MEMBER_WIN = val.MEMBER_WIN;
  });

  // aggregates
  let totalSlot = 0,
    totalSlotMemberWinFullAccumulate = 0,
    totalSlotMemberWinLossOnly = 0,
    totalCasino = 0,
    totalPoker = 0,
    totalBonus = 0;

  finishing = [];
  let rowIndex = 2;
  Object.entries(pivot).forEach(([id, v]) => {
    const memberWinRaw = v.SLOT.MEMBER_WIN || 0;
    const memberWinLossOnly = memberWinRaw < 0 ? Math.abs(memberWinRaw) : 0;
    const memberWinReal = memberWinRaw;
    const netSlot = v.SLOT.NET_TO || 0;
    const netCasino = v.CASINO.NET_TO || 0;
    const netPoker = v.POKER.NET_TO || 0;
    const slotRate = getOverrideRate(id, "SLOT", ROLL_PCT.SLOT);
    const casinoRate = getOverrideRate(id, "CASINO", ROLL_PCT.CASINO);
    const rollSlotNum = netSlot * slotRate;
    const cbSlotNum = memberWinLossOnly * CB_SLOT_PCT;
    const rollCasinoNum = netCasino * casinoRate;
    const rollPokerNum = netPoker * ROLL_PCT.POKER;
    const totalNum = rollSlotNum + cbSlotNum + rollCasinoNum + rollPokerNum;

    totalSlot += netSlot;
    totalSlotMemberWinFullAccumulate += memberWinReal;
    totalSlotMemberWinLossOnly += memberWinLossOnly;
    totalCasino += netCasino;
    totalPoker += netPoker;
    totalBonus += totalNum;

    finishing.push({
      ID: id,
      "TO SLOT": netSlot,
      memberWin: memberWinLossOnly,
      "Roll 0.3% SLOT": { f: `B${rowIndex}*${slotRate}` },
      "Cb 5% SLOT": { f: `C${rowIndex}*${CB_SLOT_PCT}` },
      "TO CASINO": netCasino,
      "Roll 0.6% CASINO": { f: `F${rowIndex}*${casinoRate}` },
      "TO POKER": netPoker,
      "Roll 0.3% POKER": { f: `H${rowIndex}*${ROLL_PCT.POKER}` },
      UserID: id,
      TOTAL: { f: `D${rowIndex}+E${rowIndex}+G${rowIndex}+I${rowIndex}` },
      slotRate,
      casinoRate,
      _numericTotal: totalNum,
      _cbSlotNum: cbSlotNum,
      RefNo: "", // placeholder; will be synced with finalData if tiket processed
    });

    rowIndex++;
  });

  summaryData = {
    slot: Math.round(totalSlot),
    casino: Math.round(totalCasino),
    poker: Math.round(totalPoker),
    slotMemberWinReal: totalSlotMemberWinFullAccumulate,
  };

  // Render Ringkasan Utama
  const agentStyle = totalSlotMemberWinFullAccumulate < 0 ? "color:#ff4444" : "color:#fff";
  const summaryEl = document.getElementById("summary");
  if (summaryEl) {
    summaryEl.innerHTML = `
      <h3>📊 Ringkasan</h3>
      <ul class="bullets" style="margin:0;padding-left:0;">
        <li>🎰 Net Turnover (SLOT): <b>${totalSlot.toLocaleString()}</b></li>
        <li>🎰 Member Win (SLOT): <b style="${agentStyle}">${totalSlotMemberWinFullAccumulate.toLocaleString()}</b></li>
        <li>🎰 Cashback SLOT (Bersih): <b>${totalSlotMemberWinLossOnly.toLocaleString()}</b></li>
        <li>🃏 Net Turnover (CASINO): <b>${totalCasino.toLocaleString()}</b></li>
        <li>♠️ Net Turnover (POKER): <b>${totalPoker.toLocaleString()}</b></li>
        <li class="bonus">💎 <u>Total Bonus</u>: <b>Rp ${Math.round(totalBonus).toLocaleString()}</b></li>
      </ul>
    `;
  }

  // Render Ringkasan Agent
  const ringkasanAgent = document.getElementById("ringkasan-agent");
  if (ringkasanAgent) {
    ringkasanAgent.innerHTML = `
      <h3>🧾 Ringkasan Agent</h3>
      <ul class="bullets">
        <li>🎰 Net Turn Over (Slot Agent): <input type="text" id="slotAgent" class="numeric-input" value="0"></li>
        <li>🎰 Member Win (Slot Agent): <input type="text" id="memberwinAgent" class="numeric-input" value="0"></li>
        <li>🃏 Net Turn Over (Casino Agent): <input type="text" id="casinoAgent" class="numeric-input" value="0"></li>
        <li>♠️ Net Turn Over (Poker Agent): <input type="text" id="pokerAgent" class="numeric-input" value="0"></li>
      </ul>
    `;
  }

  // Adding event listeners to format number on input change
document.querySelectorAll('.numeric-input').forEach(input => {
  input.addEventListener('input', (e) => {
    e.target.value = formatNumber(e.target.value);
  });
});

// Helper function to format numbers with commas
function formatNumber(value) {
  // Remove any non-numeric characters, except commas and dots
  value = value.replace(/[^\d.,]/g, '');
  
  // Split the number and decimal part
  let [integer, decimal] = value.split('.');
  
  // Format the integer part with commas
  integer = integer.replace(/\B(?=(\d{3})+(?!\d))/g, ',');

  // Join back integer and decimal part (if exists)
  return decimal ? `${integer}.${decimal}` : integer;
}

  // Render Tabel Data
  const cols = [
    "ID", "TO SLOT", "memberWin", "Roll 0.3% SLOT", "Cb 5% SLOT",
    "TO CASINO", "Roll 0.6% CASINO", "TO POKER", "Roll 0.3% POKER", "UserID", "TOTAL", "Ref No"
  ];

  let html = "<table><thead><tr>";
  cols.forEach(c => html += `<th>${c}</th>`);
  html += "</tr></thead><tbody>";

  finishing.forEach((r) => {
    html += "<tr>";
    cols.forEach((c) => {
      let val = r[c];
      if (typeof val === "number") val = val.toLocaleString();
      else if (val && val.f) val = val.f;
      else if (c === "Ref No") val = r.RefNo || "";
      const align = (c === "ID" || c === "UserID" || c === "Ref No") ? "text-align:center" : "text-align:right";
      html += `<td style="${align}">${val ?? ""}</td>`;
    });
    html += "</tr>";
  });

  html += "</tbody></table>";
  const wrap = document.getElementById("table-container");
  if (wrap) wrap.innerHTML = html;

  // Siapkan Data Download
  finalData = finishing.map((r) => ({
    ID: r.ID,
    TO_SLOT: r["TO SLOT"],
    memberWin: r.memberWin,
    RollSlotFormula: r["Roll 0.3% SLOT"].f,
    CbSlotFormula: r["Cb 5% SLOT"].f,
    TO_CASINO: r["TO CASINO"],
    RollCasinoFormula: r["Roll 0.6% CASINO"].f,
    TO_POKER: r["TO POKER"],
    RollPokerFormula: r["Roll 0.3% POKER"].f,
    UserID: r.UserID,
    TOTALFormula: r.TOTAL.f,
    slotRate: r.slotRate,
    casinoRate: r.casinoRate,
    _numericTotal: r._numericTotal,
    _cbSlotNum: r._cbSlotNum,
    RefNo: r.RefNo || "",
  }));

  // Tampilkan tombol download
  const dlBtn = document.getElementById("downloadBtn");
  if (dlBtn) {
    dlBtn.style.display = finalData.length ? "inline-block" : "none";
    dlBtn.onclick = downloadExcel;
  }

  // Render case aktif jika ada
  renderActiveCases();
}

/**
 * Sync RefNo from finalData -> finishing
 * Called after tiket processing so Excel gets updated RefNo
 */
function syncRefNoToFinishing() {
  if (!Array.isArray(finalData) || !Array.isArray(finishing)) return;
  const map = {};
  finalData.forEach(fd => {
    const uid = (fd.ID || fd.UserID || "").toString().toLowerCase();
    if (uid) map[uid] = fd.RefNo || "";
  });
  finishing = finishing.map(f => {
    const key = (f.ID || "").toString().toLowerCase();
    if (map[key]) {
      return { ...f, RefNo: map[key] };
    }
    return f;
  });
}

// =======================
// 📤 DOWNLOAD EXCEL (FINAL FIX)
// =======================
async function downloadExcel() {
  if (!finishing.length) return alert("⚠️ Tidak ada data untuk di-download.");

  const finishDate = document.getElementById("finishDate")?.value || "";
  const prefix = document.getElementById("prefixOption")?.value || "";
  const brandName = getBrandName(prefix);
  const dateRange = formatDateRange(finishDate);

  // ensure finishing has latest RefNo (in case tiket processed)
  syncRefNoToFinishing();

  // =======================
  // 1️⃣ Sortir dari TOTAL tertinggi ke terendah (pakai _numericTotal)
  // =======================
  const sortedData = [...finishing].sort((a, b) => (b._numericTotal || 0) - (a._numericTotal || 0));

  // =======================
  // 2️⃣ Buat workbook & worksheet
  // =======================
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet("Rollingan Result");

  // Judul utama (merge)
  ws.mergeCells("A1:L2");
  const titleCell = ws.getCell("A1");
  titleCell.value = `${brandName} ${dateRange}`;
  titleCell.alignment = { horizontal: "center", vertical: "middle" };
  titleCell.font = { bold: true, size: 14, color: { argb: "FFFFFFFF" } };
  titleCell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF1F3B73" } };

  // =======================
  // 3️⃣ Header kolom
  // =======================
  const headers = [
    "ID",
    "TO SLOT",
    "memberWin",
    "Roll 0.3% SLOT",
    "Cb 5% SLOT",
    "TO CASINO",
    "Roll 0.6% CASINO",
    "TO POKER",
    "Roll 0.3% POKER",
    "UserID",
    "TOTAL",
    "Ref No"
  ];
  ws.addRow(headers);

  // =======================
  // 4️⃣ Isi data
  // =======================
  sortedData.forEach((r) => {
    const row = ws.addRow([
      r.ID,
      Number(r["TO SLOT"]) || 0,
      Number(r.memberWin) || 0,
      "", // Roll slot
      "", // Cashback slot
      Number(r["TO CASINO"]) || 0,
      "", // Roll casino
      Number(r["TO POKER"]) || 0,
      "", // Roll poker
      r.UserID || "",
      "", // Total
      r.RefNo || "",
    ]);

    const rowNum = row.number;

    // Formula otomatis
    row.getCell(4).value = { formula: `=B${rowNum}*${r.slotRate || ROLL_PCT.SLOT}` };
    row.getCell(5).value = { formula: `=C${rowNum}*${CB_SLOT_PCT}` };
    row.getCell(7).value = { formula: `=F${rowNum}*${r.casinoRate || ROLL_PCT.CASINO}` };
    row.getCell(9).value = { formula: `=H${rowNum}*${ROLL_PCT.POKER}` };
    row.getCell(11).value = { formula: `=D${rowNum}+E${rowNum}+G${rowNum}+I${rowNum}` };

    // Format angka
    [2, 3, 6, 8].forEach(i => (row.getCell(i).numFmt = "#,##0"));
    [4, 5, 7, 9, 11].forEach(i => (row.getCell(i).numFmt = "#,##0"));
  });

  // =======================
  // 5️⃣ Styling header
  // =======================
  const headerRow = ws.getRow(3);
  headerRow.eachCell(cell => {
    cell.font = { bold: true, color: { argb: "FFFFFFFF" } };
    cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF1F3B73" } };
    cell.alignment = { horizontal: "center", vertical: "middle" };
    cell.border = {
      top: { style: "thin" },
      left: { style: "thin" },
      bottom: { style: "thin" },
      right: { style: "thin" }
    };
  });

 // =======================
// 6️⃣ Border semua sel
// =======================
const lastRow = ws.lastRow.number;
for (let r = 3; r <= lastRow; r++) {
  for (let c = 1; c <= 12; c++) {
    ws.getCell(r, c).border = {
      top: { style: "thin", color: { argb: "FFCCCCCC" } },
      left: { style: "thin", color: { argb: "FFCCCCCC" } },
      bottom: { style: "thin", color: { argb: "FFCCCCCC" } },
      right: { style: "thin", color: { argb: "FFCCCCCC" } },
    };
  }
}

// =======================
// 7️⃣ Auto width
// =======================
ws.columns.forEach(c => (c.width = 14));

// =======================
// 8️⃣ Tambahkan auto filter
// =======================
ws.autoFilter = { from: "A3", to: "L3" };

// =======================
// 9️⃣ Simpan file
// =======================
const buffer = await wb.xlsx.writeBuffer();
saveAs(new Blob([buffer]), `Rollingan_Result_${brandName}_${finishDate}.xlsx`);
}

// =======================
// Tiket Upload (VLOOKUP Ref No)
// =======================
function normalizeHeaderKey(k) {
  if (!k) return "";
  return String(k)
    .replace(/[\u200B-\u200D\uFEFF]/g, "")
    .replace(/\s+/g, " ")
    .trim()
    .toLowerCase();
}

document.getElementById("tiketFile")?.addEventListener("change", (e) => {
  const file = e.target.files[0];
  if (!file) {
    tiketData = [];
    return;
  }

  const reader = new FileReader();
  reader.onload = (evt) => {
    try {
      const data = new Uint8Array(evt.target.result);
      const workbook = XLSX.read(data, { type: "array" });
      const sheet = workbook.Sheets[workbook.SheetNames[0]];
      const raw = XLSX.utils.sheet_to_json(sheet, { defval: "" });

      if (!raw?.length) {
        tiketData = [];
        return console.warn("⚠️ File tiket kosong.");
      }

      const headerMap = {};
      Object.keys(raw[0]).forEach((k) => (headerMap[k] = normalizeHeaderKey(k)));

      tiketData = raw.map((row) => {
        const normalized = {};
        Object.keys(row).forEach((orig) => {
          const nk = headerMap[orig];
          normalized[nk] = row[orig];
        });
        return normalized;
      });

      console.log("✅ File tiket dibaca:", raw.length, "baris");
      console.log("🔎 Header terdeteksi:", Object.values(headerMap));

      const prefix = document.getElementById("prefixOption")?.value || "";
      const pT = document.getElementById("processTiket");
      if (finalData.length && prefix && pT) pT.disabled = false;
    } catch (err) {
      console.error("❌ Gagal membaca file tiket:", err);
      tiketData = [];
    }
  };
  reader.readAsArrayBuffer(file);
});

// =======================
// 🎫 Proses Tiket (Spinner sama dengan Confirm Data)
// =======================
document.getElementById("processTiket")?.addEventListener("click", async () => {
  const prefix = (document.getElementById("prefixOption")?.value || "").trim();
  if (!prefix) return;
  if (!Array.isArray(tiketData) || tiketData.length === 0) return;
  if (!Array.isArray(finalData) || finalData.length === 0) return;
  if (isProcessing) return;

  isProcessing = true;
  const btn = document.getElementById("processTiket");
  const originalText = btn?.innerHTML || "Proses Tiket";

  // Spinner sama seperti Confirm Data
  if (btn) {
    btn.disabled = true;
    btn.innerHTML = `
      <span class="spinner"></span> Memproses...
    `;
  }

  try {
    await new Promise((res) => setTimeout(res, 200));

    const userNameVariants = [
      "user name",
      "username",
      "user",
      "player",
      "user id",
      "userid",
      "name",
      "full name",
    ];
    const ticketVariants = [
      "ticket number",
      "ticket",
      "ticket no",
      "ticketnumber",
      "ticketno",
      "ticket_no",
      "tiket",
      "nomor tiket",
      "no tiket",
    ];

    const tiketMap = {};
    for (const r of tiketData) {
      let uname = "";
      for (const uv of userNameVariants) {
        if (r[uv] && String(r[uv]).trim() !== "") {
          uname = String(r[uv]).trim();
          break;
        }
      }
      let ticketNo = "";
      for (const tv of ticketVariants) {
        if (r[tv] && String(r[tv]).trim() !== "") {
          ticketNo = String(r[tv]).trim();
          break;
        }
      }
      if (!uname || !ticketNo) continue;
      const normUser = cleanId(uname).toLowerCase();
      if (normUser && !tiketMap[normUser]) tiketMap[normUser] = ticketNo;
    }

    let matchCount = 0;
    finalData = finalData.map((fd) => {
      const uid = cleanId(fd.UserID || fd.ID || "").toLowerCase();
      const found = tiketMap[uid];
      if (!found) return fd;
      const existingRefs = (fd.RefNo || "")
        .split("+")
        .map((x) => x.trim())
        .filter(Boolean);
      if (!existingRefs.includes(found)) {
        existingRefs.push(found);
        matchCount++;
      }
      return { ...fd, RefNo: existingRefs.join("+") };
    });

    syncRefNoToFinishing();

    // Re-render tabel utama
    const wrap = document.getElementById("table-container");
    if (wrap) {
      const cols = [
        "ID",
        "TO SLOT",
        "memberWin",
        "Roll 0.3% SLOT",
        "Cb 5% SLOT",
        "TO CASINO",
        "Roll 0.6% CASINO",
        "TO POKER",
        "Roll 0.3% POKER",
        "UserID",
        "TOTAL",
        "Ref No",
      ];
      const thead =
        "<thead><tr>" + cols.map((c) => `<th>${c}</th>`).join("") + "</tr></thead>";
      const tbody = finishing
        .map((fd) => {
          const cells = cols
            .map((c) => {
              let val = "";
              switch (c) {
                case "ID":
                  val = fd.ID;
                  break;
                case "TO SLOT":
                  val = (fd["TO SLOT"] || 0).toLocaleString();
                  break;
                case "memberWin":
                  val = (fd.memberWin || 0).toLocaleString();
                  break;
                case "Roll 0.3% SLOT":
                  val = fd["Roll 0.3% SLOT"].f || "";
                  break;
                case "Cb 5% SLOT":
                  val = fd["Cb 5% SLOT"].f || "";
                  break;
                case "TO CASINO":
                  val = (fd["TO CASINO"] || 0).toLocaleString();
                  break;
                case "Roll 0.6% CASINO":
                  val = fd["Roll 0.6% CASINO"].f || "";
                  break;
                case "TO POKER":
                  val = (fd["TO POKER"] || 0).toLocaleString();
                  break;
                case "Roll 0.3% POKER":
                  val = fd["Roll 0.3% POKER"].f || "";
                  break;
                case "UserID":
                  val = fd.UserID;
                  break;
                case "TOTAL":
                  val = fd.TOTAL.f || "";
                  break;
                case "Ref No":
                  val = fd.RefNo || "";
                  break;
              }
              const align = ["ID", "UserID", "Ref No"].includes(c)
                ? "center"
                : "right";
              return `<td style="text-align:${align}">${val ?? ""}</td>`;
            })
            .join("");
          return `<tr>${cells}</tr>`;
        })
        .join("");
      wrap.innerHTML = `<table>${thead}<tbody>${tbody}</tbody></table>`;
    }

    console.log(`✅ Proses tiket selesai. ${matchCount} user terhubung.`);
  } catch (err) {
    console.error("❌ Gagal memproses tiket:", err);
  } finally {
    if (btn) {
      btn.disabled = false;
      btn.innerHTML = originalText;
    }
    isProcessing = false;
  }
});

// =======================
// Prefix & Init Setup
// =======================
const prefixSelect = document.getElementById("prefixOption");
if (prefixSelect) {
  prefixSelect.addEventListener("change", () => {
    const prefixVal = prefixSelect.value.trim();
    const tiketBtn = document.getElementById("processTiket");
    if (tiketBtn) tiketBtn.disabled = !prefixVal;
  });
}

document.addEventListener("DOMContentLoaded", () => {
  document.querySelectorAll('input[type="file"]').forEach((input) => {
    input.addEventListener("change", () => {
      input.classList.toggle("filled", input.files.length > 0);
    });
  });

  const pT = document.getElementById("processTiket");
  const dl = document.getElementById("downloadBtn");

  if (pT) pT.disabled = true;
  if (dl) dl.style.display = "none";
});

console.log("✅ script.js final loaded (fix & stabil, spinner tiket seragam).");

// =======================
// Spinner CSS Global
// =======================
const style = document.createElement("style");
style.textContent = `
.spinner {
  border: 2px solid #fff;
  border-top: 2px solid rgba(255,255,255,0.3);
  border-radius: 50%;
  width: 14px;
  height: 14px;
  display: inline-block;
  margin-right: 6px;
  animation: spin 0.8s linear infinite;
  vertical-align: middle;
}
@keyframes spin {
  0% { transform: rotate(0deg); }
  100% { transform: rotate(360deg); }
}`;
document.head.appendChild(style);
