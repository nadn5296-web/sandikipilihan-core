// ======================================================
//  TRANSACTION PIVOT DP / WD & EVENT TANPA UNDI - FINAL REVISED VERSION
// ======================================================

// ----- Helpers -----
function formatNumber(v) {
  return typeof v === "number" ? v.toLocaleString("id-ID") : v;
}

function safeGetFiles(id) {
  const el = document.getElementById(id);
  return el ? el.files[0] : null;
}

// Global state
let prefixCasinoDone = false;
let prefixSlotDone = false;

// ======================================================
//  Generate Pivot
// ======================================================
document.getElementById("generateBtn").addEventListener("click", async () => {
  const depositFile = safeGetFiles("depositFile");
  const withdrawFile = safeGetFiles("withdrawFile");
  const adjustmentFile = safeGetFiles("adjustFile");
  const casinoFile = safeGetFiles("casinoFile");
  const slotFile = safeGetFiles("slotFile");

  if (!depositFile && !withdrawFile && !adjustmentFile && !casinoFile && !slotFile) {
    alert("Silakan upload minimal satu file!");
    return;
  }

  const [depositData, depositSummary] = depositFile
    ? await processDepositWithdraw(depositFile)
    : [[], null];
  const [withdrawData, withdrawSummary] = withdrawFile
    ? await processDepositWithdraw(withdrawFile)
    : [[], null];
  const [adjustmentData, adjustmentSummary] = adjustmentFile
    ? await processAdjustment(adjustmentFile)
    : [[], null];
  const [casinoData, casinoSummary] = casinoFile
    ? await processWinLose(casinoFile)
    : [[], null];
  const [slotData, slotSummary] = slotFile
    ? await processWinLose(slotFile)
    : [[], null];

  // Simpan ke global
  window._pivotData = {
    depositData,
    withdrawData,
    adjustmentData,
    casinoData,
    slotData,
  };

  // ======================================================
  //  Render Ringkasan
  // ======================================================
  const makeDataSection = (key, summary) => {
    if (!summary) return `<div class="text-muted">Tidak ada data.</div>`;

    if (key === "deposit" || key === "withdraw") {
      return `
        <div>Total Username: <span class="summary-value">${summary.totalUsername.toLocaleString()}</span></div>
        <div>Total Nominal: <span class="summary-value">Rp ${summary.totalNominal.toLocaleString("id-ID")}</span></div>
        <div>Total Transaksi: <span class="summary-value">${summary.totalTransaksi.toLocaleString()}</span></div>
        <div class="mt-2">
          <button class="btn btn-outline-primary btn-sm" onclick="toggleData('${key}')">Tampilkan Data</button>
          <button class="btn btn-outline-success btn-sm" onclick="downloadExcelTable('${key}')">Download Excel</button>
        </div>
        <div id="${key}Table" class="mt-2" style="display:none;"></div>
      `;
    } else if (key === "adjustment") {
      return `
        <div>Total Username: <span class="summary-value">${summary.totalUsername.toLocaleString()}</span></div>
        <div>Total Nominal: <span class="summary-value">Rp ${summary.totalNominal.toLocaleString("id-ID")}</span></div>
      `;
    } else if (key === "casino" || key === "slot") {
      return `
        <div>Total Account: <span class="summary-value">${summary.totalAccount.toLocaleString()}</span></div>
        <div>Total Net Turnover: <span class="summary-value">Rp ${summary.totalTurnover.toLocaleString("id-ID")}</span></div>
        <div>Total Member Win: <span class="summary-value">Rp ${summary.totalWin.toLocaleString("id-ID")}</span></div>
      `;
    }
  };

  document.getElementById("summaryContainer").innerHTML = `
    <div class="summary-title">📊 Ringkasan Data</div>
    <div>Brand: <span class="summary-value">${
      depositSummary?.brand ||
      withdrawSummary?.brand ||
      adjustmentSummary?.brand ||
      casinoSummary?.brand ||
      slotSummary?.brand ||
      "-"
    }</span></div>
    <hr>
    <div><strong>Deposit</strong></div>
    ${makeDataSection("deposit", depositSummary)}
    <hr>
    <div><strong>Withdraw</strong></div>
    ${makeDataSection("withdraw", withdrawSummary)}
    <hr>
    <div><strong>Adjustment</strong></div>
    ${makeDataSection("adjustment", adjustmentSummary)}
    <hr>
    <div><strong>Casino</strong></div>
    ${makeDataSection("casino", casinoSummary)}
    <hr>
    <div><strong>Slot</strong></div>
    ${makeDataSection("slot", slotSummary)}
  `;
});

// ======================================================
//  Tampilkan Data per Section
// ======================================================
function toggleData(type) {
  const dataMap = {
    deposit: window._pivotData.depositData,
    withdraw: window._pivotData.withdrawData,
  };
  const data = dataMap[type] || [];
  const container = document.getElementById(`${type}Table`);
  if (!container) return;

  if (container.style.display === "none") {
    if (!data.length) {
      container.innerHTML = "<p class='text-muted'>Tidak ada data.</p>";
    } else {
      let html = `<table class="table table-bordered table-striped mt-2"><thead class="table-dark"><tr>`;
      const keys = Object.keys(data[0]);
      keys.forEach((k) => (html += `<th>${k}</th>`));
      html += `</tr></thead><tbody>`;
      data.forEach((r) => {
        html += "<tr>";
        keys.forEach((k) => (html += `<td>${formatNumber(r[k])}</td>`));
        html += "</tr>";
      });
      html += "</tbody></table>";
      container.innerHTML = html;
    }
    container.style.display = "block";
  } else {
    container.style.display = "none";
  }
}

// ======================================================
//  Download Excel per Section
// ======================================================
function downloadExcelTable(type) {
  const dataMap = {
    deposit: window._pivotData.depositData,
    withdraw: window._pivotData.withdrawData,
  };
  const data = dataMap[type] || [];
  if (!data.length) return alert(`Tidak ada data ${type}.`);

  const sorted = [...data].sort((a, b) => (b.TotalAmount || 0) - (a.TotalAmount || 0));

  const ws = XLSX.utils.json_to_sheet(sorted);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, type.toUpperCase());
  XLSX.writeFile(wb, `${type}_data_${new Date().toISOString().slice(0, 10)}.xlsx`);
}

// ======================================================
//  Prefix Buttons Manual (Brand Dropdown)
// ======================================================
document.getElementById("prefixCasinoBtn").addEventListener("click", () => {
  const brand = document.getElementById("brandSelector").value;
  if (!brand) return alert("Pilih Brand dulu!");
  if (!window._pivotData?.casinoData?.length)
    return alert("Upload data Casino dulu!");

  prefixProcess("casino", brand);
  prefixCasinoDone = true;
  document.getElementById("casinoPrefixStatus").textContent = `✅ Prefix Casino (${brand}) selesai`;
});

document.getElementById("prefixSlotBtn").addEventListener("click", () => {
  const brand = document.getElementById("brandSelector").value;
  if (!brand) return alert("Pilih Brand dulu!");
  if (!window._pivotData?.slotData?.length)
    return alert("Upload data Slot dulu!");

  prefixProcess("slot", brand);
  prefixSlotDone = true;
  document.getElementById("slotPrefixStatus").textContent = `✅ Prefix Slot (${brand}) selesai`;
});

function prefixProcess(type, brand) {
  let data = type === "casino" ? window._pivotData.casinoData : window._pivotData.slotData;

  data.forEach(row => {
    if (row.AccountID?.startsWith(brand)) {
      row.AccountID = row.AccountID.slice(brand.length);
    }
  });

  alert(`Prefix ${brand} untuk ${type.toUpperCase()} berhasil dihapus!`);
}

// ======================================================
//  Generate Finishing Pivot
// ======================================================
document.getElementById("generateFinalBtn").addEventListener("click", () => {
  if (!prefixCasinoDone || !prefixSlotDone)
    return alert("Harap lakukan Prefix Casino & Slot terlebih dahulu!");
  if (!window._pivotData)
    return alert("Silakan generate pivot utama terlebih dahulu!");

  const {
    depositData = [],
    withdrawData = [],
    adjustmentData = [],
    casinoData = [],
    slotData = [],
  } = window._pivotData;

  const allUsers = new Set();
  [depositData, withdrawData, adjustmentData, casinoData, slotData].forEach(arr => {
    if (arr?.length) arr.forEach(r => allUsers.add(r.UserName || r.AccountID));
  });

  const finishing = Array.from(allUsers).map((user) => {
    const deposit = depositData.find((x) => x.UserName === user)?.TotalAmount || 0;
    const withdraw = withdrawData.find((x) => x.UserName === user)?.TotalAmount || 0;
    const adjustment = adjustmentData.find((x) => x.UserName === user)?.TotalAmount || 0;
    const casino = casinoData.find((x) => x.AccountID === user) || {};
    const slot = slotData.find((x) => x.AccountID === user) || {};

    return {
      UserName: user,
      Deposit: deposit,
      Withdraw: withdraw,
      Adjustment: adjustment,
      "Net Turnover Casino": casino.NetTurnover || 0,
      "Member Win Casino": casino.MemberWin || 0,
      "Net Turnover Slot": slot.NetTurnover || 0,
      "Member Win Slot": slot.MemberWin || 0,
    };
  });

  window._finishingPivot = finishing;
console.log(`✅ Finishing Pivot berhasil dibuat: ${finishing.length} baris`);
document.getElementById("finalPivotSection").classList.add("visible");
renderFinishingPivot(finishing);
});

function renderFinishingPivot(data) {
  const section = document.getElementById("finalPivotSection");
  const container = document.getElementById("finalPivotTable");
  section.classList.remove("hidden");

  if (!data?.length) {
    container.innerHTML = "<p class='text-muted'>Tidak ada data finishing.</p>";
    return;
  }

  // Urutkan berdasarkan Deposit
  data.sort((a, b) => b.Deposit - a.Deposit);
  window._finishingPivotSorted = data;

  // Tabel hitam seperti withdraw
  container.innerHTML = `
    <div id="finishingTableWrapper" style="display:none;" class="mt-3">
      <table class="table table-bordered table-striped mt-2">
        <thead class="table-dark">
          <tr>
            <th>UserName</th>
            <th>TotalAmount (Deposit)</th>
            <th>TotalAmount (Withdraw)</th>
            <th>Adjustment</th>
            <th>Net Turnover Casino</th>
            <th>Member Win Casino</th>
            <th>Net Turnover Slot</th>
            <th>Member Win Slot</th>
          </tr>
        </thead>
        <tbody>
          ${data
            .map(
              (r) => `
              <tr>
                <td>${r.UserName}</td>
                <td>${formatNumber(r.Deposit)}</td>
                <td>${formatNumber(r.Withdraw)}</td>
                <td>${formatNumber(r.Adjustment)}</td>
                <td>${formatNumber(r["Net Turnover Casino"])}</td>
                <td>${formatNumber(r["Member Win Casino"])}</td>
                <td>${formatNumber(r["Net Turnover Slot"])}</td>
                <td>${formatNumber(r["Member Win Slot"])}</td>
              </tr>`
            )
            .join("")}
        </tbody>
      </table>
    </div>
  `;

  // Tombol utama saja di atas tabel
  const toggleBtn = document.getElementById("toggleFinalBtn");
  const downloadBtn = document.getElementById("downloadFinalBtn");
  const wrapper = document.getElementById("finishingTableWrapper");

  // Tombol toggle tampil/sembunyi
  toggleBtn.onclick = () => {
    if (wrapper.style.display === "none") {
      wrapper.style.display = "block";
      toggleBtn.textContent = "Sembunyikan Data";
    } else {
      wrapper.style.display = "none";
      toggleBtn.textContent = "Tampilkan Data";
    }
  };

  // Tombol download Excel
  downloadBtn.onclick = () => {
    const sorted = window._finishingPivotSorted || data;
    const ws = XLSX.utils.json_to_sheet(sorted);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Finishing Pivot");
    XLSX.writeFile(wb, `Finishing_Pivot_${new Date().toISOString().slice(0,10)}.xlsx`);
  };
}

// ======================================================
//  Data Processors
// ======================================================
async function processDepositWithdraw(file) {
  return new Promise((resolve) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      const wb = XLSX.read(e.target.result, { type: "array" });
      const sheet = wb.Sheets[wb.SheetNames[0]];
      const data = XLSX.utils.sheet_to_json(sheet, { range: 1 });
      const parsed = data
        .map((r) => ({
          Brand: r["Brand"] || "",
          UserName: (r["User Name"] || r["Username"] || "").toString().trim(),
          Amount: parseFloat(r["Deposit Amount"] ?? r["Withdrawal Amount"] ?? 0) || 0,
        }))
        .filter((r) => r.UserName);

      const pivot = {};
      parsed.forEach((r) => {
        if (!pivot[r.UserName]) pivot[r.UserName] = { total: 0, transaksi: 0 };
        pivot[r.UserName].total += r.Amount;
        pivot[r.UserName].transaksi++;
      });

      const result = Object.entries(pivot).map(([u, v]) => ({
        UserName: u,
        TotalAmount: v.total,
        TotalTransaksi: v.transaksi,
      }));

      resolve([
        result,
        {
          brand: [...new Set(parsed.map((r) => r.Brand))].join(", ") || "-",
          totalUsername: result.length,
          totalNominal: result.reduce((s, x) => s + x.TotalAmount, 0),
          totalTransaksi: result.reduce((s, x) => s + x.TotalTransaksi, 0),
        },
      ]);
    };
    reader.readAsArrayBuffer(file);
  });
}

async function processAdjustment(file) {
  return new Promise((resolve) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      const wb = XLSX.read(e.target.result, { type: "array" });
      const sheet = wb.Sheets[wb.SheetNames[0]];
      const data = XLSX.utils.sheet_to_json(sheet, { range: 1 });

      const parsed = data
        .map((r) => {
          const rawUser = (r["User Name"] || r["Username"] || "").toString().trim();
          let rawAmt = r["Adjustment Amt."] ?? r["Adjustment Amount"] ?? 0;

          if (typeof rawAmt === "string") {
            // Hilangkan semua karakter kecuali angka, koma, titik
            rawAmt = rawAmt.replace(/[^\d.,-]/g, "").trim();

            // Jika format seperti 1.000,50 → ubah ke 1000.50
            if (rawAmt.includes(",") && rawAmt.includes(".")) {
              if (rawAmt.lastIndexOf(",") > rawAmt.lastIndexOf(".")) {
                // format Eropa: 1.000,50
                rawAmt = rawAmt.replace(/\./g, "").replace(",", ".");
              } else {
                // format biasa: 1,000.50
                rawAmt = rawAmt.replace(/,/g, "");
              }
            } else if (rawAmt.includes(".")) {
              // Jika hanya titik dan panjang digit setelah titik 3 → anggap pemisah ribuan
              const parts = rawAmt.split(".");
              if (parts.length > 1 && parts[1].length === 3) {
                rawAmt = rawAmt.replace(/\./g, "");
              }
            } else if (rawAmt.includes(",")) {
              // Format seperti 1.000 atau 2,145 → ubah koma jadi titik
              rawAmt = rawAmt.replace(/\./g, "").replace(",", ".");
            }
          }

          const amount = parseFloat(rawAmt) || 0;
          return { UserName: rawUser, Amount: amount };
        })
        .filter((r) => r.UserName && !isNaN(r.Amount));

      const pivot = {};
      parsed.forEach((r) => {
        if (!pivot[r.UserName]) pivot[r.UserName] = 0;
        pivot[r.UserName] += r.Amount;
      });

      const result = Object.entries(pivot).map(([u, v]) => ({
        UserName: u,
        TotalAmount: v,
      }));

      resolve([
        result,
        {
          totalUsername: result.length,
          totalNominal: result.reduce((s, x) => s + x.TotalAmount, 0),
        },
      ]);
    };
    reader.readAsArrayBuffer(file);
  });
}

async function processWinLose(file) {
  return new Promise((resolve) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      const wb = XLSX.read(e.target.result, { type: "array" });
      const sheet = wb.Sheets[wb.SheetNames[0]];
      const data = XLSX.utils.sheet_to_json(sheet, { range: 2 });
      const parsed = data
        .map((r) => ({
          Brand: (r["Account ID"] || "").toString().slice(0, 3),
          AccountID: (r["Account ID"] || "").toString(),
          NetTurnover: parseFloat(r["Net Turnover"] ?? 0) || 0,
          MemberWin: parseFloat(r["Member Win"] ?? 0) || 0,
        }))
        .filter((r) => r.AccountID);

      const pivot = {};
      parsed.forEach((r) => {
        if (!pivot[r.AccountID]) pivot[r.AccountID] = { turnover: 0, win: 0 };
        pivot[r.AccountID].turnover += r.NetTurnover;
        pivot[r.AccountID].win += r.MemberWin;
      });

      const result = Object.entries(pivot).map(([id, v]) => ({
        AccountID: id,
        NetTurnover: v.turnover,
        MemberWin: v.win,
      }));

      resolve([
        result,
        {
          brand: [...new Set(parsed.map((r) => r.Brand))].join(", ") || "-",
          totalAccount: result.length,
          totalTurnover: result.reduce((s, x) => s + x.NetTurnover, 0),
          totalWin: result.reduce((s, x) => s + x.MemberWin, 0),
        },
      ]);
    };
    reader.readAsArrayBuffer(file);
  });
}
// ===============================================
// 🎛️ Toggle tombol tampilkan / sembunyikan tabel
// ===============================================
const toggleFinalBtn = document.getElementById("toggleFinalBtn");
const finalPivotTable = document.getElementById("finalPivotTable");

toggleFinalBtn.addEventListener("click", () => {
  // Jika tabel sedang disembunyikan
  if (finalPivotTable.classList.contains("hidden")) {
    finalPivotTable.classList.remove("hidden");
    finalPivotTable.classList.add("visible");
    toggleFinalBtn.textContent = "Sembunyikan Data";
  } 
  // Jika tabel sedang tampil
  else {
    finalPivotTable.classList.remove("visible");
    finalPivotTable.classList.add("hidden");
    toggleFinalBtn.textContent = "Tampilkan Data";
  }
});
