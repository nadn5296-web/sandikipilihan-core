let referralData = [];
let filteredData = [];

// === Utility: normalisasi key header (lowercase + hapus spasi/underscore) ===
function normalizeKey(key) {
  return key
    .toString()
    .trim()
    .toLowerCase()
    .replace(/[_\s\.]+/g, "")
    .replace(/[^a-z0-9]/g, "");
}

// === Upload file Excel ===
function handleFileUpload(inputId) {
  const fileInput = document.getElementById(inputId);
  const file = fileInput.files[0];
  if (!file) {
    alert(`Silakan pilih file untuk ${inputId.toUpperCase()} terlebih dahulu!`);
    return;
  }

  const reader = new FileReader();
  reader.onload = (e) => {
    const data = new Uint8Array(e.target.result);
    const workbook = XLSX.read(data, { type: "array" });
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];

    // mulai baca dari baris ke-2
    const json = XLSX.utils.sheet_to_json(sheet, { range: 1, defval: "" });

    // Normalisasi header agar tidak tergantung nama kolom
    const newData = json
      .map(row => {
        const normalized = {};
        for (const k in row) {
          normalized[normalizeKey(k)] = row[k];
        }

        return {
          brand: normalized["brand"] || "-",
          username:
            normalized["username"] ||
            normalized["user"] ||
            normalized["playername"] ||
            "-",
          referralCode:
            (normalized["referralcode"] ||
              normalized["referral"] ||
              normalized["ownreferralcode"] ||
              "")
              .toString()
              .trim(),
          totalDeposit:
            parseFloat(normalized["totaldeposit"] || normalized["deposit"] || 0) || 0,
          totalWithdrawal:
            parseFloat(
              normalized["totalwithdrawal"] || normalized["withdrawal"] || 0
            ) || 0,
        };
      })
      .filter(r => r.referralCode && r.referralCode !== ""); // Menambahkan filter validasi referral code yang kosong

    if (newData.length === 0) {
      alert(`File ${file.name} tidak memiliki data Referral Code valid.`);
      return;
    }

    referralData = [...referralData, ...newData];
    filteredData = [...referralData];
    updateBrandFilterOptions();
    renderTable(filteredData);
    updateStats(filteredData);
  };

  reader.readAsArrayBuffer(file);
}

// === Render tabel data ===
function renderTable(data) {
  const tbody = document.getElementById("tableBody");
  tbody.innerHTML = "";

  if (!data || data.length === 0) {
    tbody.innerHTML = `<tr><td colspan="7" class="no-data">Tidak ada data ditemukan</td></tr>`;
    return;
  }

  data.forEach((item, idx) => {
    const tr = document.createElement("tr");

    tr.innerHTML = `
      <td>${idx + 1}</td>
      <td>${item.brand}</td>
      <td>${item.username}</td>
      <td>${item.referralCode}</td>
      <td>${item.totalDeposit > 0 ? "✅ Ya" : "❌ Tidak"}</td>
      <td>${item.totalDeposit.toLocaleString("id-ID")}</td>
      <td>${item.totalWithdrawal.toLocaleString("id-ID")}</td>
    `;
    tbody.appendChild(tr);
  });
}

// === Statistik dashboard ===
function updateStats(data) {
  const totalReferral = data.length;
  const totalDeposit = data.reduce((s, r) => s + (r.totalDeposit || 0), 0);
  const totalWithdrawal = data.reduce((s, r) => s + (r.totalWithdrawal || 0), 0);
  const belumDeposit = data.filter(r => (r.totalDeposit || 0) === 0).length;
  const sudahDeposit = totalReferral - belumDeposit;

  document.getElementById("totalReferral").textContent = totalReferral;
  document.getElementById("totalDeposit").textContent =
    totalDeposit.toLocaleString("id-ID");
  document.getElementById("totalWithdrawal").textContent =
    totalWithdrawal.toLocaleString("id-ID");
  document.getElementById("totalBelumDeposit").textContent = belumDeposit;

  // Menambahkan tanggal berdasarkan range
  const elDateRange = document.getElementById("totalDateRange");
  if (elDateRange) {
    // Cek apakah ada data
    if (data.length > 0) {
      const earliestDate = new Date(Math.min(...data.map(item => new Date(item.registrationDate))));
      const latestDate = new Date(Math.max(...data.map(item => new Date(item.registrationDate))));

      const formattedStartDate = formatDate(earliestDate);
      const formattedEndDate = formatDate(latestDate);

      elDateRange.textContent = `${formattedStartDate} - ${formattedEndDate}`;
    } else {
      elDateRange.textContent = '-'; // Tampilkan tanda "-" jika tidak ada data
    }
  }

  const elSudah = document.getElementById("totalSudahDeposit");
  if (elSudah) elSudah.textContent = sudahDeposit;
}

// === Reset semua data ===
function resetData() {
  if (!confirm("Apakah Anda yakin ingin menghapus semua data upload?")) return;
  referralData = [];
  filteredData = [];
  renderTable([]); 
  updateStats([]); 
  document.querySelectorAll('input[type="file"]').forEach(i => (i.value = "")); 
  document.getElementById("filterBrand").innerHTML = 
    '<option value="">Pilih Brand</option>'; 
  document.getElementById("filterReferral").value = ""; 
}

// === Isi dropdown Brand ===
function updateBrandFilterOptions() {
  const sel = document.getElementById("filterBrand");
  const brands = [...new Set(referralData.map(d => d.brand || "-"))].filter(
    b => b && b !== "-"
  );
  sel.innerHTML = '<option value="">Pilih Brand</option>';
  brands.sort().forEach(b => {
    const opt = document.createElement("option");
    opt.value = b;
    opt.textContent = b;
    sel.appendChild(opt);
  });
}

// === Filter data ===
function applyFilters() {
  const brand = (document.getElementById("filterBrand").value || "")
    .toString()
    .toLowerCase();
  const referral = (document.getElementById("filterReferral").value || "")
    .toString()
    .toLowerCase();

  filteredData = referralData.filter(item => {
    const brandMatch = brand
      ? (item.brand || "").toLowerCase().includes(brand)
      : true;
    const referralMatch = referral
      ? (item.referralCode || "").toLowerCase().includes(referral)
      : true;
    return brandMatch && referralMatch;
  });

  renderTable(filteredData);
  updateStats(filteredData);
}

// === Listener untuk filter ===
document.getElementById("filterBrand").addEventListener("change", applyFilters);
document.getElementById("filterReferral").addEventListener("input", applyFilters);

// === Export ke Excel ===
function exportFilteredData() {
  if (!filteredData || filteredData.length === 0) {
    alert("Tidak ada data yang bisa diexport!");
    return;
  }
  const ws = XLSX.utils.json_to_sheet(filteredData);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Filtered");
  XLSX.writeFile(
    wb,
    `Referral_Filtered_${new Date().toISOString().slice(0, 10)}.xlsx`
  );
}
