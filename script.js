// Atur worker pdf.js (wajib)
pdfjsLib.GlobalWorkerOptions.workerSrc =
  "https://cdn.jsdelivr.net/npm/pdfjs-dist@3.11.174/build/pdf.worker.min.js";

const fileInput = document.getElementById("pdfFile");
const uploadLabel = document.getElementById("uploadLabel");
const fileNameLabel = document.getElementById("fileName");
const convertBtn = document.getElementById("convertBtn");
const statusEl = document.getElementById("status");
const downloadBox = document.getElementById("downloadBox");
const downloadLink = document.getElementById("downloadLink");

// Klik area upload untuk trigger file input
uploadLabel.addEventListener("click", () => {
  fileInput.click();
});

fileInput.addEventListener("change", () => {
  const file = fileInput.files[0];
  if (file) {
    fileNameLabel.textContent = file.name;
    convertBtn.disabled = false;
    downloadBox.style.display = "none";
    statusEl.textContent = "";
    statusEl.classList.remove("error");
  } else {
    fileNameLabel.textContent = "Belum ada file dipilih";
    convertBtn.disabled = true;
  }
});

convertBtn.addEventListener("click", () => {
  const file = fileInput.files[0];
  if (!file) {
    setStatus("Silakan pilih file PDF dulu.", true);
    return;
  }

  convertBtn.disabled = true;
  setStatus("Membaca PDF, mohon tunggu...", false);
  downloadBox.style.display = "none";

  const reader = new FileReader();

  reader.onload = async (event) => {
    try {
      const arrayBuffer = event.target.result;
      await convertPdfArrayBufferToExcel(arrayBuffer, file.name);
      convertBtn.disabled = false;
    } catch (err) {
      console.error(err);
      setStatus("Terjadi kesalahan saat konversi. Coba dengan PDF lain atau cek konsol.", true);
      convertBtn.disabled = false;
    }
  };

  reader.onerror = () => {
    setStatus("Gagal membaca file PDF.", true);
    convertBtn.disabled = false;
  };

  reader.readAsArrayBuffer(file);
});

/**
 * Update status teks
 * @param {string} message
 * @param {boolean} isError
 */
function setStatus(message, isError) {
  statusEl.textContent = message || "";
  statusEl.classList.toggle("error", !!isError);
}

/**
 * Konversi PDF (ArrayBuffer) menjadi Excel dan sediakan link download
 * @param {ArrayBuffer} arrayBuffer
 * @param {string} originalFileName
 */
async function convertPdfArrayBufferToExcel(arrayBuffer, originalFileName) {
  const loadingTask = pdfjsLib.getDocument({ data: arrayBuffer });
  const pdf = await loadingTask.promise;

  const totalPages = pdf.numPages;
  setStatus(`PDF terbaca (${totalPages} halaman). Mengambil data...`, false);

  const allRows = [];

  for (let pageNum = 1; pageNum <= totalPages; pageNum++) {
    setStatus(`Memproses halaman ${pageNum} dari ${totalPages}...`, false);

    const page = await pdf.getPage(pageNum);
    const textContent = await page.getTextContent();

    const pageRows = groupTextItemsToRows(textContent.items);
    if (pageRows.length > 0) {
      // Tambahkan pemisah halaman (opsional)
      if (pageNum > 1) {
        allRows.push([]);
      }
      allRows.push(...pageRows);
    }
  }

  if (allRows.length === 0) {
    setStatus("Tidak ditemukan teks yang bisa diolah dari PDF ini.", true);
    return;
  }

  setStatus("Menyusun file Excel...", false);

  // Buat workbook & sheet
  const workbook = XLSX.utils.book_new();
  const worksheet = XLSX.utils.aoa_to_sheet(allRows);
  XLSX.utils.book_append_sheet(workbook, worksheet, "Data");

  // Export jadi ArrayBuffer
  const wbArray = XLSX.write(workbook, {
    bookType: "xlsx",
    type: "array",
  });

  const blob = new Blob([wbArray], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });

  // Buat link download
  const fileBaseName = originalFileName.replace(/\.pdf$/i, "");
  const downloadFileName = fileBaseName + ".xlsx";

  const url = URL.createObjectURL(blob);
  downloadLink.href = url;
  downloadLink.download = downloadFileName;
  downloadBox.style.display = "block";

  setStatus("Selesai. Silakan download file Excel-nya.", false);
}

/**
 * Mengelompokkan text items pdf.js menjadi baris dan kolom sederhana.
 * Cocok untuk PDF yang asalnya dari tabel/Excel.
 *
 * @param {Array} items - textContent.items dari pdf.js
 * @returns {string[][]} rows
 */
function groupTextItemsToRows(items) {
  // Peta: keyY → array of { x, text }
  const rowsMap = new Map();

  const Y_TOLERANCE = 4; // semakin besar = baris lebih "longgar"

  for (const item of items) {
    const transform = item.transform;
    const x = transform[4];
    const y = transform[5];
    const text = (item.str || "").trim();

    if (!text) continue;

    // Cari keyY yang dekat (beda y < tolerance)
    let targetKey = null;
    for (const key of rowsMap.keys()) {
      if (Math.abs(key - y) <= Y_TOLERANCE) {
        targetKey = key;
        break;
      }
    }
    if (targetKey === null) {
      targetKey = y;
      rowsMap.set(targetKey, []);
    }

    rowsMap.get(targetKey).push({ x, text });
  }

  // Ubah map → array, urutkan baris dari atas ke bawah (y besar ke kecil)
  const rows = [...rowsMap.entries()]
    .sort((a, b) => b[0] - a[0])
    .map(([_, cells]) => {
      // urutkan sel kiri ke kanan
      const sortedCells = cells.sort((c1, c2) => c1.x - c2.x);
      return sortedCells.map((c) => c.text);
    });

  return rows;
}
