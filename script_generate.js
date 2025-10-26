document.addEventListener("DOMContentLoaded", () => {
  const dropArea = document.getElementById("drop-area-generate");
  const fileInput = document.getElementById("file-input-generate");
  const fileListDiv = document.getElementById("file-list-generate");
  const processButton = document.getElementById("process-button");
  const dataOutputDiv = document.getElementById("data-output");
  const dataTableBody = document.querySelector("#data-table tbody");
  const errorMessageDiv = document.getElementById("error-message-generate");
  const downloadExcelButton = document.getElementById("download-excel-button");

  const uploadedFiles = { program: null, kegiatan: null, subkegiatan: null };

  function displayError(msg) {
    errorMessageDiv.textContent = msg;
    errorMessageDiv.style.display = "block";
  }
  function clearError() {
    errorMessageDiv.textContent = "";
    errorMessageDiv.style.display = "none";
  }
  function updateProcessButtonState() {
    processButton.disabled = !(
      uploadedFiles.program &&
      uploadedFiles.kegiatan &&
      uploadedFiles.subkegiatan
    );
  }
  function renderFileList() {
    fileListDiv.innerHTML = "";
    const filesPresent = Object.values(uploadedFiles).filter((f) => f);
    if (filesPresent.length === 0) {
      fileListDiv.style.display = "none";
      return;
    }
    fileListDiv.style.display = "block";
    const fileMap = {
      "data_program.json": "Data Program",
      "data_kegiatan.json": "Data Kegiatan",
      "data_sub_kegiatan.json": "Data Sub Kegiatan",
    };
    for (const key in uploadedFiles) {
      const file = uploadedFiles[key];
      if (file) {
        const p = document.createElement("p");
        p.textContent = `${fileMap[file.name] || file.name} (${file.name})`;
        const rm = document.createElement("span");
        rm.textContent = "✅";
        rm.className = "remove-file";
        rm.onclick = () => removeFile(key);
        p.appendChild(rm);
        fileListDiv.appendChild(p);
      }
    }
    updateProcessButtonState();
  }
  function removeFile(k) {
    uploadedFiles[k] = null;
    renderFileList();
    clearOutput();
  }
  function clearOutput() {
    dataTableBody.innerHTML = "";
    dataOutputDiv.style.display = "none";
    downloadExcelButton.style.display = "none";
    clearError();
  }
  function readFileAsText(f) {
    return new Promise((res, rej) => {
      const r = new FileReader();
      r.onload = () => res(r.result);
      r.onerror = () => rej(r.error);
      r.readAsText(f);
    });
  }
  function formatRupiahForDisplay(v) {
    if (typeof v !== "number") v = parseFloat(v);
    if (isNaN(v) || v === 0) return "-";
    return new Intl.NumberFormat("id-ID", {
      style: "currency",
      currency: "IDR",
      minimumFractionDigits: 0,
    }).format(v);
  }
  function capitalizeWords(s) {
    if (!s) return "";
    return s.toLowerCase().replace(/(^|\s)\S/g, (a) => a.toUpperCase());
  }

  // drag-drop
  ["dragenter", "dragover", "dragleave", "drop"].forEach((ev) =>
    dropArea.addEventListener(ev, preventDefaults, false)
  );
  function preventDefaults(e) {
    e.preventDefault();
    e.stopPropagation();
  }
  ["dragenter", "dragover"].forEach((ev) =>
    dropArea.addEventListener(
      ev,
      () => dropArea.classList.add("highlight"),
      false
    )
  );
  ["dragleave", "drop"].forEach((ev) =>
    dropArea.addEventListener(
      ev,
      () => dropArea.classList.remove("highlight"),
      false
    )
  );
  dropArea.addEventListener("drop", handleDrop, false);
  function handleDrop(e) {
    addFiles(e.dataTransfer.files);
  }
  fileInput.addEventListener("change", (e) => {
    addFiles(e.target.files);
    e.target.value = null;
  });

  function addFiles(list) {
    clearError();
    for (let i = 0; i < list.length; i++) {
      const f = list[i];
      if (f.type === "application/json") {
        if (f.name === "data_program.json") uploadedFiles.program = f;
        else if (f.name === "data_kegiatan.json") uploadedFiles.kegiatan = f;
        else if (f.name === "data_sub_kegiatan.json")
          uploadedFiles.subkegiatan = f;
        else displayError(`File '${f.name}' tidak dikenal.`);
      } else displayError(`File '${f.name}' bukan JSON.`);
    }
    renderFileList();
    clearOutput();
  }

  let globalProgramsData = [],
    globalKegiatanData = [],
    globalSubKegiatanData = [];

  function appendRow(program, kegiatan, sub, ang, rea, cls = "") {
    const row = dataTableBody.insertRow();
    if (cls) row.classList.add(cls);
    row.insertCell(0).textContent = program;
    row.insertCell(1).textContent = kegiatan;
    row.insertCell(2).textContent = sub;
    row.insertCell(3).textContent = ang;
    row.insertCell(4).textContent = rea;
  }

  // fungsi merge tabel
  function normalizeAndMergeTable(tbody) {
    const rows = Array.from(tbody.rows);
    const matrix = rows.map((r) => {
      const txt = [];
      r.querySelectorAll("td,th").forEach((c) =>
        txt.push(c.textContent.trim())
      );
      while (txt.length < 5) txt.push("");
      return txt.slice(0, 5);
    });
    const rc = matrix.length,
      cc = 5;
    const rowspan = Array.from({ length: rc }, () => Array(cc).fill(1));
    const skip = Array.from({ length: rc }, () => Array(cc).fill(false));

    for (let c = 0; c <= 2; c++) {
      let start = 0;
      while (start < rc) {
        if (!matrix[start][c]) {
          start++;
          continue;
        }
        let end = start + 1;
        while (end < rc && matrix[end][c] === matrix[start][c]) end++;
        const span = end - start;
        if (span > 1) {
          rowspan[start][c] = span;
          for (let r = start + 1; r < end; r++) skip[r][c] = true;
        }
        start = end;
      }
    }

    const newTbody = document.createElement("tbody");
    for (let r = 0; r < rc; r++) {
      const tr = document.createElement("tr");
      for (let c = 0; c < cc; c++) {
        if (skip[r][c]) continue;
        const td = document.createElement("td");
        const val = matrix[r][c];
        td.textContent = val && val !== "-" ? val : c >= 3 ? val : "";
        if (rowspan[r][c] > 1) td.rowSpan = rowspan[r][c];
        tr.appendChild(td);
      }
      newTbody.appendChild(tr);
    }
    tbody.parentNode.replaceChild(newTbody, tbody);
    return newTbody;
  }

  // proses utama
  processButton.addEventListener("click", async () => {
    if (
      !uploadedFiles.program ||
      !uploadedFiles.kegiatan ||
      !uploadedFiles.subkegiatan
    ) {
      displayError("Unggah semua file JSON.");
      return;
    }
    clearError();
    dataTableBody.innerHTML = "";

    try {
      const [p, k, s] = await Promise.all([
        readFileAsText(uploadedFiles.program),
        readFileAsText(uploadedFiles.kegiatan),
        readFileAsText(uploadedFiles.subkegiatan),
      ]);
      globalProgramsData = JSON.parse(p);
      globalKegiatanData = JSON.parse(k);
      globalSubKegiatanData = JSON.parse(s);
    } catch (e) {
      displayError("Gagal membaca JSON: " + e.message);
      return;
    }

    const kegiatanByProgram = new Map();
    globalKegiatanData.forEach((k) => {
      if (!kegiatanByProgram.has(k.id_program))
        kegiatanByProgram.set(k.id_program, []);
      kegiatanByProgram.get(k.id_program).push(k);
    });
    const subByGiat = new Map();
    globalSubKegiatanData.forEach((s) => {
      if (!subByGiat.has(s.id_giat)) subByGiat.set(s.id_giat, []);
      subByGiat.get(s.id_giat).push(s);
    });

    globalProgramsData.sort(
      (a, b) => (a.id_program || 0) - (b.id_program || 0)
    );
    let totalA = 0,
      totalR = 0;

    globalProgramsData.forEach((p) => {
      const prog = capitalizeWords(p.nama_program);
      totalA += +p.anggaran || 0;
      totalR += +p.realisasi_rill || 0;
      const kegs = kegiatanByProgram.get(p.id_program) || [];
      if (!kegs.length)
        appendRow(
          prog,
          "-",
          "-",
          formatRupiahForDisplay(p.anggaran),
          formatRupiahForDisplay(p.realisasi_rill)
        );
      else
        kegs.forEach((k) => {
          const keg = capitalizeWords(k.nama_giat);
          const subs = subByGiat.get(k.id_giat) || [];
          if (!subs.length)
            appendRow(
              prog,
              keg,
              "-",
              formatRupiahForDisplay(k.anggaran),
              formatRupiahForDisplay(k.realisasi_rill)
            );
          else
            subs.forEach((s) => {
              appendRow(
                prog,
                keg,
                capitalizeWords(s.nama_sub_giat),
                formatRupiahForDisplay(s.anggaran),
                formatRupiahForDisplay(s.realisasi_rill)
              );
            });
        });
    });

    const newTbody = normalizeAndMergeTable(
      document.getElementById("data-table").tBodies[0]
    );
    const totalRow = newTbody.insertRow();
    totalRow.classList.add("total-row");
    const t1 = totalRow.insertCell(0);
    t1.textContent = "Jumlah";
    t1.colSpan = 3;
    t1.style.textAlign = "center";
    t1.style.fontWeight = "bold";
    const t2 = totalRow.insertCell(1);
    t2.textContent = formatRupiahForDisplay(totalA);
    t2.style.fontWeight = "bold";
    const t3 = totalRow.insertCell(2);
    t3.textContent = formatRupiahForDisplay(totalR);
    t3.style.fontWeight = "bold";

    dataOutputDiv.style.display = "block";
    downloadExcelButton.style.display = "inline-block";
  });

  // download excel tanpa baris kosong header
  downloadExcelButton.addEventListener("click", () => {
    const date = new Date().toLocaleDateString("id-ID", {
      year: "numeric",
      month: "long",
      day: "numeric",
    });
    const cellStyle =
      "border:1px solid black;padding:4px;vertical-align:top;text-align:left;";
    const headStyle = "font-weight:bold;text-align:center;";

    const rows = Array.from(document.querySelectorAll("#data-table tr"));
    let htmlRows = "";
    rows.forEach((r) => {
      htmlRows += "<tr>";
      r.querySelectorAll("th,td").forEach((c) => {
        const rs = c.rowSpan > 1 ? `rowspan='${c.rowSpan}'` : "";
        const cs = c.colSpan > 1 ? `colspan='${c.colSpan}'` : "";
        let val = c.textContent || "";
        const isNum = /Rp|[0-9]/.test(val);
        if (
          isNum &&
          (val.includes("Rp") ||
            /^[\d\.,\s-]+$/.test(val.replace(/[Rp\s]/g, "")))
        ) {
          val = val.replace(/[^0-9-]/g, "");
          if (!val) val = "0";
          htmlRows += `<td ${rs} ${cs} style="${cellStyle}mso-number-format:'#,##0';text-align:right;">${val}</td>`;
        } else htmlRows += `<td ${rs} ${cs} style="${cellStyle}">${val}</td>`;
      });
      htmlRows += "</tr>";
    });

    const excel = `
      <html xmlns:o="urn:schemas-microsoft-com:office:office"
            xmlns:x="urn:schemas-microsoft-com:office:excel"
            xmlns="http://www.w3.org/TR/REC-html40">
      <head><meta charset="UTF-8"></head>
      <body>
      <table style="border-collapse:collapse;width:100%;">
        <tr><th colspan="5" style="${headStyle}font-size:12pt;border:none;">LAPORAN REALISASI ANGGARAN KEGIATAN</th></tr>
        <tr><th colspan="5" style="${headStyle}font-size:12pt;border:none;">DINAS PERTANIAN KABUPATEN GROBOGAN TAHUN 2025</th></tr>
        <br> 
        <tr></tr>
        <tr><th colspan="5" style="text-align:left;font-size:10pt;border:none;font-weight:normal;">Tanggal Cetak: ${date}</th></tr>
        <tr>
          <th style="${headStyle}${cellStyle}">Program</th>
          <th style="${headStyle}${cellStyle}">Kegiatan</th>
          <th style="${headStyle}${cellStyle}">Sub Kegiatan</th>
          <th style="${headStyle}${cellStyle}">Pagu Anggaran (Rp)</th>
          <th style="${headStyle}${cellStyle}">Realisasi (Rp)</th>
        </tr>
        ${htmlRows}
      </table>
      </body></html>
    `;
    const blob = new Blob([excel], {
      type: "application/vnd.ms-excel;charset=utf-8",
    });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = "Realisasi_Anggaran_Dispertan.xls";
    a.click();
    URL.revokeObjectURL(url);
  });
});
