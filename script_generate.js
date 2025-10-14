document.addEventListener("DOMContentLoaded", () => {
  // PENTING: Perubahan ID dari versi asli untuk menghindari konflik dengan Merge JSON
  const dropArea = document.getElementById("drop-area-generate"); // Diubah
  const fileInput = document.getElementById("file-input-generate"); // Diubah
  const fileListDiv = document.getElementById("file-list-generate"); // Diubah
  const processButton = document.getElementById("process-button"); // Tetap
  const dataOutputDiv = document.getElementById("data-output"); // Tetap
  const dataTableBody = document.querySelector("#data-table tbody"); // Tetap
  const errorMessageDiv = document.getElementById("error-message-generate"); // Diubah
  const downloadExcelButton = document.getElementById("download-excel-button"); // Tetap

  const uploadedFiles = {
    program: null,
    kegiatan: null,
    subkegiatan: null,
  };

  // --- Utility Functions ---
  function displayError(message) {
    errorMessageDiv.textContent = message;
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
    const filesPresent = Object.values(uploadedFiles).filter(
      (file) => file !== null
    );

    if (filesPresent.length === 0) {
      fileListDiv.style.display = "none";
      return;
    }
    fileListDiv.style.display = "block";

    const fileNamesMap = {
      "data_program.json": "Data Program",
      "data_kegiatan.json": "Data Kegiatan",
      "data_sub_kegiatan.json": "Data Sub Kegiatan",
    };

    for (const key in uploadedFiles) {
      const file = uploadedFiles[key];
      if (file) {
        const p = document.createElement("p");
        p.textContent = `${fileNamesMap[file.name] || file.name} (${
          file.name
        })`;
        const removeSpan = document.createElement("span");
        removeSpan.textContent = "✅";
        removeSpan.className = "remove-file";
        removeSpan.onclick = () => removeFile(key);
        p.appendChild(removeSpan);
        fileListDiv.appendChild(p);
      }
    }
    updateProcessButtonState();
  }

  function removeFile(fileKey) {
    uploadedFiles[fileKey] = null;
    renderFileList();
    clearOutput();
  }

  function clearOutput() {
    dataTableBody.innerHTML = "";
    dataOutputDiv.style.display = "none";
    downloadExcelButton.style.display = "none";
    clearError();
  }

  function readFileAsText(file) {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = () => resolve(reader.result);
      reader.onerror = () => reject(reader.error);
      reader.readAsText(file);
    });
  }

  // Fungsi untuk memformat angka menjadi format mata uang Rupiah untuk TAMPILAN HTML
  function formatRupiahForDisplay(amount) {
    if (typeof amount !== "number") {
      amount = parseFloat(amount);
    }
    if (isNaN(amount) || amount === 0) {
      return "-";
    }
    return new Intl.NumberFormat("id-ID", {
      style: "currency",
      currency: "IDR",
      minimumFractionDigits: 0,
    }).format(amount);
  }

  // Fungsi untuk membersihkan angka dari format Rupiah untuk EXCEL
  function cleanNumberForExcel(amount) {
    if (typeof amount !== "number") {
      amount = parseFloat(amount);
    }
    if (isNaN(amount) || amount === 0) {
      return "0";
    }
    return amount.toString();
  }

  // Fungsi untuk mengubah setiap kata menjadi Capital Case
  function capitalizeWords(str) {
    if (!str) return "";
    return str
      .toLowerCase()
      .replace(/(^|\s)\S/g, (firstLetter) => firstLetter.toUpperCase());
  }

  // --- Event Listeners for Drag and Drop ---
  ["dragenter", "dragover", "dragleave", "drop"].forEach((eventName) => {
    dropArea.addEventListener(eventName, preventDefaults, false);
  });

  function preventDefaults(e) {
    e.preventDefault();
    e.stopPropagation();
  }

  ["dragenter", "dragover"].forEach((eventName) => {
    dropArea.addEventListener(
      eventName,
      () => dropArea.classList.add("highlight"),
      false
    );
  });

  ["dragleave", "drop"].forEach((eventName) => {
    dropArea.addEventListener(
      eventName,
      () => dropArea.classList.remove("highlight"),
      false
    );
  });

  dropArea.addEventListener("drop", handleDrop, false);

  function handleDrop(e) {
    const dt = e.dataTransfer;
    const files = dt.files;
    addFiles(files);
  }

  fileInput.addEventListener("change", (e) => {
    const files = e.target.files;
    addFiles(files);
    e.target.value = null;
  });

  function addFiles(fileList) {
    clearError();
    for (let i = 0; i < fileList.length; i++) {
      const file = fileList[i];
      if (file.type === "application/json") {
        if (file.name === "data_program.json") {
          uploadedFiles.program = file;
        } else if (file.name === "data_kegiatan.json") {
          uploadedFiles.kegiatan = file;
        } else if (file.name === "data_sub_kegiatan.json") {
          uploadedFiles.subkegiatan = file;
        } else {
          displayError(
            `File '${file.name}' tidak dikenal. Harap unggah 'data_program.json', 'data_kegiatan.json', atau 'data_sub_kegiatan.json'.`
          );
        }
      } else {
        displayError(`File '${file.name}' bukan file JSON yang valid.`);
      }
    }
    renderFileList();
    clearOutput();
  }

  let globalProgramsData = [];
  let globalKegiatanData = [];
  let globalSubKegiatanData = [];

  // --- Main Processing Logic ---
  processButton.addEventListener("click", async () => {
    if (
      !uploadedFiles.program ||
      !uploadedFiles.kegiatan ||
      !uploadedFiles.subkegiatan
    ) {
      displayError(
        "Harap unggah ketiga file JSON yang diperlukan: data_program.json, merge_data_kegiatan.json, dan merge_data_subkeg.json."
      );
      return;
    }

    clearError();
    dataTableBody.innerHTML = "";

    try {
      const [programContent, kegiatanContent, subKegiatanContent] =
        await Promise.all([
          readFileAsText(uploadedFiles.program),
          readFileAsText(uploadedFiles.kegiatan),
          readFileAsText(uploadedFiles.subkegiatan),
        ]);

      globalProgramsData = JSON.parse(programContent);
      globalKegiatanData = JSON.parse(kegiatanContent);
      globalSubKegiatanData = JSON.parse(subKegiatanContent);
    } catch (error) {
      displayError(
        `Gagal membaca atau mem-parse file JSON: ${error.message}. Pastikan file valid.`
      );
      return;
    }

    const kegiatanByProgram = new Map();
    globalKegiatanData.forEach((k) => {
      if (!kegiatanByProgram.has(k.id_program)) {
        kegiatanByProgram.set(k.id_program, []);
      }
      kegiatanByProgram.get(k.id_program).push(k);
    });

    const subKegiatanByGiat = new Map();
    globalSubKegiatanData.forEach((sk) => {
      if (!subKegiatanByGiat.has(sk.id_giat)) {
        subKegiatanByGiat.set(sk.id_giat, []);
      }
      subKegiatanByGiat.get(sk.id_giat).push(sk);
    });

    let hasDataToDisplay = false;
    let totalAnggaran = 0;
    let totalRealisasi = 0;

    globalProgramsData.sort((a, b) => {
      if (
        typeof a.id_program === "string" &&
        typeof b.id_program === "string"
      ) {
        return a.id_program.localeCompare(b.id_program);
      }
      return (a.id_program || 0) - (b.id_program || 0);
    });

    globalProgramsData.forEach((program) => {
      const programName = capitalizeWords(program.nama_program);
      const programAnggaran = program.anggaran || 0;
      const programRealisasi = program.realisasi_rill || 0;
      const programAnggaranDisplay = formatRupiahForDisplay(programAnggaran);
      const programRealisasiDisplay = formatRupiahForDisplay(programRealisasi);

      // Add to totals
      totalAnggaran += parseFloat(programAnggaran);
      totalRealisasi += parseFloat(programRealisasi);

      appendRowToTable(
        programName,
        "-",
        "-",
        programAnggaranDisplay,
        programRealisasiDisplay,
        "program-row"
      );
      hasDataToDisplay = true;

      const relatedKegiatan = kegiatanByProgram.get(program.id_program) || [];
      relatedKegiatan.sort((a, b) => {
        if (typeof a.id_giat === "string" && typeof b.id_giat === "string") {
          return a.id_giat.localeCompare(b.id_giat);
        }
        return (a.id_giat || 0) - (b.id_giat || 0);
      });

      if (relatedKegiatan.length > 0) {
        relatedKegiatan.forEach((kegiatan) => {
          const kegiatanName = capitalizeWords(kegiatan.nama_giat);
          const kegiatanAnggaranDisplay = formatRupiahForDisplay(
            kegiatan.anggaran
          );
          const kegiatanRealisasiDisplay = formatRupiahForDisplay(
            kegiatan.realisasi_rill
          );
          appendRowToTable(
            "",
            kegiatanName,
            "-",
            kegiatanAnggaranDisplay,
            kegiatanRealisasiDisplay,
            "kegiatan-row"
          );
          hasDataToDisplay = true;

          const relatedSubKegiatan =
            subKegiatanByGiat.get(kegiatan.id_giat) || [];
          relatedSubKegiatan.sort((a, b) => {
            if (
              typeof a.id_sub_giat === "string" &&
              typeof b.id_sub_giat === "string"
            ) {
              return a.id_sub_giat.localeCompare(b.id_sub_giat);
            }
            return (a.id_sub_giat || 0) - (b.id_sub_giat || 0);
          });

          if (relatedSubKegiatan.length > 0) {
            relatedSubKegiatan.forEach((subkegiatan) => {
              const subKegiatanName = capitalizeWords(
                subkegiatan.nama_sub_giat
              );
              const subKegiatanAnggaranDisplay = formatRupiahForDisplay(
                subkegiatan.anggaran
              );
              const subKegiatanRealisasiDisplay = formatRupiahForDisplay(
                subkegiatan.realisasi_rill
              );
              appendRowToTable(
                "",
                "",
                subKegiatanName,
                subKegiatanAnggaranDisplay,
                subKegiatanRealisasiDisplay,
                "subkegiatan-row"
              );
              hasDataToDisplay = true;
            });
          }
        });
      }
    });

    if (hasDataToDisplay) {
      // Append the total row
      const totalRow = dataTableBody.insertRow();
      totalRow.classList.add("total-row");
      const totalCell = totalRow.insertCell(0);
      totalCell.textContent = "Jumlah";
      totalCell.colSpan = 3;
      totalCell.style.textAlign = "center";
      totalCell.style.fontWeight = "bold";

      const totalAnggaranCell = totalRow.insertCell(1);
      totalAnggaranCell.textContent = formatRupiahForDisplay(totalAnggaran);
      totalAnggaranCell.style.fontWeight = "bold";

      const totalRealisasiCell = totalRow.insertCell(2);
      totalRealisasiCell.textContent = formatRupiahForDisplay(totalRealisasi);
      totalRealisasiCell.style.fontWeight = "bold";

      dataOutputDiv.style.display = "block";
      downloadExcelButton.style.display = "inline-block";
    } else {
      displayError(
        "Tidak ada data yang cocok ditemukan untuk ditampilkan dalam tabel."
      );
      clearOutput();
    }
  });

  function appendRowToTable(
    program,
    kegiatan,
    subKegiatan,
    anggaran,
    realisasi,
    rowClass = ""
  ) {
    const row = dataTableBody.insertRow();
    if (rowClass) {
      row.classList.add(rowClass);
    }
    row.insertCell(0).textContent = program;
    row.insertCell(1).textContent = kegiatan;
    row.insertCell(2).textContent = subKegiatan;
    row.insertCell(3).textContent = anggaran;
    row.insertCell(4).textContent = realisasi;
  }

  // --- Fungsi Download Excel ---
  downloadExcelButton.addEventListener("click", () => {
    const formattedDate = new Date().toLocaleDateString("id-ID", {
      year: "numeric",
      month: "long",
      day: "numeric",
    });

    const cellStyle =
      "border: 1px solid black; padding: 4px; vertical-align: top; text-align: left;";
    const totalRowStyle = "font-weight: bold; background-color: #e0f2f1;";
    const headerStyle =
      "text-align: center; font-weight: bold; background-color: #ccfbf1;";
    const programStyle = "font-weight: bold; background-color: #f0fdf4;";
    const kegiatanStyle = "font-style: italic; background-color: #f7fee7;";
    const subKegiatanStyle = "background-color: #ffffff;";
    const titleStyle = "text-align: center; font-weight: bold;";
    const dateStyle = "text-align: right; font-size: 10px;";

    let tableRows = "";
    // Re-generate table content with inline styles for Excel
    const tableHtml = document.getElementById("data-table").outerHTML;
    const parser = new DOMParser();
    const doc = parser.parseFromString(tableHtml, "text/html");
    const rows = doc.querySelectorAll("#data-table tbody tr");

    rows.forEach((row) => {
      let rowHtml = "<tr>";
      const rowClass = row.className;
      let style = cellStyle;
      if (rowClass.includes("program-row")) style += programStyle;
      else if (rowClass.includes("kegiatan-row")) style += kegiatanStyle;
      else if (rowClass.includes("subkegiatan-row")) style += subKegiatanStyle;
      else if (rowClass.includes("total-row")) style += totalRowStyle;

      const cells = row.querySelectorAll("td");
      cells.forEach((cell, index) => {
        let content;
        let cellContent = cell.textContent;
        let cellStyleFinal = style;

        if (rowClass.includes("total-row")) {
          cellStyleFinal += "font-weight: bold;";
          if (index === 0) {
            cellStyleFinal += "text-align: center;";
          }
        }

        // Columns for Anggaran (3) and Realisasi (4) need special handling
        if (index === 3 || index === 4) {
          // Clean number for Excel (remove commas/Rp, keep only number)
          let cleanValue = cleanNumberForExcel(
            cellContent.replace(/[^0-9,-]/g, "").replace(",", ".")
          );

          content = cleanValue;
          cellStyleFinal += 'mso-number-format:"#,##0"; text-align: right;';
        } else {
          content = cellContent;
        }

        rowHtml += `<td colspan="${
          cell.colSpan || 1
        }" style="${cellStyleFinal}">${content}</td>`;
      });
      rowHtml += "</tr>";
      tableRows += rowHtml;
    });

    // Final Excel HTML structure
    let excelTableHtml = `
            <table style="border-collapse: collapse; width: 100%;">
                <tr>
                    <th colspan="5" style="${titleStyle} font-size: 16pt;">LAPORAN REALISASI ANGGARAN</th>
                </tr>
                <tr>
                    <th colspan="5" style="${titleStyle} font-size: 12pt;">DINAS PERTANIAN KABUPATEN GROBOGAN</th>
                </tr>
                <tr>
                    <th colspan="5" style="${dateStyle} font-size: 10pt;">Tanggal Cetak: ${formattedDate}</th>
                </tr>
                <tr>
                    <th style="${headerStyle} ${cellStyle}">Program</th>
                    <th style="${headerStyle} ${cellStyle}">Kegiatan</th>
                    <th style="${headerStyle} ${cellStyle}">Sub Kegiatan</th>
                    <th style="${headerStyle} ${cellStyle}">Anggaran (Rp)</th>
                    <th style="${headerStyle} ${cellStyle}">Realisasi Rill (Rp)</th>
                </tr>
                <tbody>
                    ${tableRows}
                </tbody>
            </table>
            <br>
            <table style="width: 100%; border-collapse: collapse;">
                <tr>
                    <td style="width: 33%; ${cellStyle} border: none;"></td>
                    <td style="width: 33%; ${cellStyle} border: none;"></td>
                    <td colspan="3" style="${cellStyle} border: none; text-align: center;">a.n. Kepala Dinas Pertanian Kab. Grobogan</td>
                </tr>
                <tr>
                    <td style="width: 33%; ${cellStyle} border: none;"></td>
                    <td style="width: 33%; ${cellStyle} border: none;"></td>
                    <td colspan="3" style="${cellStyle} border: none; text-align: center;">Plt. Sekretaris</td>
                </tr>
                <tr>
                    <td></td>
                    <td></td>
                    <td colspan="3" style="text-align: center;">Kepala Bidang Sarpras & Perlindungan Tanaman</td>
                </tr>
                <tr>
                    <td></td>
                    <td></td>
                    <td colspan="3" style="text-align: center; height: 70px;"></td>
                </tr>
                <tr>
                    <td></td>
                    <td></td>
                    <td colspan="3" style="text-align: center;"><u>Slamet Waluyono, S.H., M.M</u></td>
                </tr>
                <tr>
                    <td></td>
                    <td></td>
                    <td colspan="3" style="text-align: center;">Pembina Utama Muda</td>
                </tr>
                <tr>
                    <td></td>
                    <td></td>
                    <td colspan="3" style="text-align: center;">NIP. 198606132010011012</td>
                </tr>
            </table>
        `;

    const excelContent = `
            <html xmlns:o="urn:schemas-microsoft-com:office:office"
            xmlns:x="urn:schemas-microsoft-com:office:excel"
            xmlns="http://www.w3.org/TR/REC-html40">
            <head>
                <meta charset="UTF-8">
                <style>
                    table {
                        border-collapse: collapse;
                    }
                </style>
                </head>
            <body>
                ${excelTableHtml}
            </body>
            </html>
        `;

    const blob = new Blob([excelContent], {
      type: "application/vnd.ms-excel;charset=utf-8",
    });
    const url = URL.createObjectURL(blob);

    const a = document.createElement("a");
    a.href = url;
    a.download = "data_program_kegiatan_subkegiatan.xls";
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  });
});
