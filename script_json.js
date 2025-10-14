document.addEventListener("DOMContentLoaded", () => {
  // PENTING: Perubahan ID dari versi asli untuk menghindari konflik dengan Generate
  const dropArea = document.getElementById("drop-area-merge"); // Diubah
  const fileInput = document.getElementById("file-input-merge"); // Diubah
  const fileListDiv = document.getElementById("file-list-merge"); // Diubah
  const mergeButton = document.getElementById("merge-button"); // Tetap
  const outputArea = document.getElementById("output-area-merge"); // Diubah
  const jsonOutput = document.getElementById("json-output"); // Tetap
  const downloadButton = document.getElementById("download-button-json"); // Diubah
  const errorMessageDiv = document.getElementById("error-message-merge"); // Diubah

  let filesToProcess = []; // Array untuk menyimpan objek File

  // --- Fungsi Utilitas ---
  function displayError(message) {
    errorMessageDiv.textContent = message;
    errorMessageDiv.style.display = "block";
  }

  function clearError() {
    errorMessageDiv.textContent = "";
    errorMessageDiv.style.display = "none";
  }

  function updateMergeButtonState() {
    mergeButton.disabled = filesToProcess.length === 0;
  }

  function renderFileList() {
    fileListDiv.innerHTML = ""; // Kosongkan daftar sebelumnya
    if (filesToProcess.length === 0) {
      fileListDiv.style.display = "none";
      return;
    }
    fileListDiv.style.display = "block";

    filesToProcess.forEach((file, index) => {
      const p = document.createElement("p");
      p.textContent = file.name;
      const removeSpan = document.createElement("span");
      removeSpan.textContent = "✅";
      removeSpan.className = "remove-file";
      removeSpan.onclick = () => removeFile(index);
      p.appendChild(removeSpan);
      fileListDiv.appendChild(p);
    });
    updateMergeButtonState();
  }

  function removeFile(indexToRemove) {
    filesToProcess.splice(indexToRemove, 1); // Hapus file dari array
    renderFileList(); // Perbarui tampilan daftar file
    clearOutput();
  }

  function clearOutput() {
    jsonOutput.textContent = "";
    outputArea.style.display = "none";
    downloadButton.style.display = "none";
    clearError();
  }

  // --- Event Listeners untuk Drag and Drop ---

  // Mencegah perilaku default browser untuk drag events
  ["dragenter", "dragover", "dragleave", "drop"].forEach((eventName) => {
    dropArea.addEventListener(eventName, preventDefaults, false);
  });

  function preventDefaults(e) {
    e.preventDefault();
    e.stopPropagation();
  }

  // Menyorot area drop saat ada file di atasnya
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

  // Menangani file yang di-drop
  dropArea.addEventListener("drop", handleDrop, false);

  function handleDrop(e) {
    const dt = e.dataTransfer;
    const files = dt.files; // Dapatkan objek FileList

    addFiles(files);
  }

  // Menangani file yang dipilih melalui input klik
  fileInput.addEventListener("change", (e) => {
    const files = e.target.files;
    addFiles(files);
    // Reset input file agar event 'change' terpicu lagi jika file yang sama dipilih
    e.target.value = null;
  });

  function addFiles(fileList) {
    clearError();
    for (let i = 0; i < fileList.length; i++) {
      const file = fileList[i];
      if (file.type === "application/json") {
        // Cek apakah file dengan nama dan ukuran yang sama sudah ada
        if (
          !filesToProcess.some(
            (f) => f.name === file.name && f.size === file.size
          )
        ) {
          filesToProcess.push(file);
        } else {
          displayError(`File '${file.name}' sudah ada dalam daftar.`);
        }
      } else {
        displayError(`File '${file.name}' bukan file JSON yang valid.`);
      }
    }
    renderFileList();
    clearOutput();
  }

  // Helper untuk membaca file sebagai teks (menggunakan Promise untuk async/await)
  function readFileAsText(file) {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = () => resolve(reader.result);
      reader.onerror = () => reject(reader.error);
      reader.readAsText(file);
    });
  }

  // --- Logika Penggabungan File ---
  mergeButton.addEventListener("click", async () => {
    if (filesToProcess.length === 0) {
      displayError("Tidak ada file JSON untuk digabungkan.");
      return;
    }

    clearError();
    let mergedData = null;

    for (const file of filesToProcess) {
      try {
        // Baca file menggunakan FileReader
        const fileContent = await readFileAsText(file);
        const jsonData = JSON.parse(fileContent);

        if (mergedData === null) {
          // Inisialisasi mergedData dengan data dari file pertama
          mergedData = jsonData;
        } else if (Array.isArray(mergedData) && Array.isArray(jsonData)) {
          // Jika keduanya array, gabungkan array
          mergedData = mergedData.concat(jsonData);
        } else if (
          typeof mergedData === "object" &&
          !Array.isArray(mergedData) &&
          typeof jsonData === "object" &&
          !Array.isArray(jsonData)
        ) {
          // Jika keduanya object, gabungkan object (deep merge sederhana)
          mergedData = { ...mergedData, ...jsonData };
        } else {
          // Kasus tidak didukung, misalnya array digabung dengan object, dll.
          displayError(
            `Gagal menggabungkan file. Jenis data file '${file.name}' tidak kompatibel dengan data yang sudah terkumpul.`
          );
          return; // Hentikan proses penggabungan
        }
      } catch (error) {
        displayError(
          `Gagal membaca atau mem-parse file JSON: ${error.message}. Pastikan file valid.`
        );
        return; // Hentikan proses penggabungan
      }
    }

    if (mergedData !== null) {
      try {
        const outputJsonString = JSON.stringify(mergedData, null, 2);
        jsonOutput.textContent = outputJsonString;
        outputArea.style.display = "block";
        downloadButton.style.display = "inline-block";
      } catch (error) {
        displayError(`Gagal memformat hasil JSON: ${error.message}`);
      }
    } else {
      jsonOutput.textContent = "";
      outputArea.style.display = "none";
      downloadButton.style.display = "none";
    }
  });

  // --- Fungsi Download ---\
  downloadButton.addEventListener("click", () => {
    const outputJsonString = jsonOutput.textContent;
    const blob = new Blob([outputJsonString], { type: "application/json" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = "data_.json";
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  });
});
