// ==============================================================================
// VOTER SUVIDHA - APPLICATION LOGIC
// ==============================================================================

let selectedFiles = [];
let extractedData = null;
let candidatePhotoData = "";
let partySymbolData = "";

document.addEventListener('DOMContentLoaded', () => {
  setupDropzone();
  setupRadioOptions();
  setupActionButtons();
  setupCandidatePhoto();
  setupPartySymbol();
});

// Dropzone & File Selection
function setupDropzone() {
  const dropzone = document.getElementById('dropzone');
  const fileInput = document.getElementById('pdfFileInput');

  dropzone.addEventListener('dragover', (e) => {
    e.preventDefault();
    dropzone.classList.add('drag-over');
  });

  dropzone.addEventListener('dragleave', () => {
    dropzone.classList.remove('drag-over');
  });

  dropzone.addEventListener('drop', (e) => {
    e.preventDefault();
    dropzone.classList.remove('drag-over');
    if (e.dataTransfer.files && e.dataTransfer.files.length > 0) {
      handleFiles(Array.from(e.dataTransfer.files));
    }
  });

  fileInput.addEventListener('change', (e) => {
    if (e.target.files && e.target.files.length > 0) {
      handleFiles(Array.from(e.target.files));
    }
  });

  document.getElementById('btnClearFiles').addEventListener('click', () => {
    selectedFiles = [];
    renderFileList();
  });
}

function handleFiles(files) {
  const pdfs = files.filter(f => f.name.toLowerCase().endsWith('.pdf'));
  if (pdfs.length === 0) {
    alert('कृपया केवल PDF (.pdf) फाइलें चुनें।');
    return;
  }

  // Merge new files avoid duplicates
  pdfs.forEach(f => {
    if (!selectedFiles.some(existing => existing.name === f.name)) {
      selectedFiles.push(f);
    }
  });

  renderFileList();
}

function renderFileList() {
  const container = document.getElementById('fileListContainer');
  const list = document.getElementById('fileList');
  list.innerHTML = '';

  if (selectedFiles.length === 0) {
    container.style.display = 'none';
    return;
  }

  container.style.display = 'block';
  selectedFiles.forEach((f, idx) => {
    const li = document.createElement('li');
    li.className = 'file-item';
    const mb = (f.size / (1024 * 1024)).toFixed(2);
    li.innerHTML = `
      <span class="file-item-name">📄 ${f.name}</span>
      <div>
        <span class="file-item-size">${mb} MB &nbsp;</span>
        <button type="button" class="btn btn-outline" style="padding:2px 8px;font-size:0.8rem;" onclick="removeFile(${idx})">✕</button>
      </div>
    `;
    list.appendChild(li);
  });
}

function removeFile(index) {
  selectedFiles.splice(index, 1);
  renderFileList();
}

function setupRadioOptions() {
  const radios = document.querySelectorAll('input[name="slipsPerPage"]');
  radios.forEach(r => {
    r.addEventListener('change', (e) => {
      document.querySelectorAll('.layout-option').forEach(lo => lo.classList.remove('active'));
      e.target.closest('.layout-option').classList.add('active');
    });
  });
}

// Convert file to Base64
function fileToBase64(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => {
      const b64 = reader.result.split(',')[1];
      resolve(b64);
    };
    reader.onerror = reject;
    reader.readAsDataURL(file);
  });
}

// Compress and resize image using HTML5 Canvas to prevent memory overflow
function compressImage(file, maxWidth, maxHeight, quality = 0.85, mimeType = 'image/jpeg') {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      const img = new Image();
      img.onload = () => {
        let width = img.width;
        let height = img.height;

        if (width > maxWidth || height > maxHeight) {
          const ratio = Math.min(maxWidth / width, maxHeight / height);
          width = Math.round(width * ratio);
          height = Math.round(height * ratio);
        }

        const canvas = document.createElement('canvas');
        canvas.width = width;
        canvas.height = height;
        const ctx = canvas.getContext('2d');
        ctx.imageSmoothingEnabled = true;
        ctx.imageSmoothingQuality = 'high';
        ctx.drawImage(img, 0, 0, width, height);

        const outMime = (file.type === 'image/png' && mimeType === 'image/png') ? 'image/png' : mimeType;
        const dataUrl = canvas.toDataURL(outMime, quality);
        resolve(dataUrl);
      };
      img.onerror = () => reject(new Error('छवि लोड करने में असमर्थ।'));
      img.src = e.target.result;
    };
    reader.onerror = () => reject(new Error('फाइल पढ़ने में असमर्थ।'));
    reader.readAsDataURL(file);
  });
}


// Process Uploaded PDFs
async function processUploadedFiles() {
  if (selectedFiles.length === 0) {
    alert('कृपया पहले कम से कम एक वोटर लिस्ट PDF फाइल चुनें।');
    return;
  }

  const indicator = document.getElementById('processingIndicator');
  const procTitle = document.getElementById('processingTitle');
  const procDetail = document.getElementById('processingDetail');

  indicator.style.display = 'flex';
  procTitle.innerText = `मतदाता सूची (${selectedFiles.length} भाग) का विश्लेषण प्रारंभ...`;
  procDetail.innerText = 'फाइलें सर्वर पर अपलोड की जा रही हैं...';

  try {
    const filePayloads = [];
    for (let i = 0; i < selectedFiles.length; i++) {
      const f = selectedFiles[i];
      procDetail.innerText = `फाइल ${i + 1}/${selectedFiles.length}: ${f.name} पढ़ी जा रही है...`;
      const b64 = await fileToBase64(f);
      filePayloads.push({
        filename: f.name,
        data: b64
      });
    }

    procTitle.innerText = 'मतदाता नामावली से डेटा निष्कर्षण (Extraction) जारी...';
    procDetail.innerText = 'कृती देव फॉन्ट डिकोडिंग और विलोपन सूची (Deleted Voters) की छंटनी की जा रही है...';

    const resp = await fetch('/api/upload', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ files: filePayloads })
    });

    if (!resp.ok) {
      const err = await resp.text();
      throw new Error(err || 'सर्वर पर विश्लेषण में त्रुटि आई।');
    }

    const result = await resp.json();
    extractedData = result;
    renderSummary(result);

  } catch (err) {
    alert('त्रुटि: ' + err.message);
  } finally {
    indicator.style.display = 'none';
  }
}

function renderSummary(data) {
  document.getElementById('statActiveVoters').innerText = data.totalActive.toLocaleString('en-IN');
  document.getElementById('statDeletedVoters').innerText = data.totalDeleted.toLocaleString('en-IN');
  document.getElementById('statTotalSerials').innerText = data.totalSerials.toLocaleString('en-IN');
  document.getElementById('statWardNum').innerText = data.ward || '20';
  document.getElementById('statPartsCount').innerText = `${data.parts.length} भाग शामिल`;

  // Render Parts table
  const tbody = document.getElementById('partsTableBody');
  tbody.innerHTML = '';

  data.parts.forEach((p) => {
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td style="font-weight:700; text-align:center;">भाग ${p.part}</td>
      <td style="font-weight:600; color:#1a237e; line-height:1.4;">
        <span style="color:#e53935; margin-right:6px; font-size:1.1rem;">📍</span>${p.booth || 'उपलब्ध नहीं'}
      </td>
      <td style="text-align:center;">${p.totalSerials}</td>
      <td style="text-align:center; font-weight:700; color:#2e7d32;">${p.activeCount}</td>
    `;
    tbody.appendChild(tr);
  });

  // Reveal sections
  document.getElementById('summarySection').style.display = 'block';
  document.getElementById('configSection').style.display = 'block';
  document.getElementById('actionSection').style.display = 'block';

  // Scroll to summary smoothly
  document.getElementById('summarySection').scrollIntoView({ behavior: 'smooth' });
}

function setupCandidatePhoto() {
  const photoInput = document.getElementById('candidatePhotoInput');
  const previewBox = document.getElementById('photoPreviewBox');
  const previewImg = document.getElementById('candidatePhotoPreview');
  const btnRemove = document.getElementById('btnRemovePhoto');

  if (!photoInput) return;

  photoInput.addEventListener('change', async (e) => {
    if (e.target.files && e.target.files[0]) {
      try {
        const file = e.target.files[0];
        // Automatically compress candidate photo to max 400x500 at 85% JPEG to avoid out-of-memory errors
        candidatePhotoData = await compressImage(file, 400, 500, 0.85, 'image/jpeg');
        previewImg.src = candidatePhotoData;
        previewBox.style.display = 'block';
      } catch (err) {
        console.error('Photo compression error:', err);
        const b64 = await fileToBase64(e.target.files[0]);
        candidatePhotoData = `data:${e.target.files[0].type || 'image/jpeg'};base64,${b64}`;
        previewImg.src = candidatePhotoData;
        previewBox.style.display = 'block';
      }
    }
  });

  if (btnRemove) {
    btnRemove.addEventListener('click', () => {
      candidatePhotoData = '';
      photoInput.value = '';
      previewImg.src = '';
      previewBox.style.display = 'none';
    });
  }
}

function setupPartySymbol() {
  const symbolInput = document.getElementById('partySymbolInput');
  const previewBox = document.getElementById('symbolPreviewBox');
  const previewImg = document.getElementById('partySymbolPreview');
  const btnRemove = document.getElementById('btnRemoveSymbol');

  if (!symbolInput) return;

  symbolInput.addEventListener('change', async (e) => {
    if (e.target.files && e.target.files[0]) {
      try {
        const file = e.target.files[0];
        // Compress party symbol to max 300x300, preserving PNG transparency if PNG
        const mime = file.type === 'image/png' ? 'image/png' : 'image/jpeg';
        partySymbolData = await compressImage(file, 300, 300, 0.85, mime);
        previewImg.src = partySymbolData;
        previewBox.style.display = 'block';
      } catch (err) {
        console.error('Symbol compression error:', err);
        const b64 = await fileToBase64(e.target.files[0]);
        partySymbolData = `data:${e.target.files[0].type || 'image/png'};base64,${b64}`;
        previewImg.src = partySymbolData;
        previewBox.style.display = 'block';
      }
    }
  });

  if (btnRemove) {
    btnRemove.addEventListener('click', () => {
      partySymbolData = '';
      symbolInput.value = '';
      previewImg.src = '';
      previewBox.style.display = 'none';
    });
  }
}

function getConfigurationPayload() {
  const candidatePost = document.getElementById('candidatePostSelect')?.value || 'सरपंच';
  const candidateName = document.getElementById('candidateNameInput').value.trim() || 'मनोज बाबेल';
  const partyName = document.getElementById('partyNameInput').value.trim() || 'भारतीय जनता पार्टी (BJP)';
  const bottomMessage = document.getElementById('bottomMessageInput').value.trim() || 'को अपना अमूल्य वोट देकर भारी मतों से विजयी बनाएं!';
  const slipsPerPage = parseInt(document.querySelector('input[name="slipsPerPage"]:checked')?.value || '8', 10);

  // Booth addresses are automatically extracted from the uploaded PDF
  const parts = [];
  if (extractedData && extractedData.parts) {
    extractedData.parts.forEach((p) => {
      parts.push({
        part: p.part,
        booth: p.booth || ''
      });
    });
  }

  return {
    candidatePost,
    candidateName,
    partyName,
    bottomMessage,
    slipsPerPage,
    candidatePhoto: candidatePhotoData,
    partySymbol: partySymbolData,
    parts
  };
}

// Action Handlers
function setupActionButtons() {
  document.getElementById('btnProcessPdfs').addEventListener('click', processUploadedFiles);

  // Generate Excel
  document.getElementById('btnGenerateExcel').addEventListener('click', async () => {
    const config = getConfigurationPayload();
    showStatus('⏳', 'एक्सेल फाइल तैयार की जा रही है...', 'कृपया प्रतीक्षा करें, मास्टर वोटर लिस्ट का संकलन जारी है...');

    try {
      const resp = await fetch('/api/generate-excel', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(config)
      });

      if (!resp.ok) {
        throw new Error(await resp.text() || 'एक्सेल निर्माण में त्रुटि आई।');
      }

      const res = await resp.json();
      showSuccessStatus(
        '📊',
        'एक्सेल फाइल सफलतापूर्वक तैयार हो गई!',
        `कुल ${res.totalVoters.toLocaleString('en-IN')} मतदाताओं की 11-कॉलम एक्सेल लिस्ट तैयार है।`,
        res.downloadUrl,
        res.filename || 'voter_list.xlsx'
      );
    } catch (err) {
      showErrorStatus(err.message);
    }
  });

  // Preview Slips
  document.getElementById('btnPreviewSlips').addEventListener('click', async () => {
    const config = getConfigurationPayload();
    try {
      const resp = await fetch('/api/preview', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(config)
      });

      if (!resp.ok) {
        throw new Error(await resp.text() || 'नमूना पूर्वावलोकन तैयार नहीं हो सका।');
      }

      const html = await resp.text();
      const iframe = document.getElementById('previewIframe');
      iframe.srcdoc = html;

      document.getElementById('previewInfo').innerText = `लेआउट: ${config.slipsPerPage} स्लिप्स प्रति A4 पृष्ठ | पद: ${config.candidatePost} | प्रत्याशी: ${config.candidateName}`;
      document.getElementById('previewModal').style.display = 'flex';
    } catch (err) {
      alert('पूर्वावलोकन त्रुटि: ' + err.message);
    }
  });

  // Generate PDF
  document.getElementById('btnGeneratePdf').addEventListener('click', generatePdfFromPreview);
}

async function generatePdfFromPreview() {
  closePreviewModal();
  const config = getConfigurationPayload();
  showStatus('⏳', 'वोटर स्लिप PDF तैयार की जा रही है...', `प्रति पृष्ठ ${config.slipsPerPage} स्लिप्स के अनुसार सभी पेजों का उच्च-गुणवत्ता रेंडरिंग जारी है... इसमें 15-30 सेकंड लग सकते हैं।`);

  try {
    const resp = await fetch('/api/generate-pdf', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(config)
    });

    if (!resp.ok) {
      throw new Error(await resp.text() || 'PDF निर्माण में त्रुटि आई।');
    }

    const res = await resp.json();
    showSuccessStatus(
      '✅',
      'वोटर स्लिप PDF सफलतापूर्वक तैयार हो गई!',
      `कुल ${res.totalPages} A4 पेजों में ${res.totalVoters.toLocaleString('en-IN')} पर्चियां (${config.slipsPerPage} स्लिप्स/पेज) तैयार हैं।`,
      res.downloadUrl,
      res.filename || 'voter_slips.pdf'
    );
  } catch (err) {
    showErrorStatus(err.message);
  }
}

function closePreviewModal() {
  document.getElementById('previewModal').style.display = 'none';
}

function showStatus(icon, title, msg) {
  const box = document.getElementById('genStatusBox');
  const iconEl = document.getElementById('genStatusIcon');
  const titleEl = document.getElementById('genStatusTitle');
  const msgEl = document.getElementById('genStatusMsg');
  const dlArea = document.getElementById('genDownloadArea');

  box.style.display = 'flex';
  box.style.background = '#e3f2fd';
  box.style.borderColor = '#90caf9';
  iconEl.innerText = icon;
  titleEl.innerText = title;
  titleEl.style.color = '#0d47a1';
  msgEl.innerText = msg;
  msgEl.style.color = '#1565c0';
  dlArea.style.display = 'none';
  dlArea.innerHTML = '';
  box.scrollIntoView({ behavior: 'smooth' });
}

function showSuccessStatus(icon, title, msg, downloadUrl, filename) {
  const box = document.getElementById('genStatusBox');
  const iconEl = document.getElementById('genStatusIcon');
  const titleEl = document.getElementById('genStatusTitle');
  const msgEl = document.getElementById('genStatusMsg');
  const dlArea = document.getElementById('genDownloadArea');

  box.style.display = 'flex';
  box.style.background = '#e8f5e9';
  box.style.borderColor = '#81c784';
  iconEl.innerText = icon;
  titleEl.innerText = title;
  titleEl.style.color = '#1b5e20';
  msgEl.innerText = msg;
  msgEl.style.color = '#2e7d32';

  dlArea.style.display = 'block';
  dlArea.innerHTML = `
    <a href="${downloadUrl}" download="${filename}" class="btn btn-success btn-large" style="margin-right:12px;">
      ⬇️ ${filename} डाउनलोड करें
    </a>
    <a href="${downloadUrl}" target="_blank" class="btn btn-secondary btn-large">
      🖨️ सीधे ब्राउज़र में खोलें / प्रिंट करें
    </a>
  `;
  box.scrollIntoView({ behavior: 'smooth' });
}

function showErrorStatus(msg) {
  const box = document.getElementById('genStatusBox');
  const iconEl = document.getElementById('genStatusIcon');
  const titleEl = document.getElementById('genStatusTitle');
  const msgEl = document.getElementById('genStatusMsg');
  const dlArea = document.getElementById('genDownloadArea');

  box.style.display = 'flex';
  box.style.background = '#ffebee';
  box.style.borderColor = '#ef9a9a';
  iconEl.innerText = '❌';
  titleEl.innerText = 'त्रुटि हुई';
  titleEl.style.color = '#b71c1c';
  msgEl.innerText = msg;
  msgEl.style.color = '#c62828';
  dlArea.style.display = 'none';
}
