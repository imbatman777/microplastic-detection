/* ═══════════════════════════════════════════════════════════════════
   PlastiSense AI — Frontend Application Logic
   Handles drag-and-drop uploads, API calls, result rendering, history
   ═══════════════════════════════════════════════════════════════════ */

(() => {
  'use strict';

  // ── Polymer full names ─────────────────────────────────────────
  const POLYMER_NAMES = {
    PE:  'Polyethylene',
    PET: 'Polyethylene Terephthalate',
    PP:  'Polypropylene',
    PS:  'Polystyrene',
  };

  // ── State ──────────────────────────────────────────────────────
  let imageFile = null;
  let spectralFile = null;
  let lastImageResult = null;
  let lastSpectralResult = null;
  const history = [];

  // ── DOM References ─────────────────────────────────────────────
  const $ = (sel) => document.querySelector(sel);
  const statusResnet     = $('#statusResnet');
  const statusCnn        = $('#statusCnn');
  const imageDropzone    = $('#imageDropzone');
  const imageInput       = $('#imageInput');
  const imageBrowse      = $('#imageBrowse');
  const imagePreview     = $('#imagePreview');
  const imageThumb       = $('#imageThumb');
  const imageName        = $('#imageName');
  const imageSize        = $('#imageSize');
  const imageRemove      = $('#imageRemove');
  const imageAnalyzeBtn  = $('#imageAnalyzeBtn');
  const spectralDropzone = $('#spectralDropzone');
  const spectralInput    = $('#spectralInput');
  const spectralBrowse   = $('#spectralBrowse');
  const spectralPreview  = $('#spectralPreview');
  const spectralName     = $('#spectralName');
  const spectralSize     = $('#spectralSize');
  const spectralRemove   = $('#spectralRemove');
  const spectralAnalyzeBtn = $('#spectralAnalyzeBtn');
  const resultsSection   = $('#resultsSection');
  const resultPolymer    = $('#resultPolymer');
  const resultFullname   = $('#resultFullname');
  const confidenceValue  = $('#confidenceValue');
  const confidenceBar    = $('#confidenceBar');
  const statTime         = $('#statTime');
  const statModel        = $('#statModel');
  const resultsModelTag  = $('#resultsModelTag');
  const dualAgreement    = $('#dualAgreement');
  const historySection   = $('#historySection');
  const historyList      = $('#historyList');

  // ── Helpers ────────────────────────────────────────────────────
  function formatFileSize(bytes) {
    if (bytes < 1024) return bytes + ' B';
    if (bytes < 1048576) return (bytes / 1024).toFixed(1) + ' KB';
    return (bytes / 1048576).toFixed(1) + ' MB';
  }

  function showToast(message, type = 'error') {
    const existing = document.querySelectorAll('.toast');
    existing.forEach(t => t.remove());

    const toast = document.createElement('div');
    toast.className = `toast toast--${type}`;
    toast.textContent = message;
    document.body.appendChild(toast);
    setTimeout(() => {
      toast.style.opacity = '0';
      toast.style.transition = 'opacity 0.3s ease';
      setTimeout(() => toast.remove(), 300);
    }, 4000);
  }

  // ── Health Check ───────────────────────────────────────────────
  async function checkHealth() {
    try {
      const res = await fetch('/api/health');
      const data = await res.json();

      if (data.resnet_loaded) {
        statusResnet.className = 'status-badge status-badge--ready';
        statusResnet.innerHTML = '<span class="status-dot"></span>ResNet-18 — Ready';
      } else {
        statusResnet.className = 'status-badge status-badge--error';
        statusResnet.innerHTML = '<span class="status-dot"></span>ResNet-18 — Failed';
      }

      if (data.cnn_loaded) {
        statusCnn.className = 'status-badge status-badge--ready';
        statusCnn.innerHTML = '<span class="status-dot"></span>1D-CNN — Ready';
      } else {
        statusCnn.className = 'status-badge status-badge--error';
        statusCnn.innerHTML = '<span class="status-dot"></span>1D-CNN — Failed';
      }
    } catch {
      statusResnet.className = 'status-badge status-badge--error';
      statusResnet.innerHTML = '<span class="status-dot"></span>ResNet-18 — Offline';
      statusCnn.className = 'status-badge status-badge--error';
      statusCnn.innerHTML = '<span class="status-dot"></span>1D-CNN — Offline';
    }
  }

  // ── Drag-and-Drop Setup ────────────────────────────────────────
  function setupDropzone(dropzone, input, browseBtn, onFile) {
    let isOpening = false; // guard against Firefox click loops

    // Prevent browser default & set dropEffect
    ['dragenter', 'dragover', 'dragleave', 'drop'].forEach(evt => {
      dropzone.addEventListener(evt, (e) => {
        e.preventDefault();
        e.stopPropagation();
        if (evt === 'dragover' && e.dataTransfer) {
          e.dataTransfer.dropEffect = 'copy';
        }
      });
    });

    // Visual feedback
    ['dragenter', 'dragover'].forEach(evt => {
      dropzone.addEventListener(evt, () => dropzone.classList.add('dropzone--active'));
    });
    ['dragleave', 'drop'].forEach(evt => {
      dropzone.addEventListener(evt, () => dropzone.classList.remove('dropzone--active'));
    });

    // Drop handler
    dropzone.addEventListener('drop', (e) => {
      const files = e.dataTransfer && e.dataTransfer.files;
      if (files && files.length > 0) onFile(files[0]);
    });

    // Click to browse (with guard to prevent recursive loop)
    function openFilePicker() {
      if (isOpening) return;
      isOpening = true;
      input.click();
      setTimeout(() => { isOpening = false; }, 500);
    }

    browseBtn.addEventListener('click', (e) => {
      e.preventDefault();
      e.stopPropagation();
      openFilePicker();
    });
    dropzone.addEventListener('click', (e) => {
      e.preventDefault();
      openFilePicker();
    });

    input.addEventListener('change', () => {
      if (input.files[0]) onFile(input.files[0]);
      input.value = ''; // allow re-selecting same file
    });
  }

  // ── Image File Handling ────────────────────────────────────────
  function handleImageFile(file) {
    // Validate by extension (more reliable across browsers than MIME type)
    const validExts = ['.jpg', '.jpeg', '.png', '.bmp', '.tiff', '.tif'];
    const ext = '.' + file.name.split('.').pop().toLowerCase();
    if (!validExts.includes(ext)) {
      showToast('Please upload a valid image file (JPG, PNG, BMP, TIFF)');
      return;
    }
    if (file.size > 10 * 1024 * 1024) {
      showToast('Image must be under 10 MB');
      return;
    }

    imageFile = file;

    // Show preview
    const reader = new FileReader();
    reader.onload = (e) => {
      imageThumb.src = e.target.result;
    };
    reader.readAsDataURL(file);

    imageName.textContent = file.name;
    imageSize.textContent = formatFileSize(file.size);
    imagePreview.classList.add('file-preview--visible');
    imageDropzone.classList.add('dropzone--has-file');
    imageAnalyzeBtn.classList.add('analyze-btn--visible');
    imageAnalyzeBtn.disabled = false;
  }

  function clearImageFile() {
    imageFile = null;
    imagePreview.classList.remove('file-preview--visible');
    imageDropzone.classList.remove('dropzone--has-file');
    imageAnalyzeBtn.classList.remove('analyze-btn--visible');
    imageAnalyzeBtn.disabled = true;
    imageThumb.src = '';
  }

  // ── Spectral File Handling ─────────────────────────────────────
  function handleSpectralFile(file) {
    if (!file.name.toLowerCase().endsWith('.csv')) {
      showToast('Please upload a CSV file');
      return;
    }
    if (file.size > 50 * 1024 * 1024) {
      showToast('CSV must be under 50 MB');
      return;
    }

    spectralFile = file;

    spectralName.textContent = file.name;
    spectralSize.textContent = formatFileSize(file.size);
    spectralPreview.classList.add('file-preview--visible');
    spectralDropzone.classList.add('dropzone--has-file');
    spectralAnalyzeBtn.classList.add('analyze-btn--visible');
    spectralAnalyzeBtn.disabled = false;
  }

  function clearSpectralFile() {
    spectralFile = null;
    spectralPreview.classList.remove('file-preview--visible');
    spectralDropzone.classList.remove('dropzone--has-file');
    spectralAnalyzeBtn.classList.remove('analyze-btn--visible');
    spectralAnalyzeBtn.disabled = true;
  }

  // ── Display Results ────────────────────────────────────────────
  function displayResult(data, modelType) {
    const polymer = data.predicted_class;
    const confidence = data.confidence;
    const timeMs = data.time_ms;
    const probs = data.probabilities;

    // Polymer label
    resultPolymer.textContent = polymer;
    resultPolymer.className = `classification__polymer polymer--${polymer}`;
    resultFullname.textContent = POLYMER_NAMES[polymer] || polymer;

    // Confidence
    confidenceValue.textContent = (confidence * 100).toFixed(1) + '%';
    confidenceBar.className = `confidence__bar-fill confidence__bar-fill--${polymer}`;
    // Trigger reflow for animation
    confidenceBar.style.width = '0%';
    requestAnimationFrame(() => {
      requestAnimationFrame(() => {
        confidenceBar.style.width = (confidence * 100) + '%';
      });
    });

    // Stats
    statTime.textContent = timeMs.toFixed(1) + ' ms';

    if (modelType === 'image') {
      statModel.textContent = 'ResNet-18';
      resultsModelTag.className = 'results__model-tag results__model-tag--resnet';
      resultsModelTag.textContent = 'ResNet-18';
    } else {
      statModel.textContent = '1D-CNN';
      resultsModelTag.className = 'results__model-tag results__model-tag--cnn';
      resultsModelTag.textContent = '1D-CNN';
    }

    // Probabilities
    const classes = ['PE', 'PET', 'PP', 'PS'];
    classes.forEach(cls => {
      const bar = $(`#prob${cls}`);
      const val = $(`#prob${cls}Val`);
      const p = probs[cls] || 0;
      bar.style.width = '0%';
      requestAnimationFrame(() => {
        requestAnimationFrame(() => {
          bar.style.width = (p * 100) + '%';
        });
      });
      val.textContent = (p * 100).toFixed(1) + '%';
    });

    // Dual model agreement check
    if (modelType === 'image') lastImageResult = data;
    if (modelType === 'spectral') lastSpectralResult = data;

    if (lastImageResult && lastSpectralResult) {
      dualAgreement.classList.add('dual-agreement--visible');
      if (lastImageResult.predicted_class === lastSpectralResult.predicted_class) {
        dualAgreement.className = 'dual-agreement dual-agreement--visible dual-agreement--match';
        dualAgreement.textContent = `✅ Dual Model Agreement — Both models predict ${lastImageResult.predicted_class} (${POLYMER_NAMES[lastImageResult.predicted_class]})`;
      } else {
        dualAgreement.className = 'dual-agreement dual-agreement--visible dual-agreement--mismatch';
        dualAgreement.textContent = `⚠️ Model Disagreement — ResNet-18: ${lastImageResult.predicted_class} vs 1D-CNN: ${lastSpectralResult.predicted_class}`;
      }
    } else {
      dualAgreement.classList.remove('dual-agreement--visible');
    }

    // Show results section
    resultsSection.classList.add('results-section--visible');
    resultsSection.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
  }

  // ── History ────────────────────────────────────────────────────
  function addToHistory(data, modelType, fileName) {
    const entry = {
      polymer: data.predicted_class,
      confidence: data.confidence,
      model: modelType,
      fileName,
      timestamp: new Date(),
    };
    history.unshift(entry);
    renderHistory();
  }

  function renderHistory() {
    if (history.length === 0) {
      historySection.style.display = 'none';
      return;
    }
    historySection.style.display = 'block';
    historyList.innerHTML = history.map((entry) => {
      const modelClass = entry.model === 'image' ? 'resnet' : 'cnn';
      const modelLabel = entry.model === 'image' ? 'ResNet-18' : '1D-CNN';
      const timeStr = entry.timestamp.toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' });
      return `
        <div class="history-item">
          <span class="history-item__polymer polymer--${entry.polymer}">${entry.polymer}</span>
          <span class="history-item__file" title="${entry.fileName}">${entry.fileName}</span>
          <span class="history-item__confidence">${(entry.confidence * 100).toFixed(1)}%</span>
          <span class="history-item__model history-item__model--${modelClass}">${modelLabel}</span>
        </div>
      `;
    }).join('');
  }

  // ── API Calls ──────────────────────────────────────────────────
  async function analyzeImage() {
    if (!imageFile) return;

    imageAnalyzeBtn.classList.add('analyze-btn--loading');
    imageAnalyzeBtn.disabled = true;

    const formData = new FormData();
    formData.append('file', imageFile);

    try {
      const res = await fetch('/api/predict/image', {
        method: 'POST',
        body: formData,
      });

      if (!res.ok) {
        const err = await res.json();
        throw new Error(err.detail || 'Analysis failed');
      }

      const data = await res.json();
      displayResult(data, 'image');
      addToHistory(data, 'image', imageFile.name);
      showToast(`Detected: ${data.predicted_class} (${(data.confidence * 100).toFixed(1)}%)`, 'success');
    } catch (err) {
      showToast(err.message || 'Failed to analyze image');
    } finally {
      imageAnalyzeBtn.classList.remove('analyze-btn--loading');
      imageAnalyzeBtn.disabled = false;
    }
  }

  async function analyzeSpectral() {
    if (!spectralFile) return;

    spectralAnalyzeBtn.classList.add('analyze-btn--loading');
    spectralAnalyzeBtn.disabled = true;

    const formData = new FormData();
    formData.append('file', spectralFile);

    try {
      const res = await fetch('/api/predict/spectral', {
        method: 'POST',
        body: formData,
      });

      if (!res.ok) {
        const err = await res.json();
        throw new Error(err.detail || 'Analysis failed');
      }

      const data = await res.json();
      displayResult(data, 'spectral');
      addToHistory(data, 'spectral', spectralFile.name);
      showToast(`Detected: ${data.predicted_class} (${(data.confidence * 100).toFixed(1)}%)`, 'success');
    } catch (err) {
      showToast(err.message || 'Failed to analyze spectrum');
    } finally {
      spectralAnalyzeBtn.classList.remove('analyze-btn--loading');
      spectralAnalyzeBtn.disabled = false;
    }
  }

  // ── Event Bindings ─────────────────────────────────────────────
  setupDropzone(imageDropzone, imageInput, imageBrowse, handleImageFile);
  setupDropzone(spectralDropzone, spectralInput, spectralBrowse, handleSpectralFile);

  imageRemove.addEventListener('click', (e) => { e.stopPropagation(); clearImageFile(); });
  spectralRemove.addEventListener('click', (e) => { e.stopPropagation(); clearSpectralFile(); });

  imageAnalyzeBtn.addEventListener('click', analyzeImage);
  spectralAnalyzeBtn.addEventListener('click', analyzeSpectral);

  // ── Init ───────────────────────────────────────────────────────
  checkHealth();

})();
