/* main.js — CreditFile multi-file frontend */
(function () {
  'use strict';

  // ── Refs ──────────────────────────────────────────────────
  const dropZone    = document.getElementById('drop-zone');
  const fileInput   = document.getElementById('file-input');
  const analyzeBtn  = document.getElementById('analyze-btn');
  const uploadError = document.getElementById('upload-error');
  const loadingEl   = document.getElementById('loading');
  const loadingText = document.getElementById('loading-text');
  const fileQueue   = document.getElementById('file-queue');
  const fileList    = document.getElementById('file-list');
  const queueCount  = document.getElementById('queue-count');
  const clearAllBtn = document.getElementById('clear-all-btn');

  const uploadSection  = document.getElementById('upload-section');
  const resultsSection = document.getElementById('results-section');
  const resultsList    = document.getElementById('results-list');
  const resultsCount   = document.getElementById('results-count');
  const downloadAllBtn = document.getElementById('download-all-btn');
  const newBtn         = document.getElementById('new-btn');

  // files waiting to be analyzed: Map<id, File>
  let fileMap    = new Map();
  let nextId     = 0;
  let allResults = [];   // array of API result objects

  // ── Visibility helpers ────────────────────────────────────
  const show = el => { el.style.display = ''; };
  const hide = el => { el.style.display = 'none'; };

  // ── Add files to queue ────────────────────────────────────
  function addFiles(files) {
    let rejected = 0;
    for (const f of files) {
      if (!f.name.match(/\.xlsx?$/i)) { rejected++; continue; }
      // skip exact duplicates (same name + size)
      const isDupe = [...fileMap.values()].some(
        x => x.name === f.name && x.size === f.size
      );
      if (!isDupe) fileMap.set(nextId++, f);
    }
    if (rejected) showError(`${rejected} file(s) skipped — only .xlsx / .xls allowed.`);
    else clearError();
    renderQueue();
  }

  function removeFile(id) {
    fileMap.delete(id);
    renderQueue();
  }

  function clearAll() {
    fileMap.clear();
    fileInput.value = '';
    renderQueue();
  }

  function renderQueue() {
    fileList.innerHTML = '';
    if (fileMap.size === 0) {
      hide(fileQueue);
      analyzeBtn.disabled = true;
      return;
    }
    show(fileQueue);
    analyzeBtn.disabled = false;
    queueCount.textContent = `${fileMap.size} file${fileMap.size > 1 ? 's' : ''} selected`;

    for (const [id, f] of fileMap) {
      const li = document.createElement('li');
      li.className = 'file-item';
      li.innerHTML = `
        <span class="file-item-icon">&#128196;</span>
        <span class="file-item-name" title="${esc(f.name)}">${esc(f.name)}</span>
        <span class="file-item-size">${fmtSize(f.size)}</span>
        <button class="file-item-remove" aria-label="Remove ${esc(f.name)}">&times;</button>
      `;
      li.querySelector('.file-item-remove').addEventListener('click', () => removeFile(id));
      fileList.appendChild(li);
    }
  }

  // ── Drop zone events ──────────────────────────────────────
  dropZone.addEventListener('click', e => {
    if (e.target.closest('label')) return;
    fileInput.click();
  });
  dropZone.addEventListener('keydown', e => {
    if (e.key === 'Enter' || e.key === ' ') fileInput.click();
  });
  fileInput.addEventListener('change', () => {
    if (fileInput.files.length) addFiles(Array.from(fileInput.files));
    fileInput.value = '';   // reset so same file can be re-added later if needed
  });
  dropZone.addEventListener('dragover',  e => { e.preventDefault(); dropZone.classList.add('drag-over'); });
  dropZone.addEventListener('dragleave', ()  => dropZone.classList.remove('drag-over'));
  dropZone.addEventListener('drop', e => {
    e.preventDefault();
    dropZone.classList.remove('drag-over');
    if (e.dataTransfer.files.length) addFiles(Array.from(e.dataTransfer.files));
  });
  clearAllBtn.addEventListener('click', clearAll);

  // ── Analyze all ───────────────────────────────────────────
  analyzeBtn.addEventListener('click', async () => {
    if (fileMap.size === 0) return;
    clearError();
    setLoading(true);
    allResults = [];

    const files  = [...fileMap.values()];
    const total  = files.length;

    for (let i = 0; i < total; i++) {
      const f = files[i];
      loadingText.textContent = `Processing file ${i + 1} of ${total}: ${f.name}`;
      const form = new FormData();
      form.append('file', f);
      try {
        const res  = await fetch('/analyze', { method: 'POST', body: form });
        const data = await res.json();
        if (!res.ok || data.error) {
          allResults.push({ filename: f.name, error: data.error || 'Processing failed.' });
        } else {
          allResults.push(data);
        }
      } catch {
        allResults.push({ filename: f.name, error: 'Could not reach the server.' });
      }
    }

    setLoading(false);
    renderResults();
  });

  // ── Render all results ────────────────────────────────────
  function renderResults() {
    resultsList.innerHTML = '';
    resultsCount.textContent = `(${allResults.length} file${allResults.length !== 1 ? 's' : ''})`;

    allResults.forEach((data, idx) => {
      resultsList.appendChild(buildResultCard(data, idx));
    });

    hide(uploadSection);
    show(resultsSection);
    window.scrollTo({ top: 0, behavior: 'smooth' });

    // Trigger ring animations after a tick so transitions fire
    requestAnimationFrame(() => {
      document.querySelectorAll('.ring-prog[data-target]').forEach(ring => {
        ring.style.strokeDashoffset = ring.dataset.target;
      });
    });
  }

  function buildResultCard(data, idx) {
    const wrap = document.createElement('div');
    wrap.className = 'result-card';

    if (data.error) {
      // Error state
      wrap.innerHTML = `
        <div class="result-card-header" style="cursor:default">
          <div class="rch-score grade-na">!</div>
          <div class="rch-info">
            <div class="rch-name">${esc(data.filename)}</div>
            <div class="rch-grade">Failed to process</div>
          </div>
          <span class="rch-status err">Error</span>
        </div>
        <div class="result-error">${esc(data.error)}</div>
      `;
      return wrap;
    }

    const score = data.credit_score;
    const valid = data.score_valid;
    const { label, cls } = valid && score !== null ? gradeFor(score) : { label: 'N/A', cls: 'grade-na' };
    const displayScore   = valid && score !== null ? score : 'N/A';
    const statusText     = valid ? 'Scored' : 'Incomplete';
    const statusCls      = valid ? 'ok' : 'warn';
    const circ           = 263.9;
    const offset         = valid && score !== null ? circ - (score / 100) * circ : circ;

    // Card header (click to collapse/expand)
    const header = document.createElement('button');
    header.className = 'result-card-header';
    header.setAttribute('aria-expanded', 'true');
    header.innerHTML = `
      <div class="rch-score ${cls}">${displayScore}</div>
      <div class="rch-info">
        <div class="rch-name">${esc(data.filename)}</div>
        <div class="rch-grade">${label}${valid ? ` &mdash; ${score}/100` : ''}</div>
      </div>
      <span class="rch-status ${statusCls}">${statusText}</span>
      <span class="rch-caret">&#9660;</span>
    `;

    const body = document.createElement('div');
    body.className = 'result-card-body';

    header.addEventListener('click', () => {
      const open = header.getAttribute('aria-expanded') === 'true';
      header.setAttribute('aria-expanded', String(!open));
      body.style.display = open ? 'none' : '';
    });

    // Score ring
    const scoreBox = document.createElement('div');
    scoreBox.className = 'score-box';
    scoreBox.innerHTML = `
      <div class="score-ring">
        <svg viewBox="0 0 100 100" aria-hidden="true">
          <circle class="ring-bg"   cx="50" cy="50" r="42"/>
          <circle class="ring-prog ${cls.replace('grade-','ring-')}"
                  cx="50" cy="50" r="42"
                  stroke-dasharray="${circ}"
                  stroke-dashoffset="${circ}"
                  data-target="${offset}"/>
        </svg>
        <div class="score-num">${displayScore}</div>
      </div>
      <div class="score-info">
        <div class="score-label">Credit Score</div>
        <div class="score-grade ${cls}">${label}</div>
        <div class="score-note">${
          valid
            ? `${score}/100 &mdash; ${data.missing_count} missing feature(s)`
            : `Too many missing fields (${data.missing_count})`
        }</div>
      </div>
    `;
    body.appendChild(scoreBox);

    // Missing fields
    const missingBlock = document.createElement('div');
    missingBlock.className = 'block';
    missingBlock.innerHTML = `
      <h3 class="block-title">
        <span class="block-icon warn">&#9888;</span> Missing Fields
      </h3>
    `;
    const missingBody = document.createElement('div');
    renderMissingInto(missingBody, data.missing_fields);
    missingBlock.appendChild(missingBody);
    body.appendChild(missingBlock);

    // Validated data
    const dataBlock = document.createElement('div');
    dataBlock.className = 'block';
    dataBlock.innerHTML = `
      <h3 class="block-title">
        <span class="block-icon ok">&#10003;</span> Validated Data
      </h3>
    `;
    const dataTree = document.createElement('div');
    dataTree.className = 'data-tree';
    renderDataTreeInto(dataTree, data.normalized);
    dataBlock.appendChild(dataTree);
    body.appendChild(dataBlock);

    // Per-file download
    const actionsDiv = document.createElement('div');
    actionsDiv.className = 'actions-bottom';
    const dlBtn = document.createElement('button');
    dlBtn.className = 'btn btn-outline btn-sm';
    dlBtn.textContent = '↓ Download JSON';
    dlBtn.addEventListener('click', () => downloadSingle(data));
    actionsDiv.appendChild(dlBtn);
    body.appendChild(actionsDiv);

    wrap.appendChild(header);
    wrap.appendChild(body);
    return wrap;
  }

  function gradeFor(score) {
    if (score >= 80) return { label: 'Excellent', cls: 'grade-excellent' };
    if (score >= 60) return { label: 'Good',      cls: 'grade-good' };
    if (score >= 40) return { label: 'Fair',      cls: 'grade-fair' };
    return              { label: 'Poor',      cls: 'grade-poor' };
  }

  // ── Missing fields ────────────────────────────────────────
  function renderMissingInto(container, missing) {
    container.innerHTML = '';
    if (!missing || Object.keys(missing).length === 0) {
      container.innerHTML = '<p class="missing-none">&#10003; No missing fields</p>';
      return;
    }
    for (const [section, fields] of Object.entries(missing)) {
      const group = document.createElement('div');
      group.className = 'missing-group';
      const title = document.createElement('div');
      title.className   = 'missing-group-title';
      title.textContent = section.replace(/_/g, ' ');
      group.appendChild(title);
      const tags = document.createElement('div');
      tags.className = 'missing-tags';
      flattenMissing(fields).forEach(f => {
        const tag = document.createElement('span');
        tag.className   = 'missing-tag';
        tag.textContent = f.replace(/_/g, ' ');
        tags.appendChild(tag);
      });
      group.appendChild(tags);
      container.appendChild(group);
    }
  }

  function flattenMissing(val, prefix = '') {
    if (Array.isArray(val))
      return val.map(v => prefix ? `${prefix} › ${v}` : v);
    if (val && typeof val === 'object')
      return Object.entries(val).flatMap(([k, v]) =>
        flattenMissing(v, prefix ? `${prefix} › ${k}` : k));
    return [String(val)];
  }

  // ── Data tree ─────────────────────────────────────────────
  function renderDataTreeInto(container, normalized) {
    for (const [section, value] of Object.entries(normalized)) {
      if (section === 'filename' || section === 'last_modified') continue;
      if (value === null || value === undefined) continue;

      const wrap = document.createElement('div');
      wrap.className = 'tree-section';

      const hdr = document.createElement('button');
      hdr.className = 'tree-section-header';
      hdr.setAttribute('aria-expanded', 'false');   // collapsed by default
      hdr.innerHTML =
        `<span>${humanize(section)}</span><span class="tree-caret">&#9660;</span>`;

      const bdy = document.createElement('div');
      bdy.style.display = 'none';   // collapsed by default

      hdr.addEventListener('click', () => {
        const open = hdr.getAttribute('aria-expanded') === 'true';
        hdr.setAttribute('aria-expanded', String(!open));
        bdy.style.display = open ? 'none' : '';
      });

      const table = document.createElement('table');
      table.className = 'tree-table';
      buildRows(value).forEach(([key, val]) => {
        const tr  = document.createElement('tr');
        const tdK = document.createElement('td');
        const tdV = document.createElement('td');
        tdK.className   = 'td-key';
        tdV.className   = 'td-val';
        tdK.textContent = humanize(key);
        if (val === null || val === undefined || val === '') {
          tdV.innerHTML = '<span class="td-null">null</span>';
        } else {
          tdV.textContent = formatVal(val);
        }
        tr.appendChild(tdK); tr.appendChild(tdV); table.appendChild(tr);
      });

      bdy.appendChild(table);
      wrap.appendChild(hdr);
      wrap.appendChild(bdy);
      container.appendChild(wrap);
    }
  }

  function buildRows(obj, prefix = '') {
    if (obj === null || typeof obj !== 'object') return [[prefix, obj]];
    if (Array.isArray(obj)) return [[prefix, obj.join(', ')]];
    const rows = [];
    for (const [k, v] of Object.entries(obj)) {
      const key = prefix ? `${prefix} › ${k}` : k;
      if (v && typeof v === 'object' && !Array.isArray(v)) rows.push(...buildRows(v, key));
      else rows.push([key, v]);
    }
    return rows;
  }

  // ── Download helpers ──────────────────────────────────────
  function downloadSingle(data) {
    const n = data.normalized;
    // Canonical key order matching the expected JSON format
    const payload = {
      filename:              n.filename,
      last_modified:         n.last_modified,
      personal_data:         n.personal_data         || {},
      income_source_details: n.income_source_details || {},
      income_analysis:       n.income_analysis       || {},
      officer_assessment:    n.officer_assessment    || {},
      credit_score:          data.credit_score,
    };
    triggerDownload(
      JSON.stringify(payload, null, 4),
      (data.filename || 'credit_report').replace(/\.xlsx?$/i, '') + '.json'
    );
  }

  downloadAllBtn.addEventListener('click', async () => {
    const successful = allResults.filter(r => !r.error);
    if (successful.length === 0) return;

    if (successful.length === 1) {
      downloadSingle(successful[0]);
      return;
    }

    // Build a zip using JSZip (loaded from CDN inline) — if unavailable,
    // fall back to downloading files one by one.
    if (typeof JSZip !== 'undefined') {
      const zip = new JSZip();
      successful.forEach(data => {
        const payload = Object.assign({}, data.normalized, { credit_score: data.credit_score });
        const name = (data.filename || 'credit_report').replace(/\.xlsx?$/i, '') + '.json';
        zip.file(name, JSON.stringify(payload, null, 4));
      });
      const blob = await zip.generateAsync({ type: 'blob' });
      triggerDownloadBlob(blob, 'credit_results.zip');
    } else {
      // Fallback: download each file individually with a small delay
      for (let i = 0; i < successful.length; i++) {
        setTimeout(() => downloadSingle(successful[i]), i * 300);
      }
    }
  });

  function triggerDownload(text, filename) {
    triggerDownloadBlob(new Blob([text], { type: 'application/json' }), filename);
  }

  function triggerDownloadBlob(blob, filename) {
    const url = URL.createObjectURL(blob);
    const a   = document.createElement('a');
    a.href = url; a.download = filename;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  }

  // ── New batch ─────────────────────────────────────────────
  newBtn.addEventListener('click', () => {
    allResults = [];
    clearAll();
    hide(resultsSection);
    show(uploadSection);
  });

  // ── Utility ───────────────────────────────────────────────
  function setLoading(on) {
    if (on) show(loadingEl); else hide(loadingEl);
    analyzeBtn.disabled = on;
  }

  function showError(msg) {
    uploadError.textContent = '⚠ ' + msg;
    show(uploadError);
  }

  function clearError() {
    hide(uploadError);
    uploadError.textContent = '';
  }

  function humanize(str) {
    return String(str)
      .replace(/__/g, ' › ')
      .replace(/_/g, ' ')
      .replace(/\b\w/g, c => c.toUpperCase());
  }

  function formatVal(val) {
    if (Array.isArray(val)) return val.join(', ');
    if (typeof val === 'number') return val.toLocaleString();
    return String(val);
  }

  function fmtSize(bytes) {
    if (bytes < 1024) return bytes + ' B';
    if (bytes < 1024 * 1024) return (bytes / 1024).toFixed(1) + ' KB';
    return (bytes / (1024 * 1024)).toFixed(1) + ' MB';
  }

  function esc(str) {
    return String(str)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;');
  }

})();
