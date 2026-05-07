/**
 * DOCX batch generation: parse via API, queue client-side, write WAVs with File System Access API.
 * Depends on window.vctApi and window.vctShowLogin from app.js.
 */
(function () {
  const $ = (s) => document.querySelector(s);

  let batchDirHandle = null;
  let queueItems = [];
  let batchRunning = false;
  let batchPaused = false;
  let batchCancelRequested = false;
  let currentIndex = -1;

  function escapeHtml(s) {
    return (s || "").replace(/[&<>"']/g, (c) => ({
      "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;",
    }[c]));
  }

  function fsSafeSegment(name) {
    const t = (name || "").replace(/[<>:"/\\|?*\u0000-\u001f]/g, "_").trim();
    return t || "untitled";
  }

  async function fileExists(dirHandle, name) {
    try {
      await dirHandle.getFileHandle(name);
      return true;
    } catch {
      return false;
    }
  }

  async function uniqueFileName(dirHandle, baseName) {
    if (!(await fileExists(dirHandle, baseName))) return baseName;
    const dot = baseName.lastIndexOf(".");
    const stem = dot >= 0 ? baseName.slice(0, dot) : baseName;
    const ext = dot >= 0 ? baseName.slice(dot) : "";
    for (let i = 1; i < 999; i++) {
      const n = `${stem}_${i}${ext}`;
      if (!(await fileExists(dirHandle, n))) return n;
    }
    throw new Error("Could not find a unique filename");
  }

  /**
   * Create or open a subfolder. If a *file* already uses this name, try name_1, name_2…
   * (TypeMismatchError from the File System Access API).
   */
  async function getOrCreateDirectorySafe(parentHandle, safeSegment) {
    if (!safeSegment) throw new Error("Invalid folder name");
    for (let n = 0; n < 999; n++) {
      const name = n === 0 ? safeSegment : `${safeSegment}_${n}`;
      try {
        return await parentHandle.getDirectoryHandle(name, { create: true });
      } catch (err) {
        if (err && err.name === "TypeMismatchError") continue;
        throw err;
      }
    }
    throw new Error(`Could not create folder (blocked by a file?): ${safeSegment}`);
  }

  function pathPartsForEntry(entry) {
    const ext = String(entry.audio_extension || "wav").replace(/^\./, "");
    const fileSeg = `${entry.voice_label}.${ext}`;
    const g = entry.group != null && String(entry.group).trim();
    // Always use group + voice when the parser set a group (Format B), so subfolders are never skipped.
    if (g) {
      return [String(entry.group).trim(), fileSeg];
    }
    if (Array.isArray(entry.path_parts) && entry.path_parts.length > 0) {
      return entry.path_parts.map((p) => String(p));
    }
    const rp = String(entry.relative_path || "").replace(/\\/g, "/");
    return rp.split("/").filter((p) => p.length > 0);
  }

  async function writeWavToBatchRoot(entry, blob) {
    if (!batchDirHandle) throw new Error("No output folder selected");
    const rawParts = pathPartsForEntry(entry);
    if (!rawParts.length) throw new Error("Missing output path for entry");
    const safeParts = rawParts.map(fsSafeSegment);
    const fileName = safeParts[safeParts.length - 1];
    const dirParts = safeParts.slice(0, -1);
    let dh = batchDirHandle;
    for (const seg of dirParts) {
      dh = await getOrCreateDirectorySafe(dh, seg);
    }
    const unique = await uniqueFileName(dh, fileName);
    const fh = await dh.getFileHandle(unique, { create: true });
    const w = await fh.createWritable();
    await w.write(blob);
    await w.close();
    const rel = dirParts.length ? `${dirParts.join("/")}/${unique}` : unique;
    return rel;
  }

  async function fetchParseDocx(file) {
    const fd = new FormData();
    fd.append("file", file);
    const key = localStorage.getItem("vct_api_key") || "";
    const res = await fetch("/api/batch/parse-docx", {
      method: "POST",
      body: fd,
      headers: key ? { "X-API-Key": key } : {},
    });
    if (res.status === 401) {
      if (window.vctShowLogin) window.vctShowLogin();
      throw new Error("Unauthorized");
    }
    if (!res.ok) {
      let msg = `HTTP ${res.status}`;
      try {
        const j = await res.json();
        if (typeof j.detail === "string") msg = j.detail;
        else if (Array.isArray(j.detail)) msg = j.detail.map((x) => x.msg || JSON.stringify(x)).join("; ");
        else if (j.detail) msg = String(j.detail);
      } catch { /* ignore */ }
      throw new Error(msg);
    }
    return res.json();
  }

  async function generateOneWav(voiceId, text, speed) {
    const fd = new FormData();
    fd.append("voice_id", voiceId);
    fd.append("text", text);
    fd.append("speed", String(speed));
    const key = localStorage.getItem("vct_api_key") || "";
    const res = await fetch("/api/generate", {
      method: "POST",
      body: fd,
      headers: key ? { "X-API-Key": key } : {},
    });
    if (res.status === 401) {
      if (window.vctShowLogin) window.vctShowLogin();
      throw new Error("Unauthorized");
    }
    if (!res.ok) {
      let msg = `HTTP ${res.status}`;
      try {
        const j = await res.json();
        if (typeof j.detail === "string") msg = j.detail;
        else if (Array.isArray(j.detail)) msg = j.detail.map((x) => x.msg || JSON.stringify(x)).join("; ");
        else if (j.detail) msg = String(j.detail);
      } catch { /* ignore */ }
      throw new Error(msg);
    }
    return res.blob();
  }

  function renderQueue() {
    const el = $("#batchQueueBody");
    if (!el) return;
    if (!queueItems.length) {
      el.innerHTML = "<tr><td colspan=\"5\" class=\"batch-empty\">No entries yet. Upload a .docx to preview the queue.</td></tr>";
      return;
    }
    el.innerHTML = queueItems
      .map((it, idx) => {
        const prev = it.text.slice(0, 140) + (it.text.length > 140 ? "…" : "");
        const group = it.group ? escapeHtml(it.group) : "—";
        const st = it.status;
        const errHtml = it.error ? `<div class="batch-err">${escapeHtml(it.error)}</div>` : "";
        return `<tr data-idx="${idx}">
          <td>${it.order + 1}</td>
          <td>${group}</td>
          <td>${escapeHtml(it.voice_label)}</td>
          <td class="batch-preview">${escapeHtml(prev)}</td>
          <td><span class="batch-status batch-status--${st}">${st}</span>${errHtml}</td>
        </tr>`;
      })
      .join("");
  }

  function updateProgressPanel() {
    const total = queueItems.length;
    const ok = queueItems.filter((x) => x.status === "done").length;
    const bad = queueItems.filter((x) => x.status === "failed").length;
    const processed = ok + bad;
    const summary = $("#batchProgressSummary");
    const cur = $("#batchProgressCurrent");
    if (summary) {
      summary.textContent = total
        ? `${processed} of ${total} completed · ${ok} succeeded · ${bad} failed`
        : "";
    }
    if (cur) {
      if (currentIndex >= 0 && batchRunning && queueItems[currentIndex]) {
        const it = queueItems[currentIndex];
        const g = it.group ? `${it.group} · ` : "";
        cur.textContent = `Generating #${it.order + 1}: ${g}${it.voice_label}`;
      } else if (batchCancelRequested && batchRunning) {
        cur.textContent = "Stopping after current item…";
      } else {
        cur.textContent = "";
      }
    }
  }

  function setControlsEnabled() {
    const hasQueue = queueItems.length > 0;
    const canStart = hasQueue && batchDirHandle && !batchRunning;
    const startBtn = $("#batchBtnStart");
    const pauseBtn = $("#batchBtnPause");
    const resumeBtn = $("#batchBtnResume");
    const cancelBtn = $("#batchBtnCancel");
    const retryBtn = $("#batchBtnRetry");
    if (startBtn) startBtn.disabled = !canStart;
    if (pauseBtn) pauseBtn.disabled = !batchRunning || batchPaused;
    if (resumeBtn) resumeBtn.disabled = !batchRunning || !batchPaused;
    if (cancelBtn) cancelBtn.disabled = !batchRunning;
    if (retryBtn) retryBtn.disabled = batchRunning || !queueItems.some((x) => x.status === "failed");
  }

  function sleep(ms) {
    return new Promise((r) => setTimeout(r, ms));
  }

  async function waitWhilePaused() {
    while (batchPaused && !batchCancelRequested && batchRunning) {
      await sleep(120);
    }
  }

  async function runBatchLoop(retryFailedOnly) {
    if (batchRunning) return;

    const voiceSelect = $("#voiceSelect");
    const speedEl = $("#speed");
    const voiceId = voiceSelect && voiceSelect.value;
    const speed = speedEl ? parseFloat(speedEl.value) : 1;
    if (!voiceId) {
      alert("Select a voice in section 3.");
      return;
    }
    if (!batchDirHandle) {
      alert("Select an output folder first.");
      return;
    }

    let indices;
    if (retryFailedOnly) {
      indices = queueItems.map((it, i) => i).filter((i) => queueItems[i].status === "failed");
      indices.forEach((i) => {
        queueItems[i].status = "pending";
        queueItems[i].error = null;
      });
    } else {
      indices = queueItems
        .map((it, i) => i)
        .filter((i) => queueItems[i].status === "pending" || queueItems[i].status === "failed");
    }

    if (!indices.length) {
      $("#batchBatchMsg").textContent = retryFailedOnly
        ? "No failed entries to retry."
        : "Nothing to generate (all done or empty queue).";
      $("#batchBatchMsg").className = "msg err";
      return;
    }

    batchRunning = true;
    batchPaused = false;
    batchCancelRequested = false;
    $("#batchBatchMsg").textContent = "";
    $("#batchBatchMsg").className = "msg";
    setControlsEnabled();

    for (const idx of indices) {
      if (batchCancelRequested) break;
      await waitWhilePaused();
      if (batchCancelRequested) break;

      const it = queueItems[idx];
      if (it.status === "done") continue;

      currentIndex = idx;
      it.status = "generating";
      it.error = null;
      renderQueue();
      updateProgressPanel();
      setControlsEnabled();

      try {
        const blob = await generateOneWav(voiceId, it.text, speed);
        const savedPath = await writeWavToBatchRoot(it, blob);
        it.status = "done";
        it.savedAs = savedPath;
      } catch (e) {
        console.error(e);
        it.status = "failed";
        it.error = e.message || String(e);
      }

      renderQueue();
      updateProgressPanel();
    }

    currentIndex = -1;
    batchRunning = false;
    batchPaused = false;
    const cancelled = batchCancelRequested;
    batchCancelRequested = false;

    renderQueue();
    updateProgressPanel();
    setControlsEnabled();

    const ok = queueItems.filter((x) => x.status === "done").length;
    const bad = queueItems.filter((x) => x.status === "failed").length;
    const msgEl = $("#batchBatchMsg");
    if (msgEl) {
      if (cancelled) {
        msgEl.textContent = `Stopped. ${ok} succeeded, ${bad} failed so far.`;
        msgEl.className = bad ? "msg err" : "msg ok";
      } else {
        msgEl.textContent = `Batch finished. ${ok} succeeded, ${bad} failed.`;
        msgEl.className = bad ? "msg err" : "msg ok";
      }
    }
  }

  $("#batchPickFolder")?.addEventListener("click", async () => {
    if (!window.showDirectoryPicker) {
      const fp = $("#batchFolderPath");
      if (fp) {
        fp.textContent = "Folder picker requires Chrome or Edge (HTTPS or localhost).";
        fp.className = "batch-folder err";
      }
      return;
    }
    try {
      batchDirHandle = await window.showDirectoryPicker({ mode: "readwrite" });
      const fp = $("#batchFolderPath");
      if (fp) {
        fp.textContent = `Selected folder: ${batchDirHandle.name} (path not shown by browser for privacy)`;
        fp.className = "batch-folder ok";
      }
      setControlsEnabled();
    } catch (e) {
      if (e.name !== "AbortError") console.error(e);
    }
  });

  $("#batchDocxInput")?.addEventListener("change", async (e) => {
    const input = e.target;
    const f = input.files && input.files[0];
    if (!f) return;
    const pm = $("#batchParseMsg");
    if (pm) {
      pm.textContent = "Parsing…";
      pm.className = "msg";
    }
    try {
      const data = await fetchParseDocx(f);
      queueItems = data.entries.map((en) => ({
        ...en,
        audio_extension: data.audio_extension,
        status: "pending",
        error: null,
        savedAs: null,
      }));
      const fl = $("#batchFormatLabel");
      if (fl) {
        fl.textContent = `Format ${data.format} · ${data.entry_count} entries · output .${data.audio_extension}`;
      }
      renderQueue();
      updateProgressPanel();
      if (pm) {
        pm.textContent = "Parsed. Review the queue, choose output folder, then Start batch generate.";
        pm.className = "msg ok";
      }
    } catch (err) {
      if (pm) {
        pm.textContent = "Error: " + err.message;
        pm.className = "msg err";
      }
      queueItems = [];
      renderQueue();
      updateProgressPanel();
    }
    setControlsEnabled();
    input.value = "";
  });

  $("#batchBtnStart")?.addEventListener("click", () => {
    runBatchLoop(false);
  });
  $("#batchBtnRetry")?.addEventListener("click", () => {
    runBatchLoop(true);
  });

  $("#batchBtnPause")?.addEventListener("click", () => {
    if (!batchRunning) return;
    batchPaused = true;
    setControlsEnabled();
  });
  $("#batchBtnResume")?.addEventListener("click", () => {
    batchPaused = false;
    setControlsEnabled();
  });
  $("#batchBtnCancel")?.addEventListener("click", () => {
    if (!batchRunning) return;
    batchCancelRequested = true;
    batchPaused = false;
    setControlsEnabled();
  });

  renderQueue();
  updateProgressPanel();
  setControlsEnabled();
})();
