(() => {
  const form = document.getElementById("run-form");
  const envInput = document.getElementById("env");
  const modeInput = document.getElementById("mode");
  const filesField = document.getElementById("files-field");
  const fileActions = document.getElementById("file-actions");
  const agencyId = document.getElementById("agency_id");
  const reviewOnly = document.getElementById("review-only");
  const refreshFilesBtn = document.getElementById("refresh-files");
  const fileFilter = document.getElementById("file-filter");
  const fileList = document.getElementById("file-list");
  const filesEmpty = document.getElementById("files-empty");
  const selectAllBtn = document.getElementById("select-all");
  const clearSelectionBtn = document.getElementById("clear-selection");
  const selectionCount = document.getElementById("selection-count");
  const runBtn = document.getElementById("run-btn");
  const stopBtn = document.getElementById("stop-btn");
  const statusLine = document.getElementById("status-line");
  const logEl = document.getElementById("log");
  const clearLogBtn = document.getElementById("clear-log");
  const refreshOutputsBtn = document.getElementById("refresh-outputs");
  const rootFiles = document.getElementById("root-files");
  const runsEl = document.getElementById("runs");

  const STORAGE_KEY = "forecast-ui-agency-id";

  let source = null;
  let autoScroll = true;
  let statementFiles = [];

  function setStatus(text, kind) {
    statusLine.textContent = text;
    statusLine.classList.remove("running", "ok", "fail");
    if (kind) statusLine.classList.add(kind);
  }

  function setRunning(running) {
    runBtn.disabled = running;
    stopBtn.disabled = !running;
  }

  function appendLog(line) {
    const atBottom =
      logEl.scrollTop + logEl.clientHeight >= logEl.scrollHeight - 24;
    logEl.textContent += (logEl.textContent ? "\n" : "") + line;
    if (autoScroll && atBottom) {
      logEl.scrollTop = logEl.scrollHeight;
    }
  }

  function wireSegments(selector, attr, input, onChange) {
    document.querySelectorAll(selector).forEach((btn) => {
      btn.addEventListener("click", () => {
        document.querySelectorAll(selector).forEach((b) => b.classList.remove("active"));
        btn.classList.add("active");
        input.value = btn.getAttribute(attr);
        if (onChange) onChange(input.value);
      });
    });
  }

  function updateMode(mode) {
    const pickFiles = mode === "files";
    filesField.classList.toggle("hidden", !pickFiles);
    fileActions.classList.toggle("hidden", !pickFiles);
  }

  function selectedUuids() {
    return Array.from(fileList.querySelectorAll('input[type="checkbox"]:checked')).map(
      (el) => el.value
    );
  }

  function updateSelectionCount() {
    const count = selectedUuids().length;
    selectionCount.textContent = `${count} selected`;
  }

  function visibleItems() {
    return Array.from(fileList.querySelectorAll(".file-item:not(.hidden)"));
  }

  function applyFilter() {
    const q = (fileFilter.value || "").trim().toLowerCase();
    fileList.querySelectorAll(".file-item").forEach((item) => {
      const name = (item.dataset.name || "").toLowerCase();
      item.classList.toggle("hidden", Boolean(q) && !name.includes(q));
    });
  }

  function renderFiles(files) {
    statementFiles = files || [];
    fileList.innerHTML = "";
    if (!statementFiles.length) {
      const note = document.createElement("p");
      note.className = "empty-note";
      note.id = "files-empty";
      note.textContent = "No statement files found for this agency.";
      fileList.appendChild(note);
      updateSelectionCount();
      return;
    }

    statementFiles.forEach((file) => {
      const label = document.createElement("label");
      label.className = "file-item";
      label.dataset.name = file.name || "";

      const checkbox = document.createElement("input");
      checkbox.type = "checkbox";
      checkbox.value = file.uuid;
      checkbox.addEventListener("change", updateSelectionCount);

      const body = document.createElement("span");
      const name = document.createElement("div");
      name.className = "file-name";
      name.textContent = file.name || "(unnamed)";

      const meta = document.createElement("div");
      meta.className = "file-meta";
      const pill = document.createElement("span");
      pill.className = "stage-pill";
      pill.textContent = file.stage_label || `Stage ${file.stage}`;
      meta.appendChild(pill);
      const when = file.created_at ? new Date(file.created_at).toLocaleString() : "";
      meta.appendChild(document.createTextNode(when ? ` · ${when}` : ""));

      body.appendChild(name);
      body.appendChild(meta);
      label.appendChild(checkbox);
      label.appendChild(body);
      fileList.appendChild(label);
    });

    applyFilter();
    updateSelectionCount();
  }

  async function loadFiles({ silent } = {}) {
    const agency = agencyId.value.trim();
    if (!agency) {
      if (!silent) setStatus("Enter an agency ID first", "fail");
      return;
    }
    localStorage.setItem(STORAGE_KEY, agency);
    refreshFilesBtn.disabled = true;
    if (!silent) setStatus("Loading files…");
    try {
      const params = new URLSearchParams({
        env: envInput.value,
        agency_id: agency,
        review_only: reviewOnly.checked ? "true" : "false",
      });
      const res = await fetch(`/api/statements?${params}`);
      const body = await res.json().catch(() => ({}));
      if (!res.ok) {
        renderFiles([]);
        setStatus(body.detail || "Could not load files", "fail");
        return;
      }
      renderFiles(body.files || []);
      setStatus(`Loaded ${body.count || 0} file(s)`, "ok");
    } catch (err) {
      setStatus(String(err), "fail");
    } finally {
      refreshFilesBtn.disabled = false;
    }
  }

  wireSegments("[data-env]", "data-env", envInput, () => {
    if (agencyId.value.trim()) loadFiles({ silent: true });
  });
  wireSegments("[data-mode]", "data-mode", modeInput, updateMode);
  updateMode(modeInput.value);

  const savedAgency = localStorage.getItem(STORAGE_KEY);
  if (savedAgency) agencyId.value = savedAgency;

  refreshFilesBtn.addEventListener("click", () => loadFiles());
  reviewOnly.addEventListener("change", () => {
    if (agencyId.value.trim()) loadFiles({ silent: true });
  });
  agencyId.addEventListener("keydown", (ev) => {
    if (ev.key === "Enter") {
      ev.preventDefault();
      loadFiles();
    }
  });
  fileFilter.addEventListener("input", applyFilter);
  selectAllBtn.addEventListener("click", () => {
    visibleItems().forEach((item) => {
      const box = item.querySelector('input[type="checkbox"]');
      if (box) box.checked = true;
    });
    updateSelectionCount();
  });
  clearSelectionBtn.addEventListener("click", () => {
    fileList.querySelectorAll('input[type="checkbox"]').forEach((box) => {
      box.checked = false;
    });
    updateSelectionCount();
  });

  function connectStream() {
    if (source) {
      source.close();
      source = null;
    }
    source = new EventSource("/api/run/stream");
    source.addEventListener("status", (ev) => {
      const data = JSON.parse(ev.data);
      setRunning(Boolean(data.running));
      if (data.running) setStatus("Forecast running…", "running");
      else if (data.exit_code === 0) setStatus("Last run finished successfully", "ok");
      else if (data.exit_code != null) setStatus(`Last run exited with code ${data.exit_code}`, "fail");
    });
    source.addEventListener("log", (ev) => {
      appendLog(JSON.parse(ev.data));
    });
    source.addEventListener("done", (ev) => {
      const data = JSON.parse(ev.data);
      setRunning(false);
      if (data.stopped) setStatus("Stopped", "fail");
      else if (data.exit_code === 0) setStatus("Finished successfully", "ok");
      else setStatus(`Finished with exit code ${data.exit_code}`, "fail");
      loadOutputs();
    });
    source.onerror = () => {};
  }

  form.addEventListener("submit", async (event) => {
    event.preventDefault();
    const agency = agencyId.value.trim();
    if (!agency) {
      setStatus("Enter an agency ID", "fail");
      return;
    }
    localStorage.setItem(STORAGE_KEY, agency);

    const payload = {
      env: envInput.value,
      mode: modeInput.value,
      agency_id: Number(agency),
    };

    if (payload.mode === "files") {
      const uuids = selectedUuids();
      if (!uuids.length) {
        setStatus("Select one or more files", "fail");
        return;
      }
      payload.uuids = uuids;
    }

    setRunning(true);
    setStatus("Starting…", "running");
    connectStream();

    try {
      const res = await fetch("/api/run", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(payload),
      });
      const body = await res.json().catch(() => ({}));
      if (!res.ok) {
        setRunning(false);
        const detail = Array.isArray(body.detail)
          ? body.detail.map((d) => d.msg || JSON.stringify(d)).join("; ")
          : body.detail;
        setStatus(detail || "Failed to start", "fail");
        appendLog(`[ui] ${detail || res.statusText}`);
        return;
      }
      setStatus("Forecast running…", "running");
    } catch (err) {
      setRunning(false);
      setStatus(String(err), "fail");
    }
  });

  stopBtn.addEventListener("click", async () => {
    stopBtn.disabled = true;
    try {
      await fetch("/api/run/stop", { method: "POST" });
    } catch (err) {
      appendLog(`[ui] stop failed: ${err}`);
      stopBtn.disabled = false;
    }
  });

  clearLogBtn.addEventListener("click", () => {
    logEl.textContent = "";
  });

  refreshOutputsBtn.addEventListener("click", () => loadOutputs());

  function fileHref(path) {
    return `/api/outputs/file?path=${encodeURIComponent(path)}`;
  }

  function formatTime(ts) {
    if (!ts) return "";
    return new Date(ts * 1000).toLocaleString();
  }

  async function loadOutputs() {
    try {
      const res = await fetch("/api/outputs");
      const data = await res.json();
      rootFiles.innerHTML = "";
      runsEl.innerHTML = "";

      if (data.root_files && data.root_files.length) {
        data.root_files.forEach((file) => {
          const a = document.createElement("a");
          a.className = "file-link";
          a.href = fileHref(file.path);
          a.textContent = file.name;
          a.setAttribute("download", file.name);
          rootFiles.appendChild(a);
        });
      }

      if (!data.runs || !data.runs.length) {
        const note = document.createElement("p");
        note.className = "empty-note";
        note.textContent = "No forecast runs yet. Outputs will appear here after a successful run.";
        runsEl.appendChild(note);
        return;
      }

      data.runs.forEach((run) => {
        const card = document.createElement("article");
        card.className = "run";

        const title = document.createElement("p");
        title.className = "run-title";
        title.textContent = run.folder;

        const meta = document.createElement("p");
        meta.className = "run-meta";
        const bits = [formatTime(run.mtime)];
        if (run.summary) {
          if (run.summary.exceptions != null) bits.push(`${run.summary.exceptions} exceptions`);
          if (run.summary.commissions != null) bits.push(`${run.summary.commissions} commissions`);
        }
        meta.textContent = bits.join(" · ");

        const files = document.createElement("div");
        files.className = "run-files";
        (run.files || []).forEach((file) => {
          const a = document.createElement("a");
          a.className = "run-file";
          a.href = fileHref(file.path);
          a.textContent = file.name;
          a.setAttribute("download", file.name.split("/").pop());
          files.appendChild(a);
        });

        card.appendChild(title);
        card.appendChild(meta);
        if ((run.files || []).length) card.appendChild(files);
        runsEl.appendChild(card);
      });
    } catch (err) {
      runsEl.innerHTML = `<p class="empty-note">Could not load outputs: ${err}</p>`;
    }
  }

  fetch("/api/status")
    .then((r) => r.json())
    .then((snap) => {
      setRunning(Boolean(snap.running));
      (snap.recent_lines || []).forEach(appendLog);
      if (snap.running) {
        setStatus("Forecast running…", "running");
        connectStream();
      } else {
        setStatus("Ready");
      }
    })
    .catch(() => setStatus("Ready"));

  loadOutputs();
  if (agencyId.value.trim()) {
    loadFiles({ silent: true });
  }
})();
