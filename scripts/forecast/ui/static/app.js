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

  const payablesForm = document.getElementById("payables-form");
  const payablesApl = document.getElementById("payables-apl");
  const payablesIcm = document.getElementById("payables-icm");
  const payablesAplName = document.getElementById("payables-apl-name");
  const payablesIcmName = document.getElementById("payables-icm-name");
  const payablesMode = document.getElementById("payables-mode");
  const payablesTolerance = document.getElementById("payables-tolerance");
  const payablesRunBtn = document.getElementById("payables-run-btn");
  const payablesStopBtn = document.getElementById("payables-stop-btn");
  const payablesStatusLine = document.getElementById("payables-status-line");
  const payablesLogEl = document.getElementById("payables-log");
  const payablesClearLogBtn = document.getElementById("payables-clear-log");
  const payablesRefreshOutputsBtn = document.getElementById("payables-refresh-outputs");
  const payablesRunsEl = document.getElementById("payables-runs");

  const STORAGE_KEY = "forecast-ui-agency-id";
  const PAYABLES_STORAGE_KEY = "forecast-ui-payables";

  let source = null;
  let payablesSource = null;
  let autoScroll = true;
  let statementFiles = [];

  function setStatus(text, kind) {
    statusLine.textContent = text;
    statusLine.classList.remove("running", "ok", "fail");
    if (kind) statusLine.classList.add(kind);
  }

  function setPayablesStatus(text, kind) {
    payablesStatusLine.textContent = text;
    payablesStatusLine.classList.remove("running", "ok", "fail");
    if (kind) payablesStatusLine.classList.add(kind);
  }

  function setRunning(running) {
    runBtn.disabled = running;
    stopBtn.disabled = !running;
  }

  function setPayablesRunning(running) {
    payablesRunBtn.disabled = running;
    payablesStopBtn.disabled = !running;
  }

  function appendToLog(el, line) {
    const atBottom = el.scrollTop + el.clientHeight >= el.scrollHeight - 24;
    el.textContent += (el.textContent ? "\n" : "") + line;
    if (autoScroll && atBottom) {
      el.scrollTop = el.scrollHeight;
    }
  }

  function appendLog(line) {
    appendToLog(logEl, line);
  }

  function appendPayablesLog(line) {
    appendToLog(payablesLogEl, line);
  }

  function wireSegments(selector, attr, input, onChange) {
    document.querySelectorAll(selector).forEach((btn) => {
      btn.addEventListener("click", () => {
        const group = btn.parentElement;
        if (group) {
          group.querySelectorAll(".seg").forEach((b) => b.classList.remove("active"));
        }
        btn.classList.add("active");
        input.value = btn.getAttribute(attr);
        if (onChange) onChange(input.value);
      });
    });
  }

  function switchTab(tab) {
    document.querySelectorAll(".tab-btn").forEach((btn) => {
      btn.classList.toggle("active", btn.dataset.tab === tab);
    });
    document.querySelectorAll(".tab-panel").forEach((panel) => {
      const active = panel.id === `tab-${tab}`;
      panel.classList.toggle("active", active);
      panel.hidden = !active;
    });
    if (tab === "payables") {
      loadPayablesOutputs();
    }
  }

  document.querySelectorAll(".tab-btn").forEach((btn) => {
    btn.addEventListener("click", () => switchTab(btn.dataset.tab));
  });

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
  wireSegments("[data-payables-mode]", "data-payables-mode", payablesMode);
  updateMode(modeInput.value);

  const savedAgency = localStorage.getItem(STORAGE_KEY);
  if (savedAgency) agencyId.value = savedAgency;

  try {
    const savedPayables = JSON.parse(localStorage.getItem(PAYABLES_STORAGE_KEY) || "{}");
    if (savedPayables.mode === "apl" || savedPayables.mode === "icm") {
      payablesMode.value = savedPayables.mode;
      document.querySelectorAll("[data-payables-mode]").forEach((btn) => {
        btn.classList.toggle("active", btn.getAttribute("data-payables-mode") === savedPayables.mode);
      });
    }
    if (savedPayables.tolerance != null && savedPayables.tolerance !== "") {
      payablesTolerance.value = savedPayables.tolerance;
    }
  } catch (_) {
    /* ignore */
  }

  function persistPayablesForm() {
    localStorage.setItem(
      PAYABLES_STORAGE_KEY,
      JSON.stringify({
        mode: payablesMode.value,
        tolerance: payablesTolerance.value,
      })
    );
  }

  function updateChosenName(input, labelEl) {
    const file = input.files && input.files[0];
    labelEl.textContent = file ? file.name : "No file selected";
  }

  payablesApl.addEventListener("change", () => updateChosenName(payablesApl, payablesAplName));
  payablesIcm.addEventListener("change", () => updateChosenName(payablesIcm, payablesIcmName));
  updateChosenName(payablesApl, payablesAplName);
  updateChosenName(payablesIcm, payablesIcmName);

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

  function connectPayablesStream() {
    if (payablesSource) {
      payablesSource.close();
      payablesSource = null;
    }
    payablesSource = new EventSource("/api/payables/stream");
    payablesSource.addEventListener("status", (ev) => {
      const data = JSON.parse(ev.data);
      setPayablesRunning(Boolean(data.running));
      if (data.running) setPayablesStatus("Comparison running…", "running");
      else if (data.exit_code === 0) setPayablesStatus("Last comparison finished successfully", "ok");
      else if (data.exit_code != null) {
        setPayablesStatus(`Last comparison exited with code ${data.exit_code}`, "fail");
      }
    });
    payablesSource.addEventListener("log", (ev) => {
      appendPayablesLog(JSON.parse(ev.data));
    });
    payablesSource.addEventListener("done", (ev) => {
      const data = JSON.parse(ev.data);
      setPayablesRunning(false);
      if (data.stopped) setPayablesStatus("Stopped", "fail");
      else if (data.exit_code === 0) setPayablesStatus("Finished successfully", "ok");
      else setPayablesStatus(`Finished with exit code ${data.exit_code}`, "fail");
      loadPayablesOutputs();
    });
    payablesSource.onerror = () => {};
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

  payablesForm.addEventListener("submit", async (event) => {
    event.preventDefault();
    const aplFile = payablesApl.files && payablesApl.files[0];
    const icmFile = payablesIcm.files && payablesIcm.files[0];
    if (!aplFile || !icmFile) {
      setPayablesStatus("Select both APL and ICM files", "fail");
      return;
    }
    const tolerance = Number(payablesTolerance.value);
    if (Number.isNaN(tolerance) || tolerance < 0) {
      setPayablesStatus("Tolerance must be a non-negative number", "fail");
      return;
    }

    persistPayablesForm();
    setPayablesRunning(true);
    setPayablesStatus("Starting…", "running");
    connectPayablesStream();

    const formData = new FormData();
    formData.append("apl_file", aplFile, aplFile.name);
    formData.append("icm_file", icmFile, icmFile.name);
    formData.append("mode", payablesMode.value);
    formData.append("tolerance", String(tolerance));

    try {
      const res = await fetch("/api/payables/run", {
        method: "POST",
        body: formData,
      });
      const body = await res.json().catch(() => ({}));
      if (!res.ok) {
        setPayablesRunning(false);
        const detail = Array.isArray(body.detail)
          ? body.detail.map((d) => d.msg || JSON.stringify(d)).join("; ")
          : body.detail;
        setPayablesStatus(detail || "Failed to start", "fail");
        appendPayablesLog(`[payables] ${detail || res.statusText}`);
        return;
      }
      setPayablesStatus("Comparison running…", "running");
    } catch (err) {
      setPayablesRunning(false);
      setPayablesStatus(String(err), "fail");
    }
  });

  payablesStopBtn.addEventListener("click", async () => {
    payablesStopBtn.disabled = true;
    try {
      await fetch("/api/payables/stop", { method: "POST" });
    } catch (err) {
      appendPayablesLog(`[payables] stop failed: ${err}`);
      payablesStopBtn.disabled = false;
    }
  });

  payablesClearLogBtn.addEventListener("click", () => {
    payablesLogEl.textContent = "";
  });

  payablesRefreshOutputsBtn.addEventListener("click", () => loadPayablesOutputs());

  function fileHref(path) {
    return `/api/outputs/file?path=${encodeURIComponent(path)}`;
  }

  function formatTime(ts) {
    if (!ts) return "";
    return new Date(ts * 1000).toLocaleString();
  }

  function renderRunCards(container, runs, emptyMessage) {
    container.innerHTML = "";
    if (!runs || !runs.length) {
      const note = document.createElement("p");
      note.className = "empty-note";
      note.textContent = emptyMessage;
      container.appendChild(note);
      return;
    }

    runs.forEach((run) => {
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
      container.appendChild(card);
    });
  }

  async function loadOutputs() {
    try {
      const res = await fetch("/api/outputs");
      const data = await res.json();
      rootFiles.innerHTML = "";

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

      renderRunCards(
        runsEl,
        data.runs || [],
        "No forecast runs yet. Outputs will appear here after a successful run."
      );
    } catch (err) {
      runsEl.innerHTML = `<p class="empty-note">Could not load outputs: ${err}</p>`;
    }
  }

  async function loadPayablesOutputs() {
    try {
      const res = await fetch("/api/payables/outputs");
      const data = await res.json();
      renderRunCards(
        payablesRunsEl,
        data.runs || [],
        "No payables comparisons yet. Run a comparison to see Grouped_Policy_Combined and segregated errors here."
      );
    } catch (err) {
      payablesRunsEl.innerHTML = `<p class="empty-note">Could not load payables outputs: ${err}</p>`;
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

  fetch("/api/payables/status")
    .then((r) => r.json())
    .then((snap) => {
      setPayablesRunning(Boolean(snap.running));
      (snap.recent_lines || []).forEach(appendPayablesLog);
      if (snap.running) {
        setPayablesStatus("Comparison running…", "running");
        connectPayablesStream();
      } else {
        setPayablesStatus("Ready");
      }
    })
    .catch(() => setPayablesStatus("Ready"));

  loadOutputs();
  loadPayablesOutputs();
  if (agencyId.value.trim()) {
    loadFiles({ silent: true });
  }
})();
