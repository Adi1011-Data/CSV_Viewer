(function () {
  "use strict";

  const dropzone = document.getElementById("dropzone");
  const fileInput = document.getElementById("fileInput");
  const workspace = document.getElementById("workspace");
  const tabBar = document.getElementById("tabBar");
  const tabPanels = document.getElementById("tabPanels");
  const toolbar = document.getElementById("toolbar");
  const newTabBtn = document.getElementById("newTabBtn");
  const fontSizeInput = document.getElementById("fontSize");
  const fontSizeValue = document.getElementById("fontSizeValue");
  const downloadBtn = document.getElementById("downloadBtn");
  const jlcFormatSelect = document.getElementById("jlcFormatSelect");
  const jlcDialog = document.getElementById("jlcDialog");
  const jlcDialogSummary = document.getElementById("jlcDialogSummary");
  const jlcDialogDownload = document.getElementById("jlcDialogDownload");
  const jlcDialogClose = document.getElementById("jlcDialogClose");

  const EDGE = 10;
  const DEFAULTS = { fontSize: 14, rowHeight: 36, colWidth: 120 };

  const tabs = new Map();
  let activeTabId = null;
  let tabIdCounter = 0;
  let activeResize = null;
  let hoverCell = null;
  let resizeCtx = null;

  function parseCSV(text) {
    const rows = [];
    let row = [];
    let field = "";
    let i = 0;
    let inQuotes = false;

    if (text.charCodeAt(0) === 0xfeff) text = text.slice(1);

    while (i < text.length) {
      const ch = text[i];
      if (inQuotes) {
        if (ch === '"') {
          if (text[i + 1] === '"') {
            field += '"';
            i += 2;
            continue;
          }
          inQuotes = false;
          i++;
          continue;
        }
        field += ch;
        i++;
        continue;
      }
      if (ch === '"') {
        inQuotes = true;
        i++;
        continue;
      }
      if (ch === ",") {
        row.push(field);
        field = "";
        i++;
        continue;
      }
      if (ch === "\r") {
        i++;
        if (text[i] === "\n") i++;
        row.push(field);
        rows.push(row);
        row = [];
        field = "";
        continue;
      }
      if (ch === "\n") {
        row.push(field);
        rows.push(row);
        row = [];
        field = "";
        i++;
        continue;
      }
      field += ch;
      i++;
    }

    if (field.length > 0 || row.length > 0) {
      row.push(field);
      rows.push(row);
    }

    while (rows.length > 0 && rows[rows.length - 1].every((c) => c === "")) {
      rows.pop();
    }

    return rows;
  }

  function escapeCSVField(value) {
    const s = String(value ?? "");
    if (/[",\r\n]/.test(s)) return '"' + s.replace(/"/g, '""') + '"';
    return s;
  }

  function uniqueFileName(name) {
    const base = name.endsWith(".csv") ? name : name + ".csv";
    const names = new Set([...tabs.values()].map((t) => t.fileName));
    if (!names.has(base)) return base;
    const stem = base.replace(/\.csv$/i, "");
    let n = 2;
    while (names.has(`${stem} (${n}).csv`)) n++;
    return `${stem} (${n}).csv`;
  }

  function getActiveTab() {
    return activeTabId != null ? tabs.get(activeTabId) : null;
  }

  function getTabData(tab) {
    const headerCells = tab.tableHead.querySelectorAll("tr:last-child th");
    const headers = [];
    headerCells.forEach((th, idx) => {
      if (idx === 0) return;
      headers.push(th.textContent.trim());
    });

    const dataRows = [];
    tab.tableBody.querySelectorAll("tr").forEach((tr) => {
      const row = [];
      tr.querySelectorAll("td").forEach((td, idx) => {
        if (idx === 0) return;
        row.push(td.textContent.trim());
      });
      dataRows.push(row);
    });

    return { headers, dataRows };
  }

  function tableToCSV(tab) {
    const rows = [];
    const headerCells = tab.tableHead.querySelectorAll("tr:last-child th");
    const headerRow = [];
    headerCells.forEach((th, idx) => {
      if (idx === 0) return;
      headerRow.push(th.textContent.trim());
    });
    if (headerRow.length) rows.push(headerRow);

    tab.tableBody.querySelectorAll("tr").forEach((tr) => {
      const row = [];
      tr.querySelectorAll("td").forEach((td, idx) => {
        if (idx === 0) return;
        row.push(td.textContent.trim());
      });
      rows.push(row);
    });

    return rows.map((r) => r.map(escapeCSVField).join(",")).join("\r\n");
  }

  function saveTabSettings(tab) {
    const prev = tab.settings || DEFAULTS;
    tab.settings = {
      fontSize: parseInt(fontSizeInput.value, 10),
      rowHeight: prev.rowHeight,
      colWidth: prev.colWidth,
    };
    tab.colWidths = [...tab.colWidths];
    tab.rowHeights = [...tab.rowHeights];
  }

  function loadTabSettings(tab) {
    const s = tab.settings || DEFAULTS;
    fontSizeInput.value = s.fontSize;
    applyGlobalSizesToTab(tab);
  }

  function applyGlobalSizesToTab(tab) {
    const s = tab.settings || DEFAULTS;
    const fontPx = s.fontSize + "px";
    tab.panel.style.setProperty("--table-font-size", fontPx);
    fontSizeInput.value = s.fontSize;
    fontSizeValue.textContent = fontPx;
  }

  function applyColWidth(tab, colIndex, widthPx) {
    const cols = tab.colGroup.querySelectorAll("col");
    if (cols[colIndex]) cols[colIndex].style.width = widthPx + "px";
    tab.colWidths[colIndex] = widthPx;
    tab.dataTable.querySelectorAll(
      `th:nth-child(${colIndex + 1}), td:nth-child(${colIndex + 1})`
    ).forEach((el) => {
      el.style.minWidth = widthPx + "px";
      el.style.width = widthPx + "px";
      el.style.maxWidth = widthPx + "px";
    });
  }

  function applyRowHeight(tab, rowIndex, heightPx) {
    const tr =
      rowIndex === 0
        ? tab.tableHead.querySelector("tr")
        : tab.tableBody.querySelectorAll("tr")[rowIndex - 1];
    if (!tr) return;
    tr.querySelectorAll("th, td").forEach((cell) => {
      cell.style.height = heightPx + "px";
      cell.style.minHeight = heightPx + "px";
    });
    tab.rowHeights[rowIndex] = heightPx;
  }

  function buildTable(tab, rows) {
    if (!rows.length) rows = [[""]];

    const colCount = Math.max(...rows.map((r) => r.length), 1);
    const defaultColW = tab.settings?.colWidth ?? DEFAULTS.colWidth;
    const defaultRowH = tab.settings?.rowHeight ?? DEFAULTS.rowHeight;

    tab.colGroup.innerHTML = "";
    tab.tableHead.innerHTML = "";
    tab.tableBody.innerHTML = "";
    tab.colWidths = [];
    tab.rowHeights = [];

    const indexCol = document.createElement("col");
    indexCol.style.width = "48px";
    tab.colGroup.appendChild(indexCol);

    for (let c = 0; c < colCount; c++) {
      const col = document.createElement("col");
      col.style.width = defaultColW + "px";
      tab.colGroup.appendChild(col);
      tab.colWidths[c + 1] = defaultColW;
    }

    const headerTr = document.createElement("tr");
    const cornerTh = document.createElement("th");
    cornerTh.className = "row-index";
    cornerTh.textContent = "#";
    headerTr.appendChild(cornerTh);

    const headers = rows[0] || [];
    for (let c = 0; c < colCount; c++) {
      const th = document.createElement("th");
      th.textContent = headers[c] ?? `Column ${c + 1}`;
      th.contentEditable = "true";
      th.spellcheck = false;
      headerTr.appendChild(th);
    }
    tab.tableHead.appendChild(headerTr);
    tab.rowHeights[0] = defaultRowH;
    applyRowHeight(tab, 0, defaultRowH);

    const dataRows = rows.length > 1 ? rows.slice(1) : [];
    dataRows.forEach((row, rIdx) => {
      const tr = document.createElement("tr");
      const rowNum = document.createElement("td");
      rowNum.className = "row-index";
      rowNum.textContent = String(rIdx + 1);
      tr.appendChild(rowNum);

      for (let c = 0; c < colCount; c++) {
        const td = document.createElement("td");
        td.textContent = row[c] ?? "";
        td.contentEditable = "true";
        td.spellcheck = false;
        tr.appendChild(td);
      }

      tab.tableBody.appendChild(tr);
      tab.rowHeights[rIdx + 1] = defaultRowH;
      applyRowHeight(tab, rIdx + 1, defaultRowH);
    });

    tab.statsEl.textContent = `${dataRows.length} rows × ${colCount} columns`;
  }

  function getCell(target) {
    return target instanceof Element ? target.closest("th, td") : null;
  }

  function getRowIndex(tab, cell) {
    const tr = cell.parentElement;
    if (tr.parentElement === tab.tableHead) return 0;
    return Array.from(tab.tableBody.rows).indexOf(tr) + 1;
  }

  function edgeAt(cell, clientX, clientY) {
    const rect = cell.getBoundingClientRect();
    const nearRight = clientX >= rect.right - EDGE;
    const nearBottom = clientY >= rect.bottom - EDGE;
    if (!nearRight && !nearBottom) return null;
    const distRight = rect.right - clientX;
    const distBottom = rect.bottom - clientY;
    if (nearRight && (!nearBottom || distRight <= distBottom)) return "col";
    if (nearBottom) return "row";
    return null;
  }

  function clearEdgeHighlight() {
    if (hoverCell) {
      hoverCell.classList.remove("is-resize-col", "is-resize-row");
      hoverCell = null;
    }
  }

  function positionColGuide(tab, x) {
    const wrapRect = tab.tableWrap.getBoundingClientRect();
    tab.colGuide.style.left = x - wrapRect.left + tab.tableWrap.scrollLeft + "px";
    tab.colGuide.style.top = tab.tableWrap.scrollTop + "px";
    tab.colGuide.style.height = tab.tableWrap.scrollHeight + "px";
    tab.colGuide.hidden = false;
  }

  function positionRowGuide(tab, y) {
    const wrapRect = tab.tableWrap.getBoundingClientRect();
    tab.rowGuide.style.top = y - wrapRect.top + tab.tableWrap.scrollTop + "px";
    tab.rowGuide.style.left = tab.tableWrap.scrollLeft + "px";
    tab.rowGuide.style.width = tab.tableWrap.scrollWidth + "px";
    tab.rowGuide.hidden = false;
  }

  function hideGuides(tab) {
    if (!tab) return;
    tab.colGuide.hidden = true;
    tab.rowGuide.hidden = true;
  }

  function startColResize(tab, colIndex, clientX) {
    resizeCtx = tab;
    const col = tab.colGroup.querySelectorAll("col")[colIndex];
    const startW =
      tab.colWidths[colIndex] ||
      parseInt(col?.style.width, 10) ||
      (tab.settings?.colWidth ?? DEFAULTS.colWidth);
    const cell = tab.dataTable.querySelector(
      `thead th:nth-child(${colIndex + 1}), tbody td:nth-child(${colIndex + 1})`
    );
    const startRight = cell ? cell.getBoundingClientRect().right : clientX;

    activeResize = { type: "col", colIndex, startX: clientX, startW, startRight };
    tab.dataTable.classList.add("is-resizing");
    document.body.style.userSelect = "none";
    positionColGuide(tab, startRight);
  }

  function startRowResize(tab, rowIndex, clientY) {
    resizeCtx = tab;
    const tr =
      rowIndex === 0
        ? tab.tableHead.querySelector("tr")
        : tab.tableBody.querySelectorAll("tr")[rowIndex - 1];
    const startH = tab.rowHeights[rowIndex] || tab.settings?.rowHeight || DEFAULTS.rowHeight;
    const startBottom = tr
      ? tr.querySelector("td, th").getBoundingClientRect().bottom
      : clientY;

    activeResize = { type: "row", rowIndex, startY: clientY, startH, startBottom };
    tab.dataTable.classList.add("is-resizing");
    document.body.style.userSelect = "none";
    positionRowGuide(tab, startBottom);
  }

  function onResizeMove(e) {
    if (!activeResize || !resizeCtx) return;
    const tab = resizeCtx;

    if (activeResize.type === "col") {
      const dx = e.clientX - activeResize.startX;
      const w = Math.max(40, Math.min(600, activeResize.startW + dx));
      applyColWidth(tab, activeResize.colIndex, w);
      if (!tab.settings) tab.settings = { ...DEFAULTS };
      tab.settings.colWidth = Math.round(w);
      positionColGuide(tab, activeResize.startRight + dx);
    } else {
      const dy = e.clientY - activeResize.startY;
      const h = Math.max(24, Math.min(200, activeResize.startH + dy));
      applyRowHeight(tab, activeResize.rowIndex, h);
      if (!tab.settings) tab.settings = { ...DEFAULTS };
      tab.settings.rowHeight = Math.round(h);
      positionRowGuide(tab, activeResize.startBottom + dy);
    }
  }

  function endResize() {
    if (!activeResize || !resizeCtx) return;
    hideGuides(resizeCtx);
    resizeCtx.dataTable.classList.remove("is-resizing");
    activeResize = null;
    resizeCtx = null;
    document.body.style.userSelect = "";
    document.removeEventListener("mousemove", onResizeMove);
    document.removeEventListener("mouseup", endResize);
  }

  function setupBorderResize(tab) {
    tab.dataTable.addEventListener("mousemove", (e) => {
      if (activeResize) return;
      if (tab.id !== activeTabId) return;

      const cell = getCell(e.target);
      clearEdgeHighlight();

      if (!cell || !tab.dataTable.contains(cell)) {
        tab.dataTable.style.cursor = "";
        return;
      }

      const edge = edgeAt(cell, e.clientX, e.clientY);
      if (edge === "col") {
        cell.classList.add("is-resize-col");
        hoverCell = cell;
        tab.dataTable.style.cursor = "col-resize";
      } else if (edge === "row") {
        cell.classList.add("is-resize-row");
        hoverCell = cell;
        tab.dataTable.style.cursor = "row-resize";
      } else {
        tab.dataTable.style.cursor = "";
      }
    });

    tab.dataTable.addEventListener("mouseleave", () => {
      if (!activeResize) {
        clearEdgeHighlight();
        tab.dataTable.style.cursor = "";
      }
    });

    tab.dataTable.addEventListener("mousedown", (e) => {
      if (tab.id !== activeTabId) return;
      const cell = getCell(e.target);
      if (!cell || !tab.dataTable.contains(cell)) return;

      const edge = edgeAt(cell, e.clientX, e.clientY);
      if (!edge) return;

      e.preventDefault();
      e.stopPropagation();

      if (edge === "col") startColResize(tab, cell.cellIndex, e.clientX);
      else startRowResize(tab, getRowIndex(tab, cell), e.clientY);

      document.addEventListener("mousemove", onResizeMove);
      document.addEventListener("mouseup", endResize);
    });
  }

  function createTabPanel(id, fileName) {
    const panel = document.createElement("div");
    panel.className = "tab-panel";
    panel.id = `panel-${id}`;
    panel.setAttribute("role", "tabpanel");
    panel.hidden = true;

    const info = document.createElement("div");
    info.className = "viewer__info";
    const statsEl = document.createElement("span");
    statsEl.className = "viewer__stats";
    info.appendChild(statsEl);

    const tableWrap = document.createElement("div");
    tableWrap.className = "table-wrap";

    const colGuide = document.createElement("div");
    colGuide.className = "resize-guide resize-guide--col";
    colGuide.hidden = true;
    colGuide.setAttribute("aria-hidden", "true");

    const rowGuide = document.createElement("div");
    rowGuide.className = "resize-guide resize-guide--row";
    rowGuide.hidden = true;
    rowGuide.setAttribute("aria-hidden", "true");

    const dataTable = document.createElement("table");
    dataTable.className = "data-table";
    const colGroup = document.createElement("colgroup");
    const tableHead = document.createElement("thead");
    const tableBody = document.createElement("tbody");
    dataTable.append(colGroup, tableHead, tableBody);

    tableWrap.append(colGuide, rowGuide, dataTable);

    const tip = document.createElement("p");
    tip.className = "viewer__tip";
    tip.textContent =
      "Drag a column’s right border or a row’s bottom border to resize. Click a cell to edit.";

    panel.append(info, tableWrap, tip);
    tabPanels.appendChild(panel);

    return {
      id,
      fileName,
      panel,
      statsEl,
      tableWrap,
      colGuide,
      rowGuide,
      dataTable,
      colGroup,
      tableHead,
      tableBody,
      colWidths: [],
      rowHeights: [],
      settings: { ...DEFAULTS },
    };
  }

  function createTabButton(tab) {
    const btn = document.createElement("button");
    btn.type = "button";
    btn.className = "tab";
    btn.setAttribute("role", "tab");
    btn.setAttribute("aria-selected", "false");
    btn.setAttribute("aria-controls", `panel-${tab.id}`);
    btn.dataset.tabId = String(tab.id);

    const label = document.createElement("span");
    label.className = "tab__label";
    label.textContent = tab.fileName;
    label.title = tab.fileName;

    const close = document.createElement("span");
    close.className = "tab__close";
    close.setAttribute("role", "button");
    close.setAttribute("aria-label", `Close ${tab.fileName}`);
    close.innerHTML = "&times;";

    btn.append(label, close);
    tabBar.insertBefore(btn, newTabBtn);
    tab.tabBtn = btn;

    btn.addEventListener("click", (e) => {
      if (e.target.closest(".tab__close")) return;
      switchTab(tab.id);
    });

    close.addEventListener("click", (e) => {
      e.stopPropagation();
      closeTab(tab.id);
    });
  }

  function switchTab(id) {
    const prev = getActiveTab();
    if (prev) saveTabSettings(prev);
    closeJlcDialog();

    activeTabId = id;
    tabs.forEach((tab) => {
      const active = tab.id === id;
      tab.panel.hidden = !active;
      tab.tabBtn.classList.toggle("is-active", active);
      tab.tabBtn.setAttribute("aria-selected", active ? "true" : "false");
      if (!active) hideGuides(tab);
    });

    const tab = tabs.get(id);
    if (tab) {
      loadTabSettings(tab);
      syncJlcToolbar(tab);
    }
  }

  function closeTab(id) {
    const tab = tabs.get(id);
    if (!tab) return;

    hideGuides(tab);
    tab.tabBtn.remove();
    tab.panel.remove();
    tabs.delete(id);

    if (tabs.size === 0) {
      showDropzone();
      return;
    }

    if (activeTabId === id) {
      const next = tabs.keys().next().value;
      switchTab(next);
    }
  }

  function addTab(fileName, rows) {
    const id = ++tabIdCounter;
    const name = uniqueFileName(fileName);
    const tab = createTabPanel(id, name);
    tab.fileName = name;
    tabs.set(id, tab);
    createTabButton(tab);
    buildTable(tab, rows);
    setupBorderResize(tab);

    if (tabs.size === 1) showWorkspace();
    switchTab(id);
    return tab;
  }

  function showWorkspace() {
    dropzone.hidden = true;
    workspace.hidden = false;
    toolbar.hidden = false;
  }

  function showDropzone() {
    activeTabId = null;
    tabBar.querySelectorAll(".tab").forEach((el) => el.remove());
    tabPanels.innerHTML = "";
    tabs.clear();
    endResize();
    closeJlcDialog();

    dropzone.hidden = false;
    workspace.hidden = true;
    toolbar.hidden = true;
    fileInput.value = "";
    resetJlcToolbar();
  }

  function resetJlcToolbar() {
    jlcFormatSelect.value = "";
  }

  function syncJlcToolbar(tab) {
    jlcFormatSelect.value = tab.jlcFormatId || "";
  }

  function closeJlcDialog() {
    if (jlcDialog.open) jlcDialog.close();
  }

  function openJlcDialog(formatId) {
    const format = JLCFormat?.FORMATS[formatId];
    if (!format) return;
    jlcDialogSummary.textContent =
      format.label + " — " + format.description + ". Columns are mapped automatically from your CSV headers.";
    if (!jlcDialog.open) jlcDialog.showModal();
  }

  function downloadJlcCsv() {
    const tab = getActiveTab();
    const formatId = tab?.jlcFormatId;
    if (!tab || !formatId || typeof JLCFormat === "undefined") return;

    const format = JLCFormat.FORMATS[formatId];
    const { headers, dataRows } = getTabData(tab);
    if (!dataRows.length) {
      alert("No data rows to convert.");
      return;
    }

    const mapping = JLCFormat.autoMap(headers, formatId);
    const missing = JLCFormat.validateMapping(formatId, mapping);
    if (missing.length) {
      alert(
        "Could not auto-detect columns for: " +
          missing.join(", ") +
          ". Rename your CSV headers to match common names (e.g. Designator, Footprint, Mid X)."
      );
      return;
    }

    try {
      const outRows = JLCFormat.convert(headers, dataRows, formatId, mapping);
      if (outRows.length <= 1) {
        alert("No rows produced. Check that your CSV has data in the expected columns.");
        return;
      }

      const csv = JLCFormat.rowsToCSV(outRows);
      const blob = new Blob([csv], { type: "text/csv;charset=utf-8" });
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      const base = tab.fileName.replace(/\.csv$/i, "");
      a.download = base + format.fileSuffix + ".csv";
      a.click();
      URL.revokeObjectURL(url);
      closeJlcDialog();
    } catch (err) {
      alert(err.message);
    }
  }

  function loadFile(file) {
    if (!file) return;
    const isCsv =
      file.name.toLowerCase().endsWith(".csv") ||
      file.type === "text/csv" ||
      file.type === "application/vnd.ms-excel" ||
      file.type === "";
    if (!isCsv) {
      alert(`"${file.name}" is not a CSV file.`);
      return;
    }

    const reader = new FileReader();
    reader.onload = () => {
      try {
        const rows = parseCSV(reader.result);
        addTab(file.name, rows);
      } catch (err) {
        alert(`Could not parse "${file.name}": ${err.message}`);
      }
    };
    reader.onerror = () => alert(`Failed to read "${file.name}".`);
    reader.readAsText(file);
  }

  function loadFiles(fileList) {
    Array.from(fileList).forEach(loadFile);
  }

  function openFilePicker() {
    fileInput.value = "";
    fileInput.click();
  }

  dropzone.addEventListener("click", openFilePicker);

  dropzone.addEventListener("dragover", (e) => {
    e.preventDefault();
    dropzone.classList.add("is-dragover");
  });

  dropzone.addEventListener("dragleave", () => {
    dropzone.classList.remove("is-dragover");
  });

  dropzone.addEventListener("drop", (e) => {
    e.preventDefault();
    dropzone.classList.remove("is-dragover");
    if (e.dataTransfer.files.length) loadFiles(e.dataTransfer.files);
  });

  workspace.addEventListener("dragover", (e) => {
    e.preventDefault();
    e.dataTransfer.dropEffect = "copy";
  });

  workspace.addEventListener("drop", (e) => {
    e.preventDefault();
    if (e.dataTransfer.files.length) loadFiles(e.dataTransfer.files);
  });

  fileInput.addEventListener("change", () => {
    if (fileInput.files.length) loadFiles(fileInput.files);
  });

  newTabBtn.addEventListener("click", openFilePicker);

  fontSizeInput.addEventListener("input", () => {
    const tab = getActiveTab();
    if (!tab) return;
    applyGlobalSizesToTab(tab);
    tab.settings.fontSize = parseInt(fontSizeInput.value, 10);
  });

  jlcFormatSelect.addEventListener("change", () => {
    const tab = getActiveTab();
    if (!tab) return;

    const formatId = jlcFormatSelect.value;
    tab.jlcFormatId = formatId;

    if (!formatId) {
      closeJlcDialog();
      return;
    }

    if (typeof JLCFormat === "undefined") {
      alert("JLC format module failed to load.");
      jlcFormatSelect.value = "";
      tab.jlcFormatId = "";
      return;
    }

    openJlcDialog(formatId);
  });

  jlcDialogDownload.addEventListener("click", downloadJlcCsv);
  jlcDialogClose.addEventListener("click", closeJlcDialog);

  jlcDialog.addEventListener("click", (e) => {
    if (e.target === jlcDialog) closeJlcDialog();
  });

  downloadBtn.addEventListener("click", () => {
    const tab = getActiveTab();
    if (!tab) return;
    const csv = tableToCSV(tab);
    const blob = new Blob([csv], { type: "text/csv;charset=utf-8" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = tab.fileName.replace(/\.csv$/i, "") + "_edited.csv";
    a.click();
    URL.revokeObjectURL(url);
  });

  fontSizeValue.textContent = DEFAULTS.fontSize + "px";
})();
