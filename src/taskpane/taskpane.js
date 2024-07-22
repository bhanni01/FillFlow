/* FillFlow — taskpane.js  (v0.3 — preview + row selection) */

var state = {
  workbook: null,
  headers:  [],
  rows:     []
};

Office.onReady(function (info) {
  if (info.host !== Office.HostType.Word) {
    showStatus("This add-in only works in Microsoft Word.", "error");
    return;
  }
  document.getElementById("file-input").addEventListener("change", onFileChange);
  document.getElementById("btn-scan-doc").addEventListener("click", onScanDoc);
});

function onFileChange(evt) {
  var file = evt.target.files[0];
  if (!file) return;
  document.getElementById("file-name-display").textContent = file.name;

  var reader = new FileReader();
  reader.onload = function (e) {
    try {
      var data = new Uint8Array(e.target.result);
      state.workbook = XLSX.read(data, { type: "array", cellDates: true });

      var sheetName = state.workbook.SheetNames[0];
      var ws = state.workbook.Sheets[sheetName];
      var aoa = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });

      if (!aoa || aoa.length === 0) { showStatus("Worksheet appears empty.", "error"); return; }

      state.headers = aoa[0].map(function (h) { return String(h); });
      state.rows = [];
      for (var r = 1; r < aoa.length; r++) {
        var obj = {};
        for (var c = 0; c < state.headers.length; c++) {
          obj[state.headers[c]] = aoa[r][c] !== undefined ? aoa[r][c] : "";
        }
        state.rows.push(obj);
      }

      buildPreviewTable(aoa);
      document.getElementById("flat-row").value = 2;
      document.getElementById("line-start").value = 3;
      document.getElementById("line-end").value = aoa.length;
      show("step-preview");
      show("step-rows");
      hideStatus();
    } catch (err) {
      showStatus("Failed to parse file: " + err.message, "error");
    }
  };
  reader.readAsArrayBuffer(file);
}

function buildPreviewTable(aoa) {
  var tbl = document.getElementById("preview-table");
  tbl.innerHTML = "";

  var thead = document.createElement("thead");
  var trh = document.createElement("tr");
  addCell(trh, "th", "#");
  aoa[0].forEach(function (h) { addCell(trh, "th", String(h)); });
  thead.appendChild(trh);
  tbl.appendChild(thead);

  var tbody = document.createElement("tbody");
  for (var r = 1; r < aoa.length; r++) {
    var tr = document.createElement("tr");
    addCell(tr, "td", String(r + 1));
    aoa[r].forEach(function (cell) { addCell(tr, "td", formatCell(cell)); });
    tbody.appendChild(tr);
  }
  tbl.appendChild(tbody);
}

function addCell(row, tag, text) {
  var cell = document.createElement(tag);
  cell.textContent = text;
  row.appendChild(cell);
}

function formatCell(val) {
  if (val === null || val === undefined) return "";
  if (val instanceof Date) return val.toLocaleDateString();
  return String(val);
}

function onScanDoc() {
  showStatus("(document scanning not yet implemented)", "info");
}

function show(id)      { var el = document.getElementById(id); if (el) el.classList.remove("hidden"); }
function hide(id)      { var el = document.getElementById(id); if (el) el.classList.add("hidden"); }
function hideStatus()  { hide("status-area"); }

function showStatus(msg, type) {
  var area = document.getElementById("status-area");
  area.className = type || "info";
  document.getElementById("status-message").textContent = msg;
  show("status-area");
}
