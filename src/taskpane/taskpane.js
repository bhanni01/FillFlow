/* FillFlow — taskpane.js  (v0.4 — doc scanning) */

var state = {
  workbook:    null,
  headers:     [],
  rows:        [],
  docFlatTags: [],
  docItemTags: []
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
      var ws = state.workbook.Sheets[state.workbook.SheetNames[0]];
      var aoa = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
      if (!aoa || !aoa.length) { showStatus("Worksheet appears empty.", "error"); return; }
      state.headers = aoa[0].map(String);
      state.rows = [];
      for (var r = 1; r < aoa.length; r++) {
        var obj = {};
        state.headers.forEach(function (h, c) { obj[h] = aoa[r][c] !== undefined ? aoa[r][c] : ""; });
        state.rows.push(obj);
      }
      buildPreviewTable(aoa);
      document.getElementById("flat-row").value   = 2;
      document.getElementById("line-start").value = 3;
      document.getElementById("line-end").value   = aoa.length;
      show("step-preview"); show("step-rows"); hideStatus();
    } catch (err) { showStatus("Parse error: " + err.message, "error"); }
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
  thead.appendChild(trh); tbl.appendChild(thead);
  var tbody = document.createElement("tbody");
  for (var r = 1; r < aoa.length; r++) {
    var tr = document.createElement("tr");
    addCell(tr, "td", String(r + 1));
    aoa[r].forEach(function (c) { addCell(tr, "td", formatCell(c)); });
    tbody.appendChild(tr);
  }
  tbl.appendChild(tbody);
}

function onScanDoc() {
  showStatus("Scanning document…", "info");
  document.getElementById("btn-scan-doc").disabled = true;
  Word.run(function (context) {
    var controls = context.document.contentControls;
    controls.load("tag,type");
    return context.sync().then(function () {
      state.docFlatTags = [];
      state.docItemTags = [];
      for (var i = 0; i < controls.items.length; i++) {
        var tag = (controls.items[i].tag || "").trim();
        if (!tag) continue;
        if (tag.toLowerCase().startsWith("item_")) state.docItemTags.push(tag);
        else state.docFlatTags.push(tag);
      }
      state.docFlatTags = unique(state.docFlatTags);
      state.docItemTags = unique(state.docItemTags);
      showStatus(
        "Found " + state.docFlatTags.length + " flat tags, " +
        state.docItemTags.length + " item tags.",
        "success"
      );
      document.getElementById("btn-scan-doc").disabled = false;
    });
  }).catch(function (err) {
    showStatus("Scan error: " + err.message, "error");
    document.getElementById("btn-scan-doc").disabled = false;
  });
}

function unique(arr) {
  var seen = {}; return arr.filter(function (v) { if (seen[v]) return false; seen[v]=true; return true; });
}
function addCell(row, tag, text) { var c=document.createElement(tag); c.textContent=text; row.appendChild(c); }
function formatCell(v) { if(v===null||v===undefined) return ""; if(v instanceof Date) return v.toLocaleDateString(); return String(v); }
function show(id)     { var e=document.getElementById(id); if(e) e.classList.remove("hidden"); }
function hide(id)     { var e=document.getElementById(id); if(e) e.classList.add("hidden"); }
function hideStatus() { hide("status-area"); }
function showStatus(msg, type) {
  var a=document.getElementById("status-area");
  a.className = type||"info";
  document.getElementById("status-message").textContent = msg;
  show("status-area");
}
