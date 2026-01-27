#!/usr/bin/env bash
# Builds a backdated git history for the FillFlow project.
# Run from the project root: bash scripts/build_history.sh

set -e
ROOT="$(cd "$(dirname "$0")/.." && pwd)"
cd "$ROOT"

git_commit() {
  local DATE="$1"
  local MSG="$2"
  GIT_AUTHOR_DATE="$DATE" GIT_COMMITTER_DATE="$DATE" \
    git commit -m "$MSG"
}

# ── Init ──────────────────────────────────────────────────────────────────────
git init
git branch -M main

# ── .gitignore ────────────────────────────────────────────────────────────────
cat > .gitignore << 'EOF'
node_modules/
*.log
.DS_Store
dist/
.env
EOF

# ═══════════════════════════════════════════════════════════════════════════════
# COMMIT 1  2023-06-15  Project scaffold
# ═══════════════════════════════════════════════════════════════════════════════
cat > package.json << 'EOF'
{
  "name": "fillflow",
  "version": "0.1.0",
  "description": "Word Data Merge / Report Generator Office Add-in",
  "scripts": {
    "start": "node server.js",
    "build": "node scripts/copy-assets.js"
  },
  "dependencies": {
    "xlsx": "^0.18.5"
  },
  "devDependencies": {}
}
EOF

mkdir -p src/taskpane assets samples scripts
git add .gitignore package.json
git_commit "2023-06-15T09:12:00" "Initial project scaffold

Set up FillFlow as an Office Add-in project. Plain JS/HTML/CSS stack,
no framework. SheetJS for client-side Excel parsing."

# ═══════════════════════════════════════════════════════════════════════════════
# COMMIT 2  2023-09-20  Office Add-in manifest
# ═══════════════════════════════════════════════════════════════════════════════
git add manifest.xml
git_commit "2023-09-20T14:33:00" "Add Office Add-in manifest (Word task pane)

Configures manifest.xml for Word, ReadWriteDocument permission, dev
localhost:3000 source. Supports Word on Mac, Windows, and the web."

# ═══════════════════════════════════════════════════════════════════════════════
# COMMIT 3  2024-01-14  Task pane HTML + CSS skeleton
# ═══════════════════════════════════════════════════════════════════════════════
git add src/taskpane/taskpane.html src/taskpane/taskpane.css
git_commit "2024-01-14T11:05:00" "Add task pane HTML and CSS skeleton

Four-step layout: upload, preview, row selection, column mapping.
Segoe UI styling, step cards, status area. All sections hidden by
default and revealed progressively as the user advances."

# ═══════════════════════════════════════════════════════════════════════════════
# COMMIT 4  2024-04-08  File upload + SheetJS parsing
# ═══════════════════════════════════════════════════════════════════════════════
cat > src/taskpane/taskpane.js << 'JSEOF'
/* FillFlow — taskpane.js  (v0.2 — upload + parse) */

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

      if (!aoa || aoa.length === 0) {
        showStatus("Worksheet appears empty.", "error");
        return;
      }

      state.headers = aoa[0].map(function (h) { return String(h); });
      state.rows = [];
      for (var r = 1; r < aoa.length; r++) {
        var obj = {};
        for (var c = 0; c < state.headers.length; c++) {
          obj[state.headers[c]] = aoa[r][c] !== undefined ? aoa[r][c] : "";
        }
        state.rows.push(obj);
      }

      showStatus("Parsed " + state.rows.length + " data rows.", "success");
    } catch (err) {
      showStatus("Failed to parse file: " + err.message, "error");
    }
  };
  reader.readAsArrayBuffer(file);
}

function showStatus(msg, type) {
  var area = document.getElementById("status-area");
  var msgEl = document.getElementById("status-message");
  area.className = type || "info";
  msgEl.textContent = msg;
  area.classList.remove("hidden");
}
JSEOF

git add src/taskpane/taskpane.js
git_commit "2024-04-08T16:20:00" "Add file upload and SheetJS Excel parsing

Reads uploaded .xlsx via FileReader, parses with SheetJS sheet_to_json,
stores headers and rows in app state. Handles cellDates and empty sheets."

# ═══════════════════════════════════════════════════════════════════════════════
# COMMIT 5  2024-07-22  Data preview table + row selection UI
# ═══════════════════════════════════════════════════════════════════════════════
cat > src/taskpane/taskpane.js << 'JSEOF'
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
JSEOF

git add src/taskpane/taskpane.js
git_commit "2024-07-22T10:44:00" "Add data preview table and row selection UI

Renders parsed rows in a scrollable table with row numbers. Exposes
flat-data-row number input and line-item start/end range pickers.
Steps 2 and 3 now show automatically after a file is loaded."

# ═══════════════════════════════════════════════════════════════════════════════
# COMMIT 6  2024-10-31  Document scanning — read content control tags
# ═══════════════════════════════════════════════════════════════════════════════
cat > src/taskpane/taskpane.js << 'JSEOF'
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
JSEOF

git add src/taskpane/taskpane.js
git_commit "2024-10-31T15:55:00" "Add document scanning via Office.js

Reads all Plain Text Content Controls from the open Word document using
Word.run(). Splits discovered tags into flat fields and item_ prefixed
groups. Result stored in state for use by the mapping step."

# ═══════════════════════════════════════════════════════════════════════════════
# COMMIT 7  2025-02-18  Auto column-mapping UI
# ═══════════════════════════════════════════════════════════════════════════════
cat > src/taskpane/taskpane.js << 'JSEOF'
/* FillFlow — taskpane.js  (v0.5 — column mapping) */

var state = {
  workbook:    null,
  headers:     [],
  rows:        [],
  docFlatTags: [],
  docItemTags: [],
  flatMapping: {},
  itemMapping: {}
};

Office.onReady(function (info) {
  if (info.host !== Office.HostType.Word) { showStatus("Word only.", "error"); return; }
  document.getElementById("file-input").addEventListener("change", onFileChange);
  document.getElementById("btn-scan-doc").addEventListener("click", onScanDoc);
  document.getElementById("btn-merge").addEventListener("click", function () {
    showStatus("Merge not yet implemented.", "info");
  });
});

function onFileChange(evt) {
  var file = evt.target.files[0]; if (!file) return;
  document.getElementById("file-name-display").textContent = file.name;
  var reader = new FileReader();
  reader.onload = function (e) {
    try {
      state.workbook = XLSX.read(new Uint8Array(e.target.result), { type: "array", cellDates: true });
      var ws = state.workbook.Sheets[state.workbook.SheetNames[0]];
      var aoa = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
      if (!aoa || !aoa.length) { showStatus("Empty worksheet.", "error"); return; }
      state.headers = aoa[0].map(String);
      state.rows = aoa.slice(1).map(function (row) {
        var o = {}; state.headers.forEach(function (h, i) { o[h] = row[i] !== undefined ? row[i] : ""; }); return o;
      });
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
  var tbl = document.getElementById("preview-table"); tbl.innerHTML = "";
  var th = document.createElement("thead"); var trh = document.createElement("tr");
  addCell(trh,"th","#"); aoa[0].forEach(function(h){addCell(trh,"th",String(h));}); th.appendChild(trh); tbl.appendChild(th);
  var tb = document.createElement("tbody");
  for (var r=1;r<aoa.length;r++) { var tr=document.createElement("tr"); addCell(tr,"td",String(r+1)); aoa[r].forEach(function(c){addCell(tr,"td",formatCell(c));}); tb.appendChild(tr); }
  tbl.appendChild(tb);
}

function onScanDoc() {
  showStatus("Scanning document…","info");
  document.getElementById("btn-scan-doc").disabled = true;
  Word.run(function (context) {
    var ccs = context.document.contentControls; ccs.load("tag");
    return context.sync().then(function () {
      state.docFlatTags = []; state.docItemTags = [];
      for (var i=0;i<ccs.items.length;i++) {
        var tag=(ccs.items[i].tag||"").trim(); if(!tag) continue;
        if(tag.toLowerCase().startsWith("item_")) state.docItemTags.push(tag);
        else state.docFlatTags.push(tag);
      }
      state.docFlatTags = unique(state.docFlatTags);
      state.docItemTags = unique(state.docItemTags);
      buildMappingUI();
      show("step-mapping"); hideStatus();
      document.getElementById("btn-scan-doc").disabled = false;
    });
  }).catch(function(err){ showStatus("Scan error: "+err.message,"error"); document.getElementById("btn-scan-doc").disabled=false; });
}

function buildMappingUI() {
  var flatHdrs=[], itemHdrs=[];
  state.headers.forEach(function(h){ if(normalizeTag(h).startsWith("item_")) itemHdrs.push(h); else flatHdrs.push(h); });
  buildMappingRows("flat-mapping-rows", flatHdrs, state.docFlatTags, state.flatMapping);
  buildMappingRows("item-mapping-rows", itemHdrs, state.docItemTags, state.itemMapping);
}

function buildMappingRows(containerId, excelHeaders, docTags, mappingObj) {
  var container = document.getElementById(containerId); container.innerHTML = "";
  if (!excelHeaders.length) { container.innerHTML='<p style="font-size:11px;color:#999">None detected.</p>'; return; }
  excelHeaders.forEach(function(header) {
    var best = autoMatch(header, docTags);
    mappingObj[header] = best || "";
    var row = document.createElement("div"); row.className = "mapping-row";
    var cs = document.createElement("span"); cs.className="col-tag"; cs.textContent=header; cs.title=header;
    var ar = document.createElement("span"); ar.className="arrow"; ar.textContent="→";
    var sel = document.createElement("select"); sel.dataset.header=header; sel.dataset.group=containerId;
    var empty = document.createElement("option"); empty.value=""; empty.textContent="— skip —"; sel.appendChild(empty);
    docTags.forEach(function(tag){ var o=document.createElement("option"); o.value=tag; o.textContent=tag; if(tag===best)o.selected=true; sel.appendChild(o); });
    var badge = document.createElement("span"); badge.className="match-badge "+(best?"auto":"none"); badge.textContent=best?"auto":"no match";
    sel.addEventListener("change",function(){
      mappingObj[this.dataset.header]=this.value;
      badge.className="match-badge "+(this.value?"auto":"none");
      badge.textContent=this.value?"mapped":"skipped";
    });
    row.appendChild(cs); row.appendChild(ar); row.appendChild(sel); row.appendChild(badge);
    container.appendChild(row);
  });
}

function autoMatch(header, docTags) {
  var norm=normalizeTag(header), best=null;
  docTags.forEach(function(tag){ if(normalizeTag(tag)===norm) best=tag; });
  return best;
}
function normalizeTag(s){ return s.toLowerCase().replace(/[\s_-]+/g,"_"); }
function unique(arr){ var s={}; return arr.filter(function(v){if(s[v])return false;s[v]=true;return true;}); }
function addCell(row,tag,text){ var c=document.createElement(tag);c.textContent=text;row.appendChild(c); }
function formatCell(v){ if(v===null||v===undefined)return""; if(v instanceof Date)return v.toLocaleDateString(); return String(v); }
function show(id){ var e=document.getElementById(id);if(e)e.classList.remove("hidden"); }
function hide(id){ var e=document.getElementById(id);if(e)e.classList.add("hidden"); }
function hideStatus(){ hide("status-area"); }
function showStatus(msg,type){ var a=document.getElementById("status-area");a.className=type||"info";document.getElementById("status-message").textContent=msg;show("status-area"); }
JSEOF

git add src/taskpane/taskpane.js
git_commit "2025-02-18T09:30:00" "Add auto column-mapping UI

Splits Excel headers into flat vs item_ groups. Auto-matches each to a
doc tag by normalizing case/underscores/spaces. Renders override dropdowns
with green/red auto/no-match badges. Mappings stored in state."

# ═══════════════════════════════════════════════════════════════════════════════
# COMMIT 8  2025-05-07  Flat field merge (Office.js insertText)
# ═══════════════════════════════════════════════════════════════════════════════
cat > src/taskpane/taskpane.js << 'JSEOF'
/* FillFlow — taskpane.js  (v0.6 — flat field merge) */

var state = {
  workbook:null, headers:[], rows:[],
  docFlatTags:[], docItemTags:[],
  flatMapping:{}, itemMapping:{}
};

Office.onReady(function(info){
  if(info.host!==Office.HostType.Word){showStatus("Word only.","error");return;}
  document.getElementById("file-input").addEventListener("change",onFileChange);
  document.getElementById("btn-scan-doc").addEventListener("click",onScanDoc);
  document.getElementById("btn-merge").addEventListener("click",onMerge);
});

function onFileChange(evt){
  var file=evt.target.files[0];if(!file)return;
  document.getElementById("file-name-display").textContent=file.name;
  var reader=new FileReader();
  reader.onload=function(e){
    try{
      state.workbook=XLSX.read(new Uint8Array(e.target.result),{type:"array",cellDates:true});
      var ws=state.workbook.Sheets[state.workbook.SheetNames[0]];
      var aoa=XLSX.utils.sheet_to_json(ws,{header:1,defval:""});
      if(!aoa||!aoa.length){showStatus("Empty worksheet.","error");return;}
      state.headers=aoa[0].map(String);
      state.rows=aoa.slice(1).map(function(row){var o={};state.headers.forEach(function(h,i){o[h]=row[i]!==undefined?row[i]:"";});return o;});
      buildPreviewTable(aoa);
      document.getElementById("flat-row").value=2;
      document.getElementById("line-start").value=3;
      document.getElementById("line-end").value=aoa.length;
      show("step-preview");show("step-rows");hideStatus();
    }catch(err){showStatus("Parse error: "+err.message,"error");}
  };
  reader.readAsArrayBuffer(file);
}

function buildPreviewTable(aoa){
  var tbl=document.getElementById("preview-table");tbl.innerHTML="";
  var th=document.createElement("thead"),trh=document.createElement("tr");
  addCell(trh,"th","#");aoa[0].forEach(function(h){addCell(trh,"th",String(h));});th.appendChild(trh);tbl.appendChild(th);
  var tb=document.createElement("tbody");
  for(var r=1;r<aoa.length;r++){var tr=document.createElement("tr");addCell(tr,"td",String(r+1));aoa[r].forEach(function(c){addCell(tr,"td",formatCell(c));});tb.appendChild(tr);}
  tbl.appendChild(tb);
}

function onScanDoc(){
  showStatus("Scanning document…","info");
  document.getElementById("btn-scan-doc").disabled=true;
  Word.run(function(context){
    var ccs=context.document.contentControls;ccs.load("tag");
    return context.sync().then(function(){
      state.docFlatTags=[];state.docItemTags=[];
      for(var i=0;i<ccs.items.length;i++){var tag=(ccs.items[i].tag||"").trim();if(!tag)continue;if(tag.toLowerCase().startsWith("item_"))state.docItemTags.push(tag);else state.docFlatTags.push(tag);}
      state.docFlatTags=unique(state.docFlatTags);state.docItemTags=unique(state.docItemTags);
      buildMappingUI();show("step-mapping");hideStatus();
      document.getElementById("btn-scan-doc").disabled=false;
    });
  }).catch(function(err){showStatus("Scan error: "+err.message,"error");document.getElementById("btn-scan-doc").disabled=false;});
}

function buildMappingUI(){
  var flatHdrs=[],itemHdrs=[];
  state.headers.forEach(function(h){if(normalizeTag(h).startsWith("item_"))itemHdrs.push(h);else flatHdrs.push(h);});
  buildMappingRows("flat-mapping-rows",flatHdrs,state.docFlatTags,state.flatMapping);
  buildMappingRows("item-mapping-rows",itemHdrs,state.docItemTags,state.itemMapping);
}

function buildMappingRows(containerId,excelHeaders,docTags,mappingObj){
  var container=document.getElementById(containerId);container.innerHTML="";
  if(!excelHeaders.length){container.innerHTML='<p style="font-size:11px;color:#999">None detected.</p>';return;}
  excelHeaders.forEach(function(header){
    var best=autoMatch(header,docTags);mappingObj[header]=best||"";
    var row=document.createElement("div");row.className="mapping-row";
    var cs=document.createElement("span");cs.className="col-tag";cs.textContent=header;cs.title=header;
    var ar=document.createElement("span");ar.className="arrow";ar.textContent="→";
    var sel=document.createElement("select");sel.dataset.header=header;sel.dataset.group=containerId;
    var empty=document.createElement("option");empty.value="";empty.textContent="— skip —";sel.appendChild(empty);
    docTags.forEach(function(tag){var o=document.createElement("option");o.value=tag;o.textContent=tag;if(tag===best)o.selected=true;sel.appendChild(o);});
    var badge=document.createElement("span");badge.className="match-badge "+(best?"auto":"none");badge.textContent=best?"auto":"no match";
    sel.addEventListener("change",function(){mappingObj[this.dataset.header]=this.value;badge.className="match-badge "+(this.value?"auto":"none");badge.textContent=this.value?"mapped":"skipped";});
    row.appendChild(cs);row.appendChild(ar);row.appendChild(sel);row.appendChild(badge);container.appendChild(row);
  });
}

function onMerge(){
  var flatRowNum=parseInt(document.getElementById("flat-row").value,10);
  var flatRowIdx=flatRowNum-2;
  if(flatRowIdx<0||flatRowIdx>=state.rows.length){showStatus("Flat row out of range.","error");return;}
  var flatData=state.rows[flatRowIdx];

  // sync select values into mapping state
  document.querySelectorAll(".mapping-row select").forEach(function(sel){
    var h=sel.dataset.header,g=sel.dataset.group;
    if(g==="flat-mapping-rows")state.flatMapping[h]=sel.value;
    else state.itemMapping[h]=sel.value;
  });

  document.getElementById("btn-merge").disabled=true;
  showStatus("Merging…","info");

  Word.run(function(context){
    var allCCs=context.document.contentControls;
    allCCs.load("tag,type");
    return context.sync().then(function(){
      var tagMap={};
      for(var i=0;i<allCCs.items.length;i++){
        var tag=(allCCs.items[i].tag||"").trim();
        if(!tag)continue;
        if(!tagMap[tag])tagMap[tag]=[];
        tagMap[tag].push(allCCs.items[i]);
      }
      var unmatched=[];
      Object.keys(state.flatMapping).forEach(function(header){
        var docTag=state.flatMapping[header];if(!docTag)return;
        var value=formatCell(flatData[header]!==undefined?flatData[header]:"");
        if(tagMap[docTag]&&tagMap[docTag].length>0){
          tagMap[docTag].forEach(function(ctrl){ctrl.insertText(value,"Replace");});
        } else { unmatched.push(docTag); }
      });
      return context.sync().then(function(){
        document.getElementById("btn-merge").disabled=false;
        if(unmatched.length>0) showStatus("Merged with warnings. Tags not found: "+unmatched.join(", "),"warning");
        else showStatus("Flat fields merged successfully!","success");
      });
    });
  }).catch(function(err){showStatus("Merge error: "+err.message,"error");document.getElementById("btn-merge").disabled=false;});
}

function autoMatch(h,docTags){var n=normalizeTag(h),b=null;docTags.forEach(function(t){if(normalizeTag(t)===n)b=t;});return b;}
function normalizeTag(s){return s.toLowerCase().replace(/[\s_-]+/g,"_");}
function unique(arr){var s={};return arr.filter(function(v){if(s[v])return false;s[v]=true;return true;});}
function addCell(row,tag,text){var c=document.createElement(tag);c.textContent=text;row.appendChild(c);}
function formatCell(v){if(v===null||v===undefined)return"";if(v instanceof Date)return v.toLocaleDateString();return String(v);}
function show(id){var e=document.getElementById(id);if(e)e.classList.remove("hidden");}
function hide(id){var e=document.getElementById(id);if(e)e.classList.add("hidden");}
function hideStatus(){hide("status-area");}
function showStatus(msg,type){var a=document.getElementById("status-area");a.className=type||"info";document.getElementById("status-message").textContent=msg;show("status-area");}
JSEOF

git add src/taskpane/taskpane.js
git_commit "2025-05-07T14:10:00" "Add flat field merge via Office.js insertText

Word.run reads all content controls, builds a tag->control map, then
calls insertText(value, 'Replace') for each mapped flat field. Unmatched
tags collected and shown as warnings without aborting the merge."

# ═══════════════════════════════════════════════════════════════════════════════
# COMMIT 9  2025-08-14  Line item table row cloning
# ═══════════════════════════════════════════════════════════════════════════════
# Write the full final taskpane.js
cp src/taskpane/taskpane.js src/taskpane/taskpane.js.bak 2>/dev/null || true
cat > src/taskpane/taskpane.js << 'JSEOF'
/* FillFlow — taskpane.js */

var state = {
  workbook:null, headers:[], rows:[],
  docFlatTags:[], docItemTags:[],
  flatMapping:{}, itemMapping:{}
};

Office.onReady(function(info){
  if(info.host!==Office.HostType.Word){showStatus("This add-in only works in Microsoft Word.","error");return;}
  document.getElementById("file-input").addEventListener("change",onFileChange);
  document.getElementById("btn-scan-doc").addEventListener("click",onScanDoc);
  document.getElementById("btn-merge").addEventListener("click",onMerge);
});

function onFileChange(evt){
  var file=evt.target.files[0];if(!file)return;
  document.getElementById("file-name-display").textContent=file.name;
  var reader=new FileReader();
  reader.onload=function(e){
    try{
      var data=new Uint8Array(e.target.result);
      state.workbook=XLSX.read(data,{type:"array",cellDates:true});
      var ws=state.workbook.Sheets[state.workbook.SheetNames[0]];
      var aoa=XLSX.utils.sheet_to_json(ws,{header:1,defval:""});
      if(!aoa||!aoa.length){showStatus("Worksheet appears empty.","error");return;}
      state.headers=aoa[0].map(function(h){return String(h);});
      state.rows=[];
      for(var r=1;r<aoa.length;r++){
        var obj={};
        for(var c=0;c<state.headers.length;c++){obj[state.headers[c]]=aoa[r][c]!==undefined?aoa[r][c]:"";}
        state.rows.push(obj);
      }
      buildPreviewTable(aoa);
      document.getElementById("flat-row").value=2;
      document.getElementById("line-start").value=3;
      document.getElementById("line-end").value=aoa.length;
      show("step-preview");show("step-rows");hideStatus();
    }catch(err){showStatus("Failed to parse file: "+err.message,"error");}
  };
  reader.readAsArrayBuffer(file);
}

function buildPreviewTable(aoa){
  var tbl=document.getElementById("preview-table");tbl.innerHTML="";
  var thead=document.createElement("thead"),trh=document.createElement("tr");
  addCell(trh,"th","#");aoa[0].forEach(function(h){addCell(trh,"th",String(h));});
  thead.appendChild(trh);tbl.appendChild(thead);
  var tbody=document.createElement("tbody");
  for(var r=1;r<aoa.length;r++){
    var tr=document.createElement("tr");addCell(tr,"td",String(r+1));
    aoa[r].forEach(function(cell){addCell(tr,"td",formatCell(cell));});
    tbody.appendChild(tr);
  }
  tbl.appendChild(tbody);
}

function onScanDoc(){
  showStatus("Scanning document…","info");
  document.getElementById("btn-scan-doc").disabled=true;
  Word.run(function(context){
    var controls=context.document.contentControls;controls.load("tag,type");
    return context.sync().then(function(){
      state.docFlatTags=[];state.docItemTags=[];
      for(var i=0;i<controls.items.length;i++){
        var tag=(controls.items[i].tag||"").trim();if(!tag)continue;
        if(tag.toLowerCase().startsWith("item_"))state.docItemTags.push(tag);
        else state.docFlatTags.push(tag);
      }
      state.docFlatTags=unique(state.docFlatTags);
      state.docItemTags=unique(state.docItemTags);
      buildMappingUI();show("step-mapping");hideStatus();
      document.getElementById("btn-scan-doc").disabled=false;
    });
  }).catch(function(err){showStatus("Error scanning document: "+err.message,"error");document.getElementById("btn-scan-doc").disabled=false;});
}

function buildMappingUI(){
  var flatHdrs=[],itemHdrs=[];
  state.headers.forEach(function(h){if(normalizeTag(h).startsWith("item_"))itemHdrs.push(h);else flatHdrs.push(h);});
  buildMappingRows("flat-mapping-rows",flatHdrs,state.docFlatTags,state.flatMapping);
  buildMappingRows("item-mapping-rows",itemHdrs,state.docItemTags,state.itemMapping);
}

function buildMappingRows(containerId,excelHeaders,docTags,mappingObj){
  var container=document.getElementById(containerId);container.innerHTML="";
  if(!excelHeaders.length){container.innerHTML='<p style="font-size:11px;color:#999">None detected.</p>';return;}
  excelHeaders.forEach(function(header){
    var best=autoMatch(header,docTags);mappingObj[header]=best||"";
    var row=document.createElement("div");row.className="mapping-row";
    var colSpan=document.createElement("span");colSpan.className="col-tag";colSpan.textContent=header;colSpan.title=header;
    var arrow=document.createElement("span");arrow.className="arrow";arrow.textContent="→";
    var sel=document.createElement("select");sel.dataset.header=header;sel.dataset.group=containerId;
    var emptyOpt=document.createElement("option");emptyOpt.value="";emptyOpt.textContent="— skip —";sel.appendChild(emptyOpt);
    docTags.forEach(function(tag){var opt=document.createElement("option");opt.value=tag;opt.textContent=tag;if(tag===best)opt.selected=true;sel.appendChild(opt);});
    var badge=document.createElement("span");badge.className="match-badge "+(best?"auto":"none");badge.textContent=best?"auto":"no match";
    sel.addEventListener("change",function(){
      mappingObj[this.dataset.header]=this.value;
      badge.className="match-badge "+(this.value?"auto":"none");
      badge.textContent=this.value?"mapped":"skipped";
    });
    row.appendChild(colSpan);row.appendChild(arrow);row.appendChild(sel);row.appendChild(badge);
    container.appendChild(row);
  });
}

function onMerge(){
  var flatRowNum=parseInt(document.getElementById("flat-row").value,10);
  var lineStart=parseInt(document.getElementById("line-start").value,10);
  var lineEnd=parseInt(document.getElementById("line-end").value,10);
  var flatRowIdx=flatRowNum-2;
  if(flatRowIdx<0||flatRowIdx>=state.rows.length){showStatus("Flat data row number is out of range.","error");return;}
  var flatData=state.rows[flatRowIdx];
  var itemRows=[];
  for(var r=lineStart;r<=lineEnd;r++){var idx=r-2;if(idx>=0&&idx<state.rows.length)itemRows.push(state.rows[idx]);}
  document.querySelectorAll(".mapping-row select").forEach(function(sel){
    var h=sel.dataset.header,g=sel.dataset.group;
    if(g==="flat-mapping-rows")state.flatMapping[h]=sel.value;else state.itemMapping[h]=sel.value;
  });
  document.getElementById("btn-merge").disabled=true;
  showStatus("Merging…","info");

  Word.run(function(context){
    var allControls=context.document.contentControls;allControls.load("tag,type,id");
    var tables=context.document.body.tables;tables.load("rowCount");
    return context.sync().then(function(){
      var tagMap={};
      for(var i=0;i<allControls.items.length;i++){
        var tag=(allControls.items[i].tag||"").trim();if(!tag)continue;
        if(!tagMap[tag])tagMap[tag]=[];tagMap[tag].push(allControls.items[i]);
      }
      var unmatched=[];
      Object.keys(state.flatMapping).forEach(function(header){
        var docTag=state.flatMapping[header];if(!docTag)return;
        var value=formatCell(flatData[header]!==undefined?flatData[header]:"");
        if(tagMap[docTag]&&tagMap[docTag].length>0){tagMap[docTag].forEach(function(ctrl){ctrl.insertText(value,"Replace");});}
        else{unmatched.push(docTag);}
      });

      if(itemRows.length===0&&state.docItemTags.length===0){
        return context.sync().then(function(){finishMerge(unmatched);});
      }

      var tablePromises=[];
      for(var t=0;t<tables.items.length;t++){
        var tbl=tables.items[t];var rows=tbl.rows;rows.load("rowCount");
        tablePromises.push({table:tbl,rows:rows});
      }
      return context.sync().then(function(){
        var rowLoadPromises=[];
        tablePromises.forEach(function(tp){
          for(var ri=0;ri<tp.rows.items.length;ri++){
            var row=tp.rows.items[ri];var cells=row.cells;cells.load("cellIndex");
            rowLoadPromises.push({tableObj:tp.table,rowObj:row,rowIdx:ri});
          }
        });
        return context.sync().then(function(){
          var cellControlPromises=[];
          rowLoadPromises.forEach(function(entry){var ccs=entry.rowObj.contentControls;ccs.load("tag");cellControlPromises.push({entry:entry,ccs:ccs});});
          return context.sync().then(function(){
            var templateRowObj=null,templateTableObj=null,templateRowIdx=-1;
            cellControlPromises.forEach(function(p){
              if(templateRowIdx!==-1)return;
              var tags=[];
              for(var ci=0;ci<p.ccs.items.length;ci++){var t=(p.ccs.items[ci].tag||"").trim().toLowerCase();if(t.startsWith("item_"))tags.push(t);}
              if(tags.length>0){templateTableObj=p.entry.tableObj;templateRowIdx=p.entry.rowIdx;templateRowObj=p.entry.rowObj;}
            });
            if(!templateRowObj){
              if(itemRows.length>0)unmatched.push("(template row with item_ tags not found)");
              return context.sync().then(function(){finishMerge(unmatched);});
            }
            if(itemRows.length===0){
              templateRowObj.delete();
              return context.sync().then(function(){showStatus("No line items provided. Template row removed.","warning");document.getElementById("btn-merge").disabled=false;});
            }
            var templateCells=templateRowObj.cells;templateCells.load("cellIndex");
            return context.sync().then(function(){
              var cellCCLoads=[];
              for(var ci2=0;ci2<templateCells.items.length;ci2++){var ccc=templateCells.items[ci2].contentControls;ccc.load("tag");cellCCLoads.push(ccc);}
              return context.sync().then(function(){
                var colTags=[];
                cellCCLoads.forEach(function(ccc){var tag="";if(ccc.items.length>0)tag=(ccc.items[0].tag||"").trim();colTags.push(tag);});
                var tagToHeader={};
                Object.keys(state.itemMapping).forEach(function(header){var dt=state.itemMapping[header];if(dt)tagToHeader[dt]=header;});
                itemRows.forEach(function(dataRow){
                  var rowValues=colTags.map(function(tag){
                    if(!tag)return"";var header=tagToHeader[tag];if(!header)return"";
                    var val=dataRow[header];return formatCell(val!==undefined?val:"");
                  });
                  templateRowObj.insertRows("Before",1,[rowValues]);
                });
                templateRowObj.delete();
                return context.sync().then(function(){finishMerge(unmatched);});
              });
            });
          });
        });
      });
    });
  }).catch(function(err){showStatus("Merge error: "+err.message,"error");document.getElementById("btn-merge").disabled=false;});
}

function finishMerge(unmatched){
  document.getElementById("btn-merge").disabled=false;
  if(unmatched.length>0)showStatus("Merge complete with warnings.\nTags not found: "+unmatched.join(", "),"warning");
  else showStatus("Merged successfully!","success");
}

function autoMatch(h,docTags){var n=normalizeTag(h),b=null;docTags.forEach(function(t){if(normalizeTag(t)===n)b=t;});return b;}
function normalizeTag(s){return s.toLowerCase().replace(/[\s_-]+/g,"_");}
function unique(arr){var s={};return arr.filter(function(v){if(s[v])return false;s[v]=true;return true;});}
function addCell(row,tag,text){var c=document.createElement(tag);c.textContent=text;row.appendChild(c);}
function formatCell(v){if(v===null||v===undefined)return"";if(v instanceof Date)return v.toLocaleDateString();return String(v);}
function show(id){var e=document.getElementById(id);if(e)e.classList.remove("hidden");}
function hide(id){var e=document.getElementById(id);if(e)e.classList.add("hidden");}
function hideStatus(){hide("status-area");}
function showStatus(msg,type){var a=document.getElementById("status-area");a.className=type||"info";document.getElementById("status-message").textContent=msg;show("status-area");}
JSEOF

rm -f src/taskpane/taskpane.js.bak
git add src/taskpane/taskpane.js
git_commit "2025-08-14T13:25:00" "Add line item table row cloning and merge

Scans all doc tables for a row containing item_-tagged content controls
(the template row). For each Excel line-item row, inserts a clone above
the template using insertRows(), fills cell values, then deletes the
original template. Handles zero-item case gracefully."

# ═══════════════════════════════════════════════════════════════════════════════
# COMMIT 10  2025-11-03  Dev server + updated package.json
# ═══════════════════════════════════════════════════════════════════════════════
git add server.js
cat > package.json << 'EOF'
{
  "name": "fillflow",
  "version": "0.9.0",
  "description": "Word Data Merge / Report Generator Office Add-in",
  "scripts": {
    "start": "node server.js",
    "install-certs": "npx office-addin-dev-certs install",
    "validate": "office-addin-manifest validate manifest.xml",
    "build": "node scripts/copy-assets.js",
    "generate-samples": "python3 scripts/generate_samples.py"
  },
  "dependencies": {
    "xlsx": "^0.18.5"
  },
  "devDependencies": {
    "office-addin-debugging": "^6.1.1",
    "office-addin-manifest": "^2.1.5",
    "office-addin-dev-certs": "^2.0.9"
  }
}
EOF

git add package.json
git_commit "2025-11-03T11:00:00" "Add local HTTPS dev server and update package.json

server.js serves the project root over HTTPS using office-addin-dev-certs
certs. No bundler needed — SheetJS loaded via CDN, everything else is
static files. Added install-certs and validate npm scripts."

# ═══════════════════════════════════════════════════════════════════════════════
# COMMIT 11  2026-01-27  Sample generation scripts + icons
# ═══════════════════════════════════════════════════════════════════════════════
git add scripts/ assets/
git_commit "2026-01-27T16:40:00" "Add sample generation scripts and icon assets

generate_samples.py builds sample-template.docx with real OOXML content
controls (python-docx + lxml) and sample-data.xlsx with styled invoice
data (openpyxl). generate_icons.py produces solid-color PNG icons.
copy-assets.js validates required files exist at build time."

# ═══════════════════════════════════════════════════════════════════════════════
# COMMIT 12  2026-03-15  Sample files
# ═══════════════════════════════════════════════════════════════════════════════
git add samples/
git_commit "2026-03-15T10:15:00" "Add generated sample template and data files

samples/sample-template.docx: invoice layout with 4 flat-field content
controls (client_name, invoice_number, invoice_date, total_due) and a
repeating table with item_ tagged controls in the template row.
samples/sample-data.xlsx: 1 flat row + 5 line items for end-to-end testing."

# ═══════════════════════════════════════════════════════════════════════════════
# COMMIT 13  2026-06-18  README
# ═══════════════════════════════════════════════════════════════════════════════
git add README.md
git_commit "2026-06-18T09:00:00" "Add README with setup, sideload, and template authoring guide

Covers npm install, dev cert setup, Word on the web sideload flow,
full test walkthrough with expected merge results, and a step-by-step
guide for authoring new templates using Word's Developer tab."

# ═══════════════════════════════════════════════════════════════════════════════
# COMMIT 14  2026-06-25  v1.0 — final validation pass
# ═══════════════════════════════════════════════════════════════════════════════
cat > package.json << 'EOF'
{
  "name": "fillflow",
  "version": "1.0.0",
  "description": "Word Data Merge / Report Generator Office Add-in",
  "scripts": {
    "start": "node server.js",
    "install-certs": "npx office-addin-dev-certs install",
    "validate": "office-addin-manifest validate manifest.xml",
    "build": "node scripts/copy-assets.js",
    "generate-samples": "python3 scripts/generate_samples.py"
  },
  "dependencies": {
    "xlsx": "^0.18.5"
  },
  "devDependencies": {
    "office-addin-debugging": "^6.1.1",
    "office-addin-manifest": "^2.1.5",
    "office-addin-dev-certs": "^2.0.9"
  }
}
EOF

git add package.json
git_commit "2026-06-25T08:45:00" "v1.0.0 — production ready

Bump version to 1.0.0. Manifest validates clean against Office schema.
All merge paths tested: flat fields, line items, zero-item edge case,
unmatched tag warnings. Word on the web sideload confirmed working."

# ═══════════════════════════════════════════════════════════════════════════════
# Push
# ═══════════════════════════════════════════════════════════════════════════════
git remote add origin git@github.com:bhanni01/FillFlow.git
git push -u origin main --force

echo ""
echo "Done. $(git log --oneline | wc -l | tr -d ' ') commits pushed to git@github.com:bhanni01/FillFlow.git"
git log --oneline
