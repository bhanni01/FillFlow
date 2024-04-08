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
