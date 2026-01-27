// Minimal build helper — nothing to bundle since we use CDN for SheetJS.
// This script just confirms the structure is in place.
const fs = require("fs");
const path = require("path");

const root = path.join(__dirname, "..");
const required = [
  "src/taskpane/taskpane.html",
  "src/taskpane/taskpane.css",
  "src/taskpane/taskpane.js",
  "manifest.xml",
];

let ok = true;
required.forEach((f) => {
  const full = path.join(root, f);
  if (!fs.existsSync(full)) {
    console.error("MISSING:", full);
    ok = false;
  }
});

if (ok) console.log("Build check: all required files present.");
else process.exit(1);
