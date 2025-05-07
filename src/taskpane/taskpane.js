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
