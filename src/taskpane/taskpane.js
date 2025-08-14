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
