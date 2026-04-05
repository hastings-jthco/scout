// @OnlyCurrentDoc

const appDate = "05Apr26", appVers = 5

function onOpen() { createMenu() }

function onChange(e) {
// Automatically reset-cache on a manual row insertion/deletion
  if (e.changeType === 'INSERT_ROW' || e.changeType === 'DELETE_ROW') resetCache()
}

function onEdit(e) {
// Query and optionally re-cache a single block
  const cache = CacheService.getScriptCache()
  const ui = SpreadsheetApp.getUi(); 
  const ws = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet()
  const row = e.range.getRow(), col = e.range.getColumn() //ui.alert("at "+ row+"\\"+col+": "+e.value)
  if (row < stubRow || col < stubCol) return  // avoid stub; applies to panel data only
  if (e.value.toString().slice(0,6).toLowerCase() == "?fetch") { // flag to request to cache pre-check
    if (!geoIds == cache.get("A"+row) || e.value.indexOf("1(1)")==6) // attempt OR force a pre-fetch
     if(ui.alert("reCache "+row, ui.ButtonSet.YES_NO)==ui.Button.YES) reCache(row) // cache status now estb 
  }
  Utilities.sleep(500); ws.getRange(row,col).activate() // pause because fetch() runs before onEdit returns
}
/*
function onSelectionChange(e) {
// Erratic, but if working reloads entire cache (same as menu command)
const ws = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet()
const row = e.range.getRow(), col = e.range.getColumn()
  if (row == 1 && col == 2) { 
    ws.getRange("A1").activate(); ws.getRange("B1").setValue(""); 
    SpreadsheetApp.flush() // confirm to user by UI change
    resetCache(); ws.getRange("B1").activate()
  }  
}
*/

const stubRow = 2, stubCol = 4  // sheet layout  TESTing; 2 in proof

const itemListUrl = 'https://raw.githubusercontent.com/hastings-jthco/scout/refs/heads/main/ItemList.json'

const cacheTTL = 300 // requested cache duration (seconds)

var geoIds = [], idRows = []

function createMenu(s) {
  const ui = SpreadsheetApp.getUi();
  if (s) ui.removeMenu(s)
  ui.createMenu((s?s:"Fetch"))
    //.addItem((s?"Reset":"Setup")+" Sheet", 'clearSheet')
      .addItem("Clear Sheet", 'clearSheet')
      .addItem("Recalculate", 'resetCache')
      .addItem("Freeze Data", 'freezeData')
      .addToUi();
}

function clearSheet() {
// Clear the entire worksheet, outside the stub
//const props = PropertiesService.getScriptProperties();
  const ws = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet()
  const wsName = ws.getName(), lastRow = ws.getLastRow(), lastCol = ws.getLastColumn()
//const oldName = props.getProperty('currentMenuName');
  if(lastRow < stubRow && ws.getFrozenRows() == 0) { // New/empty sheet
    ws.setFrozenRows(stubRow); ws.setFrozenColumns(stubCol)
    ws.getRange("A1").setValue("CACHE")
    ws.getRange(stubRow+1,1) = "GeoID"; ws.getRange(stubRow+1,2)= "GeoName"
  //createMenu("Fetch")  // update menu [TODO: for this sheet]
  } else {
    var rg = ws.getRange(stubRow+1,stubCol+1, lastRow-stubRow, lastCol-stubCol)
    rg.clearContent(); rg.clearNote()
  }  
//props.setProperty('currentMenuName', wsName);
}

function freezeData() {
// Make UDF data permanent, and re-head rows 1-2 where known  
headupData(); return
  const ui = SpreadsheetApp.getUi()
  const ws = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet()
  const lastRow = ws.getLastRow(), lastCol = ws.getLastColumn() //; ui.alert("at freezeData "+lastRow+"\\"+lastCol)
//rg = ws.getRange("E3:E6"); rg.setValues(rg.getValues()); return //TESTing
  rg = ws.getRange(stubRow+1,stubCol+1, lastRow-stubRow, lastCol-stubCol);  //MAYBE s/b Col 1::lastCol
/*
  if (ui.alert("Really Freeze "+rg.getA1Notation(), ui.ButtonSet.OK_CANCEL) == ui.Button.CANCEL) return
  rg.setValues(rg.getValues()); rg.activate()
*/
  if (ui.alert("Reset Headings ", ui.ButtonSet.YES_NO) == ui.Button.YES) headupData()
}

function headupData() {
  const ui = SpreadsheetApp.getUi()//; ui.alert(itemListUrl.slice(0,12)+"..."+itemListUrl.slice(-5))
  const ws = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet()
  const lastRow = ws.getLastRow(), lastCol = ws.getLastColumn(); //ui.alert("at headupData "+(stubCol+1)+":"+lastCol)
  const resp = UrlFetchApp.fetch(itemListUrl, {muteHttpExceptions:true,}) //; ui.alert(resp)
  const json = JSON.parse(resp.getContentText()) //; ui.alert(json.length)
  //ui.alert(json[1][1]+";"+json[2][1]); // access contents directly
  var c, i, j, r, s,s1;
  r = stubRow+1
  c = stubCol; while (++c <= lastCol) {
    s = ws.getRange(r,c).getValue(); s1 = "" 
    i = s.indexOf("E"); if (i > 0) { s1 = s.slice(i+1); s = s.slice(0,i+1) } // trim off year, if any
    j = json.length; while (--j > 0) if (json[j][1] == s) break; ui.alert(j+" "+s) // search matching itemName
    if (j) { ws.getRange(2,c).setValue(s); with (ws.getRange(1, c)) { setValue(json[j][2]+s1); setNote(json[j][3]) } }
  }  
}

/*
function clearfix(f) {
  const ui = SpreadsheetApp.getUi()
  const wb = SpreadsheetApp.getActiveSpreadsheet(), ws = wb.getActiveSheet();
  var i, j, k, r, s, t, y;
  function clear() { ws.getRange("A2:M22").clear(); // TODO: getLastRow\Column()}
    with (ws.getRange("C1:M2")) { clearContent(); clearNote() } 
    g1s.length = 0
  }
  function fix() { 
  //ui.alert(ws.getLastColumn())
  //TODO: ?Make Cols A:B text black  
   if (1) { 
    var rg = ws.getRange("A2:E22"); rg.setValues(rg.getDisplayValues()); //return
    var col =2; while (col++ <= ws.getLastColumn()) {
     if (!ws.getRange(2, col).isBlank()) {
       with (ws.getRange(2, col)) { v = getValue(); s = v.slice(0,11); setValue(s); y = v.slice(11) }
       i = 0; while (i++ <10) if (itmList[i][1] == s) break //ui.alert(i) //lookup itmNo from itmID
       if (y) with (ws.getRange(1, col)) { setValue(itmList[i][2]+y); setNote(itmList[i][3]) } } }
   } else { 
    var col =0; while (++col <= ws.getLastColumn()) {
     if (!ws.getRange(2, col).isBlank()) {
      if (!(t = ws.getRange(2,col).getFormula())) { //ui.alert(t)
       with (ws.getRange(2, col)) { v = getValue(); s = v.slice(0,11); setValue(s); y = v.slice(11) }
       i = 0; while (i++ <10) if (itmList[i][1] == s) break //ui.alert(i) //lookup itmNo from itmID
       if (y) with (ws.getRange(1, col)) { setValue(itmList[i][2]+s); setNote(itmList[i][3]) }
        k =-1; if ((i = t.indexOf("}"))) k = t.slice(0,i).split(","); //ui.alert(i+","+k) // 1st param 
        k = (col == 1 ? 2 : (k <=0 ? 1 : k)); ui.alert(t+";"+i+","+k) // stub is always 2 cols
        var rg = ws.getRange(2,col,21,k); rg.setValues(rg.getDisplayValues()); col+= k } } }
   }
  }
//if (f) { clear(); return } // forced clear ** NOT USED **
  cacheIn(g1s); 
//ui.alert((g1s.length-1)+" rows in stub "+(g1s.length>1 ? g1s[1]+"..."+g1s[g1s[0]-1] : "")
  if (!g1s[0]) if(ui.alert("Proceed with uncached stub", ui.ButtonSet.YES_NO) == ui.Button.YES) return
  v = ui.alert("Clear (YES) or Fix (NO) the content?",ui.ButtonSet.YES_NO_CANCEL)
  if (v != ui.Button.CANCEL) { if (v == ui.Button.YES) clear(); else fix() }
  cacheStub(ws)  // all cases
  ws.getRange("A2").activate()
}
*/

function resetCache(f) {
// Reload entire cache, i.e. all fetch() blocks; f is falsey if called from menu 
  const ws = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet()
  geoIds.length = 0; reCache(0, f); 
  if (!f) { ws.insertRowAfter(2); ws.deleteRow(3) } // force UDFs to update
  return "Cached Rows: "+idRows.join(";")
}

function reCache(row, f) {
// Reload one or all cache block(s)  
  const cache = CacheService.getScriptCache()  // NB: .getUserCache() for addOn
  const ws = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet()
  function setFW(v) { return; if(rg) rg.setFontWeight(v) }
  var r = row, r0 = 0, rg, v
  if (r) { rg = ws.getRange("A"+r); rg.activate() }  // confirm to user by UI selection
  if(!f) { rg = ws.getRange("A1:A"+ws.getLastRow()); setFW("normal") }
  r -= Math.sign(r); while (++r <= ws.getLastRow()+1) { // if single block, back-off to force immed break 
    rg = ws.getRange(r,1); v = rg.getValue().toString()  // all cache data is string
  //if (!row && row>1) setFW('normal')
    if (rg.isBlank() || v.toLowerCase() == "geoid") {  // break in fetch sequence
      if (r0) { 
        cache.put("A"+r0, JSON.stringify(geoIds))  // cache prior fetch
        if (row) return; // if requested for a specific block
        idRows.push(r0) // debug: note rows of block headers
        r0 = 0; geoIds.length = 0  // prepare for next block per stub
      } 
    }
    if (!rg.isBlank()) {
      if (v.toLowerCase() == "geoid") { r0 = r; geoIds[0] = v; setFW('bold') }  // estb next block header
      else if (parseInt(v)) geoIds.push(v)  // append in-block geoIds
    }
  }
  SpreadsheetApp.flush()  // force changes; confirm to user by UI selection
}


function f1c(x) { return ["col", "col2"] } // sample 1D column UDF
function f1r(x) { return [["row1", "row2"]] } // sample 1D row UDF
function f2a(x) { return [["hdg1","hdg2"], [x, 2*x], [10*x, 2*10*x]] } // sample 2D array UDF

/**
 * Sample/Stub of fetchACS() UDF, demoing cache interactions
 *
 * @param {number} [x] - Typically a cache search value.
 * @returns {object} 1D column vector, spilled into cell(s).
 */
function fetch1(x) {
// Sample Cache lookup/status utility (as UDF)
const cache = CacheService.getScriptCache()
const ws = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet()
const row = ws.getActiveCell().getRow(), col = ws.getActiveCell().getColumn()
  if (col == 1) return [["GeoID","GeoName"]]  // convenient to relabel stub (from any row)
  if (col == 2) return  // dis-allowed
  if (row == 1) return [resetCache(1)]  // convenient to reset/monitor cache 
  geoIds = cache.get("A"+row)
  var g = (!geoIds ? ["NO Cache"] : JSON.parse(geoIds)); if (g.length <= 1) return g
  g = g.map((v,i) => { return (i==0 ? "xfetch1("+(x?x:"")+")" : (!x||v==x ? v: "")) })  // match vector
  return g
}
