// @OnlyCurrentDoc
// Date: 30Mar26
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Custom Menu')
      .addItem('Reset Cache', 'resetCache')
      .addToUi();
}


function fa(x) { return [[x, 2*x], [10*x, 2*10*x]] }

const cacheTTL = 300 // requested cache duration

function cacheIn(g) {
  var k, r, v;
  var cache = CacheService.getScriptCache();
  g.length = 0
  g[0] = cache.get("0")
  r =0; while (++r < g[0]) { k = r+""; v = cache.get(k); g[r] = v }
} 
function cacheOut(g) { 
  var k, r, v;
  var cache = CacheService.getScriptCache();
  cache.put("0", g.length+"", cacheTTL)
  r =0; while(++r<g.length) { k = r+""; v = g[r]; cache.put(k, v, cacheTTL) }  
}
var g1s = []

function resetCache() {
const ws = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet()
  g1s.length = 0; reCache("A0"); ws.getRange("B1").setValue(g1s.length+","+g1s.join("|"))
}

function reCache(aR) {
  const cache = CacheService.getScriptCache();
  const ws = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet()
  const GeoIds = [], lR = ws.getLastRow()
  var r = parseInt(aR.slice(1)), r0 = 0, rg, v
  if (r) { rg = ws.getRange("A"+r); rg.activate() }  // confirm to user
  r -= Math.sign(r); while (++r <= lR+1) { 
    rg = ws.getRange(r,1); v = rg.getValue()  // case-sensitive tests
    if (rg.isBlank() || v == "GeoID") {  // break in fetch sequence
      if (r0) { 
        cache.put("A"+r0, JSON.stringify(GeoIds))  // cache prior fetch
        if (parseInt(aR.at(1))) return; // and unless individual
        g1s.push("A"+r0)
        r0 = 0; GeoIds.length = 0  // prepare for next fetch
      } 
    }
    if (!rg.isBlank()) {
      if (v == "GeoID") { r0 = r; GeoIds[0] = v }  // preserve case
      else if (parseInt(v)) GeoIds.push(v)
    }
  }
  SpreadsheetApp.flush()
}


function onEdit(e) {
  const cache = CacheService.getScriptCache();
  const ui = SpreadsheetApp.getUi() //ui.alert("at onEdit "+e.value+"<--"+e.oldValue)
  const ws = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet()
  const row = e.range.getRow(), col = e.range.getColumn(), Ar = "A"+row //= range.getActiveCell().getA1Notation() IDW
  //ui.alert("at "+ row+"\\"+col+": "+(e.value.at(0) == "="))
  if (e.value.at(0) != "=") {
  //ui.alert(!cache.get(Ar))
   if (c = cache.get(Ar)) ui.alert(c)
   else 
    if(ui.alert("reCache "+Ar, ui.ButtonSet.YES_NO)==ui.Button.YES) { reCache(Ar); ui.alert(cache.get(Ar)) }
  }  
  Utilities.sleep(500); ws.getRange(row,col).activate()
}


function fetch() { //return [["geoID","GeoName"]]
const ws = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet()
const row = ws.getActiveCell().getRow()-1
  cacheIn(g1s); g = g1s[row]; return [[g]]
}
//*
function onSelectionChange(e) {
const ws = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet()
const row = e.range.getRow(), col = e.range.getColumn()
if (row == 1 && col == 2) {
  resetCache()
  ws.getRange(row,1).activate()
}  
}
//*/
