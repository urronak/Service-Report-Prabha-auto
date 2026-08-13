const SHEET_ID = "1jenPoMuy0po38FFFMGE9ezInz8sARPknbi-cIXNQqtM";

function doGet(e) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sh = ss.getSheetByName("AppData");
    if (!sh || sh.getLastRow() < 2) {
      return json_({success:true, data:null, updatedAt:null, message:"No synced app data yet"});
    }
    const updatedAt = sh.getRange(2,1).getValue();
    const raw = sh.getRange(2,2).getValue();
    return json_({
      success:true,
      data:raw ? JSON.parse(raw) : null,
      updatedAt: updatedAt ? new Date(updatedAt).toISOString() : null
    });
  } catch (err) {
    return json_({success:false,error:String(err.message||err)});
  }
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const data = JSON.parse((e && e.postData && e.postData.contents) || "{}");
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const now = new Date();

    // Exact app backup for phone/PC restore.
    let app = ss.getSheetByName("AppData") || ss.insertSheet("AppData");
    app.clearContents();
    app.getRange(1,1,1,2).setValues([["Updated At","Database JSON"]]);
    app.getRange(2,1,1,2).setValues([[now, JSON.stringify(data)]]);
    app.setFrozenRows(1);

    writeSales_(ss, data.sales || []);
    writePurchases_(ss, data.purchases || []);
    writePayments_(ss, data.payments || []);
    writeStock_(ss, data);
    writeSuppliers_(ss, data);

    return json_({success:true,message:"Google Sheet sync complete",updatedAt:now.toISOString()});
  } catch (err) {
    return json_({success:false,error:String(err.message||err)});
  } finally {
    try { lock.releaseLock(); } catch (_) {}
  }
}

function writeSales_(ss, sales) {
  const headers=["ID","Invoice","Date","Timestamp","Customer","Tractor","Mechanic/Service","Code","Item","Qty","Rate","Total Amount","Less","Final Amount","Payment Mode","UTR"];
  const rows=[];
  sales.forEach(x => (x.items && x.items.length ? x.items : [{code:"",qty:"",rate:""}]).forEach(i=>rows.push([
    x.id||"",x.invoice||"",x.date||"",x.timestamp||"",x.customer||"",x.tractor||"",x.machine||"",i.code||"",i.name||"",i.qty||"",i.rate||"",x.total||0,x.less||0,x.final||0,x.mode||"",x.utr||""
  ])));
  writeTable_(ss,"Sales",headers,rows);
}

function writePurchases_(ss, purchases) {
  const headers=["ID","Invoice","Date","Timestamp","Supplier","Code","Item","Qty","Rate","Amount","Photo"];
  const rows=[];
  purchases.forEach(x => (x.items && x.items.length ? x.items : [{code:"",qty:""}]).forEach(i=>rows.push([
    x.id||"",x.invoice||"",x.date||"",x.timestamp||"",x.supplier||"",i.code||"",i.name||"",i.qty||"",i.rate||"",x.amount||0,x.photo||""
  ])));
  writeTable_(ss,"Purchases",headers,rows);
}

function writePayments_(ss, payments) {
  const headers=["ID","Date","Timestamp","Party","Type","Amount","Mode","UTR","Note"];
  const rows=(payments||[]).map(x=>[x.id||"",x.date||"",x.timestamp||"",x.party||"",x.type||"",x.amount||0,x.mode||"",x.utr||"",x.note||""]);
  writeTable_(ss,"Payments",headers,rows);
}

function writeStock_(ss, data) {
  const parts=Array.isArray(data.parts)?data.parts:[];
  const deleted=new Set((data.deletedParts||[]).map(String));
  const sales=data.sales||[], purchases=data.purchases||[], issues=data.issues||[];
  const qty={};
  const master={};
  // The app's full parts master is represented by parts-data.js on the client; custom parts are in data.parts.
  parts.forEach(p=>{if(!deleted.has(String(p.code))){master[p.code]=p;qty[p.code]=Number(p.openingStock||0);}});
  // For exact stock-master export, the client sends the computed stock in data.__stock when available.
  if (Array.isArray(data.__stock)) {
    data.__stock.forEach(p=>{master[p.code]=p;qty[p.code]=Number(p.qty||0);});
  } else {
    sales.forEach(x=>(x.items||[]).forEach(i=>{if(qty[i.code]!==undefined)qty[i.code]-=Number(i.qty||0)}));
    purchases.forEach(x=>(x.items||[]).forEach(i=>{if(qty[i.code]!==undefined)qty[i.code]+=Number(i.qty||0)}));
    issues.forEach(i=>{if(qty[i.code]!==undefined)qty[i.code]-=Number(i.qty||0)});
  }
  const headers=["Code","Part Name","HSN","MRP","NDP","GST","Current Stock","Supplier"];
  const rows=Object.keys(master).map(code=>{const p=master[code];return [code,p.name||"",p.hsn||"",p.mrp||0,p.ndp||0,p.gst||"",qty[code]||0,p.supplier||""];});
  writeTable_(ss,"Stock",headers,rows);
}

function writeSuppliers_(ss,data) {
  const suppliers=((data.settings||{}).suppliers||[]);
  const headers=["Supplier","Purchase Total","Payments","Outstanding"];
  const rows=suppliers.map(name=>{
    const purchase=(data.purchases||[]).filter(x=>x.supplier===name).reduce((n,x)=>n+Number(x.amount||0),0);
    const paid=(data.payments||[]).filter(x=>x.party===name&&x.type==="Supplier Payment").reduce((n,x)=>n+Number(x.amount||0),0);
    return [name,purchase,paid,purchase-paid];
  });
  writeTable_(ss,"Suppliers",headers,rows);
}

function writeTable_(ss,name,headers,rows) {
  let sh=ss.getSheetByName(name)||ss.insertSheet(name);
  sh.clearContents();
  sh.getRange(1,1,1,headers.length).setValues([headers]);
  if(rows.length) sh.getRange(2,1,rows.length,headers.length).setValues(rows);
  sh.setFrozenRows(1);
  sh.autoResizeColumns(1,headers.length);
}

function json_(data) {
  return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(ContentService.MimeType.JSON);
}
