(()=>{"use strict";

/* PRABHA AUTO — final local business manager
   Source of truth:
   - Current Stock Excel records = opening stock master
   - PDF/parts catalogue = lookup only
   - Transactions persist in localStorage
*/

const KEY="prabha_auto_final_v3";
const OLD_KEYS=["prabha_auto_final_v2"];
const PARTS=window.PRABHA_PARTS||[];
const MAP=new Map(PARTS.map(p=>[String(p.code).trim().toLowerCase(),p]));
const INVOICE_STOCK=[
  {"code":"93354167","qty":4},{"code":"93816677","qty":2},{"code":"97200637","qty":20},
  {"code":"97201150","qty":20},{"code":"97201155","qty":15},{"code":"97201157","qty":8},{"code":"97201159","qty":2}
];
const DEFAULT_SUPPLIERS=["Ganpati Automobile","Mahaveera Agro"];
const DEFAULT_MACHINES=["Bijandra Yadav","Ajay Kumar","Vijay Patil"];
const ACCESSORIES=[
  {name:"Hitch - 242",type:"Hitch"},
  {name:"Hitch - 380",type:"Hitch"},
  {name:"Batta",type:"Accessory"},
  {name:"Hood Red",type:"Hood"},
  {name:"Hood Blue",type:"Hood"},
  {name:"Hood Silver",type:"Hood"}
];
const $=id=>document.getElementById(id);
const today=()=>{const d=new Date();return new Date(d.getTime()-d.getTimezoneOffset()*60000).toISOString().slice(0,10)};
const money=n=>"₹"+Number(n||0).toLocaleString("en-IN",{minimumFractionDigits:0,maximumFractionDigits:2});
const esc=s=>String(s??"").replace(/[&<>"']/g,c=>({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"}[c]));
const fmt=d=>d?new Date(d+"T00:00:00").toLocaleDateString("en-IN",{day:"2-digit",month:"short",year:"numeric"}):"";
const part=c=>{const k=String(c||"").trim().toLowerCase();const cp=(db?.parts||[]).find(p=>String(p.code).trim().toLowerCase()===k);return cp||MAP.get(k)};
const uid=p=>p+"-"+Date.now()+"-"+Math.random().toString(36).slice(2,8);

function emptyDB(){
 return {version:3,sales:[],issues:[],purchases:[],payments:[],
   settings:{business:"PRABHA AUTO",manager:"Chanchal Kumar",mobile:"9693124739",email:"Urronak2@gmail.com",low:2,machines:[...DEFAULT_MACHINES],suppliers:[...DEFAULT_SUPPLIERS]},
   seeded:{invoiceStock:true},parts:[],deletedParts:[],accessories:[],deletedAccessories:[]};
}
function loadDB(){
 let raw=localStorage.getItem(KEY);
 if(!raw) for(const k of OLD_KEYS){raw=localStorage.getItem(k);if(raw)break}
 let d;
 try{d=raw?JSON.parse(raw):emptyDB()}catch{d=emptyDB()}
 const e=emptyDB();
 d.sales=Array.isArray(d.sales)?d.sales:[];
 d.issues=Array.isArray(d.issues)?d.issues:[];
 d.purchases=Array.isArray(d.purchases)?d.purchases:[];
 d.payments=Array.isArray(d.payments)?d.payments:[];
 d.parts=Array.isArray(d.parts)?d.parts:[];
 d.deletedParts=Array.isArray(d.deletedParts)?d.deletedParts:[];
 d.accessories=Array.isArray(d.accessories)?d.accessories:[];
 d.deletedAccessories=Array.isArray(d.deletedAccessories)?d.deletedAccessories:[];
 d.settings={...e.settings,...(d.settings||{})};
 d.settings.machines=(d.settings.machines||DEFAULT_MACHINES).filter(Boolean);
 d.settings.suppliers=[...new Set([...(d.settings.suppliers||[]),...DEFAULT_SUPPLIERS])];
 // One-time fresh start for Ganpati ledger. User will re-enter old due as purchase and payments manually.
 if(d._ganpatiResetV1!==true){
   d.purchases=d.purchases.filter(x=>x.supplier!=="Ganpati Automobile");
   d.payments=d.payments.filter(x=>!(x.party==="Ganpati Automobile"&&x.type==="Supplier Payment"));
   d._ganpatiResetV1=true;
 }
 return d;
}
let db=loadDB();

function persist(msg="Saved",options={}){
 db._syncUpdatedAt=Date.now();
 localStorage.setItem(KEY,JSON.stringify(db));
 if($("sync")){$("sync").textContent=msg;$("sync").classList.add("saved");setTimeout(()=>{if($("sync"))$("sync").textContent="Local • Saved"},1800)}
 renderAll();
 if(options.sync!==false) scheduleCloudSync();
}
function toast(message,type="success"){
 const t=$("toast");if(!t)return;
 t.textContent=message;t.className="toast "+type;
 clearTimeout(window.__toast);window.__toast=setTimeout(()=>t.className="toast hidden",2600);
}
function closeModal(){if($("modal"))$("modal").classList.add("hidden");if($("modalBody"))$("modalBody").innerHTML=""}
function openModal(html,title=""){
 $("modalTitle").textContent=title||"";$("modalBody").innerHTML=html;$("modal").classList.remove("hidden");
}
function table(rows,cols){
 if(!rows||!rows.length)return `<div class="empty">No data found</div>`;
 const keys=cols||Object.keys(rows[0]);
 return `<div class="table"><table><thead><tr>${keys.map(k=>`<th>${esc(String(k).replace(/([A-Z])/g," $1").replace(/^./,x=>x.toUpperCase()))}</th>`).join("")}</tr></thead><tbody>${
 rows.map(r=>`<tr>${keys.map(k=>`<td>${esc(r[k]===undefined?"":typeof r[k]==="number"?money(r[k]):r[k])}</td>`).join("")}</tr>`).join("")
 }</tbody></table></div>`;
}
function excel(rows,name){
 if(!rows.length){toast("No data to export","error");return}
 const keys=Object.keys(rows[0]);
 const html=`<html><head><meta charset="UTF-8"></head><body><table border="1"><tr>${keys.map(k=>`<th>${esc(k)}</th>`).join("")}</tr>${rows.map(r=>`<tr>${keys.map(k=>`<td>${esc(r[k]??"")}</td>`).join("")}</tr>`).join("")}</table></body></html>`;
 const blob=new Blob([html],{type:"application/vnd.ms-excel;charset=utf-8"});
 const a=document.createElement("a");a.href=URL.createObjectURL(blob);a.download=`${name}-${today()}.xls`;a.click();setTimeout(()=>URL.revokeObjectURL(a.href),1000);
 toast("Excel file exported");
}
function printHTML(body,title="PRABHA AUTO"){
 body=String(body).replaceAll('src="prabha-auto-logo.png"',`src="${new URL("prabha-auto-logo.png",location.href).href}"`);
 const w=window.open("","_blank","width=1000,height=800");
 if(!w){toast("Popup blocked. Allow popups for localhost.","error");return}
 w.document.write(`<!doctype html><html><head><meta charset="utf-8"><title>${esc(title)}</title><style>
 @page{size:A4;margin:12mm}*{box-sizing:border-box}body{font-family:Arial,sans-serif;color:#17213c;font-size:12px;margin:0}
 .print-head{display:flex;align-items:center;gap:14px;border-bottom:2px solid #d71920;padding-bottom:10px;margin-bottom:14px}
 .print-head img{width:72px;height:62px;object-fit:contain}.print-head h1{margin:0;font-size:22px}.muted{color:#68758c}
 table{width:100%;border-collapse:collapse;margin-top:12px}th,td{border:1px solid #cfd6e1;padding:7px;text-align:left}th{background:#f1f4f8}
 .right{text-align:right}.total{font-size:14px;font-weight:700;margin-top:14px}.footer{margin-top:28px;border-top:1px solid #ddd;padding-top:8px}
 @media print{button{display:none!important}}
 </style></head><body>${body}</body></html>`);
 w.document.close();w.focus();setTimeout(()=>w.print(),250);
}
function printRows(rows,title){
 if(!rows.length){toast("No data to print","error");return}
 const keys=Object.keys(rows[0]);
 printHTML(`<div class="print-head"><img src="prabha-auto-logo.png"><div><h1>PRABHA AUTO</h1><div>Authorised Dealer • EICHER TRACTORS</div><div class="muted">Service Manager: ${esc(db.settings.manager||"Chanchal Kumar")} • ${esc(db.settings.mobile||"9693124739")}</div></div></div><h2>${esc(title)}</h2>${table(rows,keys)}<div class="footer">Generated on ${fmt(today())}</div>`,title);
}

function ensurePurchasePart(code){
 const key=String(code||"").trim().toLowerCase();
 if(!key)return null;
 const existing=part(key);
 const current=existing&&stock()[existing.code];
 if(current)return current;
 const master=existing||MAP.get(key);
 if(!master)return null;
 const p={code:String(master.code).trim(),name:master.name||"",hsn:master.hsn||"",mrp:Number(master.mrp||0),ndp:Number(master.ndp||0),gst:master.gst||"0",openingStock:0,stockMaster:true,supplier:master.supplier||"",source:"purchase-auto-added"};
 db.parts=db.parts||[];
 db.parts=db.parts.filter(x=>String(x.code).trim().toLowerCase()!==key);
 db.parts.push(p);
 db.deletedParts=(db.deletedParts||[]).filter(x=>String(x).trim().toLowerCase()!==key);
 return p;
}


const DEFAULT_ACCESSORIES=[
 {id:"acc-h242",code:"ACC-H242",name:"HITCH - 242",category:"Hitch",mrp:0,ndp:0,gst:"0%",qty:0,supplier:"Mahaveera Agro"},
 {id:"acc-h380",code:"ACC-H380",name:"HITCH - 380",category:"Hitch",mrp:0,ndp:0,gst:"0%",qty:0,supplier:"Mahaveera Agro"},
 {id:"acc-batta",code:"ACC-BATTA",name:"BATTA",category:"Accessory",mrp:0,ndp:0,gst:"0%",qty:0,supplier:"Mahaveera Agro"},
 {id:"acc-hood-red",code:"ACC-HOOD-RED",name:"HOOD - RED",category:"Hood",mrp:0,ndp:0,gst:"0%",qty:0,supplier:"Mahaveera Agro"},
 {id:"acc-hood-blue",code:"ACC-HOOD-BLUE",name:"HOOD - BLUE",category:"Hood",mrp:0,ndp:0,gst:"0%",qty:0,supplier:"Mahaveera Agro"},
 {id:"acc-hood-silver",code:"ACC-HOOD-SILVER",name:"HOOD - SILVER",category:"Hood",mrp:0,ndp:0,gst:"0%",qty:0,supplier:"Mahaveera Agro"}
];
function ensureAccessories(){
 const existing=new Map((db.accessories||[]).map(x=>[String(x.id||x.code||"").toLowerCase(),x]));
 DEFAULT_ACCESSORIES.forEach(d=>{if(!existing.has(d.id.toLowerCase())&&!db.deletedAccessories.includes(d.id)) db.accessories.push({...d});});
 return db.accessories;
}
function accessoryRows(){return ensureAccessories().filter(x=>!x.deleted);}
function accessoryExportRows(){return accessoryRows().map(x=>({Code:x.code||"",Item_Name:x.name||"",Category:x.category||"",MRP:Number(x.mrp||0),NDP:Number(x.ndp||0),GST:x.gst||"",Current_Stock:Number(x.qty||0),Supplier:x.supplier||""}));}
function accessoryPrint(){
 const rows=accessoryExportRows().map(x=>({...x,MRP:money(x.MRP),NDP:money(x.NDP),Current_Stock:String(x.Current_Stock)}));
 if(!rows.length)return toast("No accessories to print","error");
 printRows(rows,"PRABHA AUTO ACCESSORIES");
}
function accessoryModal(id){
 ensureAccessories();
 const x=id?db.accessories.find(v=>v.id===id):null;
 openModal(`<div class="form"><h2>${x?"✏️ Edit Accessory":"+ Add Accessory"}</h2>
 <div class="grid2"><label>Item Name<input id="acName" value="${esc(x?.name||"")}" placeholder="Hitch / Hood / Batta"></label><label>Code<input id="acCode" value="${esc(x?.code||"")}" placeholder="Accessory code"></label>
 <label>Category<input id="acCategory" value="${esc(x?.category||"")}" placeholder="Hitch / Hood / Accessory"></label><label>Supplier<select id="acSupplier">${(db.settings.suppliers||[]).map(s=>`<option ${s===(x?.supplier||"Mahaveera Agro")?"selected":""}>${esc(s)}</option>`).join("")}</select></label>
 <label>MRP<input id="acMrp" type="number" step=".01" value="${Number(x?.mrp||0)}"></label><label>NDP<input id="acNdp" type="number" step=".01" value="${Number(x?.ndp||0)}"></label>
 <label>GST %<input id="acGst" value="${esc(x?.gst||"0%")}"></label><label>Current Quantity<input id="acQty" type="number" min="0" step="1" value="${Number(x?.qty||0)}"></label></div>
 <div class="actions"><button id="acSave" class="btn primary">Save</button><button id="acCancel" class="btn">Cancel</button></div></div>`);
 $("acCancel").onclick=closeModal;
 $("acSave").onclick=()=>{
   const name=$("acName").value.trim(),code=$("acCode").value.trim();
   if(!name||!code)return toast("Item Name and Code required","error");
   const duplicate=db.accessories.some(v=>v.id!==id&&!v.deleted&&String(v.code).toLowerCase()===code.toLowerCase());
   if(duplicate)return toast("Accessory code already exists","error");
   const item={id:id||uid("acc"),code,name,category:$("acCategory").value.trim()||"Accessory",mrp:+$("acMrp").value||0,ndp:+$("acNdp").value||0,gst:$("acGst").value.trim()||"0%",qty:+$("acQty").value||0,supplier:$("acSupplier").value||"Mahaveera Agro"};
   if(id){const i=db.accessories.findIndex(v=>v.id===id);db.accessories[i]=item}else db.accessories.push(item);
   persist("Accessories Saved");closeModal();toast(id?"Accessory updated":"Accessory added");
 };
}
function accessoryView(){
 const list=accessoryRows();
 if($("accessoryTable")) $("accessoryTable").innerHTML=list.length?`<div class="table"><table><thead><tr><th>Code</th><th>Item Name</th><th>Category</th><th>MRP</th><th>NDP</th><th>GST</th><th>Stock</th><th>Supplier</th><th>Action</th></tr></thead><tbody>${list.map(x=>`<tr><td>${esc(x.code)}</td><td><b class="accessory-name">${esc(x.name)}</b></td><td>${esc(x.category)}</td><td>${money(x.mrp)}</td><td>${money(x.ndp)}</td><td>${esc(x.gst)}</td><td><b>${Number(x.qty||0)}</b></td><td>${esc(x.supplier||"")}</td><td><button class="mini" onclick="window.editAccessory('${esc(x.id)}')">✏️ Edit</button> <button class="mini dangerMini" onclick="window.deleteAccessory('${esc(x.id)}')">🗑️ Delete</button></td></tr>`).join("")}</tbody></table></div>`:`<div class="empty">No accessories</div>`;
}
window.editAccessory=id=>accessoryModal(id);
window.deleteAccessory=id=>{const x=db.accessories.find(v=>v.id===id);if(!x)return;if(!confirm(`Delete ${x.name}?`))return;db.accessories=db.accessories.filter(v=>v.id!==id);if(!db.deletedAccessories.includes(id))db.deletedAccessories.push(id);persist("Accessory Deleted");toast("Accessory deleted")};
function stock(){
 const s={};
 PARTS.filter(p=>p.stockMaster===true&&!db.deletedParts.includes(String(p.code))).forEach(p=>s[String(p.code)]={...p,qty:Number(p.openingStock||0)});
 (db.parts||[]).forEach(p=>{if(db.deletedParts.includes(String(p.code)))return;s[String(p.code)]={...s[String(p.code)],...p,qty:Number(p.openingStock??s[String(p.code)]?.qty??0),stockMaster:true}});
 if(db.seeded?.invoiceStock!==false) INVOICE_STOCK.forEach(i=>{if(s[i.code])s[i.code].qty+=Number(i.qty||0)});
 db.purchases.forEach(x=>(x.items||[]).forEach(i=>{if(s[i.code])s[i.code].qty+=Number(i.qty||0)}));
 db.sales.forEach(x=>(x.items||[]).forEach(i=>{if(s[i.code])s[i.code].qty-=Number(i.qty||0)}));
 db.issues.forEach(x=>{if(s[x.code])s[x.code].qty-=Number(x.qty||0)});
 return s;
}
function rows(){return Object.values(stock())}
function availableRows(){return rows().filter(x=>x.qty>0)}
function zeroRows(){return rows().filter(x=>x.qty<=0)}
function lowRows(){return rows().filter(x=>x.qty>0&&x.qty<=Number(db.settings.low??2))}
function currentStock(){return stock()}
function save(msg="Saved"){persist(msg)}
const low=()=>lowRows();
const zero=()=>zeroRows();

function nav(page){

  // Page show/hide
  document.querySelectorAll(".page").forEach(x => {
    x.classList.toggle("active", x.id === page);
  });

  // Desktop + other navigation
  document.querySelectorAll("[data-page]").forEach(x => {
    x.classList.toggle("active", x.dataset.page === page);
  });

  // MOBILE BOTTOM NAV — direct visual active state
  document.querySelectorAll("nav.bottom button[data-page]").forEach(btn => {

    const isActive = btn.dataset.page === page;

    btn.classList.toggle("active", isActive);
    btn.setAttribute("aria-current", isActive ? "page" : "false");

    // Direct styling — CSS/cache issue won't affect this
    if (isActive) {
      btn.style.background = "rgba(200, 16, 46, 0.14)";
      btn.style.color = "#c8102e";
      btn.style.borderRadius = "16px";
      btn.style.boxShadow = "0 4px 12px rgba(200,16,46,0.16)";
      btn.style.transform = "translateY(-3px)";
    } else {
      btn.style.background = "";
      btn.style.color = "";
      btn.style.borderRadius = "";
      btn.style.boxShadow = "";
      btn.style.transform = "";
    }
  });

  window.scrollTo({
    top: 0,
    behavior: "smooth"
  });
}
document.querySelectorAll("[data-page]").forEach(x=>x.onclick=()=>nav(x.dataset.page));
// ===== MOBILE MORE MENU FIX =====
(function initMobileMore(){

  const moreBtn = document.getElementById("mobileMore");
  const sheet = document.getElementById("mobileMoreSheet");
  const closeBtn = document.getElementById("mobileMoreX");
  const backdrop = document.getElementById("mobileMoreClose");

  if (!moreBtn || !sheet) return;

  const openMore = (e) => {
    e.preventDefault();
    e.stopPropagation();
    sheet.classList.remove("hidden");
  };

  const closeMore = (e) => {
    if (e) {
      e.preventDefault();
      e.stopPropagation();
    }
    sheet.classList.add("hidden");
  };

  moreBtn.addEventListener("click", openMore);
  closeBtn?.addEventListener("click", closeMore);
  backdrop?.addEventListener("click", closeMore);

  sheet.querySelectorAll("[data-page]").forEach(btn => {
    btn.addEventListener("click", () => {
      sheet.classList.add("hidden");
      nav(btn.dataset.page);
    });
  });

})();
// ===== END MOBILE MORE MENU FIX =====
document.addEventListener("pointerdown",e=>{const b=e.target.closest("button");if(b)b.classList.add("tap-active");});
document.addEventListener("pointerup",e=>{const b=e.target.closest("button");if(b)setTimeout(()=>b.classList.remove("tap-active"),180);});
document.addEventListener("pointercancel",e=>{const b=e.target.closest("button");if(b)b.classList.remove("tap-active");});

$("logout").onclick=()=>{localStorage.removeItem("pa");sessionStorage.removeItem("pa");location.reload()};
$("close").onclick=closeModal;
$("modal").addEventListener("click",e=>{if(e.target.id==="modal")closeModal()});

async function hash(s){return [...new Uint8Array(await crypto.subtle.digest("SHA-256",new TextEncoder().encode(s)))].map(x=>x.toString(16).padStart(2,"0")).join("")}
const IH="6e83e6ec8db725636c8f7a4f73b0910b0a41f4ec0c59ddfed91cfbc9d810ac7e", PH="e8d3d60f077c4afc2c21ac5b0414d7550ea180d2b0f0af8234e01203c5640f1c";
$("loginForm").onsubmit=async e=>{e.preventDefault();if(await hash($("loginId").value.trim())===IH&&await hash($("loginPw").value)===PH){localStorage.setItem("pa","1");sessionStorage.setItem("pa","1");$("login").classList.add("hidden");$("app").classList.remove("hidden");renderAll();startCloudSync();toast("Login successful")}else $("loginMsg").textContent="Wrong Login ID or Password"};
if(localStorage.getItem("pa")==="1"||sessionStorage.getItem("pa")==="1"){$("login").classList.add("hidden");$("app").classList.remove("hidden")}

function dateRange(select,from,to){
 const d=new Date(),t=today();
 if(select==="today")return[t,t];
 if(select==="month")return[new Date(d.getFullYear(),d.getMonth(),1).toISOString().slice(0,10),t];
 if(select==="year")return[new Date(d.getFullYear(),0,1).toISOString().slice(0,10),t];
 return[from||t,to||t];
}
const between=(d,a,b)=>String(d||"")>=a&&String(d||"")<=b;
const total=(a,k)=>a.reduce((n,x)=>n+Number(x[k]||0),0);
function dueTotal(){
 const m={};
 db.sales.filter(x=>x.mode==="Credit").forEach(x=>m[x.customer]=(m[x.customer]||0)+Number(x.final||0));
 db.payments.filter(x=>x.type==="Customer Receipt").forEach(x=>m[x.party]=(m[x.party]||0)-Number(x.amount||0));
 return Object.values(m).filter(v=>v>0).reduce((a,b)=>a+b,0);
}
function dashboard(){
 const [a,b]=dateRange($("dashRange").value,$("dashFrom").value,$("dashTo").value);
 const sales=db.sales.filter(x=>between(x.date,a,b));
 const purchases=db.purchases.filter(x=>between(x.date,a,b));
 const payments=db.payments.filter(x=>between(x.date,a,b));
 const sale=total(sales,"final");
 const cash=total(sales.filter(x=>x.mode==="Cash"),"final")+total(payments.filter(x=>x.type==="Customer Receipt"&&x.mode==="Cash"),"amount");
 const bank=total(sales.filter(x=>x.mode==="Bank"),"final")+total(payments.filter(x=>x.type==="Customer Receipt"&&x.mode==="Bank"),"amount");
 const month=today().slice(0,7);
 const monthPurchase=total(db.purchases.filter(x=>String(x.date||"").slice(0,7)===month),"amount");
 const purchase=total(purchases,"amount");
 const profit=sale-purchase;
 const serviceCount=sales.length;
 $("dSale").textContent=money(sale);$("dCash").textContent=money(cash);$("dBank").textContent=money(bank);$("dDue").textContent=money(dueTotal());
 $("dGanpati").textContent=money(ganpatiBalance());$("dPurchase").textContent=money(monthPurchase);$("dProfit").textContent=String(serviceCount);$("profitCard").className="k cyan";
 $("dLow").textContent=low().length;
 const mech={};sales.forEach(x=>{const m=x.machine||"Unassigned";if(!mech[m])mech[m]={count:0,revenue:0};mech[m].count+=1;mech[m].revenue+=Number(x.final||0)});
 const mm=Object.entries(mech).sort((a,b)=>b[1].revenue-a[1].revenue);
 $("mechanics").innerHTML=mm.length?mm.map(x=>`<div class="bar"><span>${esc(x[0])} • ${x[1].count} Service</span><b>${money(x[1].revenue)}</b><i style="--w:${Math.max(4,x[1].revenue/(mm[0][1].revenue||1)*100)}%"></i></div>`).join(""):`<div class="empty">No service for this period</div>`;
 const s=Object.values(stock()),av=s.filter(x=>x.qty>0).length,z=s.filter(x=>x.qty<=0).length,l=s.filter(x=>x.qty>0&&x.qty<=Number(db.settings.low??2)).length;
 $("stockStatus").innerHTML=`<p>🟢 In Stock <b>${av}</b></p><p>🟠 Low Stock <b>${l}</b></p><p>🔴 Out of Stock <b>${z}</b></p><small>${s.length.toLocaleString()} parts in Current Stock Excel</small>`;
 const rec=[...db.sales.map(x=>({d:x.date,t:"SALE",n:x.customer,a:x.final})),...db.purchases.map(x=>({d:x.date,t:"PURCHASE",n:x.supplier,a:x.amount})),...db.payments.map(x=>({d:x.date,t:"PAYMENT",n:x.party,a:x.amount}))].sort((a,b)=>String(b.d).localeCompare(String(a.d))).slice(0,7);
 $("recent").innerHTML=rec.length?rec.map(x=>`<div class="recent"><span>${esc(x.t)}</span><b>${esc(x.n||"-")}</b><strong>${money(x.a)}</strong></div>`).join(""):"<div class=empty>No transactions</div>";
 renderDashboardSuppliers();
 $("badge").textContent=low().length+zero().length;


}

function renderDashboardSuppliers(){
 const el=$("dashboardSuppliers");
 if(!el)return;
 const suppliers=[...new Set([...(db.settings.suppliers||[]),...DEFAULT_SUPPLIERS])];
 el.innerHTML=suppliers.map(name=>{
   const purchases=db.purchases.filter(x=>x.supplier===name);
   const payments=db.payments.filter(x=>x.party===name&&x.type==="Supplier Payment");
   const purchaseTotal=total(purchases,"amount");
   const paidTotal=total(payments,"amount");
   const balance=paidTotal-purchaseTotal;
   const label=balance>=0?"Credit / Advance":"Outstanding Due";
   return `<button type="button" class="supplier-summary-card" data-supplier="${esc(name)}"><b>${esc(name)}</b><span>Purchase ${money(purchaseTotal)} • Paid ${money(paidTotal)}</span><strong>${label}: ${money(Math.abs(balance))}</strong></button>`;
 }).join("");
 el.querySelectorAll("[data-supplier]").forEach(b=>b.onclick=()=>{if(typeof window.supplierDetail==="function")window.supplierDetail(b.dataset.supplier);else nav("suppliers")});
}

function fillMachines(){
 const m=db.settings.machines?.length?db.settings.machines:DEFAULT_MACHINES;
 $("saleMachine").innerHTML=m.map(x=>`<option>${esc(x)}</option>`).join("");
 $("settingLow").value=db.settings.low??2;$("lowLimit").value=db.settings.low??2;
 $("business").value=db.settings.business??"PRABHA AUTO";$("manager").value=db.settings.manager??"Chanchal Kumar";$("mobile").value=db.settings.mobile??"9693124739";$("email").value=db.settings.email??"Urronak2@gmail.com";
 ["m1","m2","m3"].forEach((id,i)=>$(id).value=m[i]||"");
}

function recalcSale(){
 let t=0;
 document.querySelectorAll(".item").forEach(e=>{
   const p=part(e.querySelector(".icode").value),q=Number(e.querySelector(".iqty").value)||0;
   const v=p?q*Number(p.mrp||0):0;t+=v;
   e.querySelector(".iamt").textContent=money(v);
 });
 $("saleTotal").value=t.toFixed(2);$("saleFinal").value=Math.max(0,t-(Number($("saleLess").value)||0)).toFixed(2);
}
function addItem(code="",qty=1){
 const el=document.createElement("div");el.className="item";
 el.innerHTML=`<input class="icode" placeholder="Part Code" value="${esc(code)}"><div class="idesc"><b>Enter part code</b><small>Search by code</small></div><input class="iqty" type="number" min="1" value="${qty}"><b class="iamt">₹0</b><button class="del" title="Remove">×</button>`;
 $("items").appendChild(el);
 el.querySelector(".icode").oninput=()=>fillItem(el);
 el.querySelector(".iqty").oninput=recalcSale;
 el.querySelector(".del").onclick=()=>{el.remove();recalcSale()};
 if(code)fillItem(el);
}
function fillItem(el){
 const p=part(el.querySelector(".icode").value),s=stock();
 if(!p){el.querySelector(".idesc").innerHTML="<b>Part not found</b><small>Check code from Stock</small>";recalcSale();return}
 const cur=s[p.code];
 el.querySelector(".idesc").innerHTML=`<b>${esc(p.name)}</b><small>HSN ${esc(p.hsn||"-")} • MRP ${money(p.mrp)} • NDP ${money(p.ndp)} • GST ${esc(p.gst||"-")} • ${cur?"Current Stock: "+cur.qty:"PDF lookup only — not in Current Stock Excel"}</small>`;
 recalcSale();
}
function resetSale(){
 $("saleCustomer").value="";$("saleTractor").value="";$("saleDate").value=today();$("saleLess").value="0";$("saleMode").value="Cash";$("saleUtr").value="";$("utrBox").classList.add("hidden");$("items").innerHTML="";addItem();recalcSale();
 $("saleMsg").textContent="New sale ready";$("saveSale").disabled=false;
}
$("addItem").onclick=()=>addItem();
$("saleLess").oninput=recalcSale;
$("saleMode").onchange=()=>{$("utrBox").classList.toggle("hidden",$("saleMode").value!=="Bank")};

$("saveSale").onclick=()=>{
 const btn=$("saveSale");if(btn.disabled)return;
 const rows=[...document.querySelectorAll(".item")];
 const items=rows.map(e=>{const code=e.querySelector(".icode").value.trim(),q=Number(e.querySelector(".iqty").value)||0,p=part(code);return {code,qty:q,rate:Number(p?.mrp||0)}}).filter(x=>x.code);
 if(!$("saleCustomer").value.trim())return toast("Customer name required","error");
 if(!items.length)return toast("Add at least one part","error");
 const s=stock();
 for(const i of items){const p=part(i.code);if(!p)return toast("Invalid Part Code: "+i.code,"error");if(!s[p.code])return toast("This part is PDF lookup only, not in Current Stock Excel","error");if(i.qty>s[p.code].qty)return toast(`Insufficient stock: ${p.name} (Current ${s[p.code].qty})`,"error")}
 if($("saleMode").value==="Bank"&&!$("saleUtr").value.trim())return toast("Bank sale requires UTR No.","error");
 const x={id:uid("sale"),invoice:uid("INV"),timestamp:new Date().toISOString(),date:$("saleDate").value||today(),customer:$("saleCustomer").value.trim(),tractor:$("saleTractor").value.trim(),machine:$("saleMachine").value,items,total:Number($("saleTotal").value)||0,less:Number($("saleLess").value)||0,final:Number($("saleFinal").value)||0,mode:$("saleMode").value,utr:$("saleUtr").value.trim()};
 btn.disabled=true;db.sales.push(x);persist("Sale Saved");showSuccess("Sale saved successfully",`Invoice ${x.invoice}`);resetSale();setTimeout(()=>nav("sale"),400);
};
function showSuccess(title,sub){
 openModal(`<div class="success-box"><div class="success-icon">✓</div><h2>${esc(title)}</h2><p>${esc(sub)}</p><div class="actions"><button class="btn primary" id="successInvoice">View Invoice</button><button class="btn" id="successClose">Close</button></div></div>`);
 $("successInvoice").onclick=()=>{closeModal();showInvoice(db.sales.at(-1))};$("successClose").onclick=closeModal;
}

function invoiceBody(x){
 const rows=(x.items||[]).map(i=>{const p=part(i.code);const rate=Number(i.rate??p?.mrp??0);return `<tr><td>${esc(i.code)}</td><td>${esc(p?.name||"")}</td><td>${i.qty}</td><td>${money(rate)}</td><td>${money(Number(i.qty)*rate)}</td></tr>`}).join("");
 return `<div class="print-head"><img src="prabha-auto-logo.png"><div><h1>PRABHA AUTO</h1><div>Authorised Dealer • EICHER TRACTORS</div><div class="muted">Service Manager: ${esc(db.settings.manager||"Chanchal Kumar")} • ${esc(db.settings.mobile||"9693124739")}</div></div></div>
 <div class="invoice-meta"><b>SALE INVOICE</b><span>Invoice: ${esc(x.invoice||x.id)}</span><span>Date: ${fmt(x.date)}</span></div>
 <p><b>Customer:</b> ${esc(x.customer||"-")} &nbsp; <b>Tractor:</b> ${esc(x.tractor||"-")} &nbsp; <b>Machine:</b> ${esc(x.machine||"-")}</p>
 <table><thead><tr><th>Code</th><th>Part</th><th>Qty</th><th>Rate</th><th>Amount</th></tr></thead><tbody>${rows}</tbody></table>
 <div class="invoice-total"><div>Total: ${money(x.total)}</div><div>Less: ${money(x.less)}</div><strong>Final: ${money(x.final)}</strong></div>
 <div class="footer"><div>Payment Mode: ${esc(x.mode)}${x.utr?` • UTR: ${esc(x.utr)}`:""}</div><div>Service Manager: ${esc(db.settings.manager||"Chanchal Kumar")} • ${esc(db.settings.mobile||"9693124739")}</div></div>`;
}
function showInvoice(x){
 if(!x){toast("No saved sale invoice yet","error");return}
 openModal(`${invoiceBody(x)}<div class="actions invoice-actions"><button id="printInv" class="btn primary">Print / Save PDF (A4)</button><button id="closeInv" class="btn">Close</button></div>`);
 $("printInv").onclick=()=>printHTML(invoiceBody(x),"PRABHA AUTO SALE INVOICE");
 $("closeInv").onclick=closeModal;
}
$("lastInvoice").onclick=()=>showInvoice(db.sales.at(-1));
$("shareSale").onclick=()=>{const x=db.sales.at(-1);if(!x)return toast("Save a sale first","error");const text=encodeURIComponent(["PRABHA AUTO","SALE RECEIPT",`Invoice: ${x.invoice}`,`Customer: ${x.customer}`,`Date: ${fmt(x.date)}`,`Machine: ${x.machine}`,`Tractor: ${x.tractor}`,...x.items.map(i=>`${i.code} x ${i.qty} @ ${money(i.rate)}`),`Total: ${money(x.total)}`,`Less: ${money(x.less)}`,`Final: ${money(x.final)}`,`Payment: ${x.mode}`,x.utr?`UTR: ${x.utr}`:""].filter(Boolean).join("\n"));window.open("https://wa.me/?text="+text,"_blank")};

function saleHistory(){
 const q=($("saleSearch").value||"").toLowerCase();
 const rows=db.sales.filter(x=>(`${x.customer} ${x.tractor} ${x.machine} ${x.invoice||""} ${(x.items||[]).map(i=>i.code).join(" ")}`).toLowerCase().includes(q)).sort((a,b)=>String(b.date).localeCompare(String(a.date)));
 $("saleHistory").innerHTML=rows.length?`<div class="table"><table><thead><tr><th>Date</th><th>Invoice</th><th>Customer</th><th>Machine</th><th>Final</th><th>Mode</th><th>Action</th></tr></thead><tbody>${rows.map(x=>`<tr><td>${fmt(x.date)}</td><td>${esc(x.invoice||x.id)}</td><td>${esc(x.customer)}</td><td>${esc(x.machine)}</td><td>${money(x.final)}</td><td>${esc(x.mode)}</td><td><button class="mini" onclick="window.openSaleInvoice('${esc(x.id)}')">Invoice</button> <button class="mini" onclick="window.editSaleInvoice('${esc(x.id)}')">✏️ Edit</button> <button class="mini dangerMini" onclick="window.deleteSaleInvoice('${esc(x.id)}')">🗑️ Delete</button></td></tr>`).join("")}</tbody></table></div>`:`<div class=empty>No sales history</div>`;
}
window.openSaleInvoice=id=>showInvoice(db.sales.find(x=>x.id===id));

function editSaleInvoice(id){
 const x=db.sales.find(v=>v.id===id);if(!x)return alert("Invoice not found.");
 openModal(`<div class="form"><h2>✏️ Edit Invoice ${esc(x.invoice||"")}</h2>
 <div class="grid2"><label>Customer Name<input id="eiCustomer" value="${esc(x.customer)}"></label><label>Tractor Model<input id="eiTractor" value="${esc(x.tractor||"")}"></label><label>Date<input id="eiDate" type=date value="${esc(x.date)}"></label><label>Machine<select id="eiMachine">${db.settings.machines.map(m=>`<option ${m===x.machine?"selected":""}>${esc(m)}</option>`).join("")}</select></label></div>
 <h3>Items</h3><div id="eiItems"></div><button id="eiAdd" class="btn">+ Add Item</button>
 <div class="grid2"><label>Less / Discount<input id="eiLess" type=number step=.01 value="${Number(x.less||0)}"></label><label>Payment Mode<select id="eiMode"><option ${x.mode==="Cash"?"selected":""}>Cash</option><option ${x.mode==="Bank"?"selected":""}>Bank</option><option ${x.mode==="Credit"?"selected":""}>Credit</option></select></label><label>UTR<input id="eiUtr" value="${esc(x.utr||"")}"></label></div>
 <div id="eiTotal" class="edit-total"></div><div class="actions"><button id="eiSave" class="btn primary">Save Changes</button><button id="eiCancel" class="btn">Cancel</button></div></div>`);
 const add=(i={code:"",qty:1,rate:0})=>{const d=document.createElement("div");d.className="purchase-line edit-sale-line";d.innerHTML=`<input class="eicode" placeholder="Part Code" value="${esc(i.code)}"><input class="eiqty" type=number min=1 value="${Number(i.qty||1)}"><input class="eirate" type=number min=0 step=.01 value="${Number(i.rate||0)}"><span class="eipart muted">${esc(part(i.code)?.name||"")}</span><button class="del" type=button>×</button>`;$("eiItems").appendChild(d);d.querySelector(".eicode").oninput=()=>{d.querySelector(".eipart").textContent=part(d.querySelector(".eicode").value)?.name||"Part not found";calc()};d.querySelector(".eiqty").oninput=calc;d.querySelector(".eirate").oninput=calc;d.querySelector(".del").onclick=()=>{d.remove();calc()};};
 x.items.forEach(add);$("eiAdd").onclick=()=>add();$("eiLess").oninput=calc;
 function calc(){let t=0;document.querySelectorAll(".edit-sale-line").forEach(d=>t+=(+d.querySelector(".eiqty").value||0)*(+d.querySelector(".eirate").value||0));$("eiTotal").textContent=`Total ${money(t)} • Final ${money(Math.max(0,t-(+$("eiLess").value||0)))}`}
 calc();
 $("eiCancel").onclick=closeModal;
 $("eiSave").onclick=()=>{
   const items=[...document.querySelectorAll(".edit-sale-line")].map(d=>({code:d.querySelector(".eicode").value.trim(),qty:+d.querySelector(".eiqty").value||0,rate:+d.querySelector(".eirate").value||0})).filter(i=>i.code);
   if(!items.length)return alert("At least one item is required.");
   const idx=db.sales.findIndex(v=>v.id===id),old=db.sales[idx];db.sales.splice(idx,1);
   const st=currentStock();let error="";
   for(const i of items){const p=part(i.code);if(!p){error="Invalid Part Code: "+i.code;break}if(!st[p.code]){error="Part is not in Current Stock: "+i.code;break}if(i.qty>st[p.code].qty){error=`Insufficient stock for ${p.name}. Available: ${st[p.code].qty}`;break}}
   if(error){db.sales.splice(idx,0,old);return alert(error)}
   const total=items.reduce((n,i)=>n+i.qty*i.rate,0),less=+$("eiLess").value||0;
   db.sales.splice(idx,0,{...old,customer:$("eiCustomer").value.trim(),tractor:$("eiTractor").value.trim(),date:$("eiDate").value,machine:$("eiMachine").value,items,total,less,final:Math.max(0,total-less),mode:$("eiMode").value,utr:$("eiUtr").value.trim()});
   save();closeModal();toast("Invoice updated successfully","success");
 };
}
function deleteSaleInvoice(id){
 const x=db.sales.find(v=>v.id===id);if(!x)return alert("Invoice not found.");
 if(!confirm(`Delete ${x.invoice||"this invoice"}?\\nSold stock will be restored automatically.`))return;
 db.sales=db.sales.filter(v=>v.id!==id);save();toast("Invoice deleted and stock restored","success");
}
window.editSaleInvoice=editSaleInvoice;window.deleteSaleInvoice=deleteSaleInvoice;


function stockView(){
 let all=rows(),q=($("stockSearch").value||"").trim().toLowerCase(),f=$("stockFilter").value;
 let filtered=all.filter(x=>(!q||String(x.code).toLowerCase()===q||String(x.name).toLowerCase().includes(q))&&(f==="All"||(f==="Available"&&x.qty>0)||(f==="Low"&&x.qty>0&&x.qty<=Number(db.settings.low||2))||(f==="Zero"&&x.qty<=0)));
 $("sp").textContent=all.length;$("sa").textContent=availableRows().length;$("sz").textContent=zeroRows().length;$("sl").textContent=lowRows().length;$("sq").textContent=all.reduce((n,x)=>n+Number(x.qty||0),0);
 if(!filtered.length){$("stockTable").innerHTML='<div class="empty">No data found.</div>';return}
 $("stockTable").innerHTML=`<div class="table"><table><thead><tr><th>Select</th><th>Code</th><th>Part Name</th><th>HSN</th><th>MRP</th><th>NDP</th><th>GST</th><th>Current Stock</th><th>Action</th></tr></thead><tbody>${
 filtered.map(x=>`<tr><td><input class="partPick" type="radio" name="partPick" value="${esc(x.code)}"></td><td>${esc(x.code)}</td><td>${esc(x.name)}</td><td>${esc(x.hsn||"-")}</td><td>${money(x.mrp)}</td><td>${money(x.ndp)}</td><td>${esc(x.gst||"-")}</td><td><b>${Number(x.qty||0)}</b></td><td><button class="mini editPartRow" data-code="${esc(x.code)}">✏️ Edit</button></td></tr>`).join("")
 }</tbody></table></div>`;
 document.querySelectorAll(".editPartRow").forEach(b=>b.onclick=()=>editPart(b.dataset.code));
}
function selectedPartCode(){return document.querySelector(".partPick:checked")?.value||""}
function addPart(){
 openModal(`<h2>+ Add Part</h2><p class="muted">Part Code se Current Stock/PDF catalogue master ka data auto-fill hoga. PDF sirf lookup hai; Save Part ke baad hi stock me add hoga.</p>
 <label>Part Code<input id="partCode" placeholder="Example: 97201155" autocomplete="off"></label>
 <label>Part Name<input id="partName"></label>
 <label>HSN<input id="partHsn"></label>
 <div class="form-grid"><label>MRP<input id="partMrp" type="number" step=".01"></label><label>NDP<input id="partNdp" type="number" step=".01"></label><label>GST %<input id="partGst" type="text"></label><label>Current Stock<input id="partQty" type="number" step="1" min="0"></label></div>
 <label>Supplier<select id="partSupplier">${(db.settings.suppliers||[]).map(x=>`<option>${esc(x)}</option>`).join("")}</select></label>
 <div id="partLookupMsg" class="muted"></div><button id="saveNewPart" class="btn primary wide">Save Part</button>`);
 const fill=()=>{const code=$("partCode").value.trim();const p=part(code);const st=stock()[code];if(!code){$("partLookupMsg").textContent="Enter a Part Code";return}if(p){$("partName").value=p.name||"";$("partHsn").value=p.hsn||"";$("partMrp").value=Number(p.mrp||0);$("partNdp").value=Number(p.ndp||0);$("partGst").value=p.gst||"";$("partQty").value=st?Number(st.qty||0):0;$("partLookupMsg").textContent=st?`✓ Master found • Current Stock: ${st.qty}`:"✓ PDF/Catalogue found • Not currently in stock";}else{$("partLookupMsg").textContent="Part not found in master/PDF. Enter details manually."}};
 $("partCode").oninput=fill;
 $("saveNewPart").onclick=()=>{
   let code=$("partCode").value.trim(),name=$("partName").value.trim();
   if(!code||!name)return alert("Part Code and Part Name are required.");
   const existing=part(code);
   if(existing && stock()[code])return alert("This Part Code is already in Current Stock. Use Edit instead.");
   let p={code,name,hsn:$("partHsn").value.trim(),mrp:+$("partMrp").value||0,ndp:+$("partNdp").value||0,gst:$("partGst").value||"0",openingStock:+$("partQty").value||0,stockMaster:true,supplier:$("partSupplier").value||"",source:existing?"catalogue-added":"manual-added"};
   db.parts.push(p);db.deletedParts=db.deletedParts.filter(x=>String(x).toLowerCase()!==String(code).toLowerCase());save("Part Saved");closeModal();toast("Part added successfully");
 };
}
function editPart(code){
 let p=part(code);if(!p)return alert("Part not found.");
 openModal(`<h2>✏️ Edit Part</h2><label>Part Code<input id="partCode" value="${esc(p.code)}" disabled></label><label>Part Name<input id="partName" value="${esc(p.name)}"></label><label>HSN<input id="partHsn" value="${esc(p.hsn||"")}"></label>
 <div class="form-grid"><label>MRP<input id="partMrp" type="number" step=".01" value="${Number(p.mrp||0)}"></label><label>NDP<input id="partNdp" type="number" step=".01" value="${Number(p.ndp||0)}"></label><label>GST %<input id="partGst" type="number" step=".01" value="${esc(p.gst||"0")}"></label><label>Opening/Manual Stock<input id="partQty" type="number" step="1" min="0" value="${Number(p.openingStock||0)}"></label></div>
 <label>Supplier<select id="partSupplier">${(db.settings.suppliers||[]).map(x=>`<option ${x===(p.supplier||"")?"selected":""}>${esc(x)}</option>`).join("")}</select></label>
 <button id="saveEditPart" class="btn primary wide">Save Changes</button>`);
 $("saveEditPart").onclick=()=>{
   p.name=$("partName").value.trim()||p.name;p.hsn=$("partHsn").value.trim();p.mrp=+$("partMrp").value||0;p.ndp=+$("partNdp").value||0;p.gst=$("partGst").value||"0";p.openingStock=+$("partQty").value||0;p.supplier=$("partSupplier").value||"";
   let cp=db.parts.find(x=>String(x.code).toLowerCase()===String(p.code).toLowerCase());if(cp)Object.assign(cp,p);else db.parts.push({...p});save();closeModal();alert("Part updated successfully.");
 };
}
function deletePart(code){
 let p=part(code);if(!p)return alert("Part not found.");
 if(!confirm(`Delete ${p.code} - ${p.name}?`))return;
 if(db.sales.some(x=>(x.items||[]).some(i=>i.code===p.code))||db.purchases.some(x=>(x.items||[]).some(i=>i.code===p.code)))return alert("This part is already used in a Sale/Purchase and cannot be deleted. Edit it instead.");
 db.parts=db.parts.filter(x=>String(x.code).toLowerCase()!==String(code).toLowerCase());if(PARTS.some(x=>String(x.code).toLowerCase()===String(code).toLowerCase())){if(!db.deletedParts.includes(p.code))db.deletedParts.push(p.code)};save();alert("Part deleted.");
}

function orderView(){
 const f=$("orderFilter").value;const rows=Object.values(stock()).filter(x=>f==="Zero"?x.qty<=0:f==="Low"?x.qty>0&&x.qty<=Number(db.settings.low??2):x.qty<=Number(db.settings.low??2));
 $("orderTable").innerHTML=rows.length?`<div class="table"><table><thead><tr><th>Code</th><th>Part</th><th>Current</th><th>NDP</th><th>GST</th><th>Order Qty</th><th>Payable</th></tr></thead><tbody>${rows.map(x=>`<tr><td>${esc(x.code)}</td><td>${esc(x.name)}</td><td>${x.qty}</td><td>${money(x.ndp)}</td><td>${esc(x.gst||"-")}</td><td><input class="oq" data-code="${esc(x.code)}" type="number" min="0" value="0"></td><td class="opay">₹0</td></tr>`).join("")}</tbody></table></div>`:`<div class=empty>No low/zero stock parts</div>`;
 document.querySelectorAll(".oq").forEach(e=>e.oninput=orderTotals);orderTotals();
}
function orderTotals(){
 let ndp=0,gst=0;
 document.querySelectorAll(".oq").forEach(e=>{const p=part(e.dataset.code),q=Number(e.value)||0,n=q*Number(p?.ndp||0),g=n*(parseFloat(p?.gst)||0)/100;ndp+=n;gst+=g;e.closest("tr").querySelector(".opay").textContent=money(n+g)});
 $("ordNdp").textContent=money(ndp);$("ordGst").textContent=money(gst);$("ordTotal").textContent=money(ndp+gst);
}
const stockExportRows=()=>rows().map(x=>({Code:x.code,Part_Name:x.name,HSN:x.hsn||"",MRP:x.mrp||0,NDP:x.ndp||0,GST:x.gst||"",Current_Stock:Number(x.qty||0),Supplier:x.supplier||""}));
$("stockCsv").onclick=()=>excel(stockExportRows(),"prabha-current-stock");
$("stockPrint").onclick=()=>printRows(stockExportRows().map(x=>({...x,Current_Stock:String(x.Current_Stock)})),"PRABHA AUTO CURRENT STOCK");
$("orderFilter").onchange=orderView;$("lowLimit").oninput=()=>{db.settings.low=Math.max(0,Number($("lowLimit").value)||0);localStorage.setItem(KEY,JSON.stringify(db));renderAll()};
function orderRows(){
 return [...document.querySelectorAll(".oq")].map(e=>{const p=part(e.dataset.code),q=Number(e.value)||0;return q?{code:p.code,name:p.name,current:stock()[p.code].qty,ndp:p.ndp,gst:p.gst,orderQty:q,payable:q*p.ndp*(1+(parseFloat(p.gst)||0)/100)}:null}).filter(Boolean);
}
$("orderCsv").onclick=()=>excel(orderRows(),"prabha-order-list");
$("orderPrint").onclick=()=>printRows(orderRows(),"PRABHA AUTO ORDER LIST");

function purchaseModal(){
 const opts=db.settings.suppliers.map(x=>`<option>${esc(x)}</option>`).join("");
 openModal(`<div class="form"><h2>Add Purchase</h2><div class="grid2"><label>Invoice No.<input id="pinv"></label><label>Date<input id="pdate" type=date value="${today()}"></label><label>Supplier<select id="psup">${opts}</select></label><label>Total Amount<input id="pamt" type=number step=.01></label></div>
 <div class="rowhead"><h3>Purchase Items (optional but recommended for stock)</h3><button class="btn" id="padd">+ Add Part</button></div><div id="pitems"></div>
 <label>Invoice Photo<input id="photo" type=file accept="image/*" capture="environment"></label><div class="actions"><button id="psave" class="btn primary">Save Purchase</button><button id="pcancel" class="btn">Cancel</button></div></div>`);
 const addP=()=>{const d=document.createElement("div");d.className="purchase-line";d.innerHTML=`<input class="pcode" placeholder="Part Code"><input class="pqty" type=number min=1 value=1><span class="pname muted">Part name</span><button class="del">×</button>`;$("pitems").appendChild(d);d.querySelector(".pcode").oninput=()=>{const p=part(d.querySelector(".pcode").value);d.querySelector(".pname").textContent=p?p.name:"Part not found"};d.querySelector(".del").onclick=()=>d.remove()};
 $("padd").onclick=addP;addP();
 $("pcancel").onclick=closeModal;
 $("psave").onclick=()=>{
   const f=$("photo").files[0],done=photo=>{
     const items=[...document.querySelectorAll(".purchase-line")].map(e=>({code:e.querySelector(".pcode").value.trim(),qty:Number(e.querySelector(".pqty").value)||0})).filter(x=>x.code);
     for(const i of items){
       if(!i.qty||i.qty<1)return toast("Purchase quantity must be greater than 0","error");
       const p=part(i.code);
       if(!p)return toast(`Part Code ${i.code} not found in the Parts Master`,"error");
       if(!stock()[p.code])ensurePurchasePart(p.code);
     }
     const x={id:uid("purchase"),invoice:$("pinv").value.trim(),date:$("pdate").value,supplier:$("psup").value,amount:Number($("pamt").value)||0,items,photo,timestamp:new Date().toISOString()};
     if(!x.amount&&!items.length)return toast("Enter amount or purchase items","error");
     db.purchases.push(x);persist("Purchase Saved");closeModal();showSuccess("Purchase saved",`Supplier: ${x.supplier}`);
   };
   if(f){const r=new FileReader();r.onload=()=>done(r.result);r.readAsDataURL(f)}else done("");
 };
}
$("addPurchase").onclick=purchaseModal;
$("purchaseSearch").oninput=()=>purchaseHistoryView();$("purchaseFrom").onchange=()=>purchaseHistoryView();$("purchaseTo").onchange=()=>purchaseHistoryView();
function purchaseHistoryView(){
 const q=($("purchaseSearch").value||"").toLowerCase(),from=$("purchaseFrom").value||"0000-01-01",to=$("purchaseTo").value||"9999-12-31";
 const rows=db.purchases.filter(x=>between(x.date,from,to)&&(`${x.invoice} ${x.supplier} ${(x.items||[]).map(i=>i.code).join(" ")}`).toLowerCase().includes(q)).sort((a,b)=>String(b.date).localeCompare(String(a.date)));
 $("recentPurchases").innerHTML=rows.length?table(rows.slice(0,5).map(x=>({invoice:x.invoice||"-",date:fmt(x.date),supplier:x.supplier,amount:money(x.amount),photo:x.photo?"Yes":"No"}))):`<div class=empty>No recent purchases</div>`;
 $("purchaseHistory").innerHTML=rows.length?`<div class="table"><table><thead><tr><th>Invoice</th><th>Date</th><th>Supplier</th><th>Amount</th><th>Items</th><th>Actions</th></tr></thead><tbody>${rows.map(x=>`<tr><td>${esc(x.invoice||"-")}</td><td>${fmt(x.date)}</td><td>${esc(x.supplier)}</td><td>${money(x.amount)}</td><td>${(x.items||[]).length}</td><td><button class="mini" onclick="window.viewPurchase('${esc(x.id)}')">Open</button> <button class="mini" onclick="window.editPurchase('${esc(x.id)}')">✏️ Edit</button> <button class="mini dangerMini" onclick="window.deletePurchase('${esc(x.id)}')">🗑️ Delete</button></td></tr>`).join("")}</tbody></table></div>`:`<div class=empty>No purchase history</div>`;
}
window.viewPurchase=id=>{const x=db.purchases.find(p=>p.id===id);if(!x)return;openModal(`<h2>Purchase ${esc(x.invoice||"-")}</h2><p><b>Supplier:</b> ${esc(x.supplier)} • <b>Date:</b> ${fmt(x.date)} • <b>Amount:</b> ${money(x.amount)}</p>${x.items?.length?table(x.items.map(i=>({code:i.code,qty:i.qty,name:part(i.code)?.name||"Unknown"}))):""}${x.photo?`<img class="photo" src="${x.photo}" alt="Invoice">`:"<p class=muted>No invoice photo saved.</p>"}<div class=actions><button class="btn" id="pprint">Print</button><button class="btn" id="pclose">Close</button></div>`);$("pprint").onclick=()=>printHTML(`<div class="print-head"><img src="prabha-auto-logo.png"><div><h1>PRABHA AUTO</h1><div>Purchase Invoice Record</div><div class="muted">Supplier: ${esc(x.supplier)} • Invoice: ${esc(x.invoice||"-")} • ${fmt(x.date)}</div></div></div>${x.items?.length?table(x.items):""}<div class="total right">Amount: ${money(x.amount)}</div>`,"PRABHA AUTO PURCHASE");$("pclose").onclick=closeModal};

function purchaseEditor(x){
 const opts=db.settings.suppliers.map(s=>`<option ${s===x.supplier?"selected":""}>${esc(s)}</option>`).join("");
 openModal(`<div class="form"><h2>✏️ Edit Purchase</h2><div class="grid2"><label>Invoice No.<input id="epinv" value="${esc(x.invoice||"")}"></label><label>Date<input id="epdate" type=date value="${esc(x.date||today())}"></label><label>Supplier<select id="epsup">${opts}</select></label><label>Total Amount<input id="epamt" type=number step=.01 value="${Number(x.amount||0)}"></label></div><div class="rowhead"><h3>Purchase Items</h3><button class="btn" id="epadd">+ Add Part</button></div><div id="epitems"></div><p class="muted">Edit/Delete automatically recalculates Current Stock.</p><div class="actions"><button id="epsave" class="btn primary">Save Changes</button><button id="epcancel" class="btn">Cancel</button></div></div>`);
 const addLine=(item={code:"",qty:1})=>{const d=document.createElement("div");d.className="purchase-line";d.innerHTML=`<input class="epcode" placeholder="Part Code" value="${esc(item.code)}"><input class="epqty" type=number min=1 value="${Number(item.qty)||1}"><span class="epname muted">${esc(part(item.code)?.name||"Part name")}</span><button class="del">×</button>`;$("epitems").appendChild(d);d.querySelector(".epcode").oninput=()=>{const p=part(d.querySelector(".epcode").value);d.querySelector(".epname").textContent=p?p.name:"Part not found"};d.querySelector(".del").onclick=()=>d.remove()};
 (x.items||[]).forEach(addLine); if(!(x.items||[]).length)addLine();
 $("epadd").onclick=()=>addLine(); $("epcancel").onclick=closeModal;
 $("epsave").onclick=()=>{const items=[...document.querySelectorAll("#epitems .purchase-line")].map(e=>({code:e.querySelector(".epcode").value.trim(),qty:Number(e.querySelector(".epqty").value)||0})).filter(i=>i.code);for(const i of items){const p=part(i.code);if(!p)return toast(`Part Code ${i.code} not found in the Parts Master`,"error");if(!stock()[p.code])ensurePurchasePart(p.code)}x.invoice=$("epinv").value.trim();x.date=$("epdate").value||today();x.supplier=$("epsup").value;x.amount=Number($("epamt").value)||0;x.items=items;persist("Purchase Updated");closeModal();toast("Purchase updated")};
}
window.editPurchase=id=>{const x=db.purchases.find(p=>p.id===id);if(x)purchaseEditor(x)};
window.deletePurchase=id=>{const x=db.purchases.find(p=>p.id===id);if(!x)return;if(!confirm(`Delete purchase ${x.invoice||"-"}? Current Stock will be recalculated.`))return;db.purchases=db.purchases.filter(p=>p.id!==id);persist("Purchase Deleted");toast("Purchase deleted and stock recalculated")};


function supplierView(){
 const suppliers=db.settings.suppliers;
 $("supplierCards").innerHTML=suppliers.map(s=>{
   const purchases=db.purchases.filter(x=>x.supplier===s),payments=db.payments.filter(x=>x.party===s&&x.type==="Supplier Payment");
   const pv=total(purchases,"amount"),paid=total(payments,"amount"),bal=paid-pv;
   const label=bal>=0?"Credit / Advance":"Due";
   return `<div class="supplier-card"><div><b>${esc(s)}</b><small>Purchase ${money(pv)} • Paid ${money(paid)}</small></div><strong>${label} ${money(Math.abs(bal))}</strong><button class="mini" onclick="window.supplierDetail('${esc(s)}')">View</button></div>`;
 }).join("");
 $("paymentHistory").innerHTML=table(db.payments.filter(x=>x.type==="Supplier Payment").sort((a,b)=>String(b.date).localeCompare(String(a.date))).map(x=>({date:fmt(x.date),supplier:x.party,amount:money(x.amount),mode:x.mode,utr:x.utr||"-"})));
}
window.supplierDetail=name=>{
 const purchases=db.purchases.filter(x=>x.supplier===name).sort((a,b)=>String(b.date).localeCompare(String(a.date)));
 const payments=db.payments.filter(x=>x.party===name&&x.type==="Supplier Payment");
 const pv=total(purchases,"amount"),paid=total(payments,"amount"),bal=paid-pv;
 const label=bal>=0?"Credit / Advance":"Due";
 openModal(`<h2>${esc(name)}</h2><div class="ledger total"><b>Purchase</b><strong>${money(pv)}</strong></div><div class="ledger"><b>Paid</b><strong>${money(paid)}</strong></div><div class="ledger"><b>${label}</b><strong>${money(Math.abs(bal))}</strong></div><h3>Purchase History</h3>${table(purchases.map(x=>({invoice:x.invoice||"-",date:fmt(x.date),amount:money(x.amount)})))}<h3>Payment History</h3>${table(payments.map(x=>({date:fmt(x.date),amount:money(x.amount),mode:x.mode,utr:x.utr||"-"})))}<div class=actions><button class="btn" id="sdExcel">Excel</button><button class="btn" id="sdPdf">PDF / Print</button><button class="btn" id="sdClose">Close</button></div>`);
 $("sdExcel").onclick=()=>excel([...purchases.map(x=>({date:fmt(x.date),type:"Purchase",invoice:x.invoice||"-",amount:x.amount})),...payments.map(x=>({date:fmt(x.date),type:"Payment",invoice:"-",amount:x.amount}))],`supplier-${name.replace(/[^a-z0-9]+/gi,"-").toLowerCase()}`);
 $("sdPdf").onclick=()=>printRows([...purchases.map(x=>({date:fmt(x.date),type:"Purchase",invoice:x.invoice||"-",amount:money(x.amount)})),...payments.map(x=>({date:fmt(x.date),type:"Payment",invoice:"-",amount:money(x.amount)}))],`PRABHA AUTO - ${name}`);
 $("sdClose").onclick=closeModal;
};

function addSupplier(){
 openModal(`<h2>New Supplier</h2><label>Supplier Name<input id="sname" placeholder="Supplier name"></label><div class=actions><button id="saveSup" class="btn primary">Add Supplier</button><button id="scancel" class="btn">Cancel</button></div>`);
 $("saveSup").onclick=()=>{const n=$("sname").value.trim();if(!n)return toast("Supplier name required","error");if(db.settings.suppliers.includes(n))return toast("Supplier already exists","error");db.settings.suppliers.push(n);persist("Supplier Saved");closeModal();toast("New supplier added")};$("scancel").onclick=closeModal;
}
function supplierPayment(){
 const opts=db.settings.suppliers.map(x=>`<option>${esc(x)}</option>`).join("");
 openModal(`<h2>Supplier Payment</h2><div class=form><label>Supplier<select id="pparty">${opts}</select></label><label>Date<input id="pdate2" type=date value="${today()}"></label><label>Amount<input id="pamount" type=number min=0 step=.01></label><label>Mode<select id="pmode"><option>Cash</option><option>Bank</option></select></label><label>UTR No.<input id="putr"></label><div class=actions><button id="savePay" class="btn primary">Save Payment</button><button id="payCancel" class="btn">Cancel</button></div></div>`);
 $("savePay").onclick=()=>{const a=Number($("pamount").value)||0;if(!a)return toast("Amount required","error");db.payments.push({id:uid("pay"),party:$("pparty").value,type:"Supplier Payment",date:$("pdate2").value,amount:a,mode:$("pmode").value,utr:$("putr").value.trim()});persist("Payment Saved");closeModal();toast("Supplier payment saved")};$("payCancel").onclick=closeModal;
}
$("addSupplier").onclick=addSupplier;$("addPayment").onclick=supplierPayment;
$("quickPayment").onclick=supplierPayment;

function supplierNetBalance(name){
 const purchases=db.purchases.filter(x=>x.supplier===name).reduce((n,x)=>n+Number(x.amount||0),0);
 const payments=db.payments.filter(x=>x.party===name&&x.type==="Supplier Payment").reduce((n,x)=>n+Number(x.amount||0),0);
 // Supplier account starts at zero: payment creates advance/credit, purchase consumes it.
 return payments-purchases;
}
function ganpatiBalance(){
 return supplierNetBalance("Ganpati Automobile");
}

function reportRows(type,a,b){
 const sales=db.sales.filter(x=>between(x.date,a,b));
 const purchases=db.purchases.filter(x=>between(x.date,a,b));
 if(type==="Sales")return sales.map(x=>({date:x.date,invoice:x.invoice,customer:x.customer,machine:x.machine,final:x.final,mode:x.mode,utr:x.utr||""}));
 if(type==="Purchase")return purchases.map(x=>({date:x.date,invoice:x.invoice,supplier:x.supplier,amount:x.amount,items:(x.items||[]).length}));
 if(type==="Payments")return db.payments.filter(x=>between(x.date,a,b)).map(x=>({date:x.date,party:x.party,type:x.type,amount:x.amount,mode:x.mode,utr:x.utr||""}));
 if(type==="Stock")return Object.values(stock()).map(x=>({code:x.code,name:x.name,mrp:x.mrp,ndp:x.ndp,gst:x.gst,currentStock:Number(x.qty||0),supplier:x.supplier||""}));
 if(type==="Order"){
   const seen=new Set();
   return Object.values(stock())
     .filter(x=>x.qty<=Number(db.settings.low??2))
     .filter(x=>{const k=String(x.code).trim().toLowerCase();if(!k||seen.has(k))return false;seen.add(k);return true;})
     .map(x=>({code:x.code,name:x.name,current:x.qty,ndp:x.ndp,gst:x.gst||"",supplier:x.supplier||""}));
 }
 if(type==="Ganpati")return[
   ...db.payments.filter(x=>x.party==="Ganpati Automobile"&&x.type==="Supplier Payment"&&between(x.date,a,b)).map(x=>({entry:"Payment / Advance",amount:x.amount,date:x.date})),
   ...purchases.filter(x=>x.supplier==="Ganpati Automobile").map(x=>({entry:`Purchase ${x.invoice||"-"}`,amount:-Number(x.amount||0),date:x.date}))
 ];
 return [];
}

// Stock is available as its own report, but is intentionally NOT included in Complete.
// This prevents the full spare-parts master from flooding the normal business report.
const REPORT_TYPES=["Sales","Purchase","Payments","Order","Ganpati","Stock"];
const COMPLETE_REPORT_TYPES=["Sales","Purchase","Payments","Order","Ganpati"];

function selectedReportTypes(){
 const checks=[...document.querySelectorAll(".reportCheck:checked")].map(x=>x.value);
 if(checks.includes("Complete"))return COMPLETE_REPORT_TYPES.slice();
 return [...new Set(checks.filter(x=>REPORT_TYPES.includes(x)))];
}

function updateReportPickerLabel(){
 const selected=selectedReportTypes();
 const complete=document.querySelector('.reportCheck[value="Complete"]')?.checked;
 if($("reportPickerBtn"))$("reportPickerBtn").innerHTML=`${complete?"✓ ":""}${complete?"Complete":selected.length?selected.join(", "):"Select reports"} <span>▾</span>`;
}

const REPORT_COLUMNS={
 Sales:["date","invoice","customer","machine","final","mode","utr"],
 Purchase:["date","invoice","supplier","amount","items"],
 Payments:["date","party","type","amount","mode","utr"],
 Order:["code","name","current","ndp","gst","supplier"],
 Ganpati:["date","entry","amount"],
 Stock:["code","name","mrp","ndp","gst","currentStock","supplier"]
};
const REPORT_LABELS={date:"Date",invoice:"Invoice",customer:"Customer",machine:"Machine",final:"Final",mode:"Mode",utr:"UTR",
 party:"Party",type:"Type",amount:"Amount",items:"Items",code:"Code",name:"Name",current:"Current Stock",mrp:"MRP",
 ndp:"NDP",gst:"GST",supplier:"Supplier",entry:"Entry",currentStock:"Current Stock"};

function reportTable(type,rows){
 if(!rows.length)return `<div class="empty">No ${esc(type)} data for selected dates</div>`;
 const cols=REPORT_COLUMNS[type]||Object.keys(rows[0]);
 return `<div class="report-section"><div class="report-section-title"><b>${esc(type)}</b><span>${rows.length} row${rows.length===1?"":"s"}</span></div><div class="table"><table><thead><tr>${
   cols.map(k=>`<th>${esc(REPORT_LABELS[k]||k)}</th>`).join("")
 }</tr></thead><tbody>${
   rows.map(r=>`<tr>${cols.map(k=>{
     const v=r[k];
     if(v===undefined||v===null||v==="")return "<td></td>";
     if(["final","amount","mrp","ndp"].includes(k))return `<td>${money(v)}</td>`;
     return `<td>${esc(String(v))}</td>`;
   }).join("")}</tr>`).join("")
 }</tbody></table></div></div>`;
}

let lastReport=[];
function makeReport(){
 const [a,b]=[$("rFrom").value||"1900-01-01",$("rTo").value||today()];
 const types=selectedReportTypes();
 lastReport=[];
 const sections=[];
 types.forEach(type=>{
   const rows=reportRows(type,a,b);
   rows.forEach(row=>lastReport.push({Report:type,...row}));
   sections.push(reportTable(type,rows));
 });
 $("report").innerHTML=sections.length?sections.join(""):`<div class="empty">Select at least one report</div>`;
 updateReportPickerLabel();
}
$("makeReport").onclick=makeReport;
$("rAllDates").onclick=()=>{$("rFrom").value="1900-01-01";$("rTo").value=today();makeReport()};
$("reportCsv").onclick=()=>excel(lastReport,"prabha-report");
$("reportPrint").onclick=()=>{
 const [a,b]=[$("rFrom").value||"1900-01-01",$("rTo").value||today()];
 const types=selectedReportTypes();
 if(!types.length)return toast("Select at least one report","error");
 const body=types.map(type=>{
   const rows=reportRows(type,a,b);
   if(!rows.length)return "";
   const cols=REPORT_COLUMNS[type]||Object.keys(rows[0]);
   const mapped=rows.map(r=>{const o={};cols.forEach(k=>{o[REPORT_LABELS[k]||k]=(["final","amount","mrp","ndp"].includes(k)?money(r[k]):(r[k]??""))});return o;});
   return `<h2>${esc(type)}</h2>${table(mapped,Object.keys(mapped[0]))}`;
 }).join("");
 if(!body)return toast("No report data to print","error");
 printHTML(`<div class="print-head"><img src="prabha-auto-logo.png"><div><h1>PRABHA AUTO</h1><div>Authorised Dealer • EICHER TRACTORS</div><div class="muted">Service Manager: ${esc(db.settings.manager||"Chanchal Kumar")} • ${esc(db.settings.mobile||"9693124739")}</div></div></div>${body}<div class="footer">Generated on ${fmt(today())}</div>`,"PRABHA AUTO REPORT");
};

function initReportPicker(){
 const btn=$("reportPickerBtn"),menu=$("reportPickerMenu");if(!btn||!menu)return;
 btn.onclick=e=>{e.stopPropagation();menu.classList.toggle("hidden")};
 document.querySelectorAll(".reportCheck").forEach(c=>c.addEventListener("change",()=>{
   if(c.value==="Complete"&&c.checked){document.querySelectorAll(".reportCheck").forEach(x=>{if(x!==c)x.checked=false});}
   else if(c.value!=="Complete"&&c.checked){const complete=document.querySelector('.reportCheck[value="Complete"]');if(complete)complete.checked=false;}
   updateReportPickerLabel();makeReport();
 }));
 $("reportSelectAll").onclick=e=>{e.stopPropagation();document.querySelectorAll(".reportCheck").forEach(x=>x.checked=x.value!=="Complete");updateReportPickerLabel();makeReport()};
 $("reportClearAll").onclick=e=>{e.stopPropagation();document.querySelectorAll(".reportCheck").forEach(x=>x.checked=false);updateReportPickerLabel();makeReport()};
 menu.addEventListener("click",e=>e.stopPropagation());
 document.addEventListener("click",e=>{if(!e.target.closest("#reportPicker"))menu.classList.add("hidden")});
 updateReportPickerLabel();
}


$("saveSettings").onclick=()=>{db.settings.business=$("business").value.trim()||"PRABHA AUTO";db.settings.manager=$("manager").value.trim()||"Chanchal Kumar";db.settings.mobile=$("mobile").value.trim();db.settings.email=$("email").value.trim();db.settings.low=Math.max(0,Number($("settingLow").value)||0);persist("Settings Saved");toast("Settings saved")};
$("saveMachines").onclick=()=>{db.settings.machines=[$("m1").value.trim(),$("m2").value.trim(),$("m3").value.trim()].filter(Boolean);persist("Machines Saved");toast("Machine names saved")};
$("backup").onclick=()=>{const a=document.createElement("a");a.href=URL.createObjectURL(new Blob([JSON.stringify(db,null,2)],{type:"application/json"}));a.download=`prabha-auto-backup-${today()}.json`;a.click();toast("Backup downloaded")};
$("restore").onchange=e=>{const f=e.target.files[0];if(!f)return;const r=new FileReader();r.onload=()=>{try{const x=JSON.parse(r.result);if(!x||!Array.isArray(x.sales))throw Error();db={...emptyDB(),...x,settings:{...emptyDB().settings,...(x.settings||{})}};persist("Restored");toast("Backup restored")}catch{toast("Invalid backup file","error")}};r.readAsText(f)};
let cloudSyncTimer=null;
let cloudPollTimer=null;
let cloudSyncBusy=false;
const CLOUD_POLL_MS=30000;

function cloudUrl(){
 return window.PRABHA_CONFIG?.GOOGLE_SCRIPT_URL||"";
}
function hasLocalTransactions(){
 return !!(db.sales.length||db.purchases.length||db.payments.length||db.issues.length||db.parts.length||db.accessories.length||db.deletedParts.length||db.deletedAccessories.length);
}
function hasLocalDatabase(){
 return hasLocalTransactions() || !!db._syncUpdatedAt || !!db._ganpatiResetV1;
}
function setSyncStatus(text){
 if($("syncStatus"))$("syncStatus").textContent=text;
 if($("sync"))$("sync").textContent=text;
}
function scheduleCloudSync(){
 clearTimeout(cloudSyncTimer);
 cloudSyncTimer=setTimeout(()=>syncToSheet({silent:true}),800);
}
async function syncToSheet(options={}){
 const u=cloudUrl();
 if(!u){if(!options.silent)toast("Set Google Apps Script Web App URL in config.js first","error");return false}
 if(cloudSyncBusy)return false;
 cloudSyncBusy=true;
 try{
   const payload={...db,__stock:rows().map(x=>({code:x.code,name:x.name,hsn:x.hsn||"",mrp:x.mrp||0,ndp:x.ndp||0,gst:x.gst||"",qty:x.qty||0,supplier:x.supplier||""}))};
   const r=await fetch(u,{method:"POST",headers:{"Content-Type":"text/plain;charset=utf-8"},body:JSON.stringify(payload),cache:"no-store"});
   const j=await r.json();
   if(!j.success)throw Error(j.error||"Sync failed");
   db._syncUpdatedAt=Date.parse(j.updatedAt)||Date.now();
   localStorage.setItem(KEY,JSON.stringify(db));
   setSyncStatus("☁ Synced • "+new Date(db._syncUpdatedAt).toLocaleTimeString("en-IN",{hour:"2-digit",minute:"2-digit"}));
   if(!options.silent)toast("Google Sheet sync complete");
   return true;
 }catch(e){
   if(!options.silent)toast("Google Sheet sync failed","error");
   if($("syncStatus"))$("syncStatus").textContent="Sync pending • "+String(e.message||e);
   return false;
 }finally{cloudSyncBusy=false}
}
async function loadFromSheet(options={}){
 const u=cloudUrl();
 if(!u){if(!options.silent)toast("Set Google Apps Script Web App URL in config.js first","error");return false}
 try{
   const r=await fetch(u+"?action=load&t="+Date.now(),{cache:"no-store"});
   const j=await r.json();
   if(!j.success)throw Error(j.error||"Load failed");
   if(!j.data){
     if(hasLocalDatabase()) return await syncToSheet({silent:true});
     if(!options.silent)toast("Google Sheet has no saved database yet","error");
     return false;
   }
   const remoteTime=Date.parse(j.updatedAt)||0;
   const localTime=Number(db._syncUpdatedAt||0);
   if(options.force || remoteTime>localTime || !hasLocalTransactions()){
     const merged={...emptyDB(),...j.data,settings:{...emptyDB().settings,...(j.data.settings||{})}};
     let ganpatiReset=false;
     if(merged._ganpatiResetV1!==true){
       merged.purchases=(merged.purchases||[]).filter(x=>x.supplier!=="Ganpati Automobile");
       merged.payments=(merged.payments||[]).filter(x=>!(x.party==="Ganpati Automobile"&&x.type==="Supplier Payment"));
       merged._ganpatiResetV1=true;
       ganpatiReset=true;
     }
     db=merged;
     db._syncUpdatedAt=remoteTime||Date.now();
     localStorage.setItem(KEY,JSON.stringify(db));
     renderAll();
     if(ganpatiReset) setTimeout(()=>syncToSheet({silent:true}),300);
     setSyncStatus("☁ Loaded • "+new Date(db._syncUpdatedAt).toLocaleTimeString("en-IN",{hour:"2-digit",minute:"2-digit"}));
     if(!options.silent)toast("Latest data loaded from Google Sheet");
     return true;
   }
   if(localTime>remoteTime) return await syncToSheet({silent:true});
   return false;
 }catch(e){
   if(!options.silent)toast("Could not load Google Sheet data","error");
   return false;
 }
}
async function refreshDashboard(){
 await loadFromSheet({silent:true});
 renderAll();
 toast("Dashboard refreshed");
}
function startCloudSync(){
 if(cloudPollTimer)return;
 loadFromSheet({silent:true}).then(()=>{
   cloudPollTimer=setInterval(()=>loadFromSheet({silent:true}),CLOUD_POLL_MS);
 });
}
$("syncNow").onclick=()=>syncToSheet();
$("loadSync").onclick=()=>loadFromSheet({force:true});
$("forgot").onclick=$("otp").onclick=()=>toast("Recovery endpoint is not configured. Use the configured recovery email.","error");
$("forgot").onclick=$("otp").onclick=()=>toast("Recovery endpoint is not configured. Use the configured recovery email.","error");
$("dashRange").onchange=()=>{const c=$("dashRange").value==="custom";$("dashFrom").classList.toggle("hidden",!c);$("dashTo").classList.toggle("hidden",!c);dashboard()};
$("dashApply").onclick=dashboard;
$("dashRefresh").onclick=async()=>{const b=$("dashRefresh");b.classList.add("tap-active");b.textContent="↻ Refreshing…";try{await refreshDashboard()}finally{b.textContent="↻ Refresh";setTimeout(()=>b.classList.remove("tap-active"),250)}};
$("saleSearch").oninput=saleHistory;
$("stockSearch").oninput=stockView;
$("stockFilter").onchange=stockView;
$("addPart").onclick=addPart;
$("editSelectedPart").onclick=()=>{const c=selectedPartCode();c?editPart(c):toast("Select a part first","error")};
$("deleteSelectedPart").onclick=()=>{const c=selectedPartCode();c?deletePart(c):toast("Select a part first","error")};

function initMobileMore(){
  const sheet = $("mobileMoreSheet");
  const moreBtn = $("mobileMore");
  const closeBtn = $("mobileMoreClose");
  const xBtn = $("mobileMoreX");

  if(!sheet || !moreBtn) return;

  const openMore = (e) => {
    e?.preventDefault();
    e?.stopPropagation();

    sheet.classList.remove("hidden");
    moreBtn.classList.add("active");
  };

  const closeMore = (e) => {
    e?.preventDefault();
    e?.stopPropagation();

    sheet.classList.add("hidden");
    moreBtn.classList.remove("active");
  };

  moreBtn.onclick = openMore;

  closeBtn && (closeBtn.onclick = closeMore);
  xBtn && (xBtn.onclick = closeMore);

  sheet.querySelectorAll(".mobile-more-card [data-page]").forEach(btn => {
    btn.onclick = (e) => {
      e.preventDefault();
      e.stopPropagation();

      const page = btn.dataset.page;

      closeMore();

      if(page) nav(page);
    };
  });
}

$("addAccessory")?.addEventListener("click",()=>accessoryModal());
$("accessoryExcel")?.addEventListener("click",()=>excel(accessoryExportRows(),"prabha-accessories"));
$("accessoryPrint")?.addEventListener("click",accessoryPrint);
ensureAccessories();
function renderAll(){fillMachines();dashboard();stockView();orderView();purchaseHistoryView();supplierView();saleHistory();accessoryView();makeReport()}
initReportPicker();initMobileMore();$("saleDate").value=today();$("rFrom").value=new Date(new Date().getFullYear(),new Date().getMonth(),1).toISOString().slice(0,10);$("rTo").value=today();addItem();renderAll();
if($("app")&&!$("app").classList.contains("hidden"))startCloudSync();
})();