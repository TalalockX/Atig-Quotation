/* Loads each library from the first CDN that works. */
function loadChain(urls, ready){
  return new Promise(resolve => {
    let i = 0;
    const next = () => {
      if (ready()) return resolve(true);
      if (i >= urls.length) return resolve(false);
      const s = document.createElement("script"); s.src = urls[i++];
      s.onload = () => ready() ? resolve(true) : next();
      s.onerror = next;
      document.head.appendChild(s);
    };
    next();
  });
}


/* Set these after deploying Code.gs. Leave SCRIPT_URL empty to run as a standalone demo. */
const CONFIG = { SCRIPT_URL: "", SECRET: "", SHEET_URL: "" };
/* The Connect panel at the bottom of the page saves these on the tablet, so no file editing is needed. */
try{ Object.assign(CONFIG, JSON.parse(localStorage.getItem("gqs-config") || "{}")); }catch(e){}

/* The gallery's real items (items.js, made from the Excel): serial -> [Arabic name, selling price, stock]. */
const ITEMS = window.GQS_ITEMS || {};
const itemBySerial = s => { const r = ITEMS[s]; return r ? { serial:s, en:"", ar:r[0] || s, price:r[1] || 0, cur:"SAR" } : null; };
const stockOf = serial => ITEMS[serial] ? ITEMS[serial][2] : null;
/* Units of this item on every quotation (saved or draft, never deleted ones) - the open one counted live. */
function usedQty(serial){
  let n = 0; const cur = state.quote && state.quote.id;
  Object.keys(state.saved).forEach(id => { if (id !== cur && (state.saved[id].quote.savedAt || 0) >= itemsStamp) state.saved[id].lines.forEach(l => { if (l.serial === serial) n += l.qty; }); });
  state.lines.forEach(l => { if (l.serial === serial) n += l.qty; });
  return n;
}
const leftQty = serial => stockOf(serial) === null ? null : stockOf(serial) - usedQty(serial);
function stockHtml(serial){
  const left = leftQty(serial); if (left === null) return "";
  return `<span class="stock${left < 0 ? " bad" : left === 0 ? " zero" : ""}" data-stock="${esc(serial)}">المتبقي / Left: ${left} <small>من / of ${stockOf(serial)}</small></span>`;
}
function updateStock(){
  document.querySelectorAll("#lines [data-stock]").forEach(el => { el.outerHTML = stockHtml(el.dataset.stock); });
}

/* The page is split in two (index.html = quotations, settings.html = Google Sheet + data update).
   Boxes that exist on the other page resolve to a detached dummy element so shared code never crashes. */
const _stub = {};
const $ = id => document.getElementById(id) || (_stub[id] || (_stub[id] = document.createElement("div")));
const IS_MAIN = !!document.getElementById("lines");

const fmt = n => n.toLocaleString("en-US",{minimumFractionDigits:2,maximumFractionDigits:2});
const r2 = n => Math.round((n + Number.EPSILON) * 100) / 100;
const esc = s => String(s).replace(/[&<>"']/g, c => ({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"}[c]));

const state = { quote:null, lines:[], saved:{} };
let armed = false;   // autosave stays off until the page has finished starting up

/* ---------- label format ---------- */
function parseLabel(raw){
  const p = String(raw).trim().split("|");
  if (p.length !== 5) return {error:"Not a gallery label. Expected 5 fields separated by |."};
  const [serial,en,ar,priceStr,cur] = p.map(s => s.trim());
  if (!/^[A-Z]{2,4}-\d{3,}$/.test(serial)) return {error:"Serial number looks wrong: " + serial};
  const price = Number(priceStr);
  if (!isFinite(price) || price < 0 || priceStr === "") return {error:"Price looks wrong: " + priceStr};
  if (!/^[A-Z]{3}$/.test(cur)) return {error:"Currency looks wrong: " + cur};
  if (!en) return {error:"Item name is missing."};
  return {item:{serial,en,ar,price,cur}};
}

/* ---------- quotation ---------- */
/* Quote numbers never repeat: the next one is above every saved quotation and every deleted one. */
/* The Google Sheet tab (and the Quotations index row) is named by the document number the user typed; QT number if empty. */
const tabName = q => { const n = String(q.docNo || "").replace(/[:\\\/?*\[\]]/g, "-").trim().slice(0, 100); return n || q.id; };
const quoteNo = id => parseInt((String(id).match(/(\d+)$/) || [0, 0])[1], 10) || 0;   // QT-0012 and older QT-2026-0012 both give 12
function nextQuoteNo(){
  let max = 0;
  try{ max = Number(localStorage.getItem("gqs-counter")) || 0; }catch(e){}
  Object.keys(state.saved).forEach(id => { const n = quoteNo(id); if (n > max) max = n; });
  return max + 1;
}
function newQuote(){
  const n = nextQuoteNo();
  $("delConfirm").hidden = true;
  const d = new Date();
  let branch = ""; try{ branch = localStorage.getItem("gqs-branch") || ""; }catch(e){}
  state.quote = { id:"QT-" + String(n).padStart(4,"0"), customer:"", branch, custNo:"", custVat:"", custCity:"", notes:"", custPhone:"", docNo:"", validity:"",custAddress:"",custCommReg:"",
    date:d.toISOString().slice(0,10) };
  state.lines = [];
  $("cust").value = "";
  syncExtra();
  renderAll();
}
function lineTotal(l){ return r2(l.price * l.qty * (1 - l.disc/100)); }
const vatRate = () => $("vatOn").checked ? 15 : 0;
/* Per-line figures, as the Sheet shows them. Taxable amount is after discount and before VAT. */
function lineCalc(l, vatPct){
  const gross = r2(l.price * l.qty);
  const taxable = r2(gross * (1 - l.disc/100));
  const discAmt = r2(gross - taxable);
  const vatAmt = r2(taxable * vatPct / 100);
  return { gross, taxable, discAmt, vatPct, vatAmt, total:r2(taxable + vatAmt) };
}
function calc(){
  let gross=0, net=0, vat=0;
  state.lines.forEach(l => { const c = lineCalc(l, vatRate()); gross += c.gross; net += c.taxable; vat += c.vatAmt; });
  gross=r2(gross); net=r2(net); vat=r2(vat);
  return {gross, discount:r2(gross-net), net, vat, total:r2(net+vat)};
}
/* The 11 columns of the customer quotation tab, in this order. */
const LINE_COLS = ["الباركود على العلبة\nBarcode on box","اسم الصنف عربي\nItem name Arabic","رقم الفرع\nBranch no","العدد\nQty","السعر\nPrice","المبلغ\nAmount","نسبة الخصم\nDiscount %","قيمة الخصم\nDiscount value","نسبة الضريبة\nVAT %","قيمة الضريبة\nVAT value","الإجمالي\nTotal"];
function sheetRows(snap){
  return snap.lines.map(l => { const c = lineCalc(l, snap.vatPct);
    return { barcode:l.serial, name:l.ar || l.en, qty:l.qty, unitPrice:l.price, amount:c.gross,
      discPct:l.disc, discAmt:c.discAmt, vatPct:c.vatPct, vatAmt:c.vatAmt, subtotal:c.total }; });
}
function addItem(item){
  const ex = state.lines.find(l => l.serial === item.serial);
  const nm = item.ar || item.en;
  if (ex){ ex.qty += 1; renderLines(item.serial); }
  else { state.lines.push({...item, qty:1, disc:0, manualPrice: !item.price}); renderLines(item.serial); }
  renderTotals();
  const left = leftQty(item.serial);
  let msg = nm + (ex ? " – تمت الإضافة (الكمية " + ex.qty + ")" : " – تمت الإضافة") + "\n" + (ex ? "added (qty " + ex.qty + ")" : "added");
  if (left !== null) msg += " · المتبقي " + left + " / " + left + " left";
  if (left !== null && left < 0) msg += " · تجاوز المخزون! / over stock!";
  if (!item.price) msg += "\nلا يوجد سعر بيع، أدخله / no selling price, enter it";
  toast(msg, 5000);
  queueAutosave();
}
/* fromCamera = true: only private labels (AQ1...) are accepted. Typed or pasted entries may also be a plain serial. */
function handleScan(raw, fromCamera){
  const t = String(raw).trim();
  if (!t) return;
  const serial = window.GQS_QR ? GQS_QR.decode(t) : null;
  if (serial){
    const it = itemBySerial(serial);
    return it ? addItem(it) : toast("الصنف " + serial + " غير موجود في قائمة الأصناف.\nItem " + serial + " is not in the item list.");
  }
  if (fromCamera) return toast("هذا ليس ملصق المعرض.\nNot a gallery label.");
  if (t.includes("|")){                       // older 5-field labels still work when pasted
    const r = parseLabel(t);
    if (r.error) return toast("ملصق غير صالح: " + r.error + "\nNot a valid label");
    return addItem(r.item);
  }
  const it = itemBySerial(t);
  if (!it) return toast("الصنف " + t.slice(0,40) + " غير موجود في قائمة الأصناف.\nItem " + t.slice(0,40) + " is not in the item list.");
  addItem(it);
}

/* ---------- rendering ---------- */
/* Shows at most n rows of a table and scrolls the rest; the header stays put. */
function capRows(table, n){
  if (!table) return;
  const wrap = table.closest(".tablewrap"); if (!wrap) return;
  wrap.classList.add("cap");
  const rows = table.tBodies[0] ? [...table.tBodies[0].rows].filter(r => !r.querySelector(".empty")) : [];
  if (rows.length <= n || !table.tHead || !table.tHead.offsetHeight){ wrap.style.maxHeight = ""; wrap.style.overflowY = ""; return; }
  let h = table.tHead.offsetHeight; for (let i = 0; i < n; i++) h += rows[i].offsetHeight;
  wrap.style.maxHeight = (h + 2) + "px"; wrap.style.overflowY = "auto";
}
function renderLines(flashSerial){
  $("empty").hidden = state.lines.length > 0;
  $("lines").innerHTML = state.lines.map((l,i) => `
    <tr data-i="${i}" class="${l.serial===flashSerial?"flash":""}">
      <td class="item"><b dir="auto">${esc(l.ar || l.en)}</b><span class="sub"><code>${esc(l.serial)}</code>${l.en ? `<span>${esc(l.en)}</span>` : ""}</span>${stockHtml(l.serial)}</td>
      <td class="num">${l.manualPrice ? `<div class="disc"><input type="number" inputmode="decimal" min="0" step="0.01" value="${l.price}" data-f="price" aria-label="السعر / Price"></div>` : fmt(l.price)}</td>
      <td><div class="qty">
        <button type="button" data-act="dec" aria-label="إنقاص الكمية / Decrease quantity">−</button>
        <input type="number" inputmode="numeric" min="1" step="1" value="${l.qty}" data-f="qty" aria-label="الكمية / Quantity">
        <button type="button" data-act="inc" aria-label="زيادة الكمية / Increase quantity">+</button></div></td>
      <td class="num"><div class="disc"><input type="number" inputmode="decimal" min="0" max="100" step="0.5" value="${l.disc}" data-f="disc" aria-label="نسبة الخصم / Discount percent"><button type="button" class="applyAll" data-act="applyall" hidden>تطبيق على الكل <small>Apply to all</small></button></div></td>
      <td class="num" data-total>${fmt(lineTotal(l))}</td>
      <td><button type="button" data-act="del" aria-label="حذف الصنف / Remove item">✕</button></td>
    </tr>`).join("");
  capRows($("lines").closest("table"), 6);
  const fl = $("lines").querySelector("tr.flash"); if (fl) fl.scrollIntoView({ block:"nearest" });
}
function renderTotals(){
  const t = calc();
  $("totals").innerHTML =
    `<div><span>المجموع قبل الخصم / Subtotal</span><span>${fmt(t.gross)}</span></div>` +
    `<div><span>الخصم / Discount</span><span>−${fmt(t.discount)}</span></div>` +
    `<div><span>ضريبة القيمة المضافة 15% / VAT 15%</span><span>${$("vatOn").checked ? fmt(t.vat) : "متوقفة / off"}</span></div>` +
    `<div class="grand"><span>الإجمالي (ر.س) / Total SAR</span><span>${fmt(t.total)}</span></div>`;
}
function renderHeader(){
  const q = state.quote;
  $("qid").value = q.id; $("qdate").value = q.date; $("branch").value = q.branch || "";
  ["status","finBtn"].forEach(id => { const el = $(id); if (el) el.hidden = true; });   // old buttons, if the page still has them
  $("delBtn").disabled = !(state.lines.length || state.saved[q.id]);
  renderQuoteLinks();
}
function renderSaved(){
  const ids = Object.keys(state.saved);
  $("savedPanel").hidden = !ids.length;
  $("saved").innerHTML = ids.map(id => `<button type="button" data-open="${id}">${esc(tabName(state.saved[id].quote))} · ${esc(state.saved[id].quote.customer||"بدون عميل / No customer")}</button>`).join("");
  renderPreview();
}
/* The preview shows every saved quotation, with the one being edited shown live. */
function renderPreview(){
  const map = {...state.saved};
  if (state.quote && (state.lines.length || state.saved[state.quote.id])) map[state.quote.id] = liveSnap();
  const q = Object.keys(map).map(id => map[id]);
  const head = c => "<thead><tr>" + c.map(x => `<th>${x.replace(/\n/g,"<br>")}</th>`).join("") + "</tr></thead>";
  const none = n => `<tr><td colspan="${n}" class="empty">لا يوجد شيء محفوظ بعد / Nothing saved yet.</td></tr>`;
  if (!$("tQuotes")) return;
  $("tQuotes").innerHTML = head(["رقم المستند\nDoc no.","العميل\nCustomer","فرع\nBranch","التاريخ\nDate","المجموع قبل الخصم\nSubtotal","الخصم\nDiscount","الضريبة\nVAT","الإجمالي\nTotal"]) + "<tbody>" +
    (q.length ? q.map(s => { const t=s.totals; return `<tr><td>${esc(tabName(s.quote))}</td><td>${esc(s.quote.customer)}</td><td>${esc(s.quote.branch || "")}</td><td>${s.quote.date}</td><td class="num">${fmt(t.gross)}</td><td class="num">${fmt(t.discount)}</td><td class="num">${fmt(t.vat)}</td><td class="num">${fmt(t.total)}</td></tr>`; }).join("") : none(8)) + "</tbody>";
  capRows($("tQuotes"), 4);
  /* The open quotation's own tab: details, the 11-column table and the totals. */
  const cur = state.quote && (state.lines.length || state.saved[state.quote.id]) ? liveSnap() : null;
  if (!$("tLines")) return;   // the page no longer shows the per-quotation tab preview
  $("tabTitle").textContent = cur ? "ورقة " + tabName(cur.quote) + " / Tab " + tabName(cur.quote) : "ورقة هذا العرض / This quotation's tab";
  if (!cur){ $("tabInfo").innerHTML = ""; $("tLines").innerHTML = head(LINE_COLS) + "<tbody>" + none(11) + "</tbody>"; $("tTotals").innerHTML = ""; capRows($("tLines"), 6); return; }
  const qq = cur.quote, t = cur.totals, net = t.net !== undefined ? t.net : r2(t.gross - t.discount);
  $("tabInfo").innerHTML = [["رقم المستند / Doc no.", tabName(qq)],["العميل / Customer", qq.customer],["الفرع / Branch", qq.branch],["التاريخ / Date", qq.date]]
    .map(([k,v]) => `<div><span>${k}</span><b dir="auto">${esc(v || "—")}</b></div>`).join("");
  const rows = sheetRows(cur).map(r => `<tr><td>${esc(r.barcode)}</td><td dir="auto">${esc(r.name)}</td><td>${esc(qq.branch || "")}</td><td class="num">${r.qty}</td><td class="num">${fmt(r.unitPrice)}</td><td class="num">${fmt(r.amount)}</td><td class="num">${r.discPct}%</td><td class="num">${fmt(r.discAmt)}</td><td class="num">${r.vatPct}%</td><td class="num">${fmt(r.vatAmt)}</td><td class="num">${fmt(r.subtotal)}</td></tr>`);
  const tot = (label, v, cls) => `<div class="${cls || ""}"><span>${label}</span><b>${fmt(v)}</b></div>`;
  $("tLines").innerHTML = head(LINE_COLS) + "<tbody>" + (rows.length ? rows.join("") : none(11)) + "</tbody>";
  $("tTotals").innerHTML =
    tot("المجموع قبل الخصم / Subtotal", t.gross) + tot("إجمالي الخصم / Total discount", t.discount) +
    tot("المبلغ الخاضع للضريبة / Taxable amount", net) + tot("إجمالي الضريبة / Total VAT", t.vat) +
    tot("الإجمالي شامل الضريبة / Grand total", t.total, "grand");
  capRows($("tLines"), 6);
}
function renderAll(){ renderHeader(); renderLines(); renderTotals(); renderSaved(); }

/* ---------- editing ---------- */
$("lines").addEventListener("input", e => {
  const tr = e.target.closest("tr"); if (!tr || !e.target.dataset.f) return;
  const l = state.lines[+tr.dataset.i], f = e.target.dataset.f;
  let v = parseFloat(e.target.value);
  if (f === "qty") l.qty = isFinite(v) && v >= 1 ? Math.floor(v) : 1;
  else if (f === "price") l.price = isFinite(v) && v >= 0 ? r2(v) : 0;
  else l.disc = isFinite(v) ? Math.min(100, Math.max(0, v)) : 0;
  tr.querySelector("[data-total]").textContent = fmt(lineTotal(l));
  if (f === "disc"){ const ab = tr.querySelector(".applyAll"); if (ab) ab.hidden = !(+tr.dataset.i < state.lines.length - 1); }
  updateStock();
  renderTotals();
  queueAutosave();
});
$("lines").addEventListener("click", e => {
  const b = e.target.closest("button[data-act]"); if (!b) return;
  const i = +b.closest("tr").dataset.i, l = state.lines[i];
  if (b.dataset.act === "applyall"){
    for (let k = i + 1; k < state.lines.length; k++) state.lines[k].disc = l.disc;
    renderLines(); renderTotals(); queueAutosave();
    return toast("تم تطبيق الخصم \u2066" + l.disc + "%\u2069 على الأصناف التالية\nDiscount " + l.disc + "% applied to the items below");
  }
  if (b.dataset.act === "inc") l.qty += 1;
  if (b.dataset.act === "dec") l.qty = Math.max(1, l.qty - 1);
  if (b.dataset.act === "del"){ state.lines.splice(i,1); }
  renderLines(); renderTotals(); renderHeader();
  queueAutosave();
});
$("cust").addEventListener("input", e => { state.quote.customer = e.target.value; showCustMatches(); queueAutosave(); });
$("branch").addEventListener("change", e => {
  state.quote.branch = e.target.value;
  try{ localStorage.setItem("gqs-branch", e.target.value); }catch(err){}   // the tablet remembers its branch
  queueAutosave();
});
$("vatOn").addEventListener("change", () => { renderTotals(); queueAutosave(); });

/* ---------- save ---------- */
let pushChain = Promise.resolve(), autoT;
function liveSnap(){ return { quote:{...state.quote}, lines:state.lines.map(l => ({...l})), totals:calc(), vatPct:vatRate() }; }
function renderQuoteLinks(){
  const s = state.quote && state.saved[state.quote.id];
  const ok = !!(s && validSheetUrl(s.tabUrl));
  $("quoteLinks").hidden = !ok;
  if (ok){ $("lnkTab").href = s.tabUrl; }
}
function setSync(msg){ $("syncState").textContent = msg; }
/* Every change updates the preview straight away and is sent to the Sheet about a second later. */
function queueAutosave(){
  renderPreview(); renderHeader();
  if (!armed || (!state.lines.length && !state.saved[state.quote.id])) return;
  setSync("تغييرات غير محفوظة… / Unsaved changes…");
  clearTimeout(autoT);
  autoT = setTimeout(() => save({ auto:true }), 900);
}
async function save(opts = {}){
  clearTimeout(autoT);
  const q = state.quote, auto = !!opts.auto;
  if (!state.lines.length && !auto) return toast("أضف صنفاً واحداً على الأقل.\nAdd at least one item first.");
  q.savedAt = Date.now();
  const snap = liveSnap();
  const prev = state.saved[q.id];
  if (prev){ snap.tabUrl = prev.tabUrl; snap.pdfUrl = prev.pdfUrl; }   // keep this quotation's Sheet links
  state.saved[q.id] = snap;
  if (auto){ renderHeader(); renderSaved(); } else renderAll();   // autosave must not redraw the lines while someone is typing
  persist();
  if (!CONFIG.SCRIPT_URL){
    setSync("محفوظ على هذا الجهاز فقط. اربط جدول Google أدناه لإرساله. / Kept on this tablet only. Connect a Google Sheet below to send it.");
    if (!auto) toast("تم الحفظ على هذا الجهاز. اربط جدول Google لإرساله.\nSaved on this tablet. Connect a Google Sheet to send it.");
    return;
  }
  setSync("جارٍ الحفظ في Google Sheets… / Saving to Google Sheets…");
  const payload = JSON.stringify({ key:CONFIG.SECRET, quote:snap.quote, totals:snap.totals, lines:sheetRows(snap) });
  pushChain = pushChain.then(async () => {   // one at a time, in order
    try{
      const res = await fetch(CONFIG.SCRIPT_URL, { method:"POST", headers:{"Content-Type":"text/plain;charset=utf-8"}, body:payload });
      const out = await res.json();
      noteSheetUrl(out);
      if (out.ok && !out.tabUrl){
        setSync("نص Apps Script قديم ولم يحفظ بالتنسيق الجديد. الصق Code.gs الجديد وانشر إصداراً جديداً. / Your Apps Script is out of date: it saved nothing in the new layout. Paste the new Code.gs and deploy a New version.");
        toast("نص Apps Script قديم. حدّث Code.gs.\nYour Apps Script is out of date. Update Code.gs.");
      } else if (out.ok){
        if (validSheetUrl(out.tabUrl) && state.saved[snap.quote.id]){
          state.saved[snap.quote.id].tabUrl = out.tabUrl; persist(); renderQuoteLinks();
        }
        const flag = out.flagged ? " " + out.flagged + " خلية مظللة للمراجعة / cell(s) highlighted for review." : "";
        setSync("تم الحفظ في Google Sheets الساعة " + new Date().toLocaleTimeString([], {hour:"2-digit", minute:"2-digit"}) + (out.tabName ? " (الورقة: " + out.tabName + ")" : "") + "." + flag + " / Saved to Google Sheets");
        if (!auto) toast("تم الحفظ في Google Sheets.\nSaved to Google Sheets." + flag);
      } else { setSync("رفض الجدول الحفظ: " + out.error + " / The Sheet rejected the save"); toast("رفض الجدول الحفظ: " + out.error + "\nSheet rejected the save"); }
    }catch(err){ setSync("تعذر الوصول إلى الجدول. التغييرات محفوظة على هذا الجهاز. / Could not reach the Sheet. Changes are kept on this tablet."); if (!auto) toast("تعذر الوصول إلى الجدول. العرض محفوظ على هذا الجهاز.\nCould not reach the Sheet. The quotation is kept on this tablet."); }
  });
}
$("saveBtn").onclick = () => save();
if ($("exportBtn")) $("exportBtn").onclick = () => exportQuotation();
$("newBtn").onclick = newQuote;

/* ---------- delete a quotation (tablet and Google Sheet) ---------- */
$("delBtn").onclick = () => {
  const q = state.quote;
  $("delMsg").textContent = "حذف " + q.id + (q.customer.trim() ? " (" + q.customer.trim() + ")" : "") +
    (CONFIG.SCRIPT_URL ? " وصفوفه في Google Sheet؟" : "؟") + " لا يمكن التراجع. / Delete" + (CONFIG.SCRIPT_URL ? " incl. Sheet rows" : "") + "? This cannot be undone.";
  $("delConfirm").hidden = false;
};
$("delNo").onclick = () => { $("delConfirm").hidden = true; };
$("delYes").onclick = () => deleteQuote();
function deleteQuote(){
  const id = state.quote.id, wasSaved = !!state.saved[id];
  clearTimeout(autoT);
  delete state.saved[id];
  try{ const n = quoteNo(id), c = Number(localStorage.getItem("gqs-counter")) || 0; if (n > c) localStorage.setItem("gqs-counter", String(n)); }catch(e){}
  persist();
  newQuote();
  toast("تم حذف " + id + "\n" + id + " deleted");
  if (!CONFIG.SCRIPT_URL || !wasSaved){ setSync("تم حذف " + id + " من هذا الجهاز. / " + id + " deleted from this tablet."); return; }
  setSync("جارٍ حذف " + id + " من Google Sheets… / Deleting " + id + " from Google Sheets…");
  const payload = JSON.stringify({ key:CONFIG.SECRET, delete:true, quoteId:id });
  pushChain = pushChain.then(async () => {
    try{
      const res = await fetch(CONFIG.SCRIPT_URL, { method:"POST", headers:{"Content-Type":"text/plain;charset=utf-8"}, body:payload });
      const out = await res.json();
      noteSheetUrl(out);
      if (!out.ok) setSync(out.error === "Bad quote id"
        ? "نص Apps Script قديم، لم يُحذف " + id + " من الجدول. الصق Code.gs الجديد وانشر إصداراً جديداً. / Apps Script is out of date, " + id + " NOT removed from the Sheet."
        : "تعذر حذف " + id + " من الجدول: " + out.error + " / The Sheet could not delete it");
      else if (!out.tab && !out.rows && !out.summary) setSync(id + " غير موجود في الجدول. / not found in the Sheet (nothing to remove).");
      else setSync("تم حذف " + id + " من Google Sheets" + (out.tab ? " (ورقته وصفه في Quotations)." : ".") + " / deleted from Google Sheets");
    }catch(err){ setSync("تعذر الوصول إلى الجدول. حُذف " + id + " هنا لكن صفوفه ما زالت في الجدول. / Could not reach the Sheet; its rows are still there."); }
  });
}
$("saved").addEventListener("click", e => {
  const b = e.target.closest("[data-open]"); if (!b) return;
  const s = state.saved[b.dataset.open];
  $("delConfirm").hidden = true;
  state.quote = {...s.quote}; state.lines = s.lines.map(l => ({...l}));
  $("vatOn").checked = s.vatPct !== 0;
  $("cust").value = state.quote.customer; syncExtra(); renderAll();
  toast("تم فتح " + state.quote.id + "\nOpened " + state.quote.id);
});
function persist(){ try{ localStorage.setItem("gqs-saved", JSON.stringify(state.saved)); }catch(e){} }
function restore(){ try{ const s = JSON.parse(localStorage.getItem("gqs-saved")||"{}"); if (s && typeof s==="object") state.saved = s; }catch(e){} }


/* ---------- customer quotation PDF (print layout) ---------- */
const COMPANY = {
  nameAr:"مؤسسة بدر عمر الخنبشي التجارية", addrAr:"جدة - شارع ابوبكر الصديق - الشرفية", crAr:"سجل تجاري 4030197794",
  nameEn:"Badr O. Alkhanbashi Est. for Trading", addrEn:"Jeddah, Abu Bakr as Siddiq Street", crEn:"CR 4030197794",
  vat:"300234539600003", logo:""
};
const EXTRA = { custNo:"custNo", custVat:"custVat", custCity:"custCity", notes:"notes", custPhone:"custPhone", docNo:"docNo", validity:"validity",custAddress:"custAddress",custCommReg:"custCommReg" };
function syncExtra(){ Object.keys(EXTRA).forEach(k => { $(EXTRA[k]).value = state.quote[k] || ""; }); }
Object.keys(EXTRA).forEach(k => $(EXTRA[k]).addEventListener("input", e => { state.quote[k] = e.target.value; queueAutosave(); }));

/* Arabic amount in words, riyals and halalas. */
/* Arabic word for "day" that matches the number: 1 يوم, 2 يومان, 3-10 أيام, 11+ يوماً */
function dayWord(v){
  const n = parseInt(String(v || "").replace(/[٠-٩]/g, d => "٠١٢٣٤٥٦٧٨٩".indexOf(d)), 10);
  if (!isFinite(n) || n === 1) return "يوم";
  if (n === 2) return "يومان";
  if (n >= 3 && n <= 10) return "أيام";
  return "يوماً";
}
function tafqeet(amount){
  const ones = ["","واحد","اثنان","ثلاثة","أربعة","خمسة","ستة","سبعة","ثمانية","تسعة","عشرة","أحد عشر","اثنا عشر","ثلاثة عشر","أربعة عشر","خمسة عشر","ستة عشر","سبعة عشر","ثمانية عشر","تسعة عشر"];
  const tens = ["","","عشرون","ثلاثون","أربعون","خمسون","ستون","سبعون","ثمانون","تسعون"];
  const hund = ["","مائة","مئتان","ثلاثمائة","أربعمائة","خمسمائة","ستمائة","سبعمائة","ثمانمائة","تسعمائة"];
  const u1000 = n => { const p = [], h = Math.floor(n/100), r = n % 100;
    if (h) p.push(hund[h]);
    if (r) { if (r < 20) p.push(ones[r]); else { if (r % 10) p.push(ones[r % 10]); p.push(tens[Math.floor(r/10)]); } }
    return p.join(" و"); };
  const scale = (n, one, two, few, many) => n === 1 ? one : n === 2 ? two : (n >= 3 && n <= 10) ? u1000(n) + " " + few : u1000(n) + " " + many;
  const words = n => { if (n === 0) return "صفر";
    const p = [], m = Math.floor(n / 1e6), t = Math.floor((n % 1e6) / 1000), r = n % 1000;
    if (m) p.push(scale(m, "مليون", "مليونان", "ملايين", "مليون"));
    if (t) p.push(scale(t, "ألف", "ألفان", "آلاف", "ألف"));
    if (r) p.push(u1000(r));
    return p.join(" و"); };
  const total = Math.round(amount * 100), riyal = Math.floor(total / 100), halala = total % 100;
  const parts = [];
  if (riyal || !halala) parts.push(words(riyal) + " ريال");
  if (halala) parts.push(words(halala) + " هللة");
  return parts.join(" و") + " فقط لا غير";
}

const SAR = '<svg class="sar" viewBox="0 0 1124.14 1256.39" aria-label="SAR"><path fill="currentColor" d="M699.62,1113.02h0c-20.06,44.48-33.32,92.75-38.4,143.37l424.51-90.24c20.06-44.47,33.31-92.75,38.4-143.37l-424.51,90.24Z"/><path fill="currentColor" d="M1085.73,895.8c20.06-44.47,33.32-92.75,38.4-143.37l-330.68,70.33v-135.2l292.27-62.11c20.06-44.47,33.32-92.75,38.4-143.37l-330.68,70.27V66.13c-50.67,28.45-95.67,66.32-132.25,110.99v403.35l-132.25,28.11V0c-50.67,28.44-95.67,66.32-132.25,110.99v525.69l-295.91,62.88c-20.06,44.47-33.33,92.75-38.42,143.37l334.33-71.05v170.26l-358.3,76.14c-20.06,44.47-33.32,92.75-38.4,143.37l375.04-79.7c30.53-6.35,56.77-24.4,73.83-49.24l68.78-101.97v-.02c7.14-10.55,11.3-23.27,11.3-36.97v-149.98l132.25-28.11v270.4l424.53-90.28Z"/></svg>';
const money = n => SAR + " " + fmt(n);
const mf = n => `<div class="mf"><span class="sym">${SAR}</span><span class="n">${fmt(n)}</span></div>`;
function printHtml(){
  const q = state.quote, t = calc(), vp = vatRate();
  const rows = state.lines.map((l, i) => { const c = lineCalc(l, vp);
    return `<tr class="bx">
      <td class="sub">${fmt(c.total)}</td><td class="nw">${fmt(c.vatAmt)}</td><td class="nw">${fmt(c.discAmt)}</td><td class="nw">${l.disc}%</td>
      <td class="nw">${fmt(c.taxable)}</td><td class="nw">${fmt(l.price)}</td><td class="nw">${l.qty}</td><td class="pic"><img class="pic" data-serial="${esc(l.serial)}" alt=""></td>
      <td class="desc">${esc(l.ar || l.en)}</td><td class="nw">${esc(l.serial)}</td><td class="nw">${i + 1}</td></tr>`; }).join("");
  const W0 = ["10%","8.8%","9.8%","6%","9.8%","7.6%","5%","20%","12.8%","8.6%","3.8%"];   // widths, left to right: Subtotal, Tax, Discount, %, Taxable, Unit price, QTY, Picture, Description, Barcode, #  (any numbers: they are scaled to add up to 100%)
  const W = W0.map(w => (parseFloat(w) * 100 / W0.reduce((s, x) => s + parseFloat(x), 0)).toFixed(3) + "%");
  // The blue boxes (title + customer labels) start exactly where the Description column starts below them
  const blueW = parseFloat(W[8]) + parseFloat(W[9]) + parseFloat(W[10]);
  const remW = 100 - blueW, kk = remW / 79, sc = x => (x * kk).toFixed(3) + "%";
  const tot = [
    [t.gross, "Total (Excluding VAT)", "الإجمالي ( غير شامل ضريبة القيمة المضافة )"],
    [t.discount, "Discount", "مجموع الخصومات"],
    [t.net, "Total Taxable Amount (Excluding VAT)", "الإجمالي الخاضع للضريبة ( غير شامل ضريبة القيمة المضافة )"],
    [t.vat, "Total VAT", "مجموع ضريبة القيمة المضافة"],
    [t.total, "Total Amount Due", "إجمالي المبلغ المستحق"]
  ].map(r => `<tr><td class="v" style="width:22%">${mf(r[0])}</td><td class="en" style="width:34%">${r[1]}</td><td class="ar">${r[2]}</td></tr>`).join("");
  const header = `
  <table class="head"><tr>
    <td class="en" style="width:50%"><div class="big">${COMPANY.nameEn}</div><div>${COMPANY.addrEn}</div><div>${COMPANY.crEn}</div></td>
    <td class="ar" style="width:50%"><div class="big">${COMPANY.nameAr}</div><div>${COMPANY.addrAr}</div><div>${COMPANY.crAr}</div></td>
  </tr></table>
 <table style="margin:2px 0; width:100%; text-align:center;">
  <tr>
   <td style="width:${(remW*0.543).toFixed(3)}%; padding:0; vertical-align:bottom; text-align:center;">

   <table class="kv" style="width:100%; text-align:center; border-collapse: separate; border-radius: 6px; border-spacing: 0; box-shadow: none;">
    <tr>
      <!-- Decreased to 6px border radius -->
      <td class="en" style="text-align:center; border-top-left-radius: 6px; padding:3px;">Issue Date</td>
      <td class="c" style="text-align:center; padding:3px;">${esc(q.date)}</td>
      <td class="ar" style="text-align:center; border-top-right-radius: 6px; padding:3px;">التاريخ</td>
    </tr>
    <tr>
      <td class="en" style="text-align:center; padding:3px;">Doc No.</td>
      <td class="c" style="text-align:center; color:#c00000; padding:3px;">${esc(q.docNo || "")}</td>
      <td class="ar" style="text-align:center; padding:3px;"><bdi>رقم المستند</bdi></td>
    </tr>
    <tr>
      <!-- Decreased to 6px border radius -->
      <td class="en" style="text-align:center; border-bottom-left-radius: 6px; padding:3px;">VAT NO.</td>
      <td class="c" style="text-align:center; color:#0094d6; padding:3px;">${COMPANY.vat}</td>
      <td class="ar" style="text-align:center; border-bottom-right-radius: 6px; padding:3px;">الرقم الضريبي</td>
    </tr>
  </table>
</td>


    <td class="c" style="width:${(remW*0.457).toFixed(3)}%; text-align:center; vertical-align:middle;">
      <img src="${COMPANY.logo}" alt=""
        style="display:block; margin:0 auto; width:auto; height:auto; max-width:96%; max-height:22mm">
    </td>

    <td id= "quoteHeader" style="width:${blueW.toFixed(3)}%; padding:0; vertical-align:bottom; text-align:center;">
      <div class="title-ar" style="text-align:center; font-size:19px;">عرض سعر</div>
      <div class="title-en" style="text-align:center;font-size:18px;">Quotation</div>
    </td>
  </tr>
</table>
  <table class="bx cu">
    <tr><td class="vn" style="width:${sc(15)}">${esc(q.custNo || "")}</td><td class="cl" style="width:${sc(12)}">رقم العميل<br>Customer No</td>
      <td class="v ar cn" colspan="3" style="width:${sc(52)}">${esc(q.customer || "")}</td><td class="cla" dir="rtl" style="width:${blueW.toFixed(3)}%">اسم العميل <br>Customer Name</td></tr>
    <tr><td class="v" style="font-size:9px">${esc(q.custCity || "")}</td><td class="cl">البلد - المدينة<br>Country - City</td>
      <td class="v" style="width:${sc(17)}">${esc(q.custVat || "")}</td><td class="cl" style="width:${sc(12)}">الرقم الضريبي<br>VAT Number</td>
      <td class="v" style="width:${sc(23)}; font-weight: bold">${esc(q.custPhone || "")}</td><td class="cla" dir="rtl">رقم الجوال <br> Mobile No</td></tr>
      <tr>
      <td class="v" style="width:${sc(12)}">${esc(q.custCommReg || "-")}</td><td class="cla" dir="rtl">السجل التجاري <br>CR</td>
      
      <td class="v ar" colspan="3" style="width:${sc(52)}">${esc(q.custAddress || "")}</td><td class="cla" dir="rtl" style="width:${blueW.toFixed(3)}%">العنوان<br>Address</td></tr>

  </table>
  <div class="hgap"></div>`;
  const body = `
  <table class="items bx"><colgroup>${W.map(w => `<col style="width:${w}">`).join("")}</colgroup>
    <thead>
<tr>
  <th class="pu" style="text-align:center;">المجموع شامل<br>الضريبة</th>
  <th class="pu sm" style="text-align:center;">ضريبة القيمة<br>المضافة <bdi>${vp}%</bdi></th>
  <th class="pu sm" style="text-align:center;" colspan="2">الخصم</th>
  <th class="pu sm" style="text-align:center;">المبلغ الخاضع<br>للضريبة</th>
  <th class="pu" style="text-align:center;">سعر الوحدة</th>
  <th class="ma" style="text-align:center;">الكمية</th>
  <th class="ma" style="text-align:center;">الصورة</th>
  <th class="ma" style="text-align:center;">الوصف</th>
  <th class="ma" style="text-align:center;">رقم الصنف</th>
  <th class="ma" rowspan="2" style="text-align:center;">#</th>
</tr>
    <tr>
  <th class="lb" style="text-align:center; font-weight: bold;">Item Subtotal<br>(Including VAT)</th>
  <th class="lb" style="text-align:center; font-weight: bold;">Tax ${vp} %</th>
  <th class="lb" style="text-align:center; font-weight: bold;">Discount</th>
  <th class="lb" style="text-align:center; font-weight: bold;">%</th>
  <th class="lb" style="text-align:center; font-weight: bold;">Taxable Amount</th>
  <th class="lb" style="text-align:center; font-weight: bold;">Unit price</th>
  <th class="gb" style="text-align:center; font-weight: bold;">QTY</th>
  <th class="gb" style="text-align:center; font-weight: bold;">Picture</th>
  <th class="gb" style="text-align:center; font-weight: bold;">Description</th>
  <th class="gb" style="text-align:center; font-weight: bold;">Barcode</th>
</tr>
      </thead><tbody>${rows}</tbody></table>
  <div class="foot">
  <br>
  <table class="tot cnt"><tr><td class="en" style="width:${W.slice(0,6).reduce((s,w)=>s+parseFloat(w),0).toFixed(3)}%; font-size:10px;" >Total Quantity</td><td class="v" style="width:${W[6]}">${state.lines.reduce((s, l) => s + (+l.qty || 0), 0)}</td><td class="ar">إجمالي الكمية</td></tr></table>
  <table class="tot" style="margin-top:3mm">${tot}</table>
  <div class="words ar">${tafqeet(t.total)}</div>
  <table class="notes">
    <tr><td rowspan="2" style="width:26%"><div style="height:30mm;overflow:hidden;text-align:center;padding-top:1mm"><img src="stamp.png" alt="" style="height:27mm;width:auto;max-width:94%;object-fit:contain"></div></td>
      <td class="ar nt" style="width:60%;height:18mm">${esc(q.notes || "")}</td><td class="ar c" style="width:14%"><bdi>ملاحظات</bdi><br>Notes</td></tr>
    <tr><td class="ar vd" style="height:10mm">العرض ساري لمدة ( <b style="display:inline-block;min-width:7mm;text-align:center">${esc(q.validity || "")}</b> ) ${dayWord(q.validity)}</td><td class="ar c"><bdi>صلاحية العرض</bdi><br>Validity</td></tr>
  </table></div>`;
  return `<table class="page"><thead><tr><td>${header}</td></tr></thead><tbody><tr><td>${body}</td></tr></tbody></table>`;
}
$("pdfBtn").addEventListener("click", () => {
if (!state.lines.length)
  return toast("أضف صنفاً واحداً على الأقل.\nAdd at least one item first.");
  $("printSheet").innerHTML = printHtml();
  /* Arabic stays bold; English words (labels, headings, address) are set in normal weight. Values stay bold. */
  const tw = document.createTreeWalker($("printSheet"), NodeFilter.SHOW_TEXT), nodes = [];
  while (tw.nextNode()) nodes.push(tw.currentNode);
  nodes.forEach(n => {
    if (!/[A-Za-z]/.test(n.nodeValue) || n.parentElement.closest(".items tbody, .v, .vn, .c, .stamp, .notes")) return;
    const span = document.createElement("span");
    span.innerHTML = esc(n.nodeValue).replace(/[A-Za-z(][A-Za-z0-9 .,()\-%'&]*/g, m => '<span class="nb">' + m + '</span>');
    n.replaceWith(...span.childNodes);
  });
  document.body.classList.add("pq");
  /* "Save as PDF" proposes the page title as the file name, so use the Doc no. while printing. */
  const oldTitle = document.title;
  document.title = tabName(state.quote).replace(/[\\\/:*?"<>|\[\]]/g, "-");
  const done = () => { document.body.classList.remove("pq"); document.title = oldTitle; window.removeEventListener("afterprint", done); };
  window.addEventListener("afterprint", done);
  /* Item pictures: items/<serial>.jpg (also .jpeg .png .webp). Missing pictures leave the space empty. */
  const exts = ["jpg","jpeg","png","webp"];
  const settle = im => new Promise(res => {
    let k = 0; const serial = encodeURIComponent(im.dataset.serial);
    im.onload = () => res(); 
    im.onerror = () => { k++; if (k < exts.length) im.src = "items/" + serial + "." + exts[k]; else { im.style.display = "none"; res(); } };
    im.src = "items/" + serial + "." + exts[0];
  });
  const fixed = [...$("printSheet").querySelectorAll("img:not(.pic)")].map(im => im.complete ? 0 : new Promise(r => { im.onload = r; im.onerror = r; }));
  const pics = [...$("printSheet").querySelectorAll("img.pic")].map(settle);
  Promise.race([Promise.all([...fixed, ...pics]), new Promise(r => setTimeout(r, 15000))]).then(() => setTimeout(() => window.print(), 60));
});

/* ---------- toast ---------- */
let tt, tf;
function toast(m, ms = 3200){
  const t = $("toast"); t.textContent = m; t.classList.remove("out"); t.hidden = false;
  clearTimeout(tt); clearTimeout(tf);
  tf = setTimeout(() => t.classList.add("out"), ms - 600);   // fade over the last 0.6 s
  tt = setTimeout(() => { t.hidden = true; t.classList.remove("out"); }, ms);
}

/* ---------- search the item list ---------- */
const norm = s => String(s).toLowerCase().replace(/[\u064B-\u0652\u0640]/g,"").replace(/[أإآ]/g,"ا").replace(/ة/g,"ه").replace(/ى/g,"ي");
let searchIndex = null;
function searchItems(q){
  if (!searchIndex) searchIndex = Object.keys(ITEMS).map(s => [s, norm(ITEMS[s][0]), s.toLowerCase()]);
  const toks = norm(q).split(/\s+/).filter(Boolean); if (!toks.length) return [];
  const out = [];
  for (const [s, n, sl] of searchIndex){
    if (toks.every(t => n.includes(t) || sl.includes(t))){ out.push(s); if (out.length >= 12) break; }
  }
  return out;
}
$("searchIn").addEventListener("input", () => {
  const list = searchItems($("searchIn").value);
  $("searchOut").innerHTML = list.map(s => { const r = ITEMS[s], left = leftQty(s);
    return `<button type="button" class="hit" data-s="${esc(s)}"><b dir="auto">${esc(r[0] || s)}</b><span><code>${esc(s)}</code> · ${r[1] ? fmt(r[1]) + " ر.س" : "بدون سعر"} · المتبقي / Left: ${left}</span></button>`; }).join("")
    || ($("searchIn").value.trim() ? '<p class="hint">لا يوجد صنف / No item found.</p>' : "");
});
$("searchOut").addEventListener("click", e => {
  const b = e.target.closest("[data-s]"); if (!b) return;
  addItem(itemBySerial(b.dataset.s)); $("searchIn").value = ""; $("searchOut").innerHTML = "";
});
$("pasteForm").addEventListener("submit", e => {
  e.preventDefault();
  const v = $("searchIn").value.trim(); if (!v) return;
  const direct = (window.GQS_QR && GQS_QR.decode(v)) || ITEMS[v] || v.includes("|");   // a scanned/typed code or exact serial
  if (direct) handleScan(v);
  else { const list = searchItems(v); if (list.length) addItem(itemBySerial(list[0])); else toast("لا يوجد صنف باسم \"" + v.slice(0,40) + "\"\nNo item found"); }
  $("searchIn").value = ""; $("searchOut").innerHTML = "";
});


/* ---------- update items from an Excel file (shared through the Google Sheet) ---------- */
const ITEMS_KEY = "gqs-items";
let itemsStamp = 0, pending = null;
function setItems(map, stamp){
  Object.keys(ITEMS).forEach(k => delete ITEMS[k]); Object.assign(ITEMS, map);
  searchIndex = null; itemsStamp = stamp || 0;
}
function saveItemsLocal(){ try{ localStorage.setItem(ITEMS_KEY, JSON.stringify({ updatedAt:itemsStamp, items:ITEMS })); }catch(e){} }
try{ const s = JSON.parse(localStorage.getItem(ITEMS_KEY) || "null"); if (s && s.items && Object.keys(s.items).length) setItems(s.items, s.updatedAt || 1); }catch(e){}
/* "last update: date // time" and "total" lines under each update box */
function updInfoHtml(stamp, builtIn, label, n, extra){
  const d = new Date(stamp);
  const when = stamp > 1
    ? d.toLocaleDateString("en-GB") + " // " + d.toLocaleTimeString("en-GB", { hour:"2-digit", minute:"2-digit" })
    : builtIn;
  return `<div><span>آخر تحديث / Last update:</span> <b dir="auto">${when}</b></div>` +
         `<div><span>${label}:</span> <b>${n}</b></div>` + (extra ? `<div class="hint">${extra}</div>` : "");
}
function showItemsInfo(extra){
  const n = Object.keys(ITEMS).length;
  $("itemsInfo").innerHTML = updInfoHtml(itemsStamp, "القائمة المدمجة / built-in list", "مجموع الأصناف / Total items", n, extra ? esc(extra) : "");
}
const toNum = v => { const n = Number(String(v == null ? "" : v).replace(/[^\d.\-]/g, "")); return isFinite(n) ? n : 0; };
async function readExcel(file){
  const ok = await loadChain(["https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js",
                              "https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js",
                              "https://unpkg.com/xlsx@0.18.5/dist/xlsx.full.min.js"], () => !!window.XLSX);
  if (!ok) throw new Error("تعذر تحميل قارئ إكسل. تحقق من الإنترنت وحاول مرة أخرى. / The Excel reader did not load.");
  const wb = XLSX.read(await file.arrayBuffer(), { type:"array" });
  const rows = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { header:1, raw:false, defval:"" });
  const hi = rows.findIndex(r => r.some(c => String(c).trim() === "الباركود"));
  if (hi < 0) throw new Error('Could not find the column "الباركود" in the first sheet.');
  const col = {}; rows[hi].forEach((c, i) => { if (String(c).trim()) col[String(c).trim()] = i; });
  for (const need of ["الباركود", "الصنف", "سعر البيع", "العدد"]) if (!(need in col)) throw new Error('Missing column: ' + need);
  const map = {}; let skipped = 0, noPrice = 0, noName = 0, inStock = 0;
  rows.slice(hi + 1).forEach(r => {
    const s = String(r[col["الباركود"]]).trim();
    if (!s) { skipped++; return; }
    const name = String(r[col["الصنف"]]).trim(), price = Math.round(toNum(r[col["سعر البيع"]]) * 100) / 100, qty = Math.round(toNum(r[col["العدد"]]));
    if (!price) noPrice++; if (!name) noName++; if (qty > 0) inStock++;
    map[s] = [name, price, qty];
  });
  const count = Object.keys(map).length;
  if (!count) throw new Error("لا توجد أصناف في هذا الملف. / No items found in that file.");
  return { map, count, skipped, noPrice, noName, inStock };
}
async function pushItems(map){
  const items = Object.keys(map).map(s => [s, map[s][0], map[s][1], map[s][2]]);
  const res = await fetch(CONFIG.SCRIPT_URL, { method:"POST", headers:{"Content-Type":"text/plain;charset=utf-8"}, body:JSON.stringify({ key:CONFIG.SECRET, putItems:true, items }) });
  const out = await res.json();
  if (!out.ok) throw new Error(out.error === "Bad quote id" ? "Your Apps Script is out of date. Paste the new Code.gs and deploy a New version." : out.error);
  return out.updatedAt;
}
$("itemsFile").addEventListener("change", async e => {
  const f = e.target.files[0]; e.target.value = ""; if (!f) return;
  $("itemsApply").hidden = $("itemsCancel").hidden = true; pending = null;
  $("itemsMsg").textContent = "جارٍ قراءة " + f.name + "… / Reading…";
  try{
    pending = await readExcel(f);
    const p = pending;
    $("itemsMsg").textContent = p.count + " items found (" + p.inStock + " in stock" + (p.noPrice ? ", " + p.noPrice + " without a price" : "") + (p.noName ? ", " + p.noName + " without a name" : "") + (p.skipped ? "; " + p.skipped + " rows without a barcode skipped" : "") + "). This replaces the current list. Apply?";
    $("itemsApply").hidden = $("itemsCancel").hidden = false;
  }catch(err){ $("itemsMsg").textContent = err.message || "تعذرت قراءة هذا الملف. / Could not read that file."; }
});
$("itemsCancel").onclick = () => { pending = null; $("itemsApply").hidden = $("itemsCancel").hidden = true; $("itemsMsg").textContent = "تم الإلغاء. لم تتغير قائمة الأصناف. / Cancelled. The item list was not changed."; };
$("itemsApply").onclick = async () => {
  if (!pending) return;
  const map = pending.map; pending = null; $("itemsApply").hidden = $("itemsCancel").hidden = true;
  setItems(map, Date.now()); saveItemsLocal(); renderLines(); showItemsInfo();
  if (!CONFIG.SCRIPT_URL){ $("itemsMsg").textContent = "تم تحديث قائمة الأصناف على هذا الجهاز. اربط Google Sheet لمشاركتها مع الأجهزة الأخرى. / Item list updated on this tablet. Connect the Google Sheet to share it."; return; }
  $("itemsMsg").textContent = "تم التحديث هنا. جارٍ المشاركة مع الأجهزة الأخرى… / Updated here. Sharing with the other tablets…";
  try{
    const t = await pushItems(map); itemsStamp = t; saveItemsLocal(); showItemsInfo();
    $("itemsMsg").textContent = "تم تحديث قائمة الأصناف ومشاركتها. ستصل إلى الأجهزة الأخرى عند فتح الصفحة. / Item list updated and shared.";
  }catch(err){ $("itemsMsg").textContent = "تم التحديث على هذا الجهاز فقط. تعذرت المشاركة: " + (err.message || err); }
};
/* Other tablets pick up the newest shared list when the page opens. */
async function pullItems(){
  if (!CONFIG.SCRIPT_URL) return;
  try{
    const res = await fetch(CONFIG.SCRIPT_URL, { method:"POST", headers:{"Content-Type":"text/plain;charset=utf-8"}, body:JSON.stringify({ key:CONFIG.SECRET, getItems:true, since:itemsStamp }) });
    const out = await res.json();
    if (!out.ok || out.unchanged || !out.items || !out.items.length || out.updatedAt <= itemsStamp) return;
    const map = {}; out.items.forEach(r => { map[r[0]] = [r[1], r[2], r[3]]; });
    setItems(map, out.updatedAt); saveItemsLocal(); renderAll(); showItemsInfo("from the Google Sheet");
  }catch(e){}
}


/* ---------- export the quotation as an Excel file (same table as its Google Sheet tab) ---------- */
const XLSX_URLS = ["https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js",
                   "https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js",
                   "https://unpkg.com/xlsx@0.18.5/dist/xlsx.full.min.js"];
const loadXlsx = () => loadChain(XLSX_URLS, () => !!window.XLSX);
async function exportQuotation(){
  if (!state.lines.length) return toast("أضف صنفاً واحداً على الأقل.\nAdd at least one item first.");
  if (!(await loadXlsx())) return toast("تعذر تحميل أداة إكسل. تحقق من الإنترنت.\nThe Excel writer did not load.");
  const q = state.quote, snap = liveSnap();
  /* Same layout as the Google Sheet tab and the other program's export: A = checkbox column, B..L = the 11 columns,
     M = notes, N = spare, totals row under the items. Three headers start with a space on purpose (the other program matches them). */
  const heads = [" ", " الباركود على العلبة", "اسم الصنف عربي", " رقم الفرع", " العدد", "السعر", "المبلغ", "نسبة الخصم", "قيمة الخصم", "نسبة الضريبة", "قيمة الضريبة", "الإجمالي", "ملاحظات", ""];
  const rows = sheetRows(snap).map(r => [false, String(r.barcode), r.name, q.branch || "", r.qty, r.unitPrice, r.amount, r.discPct / 100, r.discAmt, r.vatPct / 100, r.vatAmt, r.subtotal, "", ""]);
  const ws = XLSX.utils.aoa_to_sheet([heads, ...rows]);
  const fm = ["General", "@", "@", "@", "#,##0.00", "#,##0.00", "#,##0.00", "#,##0.00%", "#,##0.00", "#,##0.00%", "#,##0.00", "#,##0.00", "General", "General"];
  rows.forEach((_, i) => fm.forEach((f, c) => { const a = XLSX.utils.encode_cell({ r:i + 1, c }); if (ws[a]) ws[a].z = f; }));
  const last = rows.length + 1, tr = last + 1;                       // 1-based sheet rows: items 2..last, totals row tr
  const sum = col => rows.reduce((s, r) => s + (+r[col] || 0), 0);
  [[4, "E"], [6, "G"], [8, "I"], [10, "K"]].forEach(([c, L]) => { ws[L + tr] = { t:"n", v:sum(c), f:"SUBTOTAL(9," + L + "2:" + L + last + ")", z:"#,##0.00" }; });
  ws["L" + tr] = { t:"s", v:"SUM=" + fmt(sum(11)), f:'CONCATENATE("SUM=",TEXT(SUBTOTAL(9,L2:L' + last + '), "#,##0.00"),"")' };
  ws["!ref"] = "A1:N" + tr;
  ws["!autofilter"] = { ref:"A1:N" + last };
  ws["!cols"] = [21, 132, 216, 74, 73, 55, 89, 83, 86, 102, 92, 74, 80, 29].map(w => ({ wpx:w }));
  const nm = tabName(q).replace(/[\\\/:*?"<>|\[\]]/g, "-");
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, nm.slice(0, 31) || "Quotation");
  wb.Workbook = { Views:[{ RTL:true }] };   // column A on the right, like the Google Sheet tab
  XLSX.writeFile(wb, nm + ".xlsx");
  toast("تم التصدير " + nm + ".xlsx\nExported");
}

/* ---------- import an Excel file into the quotation: serial (barcode) + qty + discount % per row ---------- */
async function importQuoteExcel(file){
  if (!(await loadXlsx())) throw new Error("تعذر تحميل قارئ إكسل. تحقق من الإنترنت.\nThe Excel reader did not load.");
  const wb = XLSX.read(await file.arrayBuffer(), { type:"array" });
  const ws = wb.Sheets[wb.SheetNames[0]];
  const txt = XLSX.utils.sheet_to_json(ws, { header:1, raw:false, defval:"" });   // as displayed (barcodes stay text)
  const raw = XLSX.utils.sheet_to_json(ws, { header:1, raw:true, defval:"" });    // real numbers (discount 0.1 vs "10%")
  const clean = s => String(s).replace(/\s+/g, " ").trim();
  const find = (head, names) => head.findIndex(c => names.some(n => clean(c).startsWith(n)));
  const hi = txt.findIndex(r => find(r, ["الباركود"]) >= 0);
  if (hi < 0) throw new Error("لم أجد عمود الباركود في الملف.\nCould not find the barcode column.");
  const head = txt[hi], cB = find(head, ["الباركود"]), cQ = find(head, ["العدد", "الكمية"]), cD = find(head, ["نسبة الخصم", "الخصم"]);
  let added = 0, merged = 0; const missing = [];
  for (let i = hi + 1; i < txt.length; i++){
    const serial = clean(txt[i][cB]); if (!serial || serial === "false" || serial.startsWith("SUM")) continue;
    const it = itemBySerial(serial);
    if (!it){ missing.push(serial); continue; }
    let qty = cQ >= 0 ? Math.round(toNum(raw[i][cQ])) : 1; if (!(qty > 0)) qty = 1;
    let disc = 0;
    if (cD >= 0){
      const rv = raw[i][cD];
      if (typeof rv === "number") disc = rv > 0 && rv <= 1 ? rv * 100 : rv;     // 0.1 stored as 10%
      else disc = toNum(rv);                                                    // text such as "10%" or "10"
    }
    disc = Math.min(100, Math.max(0, Math.round(disc * 100) / 100));
    const ex = state.lines.find(l => l.serial === serial);
    if (ex){ ex.qty += qty; ex.disc = disc; merged++; }
    else { state.lines.push({ ...it, qty, disc, manualPrice: !it.price }); added++; }
  }
  return { added, merged, missing };
}
if ($("importBtn")) $("importBtn").onclick = () => $("importFile").click();
if ($("importFile")) $("importFile").addEventListener("change", async e => {
  const f = e.target.files[0]; e.target.value = ""; if (!f) return;
  toast("جارٍ قراءة " + f.name + "…\nReading…");
  try{
    const { added, merged, missing } = await importQuoteExcel(f);
    if (!added && !merged){
      return toast(missing.length ? "لم يوجد أي صنف من الملف في قائمة الأصناف (" + missing.length + ").\nNone of the file's items were found in the item list." : "لا توجد أصناف في هذا الملف.\nNo items found in that file.", 5000);
    }
    renderLines(); renderTotals(); queueAutosave();
    toast("تمت إضافة " + (added + merged) + " صنف إلى العرض" + (missing.length ? " · غير موجود: " + missing.length : "") + "\nAdded " + (added + merged) + " items to the quotation" + (missing.length ? " · not found: " + missing.slice(0, 5).join(", ") + (missing.length > 5 ? "…" : "") : ""), 7000);
  }catch(err){ toast(err.message || "تعذرت قراءة هذا الملف.\nCould not read that file.", 5000); }
});

/* ---------- customers: search by name, tap to fill in the rest ---------- */
/* Each row: [customer no, name, VAT no, address, commercial registration, country, mobile]. */
const CUST_KEY = "gqs-customers";
let CUSTOMERS = [], custStamp = 0, custPending = null, custIndex = null;
function setCustomers(rows, stamp){ CUSTOMERS = rows; custStamp = stamp || 0; custIndex = null; }
function saveCustLocal(){ try{ localStorage.setItem(CUST_KEY, JSON.stringify({ updatedAt:custStamp, rows:CUSTOMERS })); }catch(e){ toast("تعذر حفظ قائمة العملاء في هذا المتصفح.\nThis browser could not store the customer list."); } }
try{ const s = JSON.parse(localStorage.getItem(CUST_KEY) || "null"); if (s && Array.isArray(s.rows)) setCustomers(s.rows, s.updatedAt || 1); }catch(e){}
function showCustInfo(extra){
  const el = $("custInfo"); if (!el) return;
  el.innerHTML = updInfoHtml(custStamp, "—", "مجموع العملاء / Total customers", CUSTOMERS.length, extra ? esc(extra) : "");
}
function searchCustomers(q){
  if (!custIndex) custIndex = CUSTOMERS.map(r => [norm(r[1]), norm(r[0])]);
  const toks = norm(q).split(/\s+/).filter(Boolean); if (!toks.length) return [];
  const out = [];
  for (let i = 0; i < custIndex.length; i++){
    if (toks.every(t => custIndex[i][0].includes(t) || custIndex[i][1].includes(t))){ out.push(i); if (out.length >= 30) break; }
  }
  return out;
}
/* Shows only the first n rows of a list and scrolls the rest. */
function capList(el, n){
  el.style.maxHeight = ""; el.style.overflowY = "";
  const k = [...el.children];
  if (k.length <= n || !el.offsetHeight) return;
  let h = 0; for (let i = 0; i < n; i++) h += k[i].offsetHeight;
  const cs = getComputedStyle(el), gap = parseFloat(cs.rowGap) || 0, pad = (parseFloat(cs.paddingTop) || 0) + (parseFloat(cs.paddingBottom) || 0);
  el.style.maxHeight = Math.ceil(h + gap * (n - 1) + pad + 2) + "px"; el.style.overflowY = "auto";
}
function showCustMatches(){
  const box = $("custOut"); if (!box) return;
  const v = $("cust").value.trim();
  const list = v && CUSTOMERS.length ? searchCustomers(v) : [];
  box.innerHTML = list.map(i => { const r = CUSTOMERS[i];
    return `<button type="button" class="hit" data-c="${i}"><b dir="auto">${esc(r[1])}</b><span>${esc([r[0], r[5], r[6]].filter(Boolean).join(" · "))}</span></button>`; }).join("");
  box.hidden = !list.length;
  capList(box, 2);
}
function applyCustomer(r){
  const q = state.quote;
  q.customer = r[1]; q.custNo = r[0]; q.custVat = r[2]; q.custAddress = r[3]; q.custCommReg = r[4]; q.custCity = r[5]; q.custPhone = r[6];
  $("cust").value = q.customer; syncExtra();
  const box = $("custOut"); if (box){ box.hidden = true; box.innerHTML = ""; }
  queueAutosave();
  toast(r[1] + ": تم تعبئة بيانات العميل\nCustomer details filled in");
}
if ($("custOut")) $("custOut").addEventListener("click", e => { const b = e.target.closest("[data-c]"); if (b && CUSTOMERS[+b.dataset.c]) applyCustomer(CUSTOMERS[+b.dataset.c]); });
document.addEventListener("click", e => { const box = $("custOut"); if (box && !e.target.closest(".custWrap")) box.hidden = true; });
/* Enter in the customer name box picks the top suggestion (the first one in the list). */
$("cust").addEventListener("keydown", e => {
  if (e.key !== "Enter") return;
  const box = $("custOut"), first = box && !box.hidden && box.querySelector("[data-c]");
  if (first && CUSTOMERS[+first.dataset.c]){ e.preventDefault(); applyCustomer(CUSTOMERS[+first.dataset.c]); }
});

/* Columns are found by their Arabic heading, in any order. Each field accepts the headings below (first found wins).
   The address is taken from "العنوان", or built from المنطقة - الحي - رقم المبنى الشارع when there is no such column. */
const CUST_COLS = {
  no:["رقم العميل", "رقم وحدة العمل"], name:["اسم العميل", "الاسم", "اسم الحساب"], name2:["وحدة العمل"],
  vat:["الرقم الضريبي"], addr:["العنوان"], reg:["السجل التجاري", "رقم السجل"], country:["البلد", "الدولة"], phone:["الجوال", "رقم الجوال"],
  region:["المنطقة"], district:["الحي"], street:["الشارع"], bldg:["رقم المبنى"]
};
async function readCustomers(file){
  if (!(await loadXlsx())) throw new Error("تعذر تحميل قارئ إكسل. تحقق من الإنترنت وحاول مرة أخرى. / The Excel reader did not load.");
  const wb = XLSX.read(await file.arrayBuffer(), { type:"array" });
  const rows = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { header:1, raw:false, defval:"" });
  const clean = s => norm(String(s).replace(/\s+/g, " ").trim());
  const hi = rows.findIndex(r => r.some(c => CUST_COLS.name.some(a => clean(c) === clean(a))));
  if (hi < 0) throw new Error('Could not find a customer name column (اسم العميل / اسم الحساب) in the first sheet.');
  const head = rows[hi], col = {};
  Object.keys(CUST_COLS).forEach(k => {
    let idx = -1;
    for (const a of CUST_COLS[k]){ idx = head.findIndex(c => clean(c) === clean(a)); if (idx >= 0) break; }
    if (idx >= 0) col[k] = idx;
  });
  const get = (r, k) => k in col ? String(r[col[k]] == null ? "" : r[col[k]]).trim() : "";
  const missing = [];
  [["no", "رقم العميل"], ["vat", "الرقم الضريبي"], ["reg", "السجل التجاري"], ["country", "البلد"], ["phone", "الجوال"]].forEach(([k, label]) => { if (!(k in col)) missing.push(label); });
  if (!("addr" in col) && !["region", "district", "street", "bldg"].some(k => k in col)) missing.push("العنوان");
  const address = r => get(r, "addr") || [get(r, "region"), get(r, "district"), (get(r, "bldg") + " " + get(r, "street")).trim()].filter(Boolean).join(" - ");
  const seen = {}, out = []; let skipped = 0;
  rows.slice(hi + 1).forEach(r => {
    const name = get(r, "name") || get(r, "name2");
    if (!name) { skipped++; return; }
    const row = [get(r, "no"), name, get(r, "vat"), address(r), get(r, "reg"), get(r, "country"), get(r, "phone")];
    const key = row.join("|"); if (seen[key]) return; seen[key] = 1;
    out.push(row);
  });
  if (!out.length) throw new Error("لا يوجد عملاء في هذا الملف. / No customers found in that file.");
  return { rows:out, count:out.length, skipped, missing };
}
async function pushCustomers(rows){
  const res = await fetch(CONFIG.SCRIPT_URL, { method:"POST", headers:{"Content-Type":"text/plain;charset=utf-8"}, body:JSON.stringify({ key:CONFIG.SECRET, putCustomers:true, customers:rows }) });
  const out = await res.json();
  if (!out.ok) throw new Error(out.error === "Bad quote id" ? "Your Apps Script is out of date. Paste the new Code.gs and deploy a New version." : out.error);
  return out.updatedAt;
}
async function pullCustomers(){
  if (!CONFIG.SCRIPT_URL) return;
  try{
    const res = await fetch(CONFIG.SCRIPT_URL, { method:"POST", headers:{"Content-Type":"text/plain;charset=utf-8"}, body:JSON.stringify({ key:CONFIG.SECRET, getCustomers:true, since:custStamp }) });
    const out = await res.json();
    if (!out.ok || out.unchanged || !out.customers || !out.customers.length || out.updatedAt <= custStamp) return;
    setCustomers(out.customers, out.updatedAt); saveCustLocal(); showCustInfo("from the Google Sheet");
  }catch(e){}
}
function on(id, ev, fn){ const el = $(id); if (el) el.addEventListener(ev, fn); }
on("custFile", "change", async e => {
  const f = e.target.files[0]; e.target.value = ""; if (!f) return;
  $("custApply").hidden = $("custCancel").hidden = true; custPending = null;
  $("custMsg").textContent = "جارٍ قراءة " + f.name + "… / Reading…";
  try{
    custPending = await readCustomers(f);
    const p = custPending;
    $("custMsg").textContent = p.count + " customers found" + (p.missing.length ? " (no column found for: " + p.missing.join("، ") + ")" : "") + (p.addrFromParts ? " (address built from building, street, district and region)" : "") + (p.skipped ? "; " + p.skipped + " rows without a name skipped" : "") + ". This replaces the current list. Apply?";
    $("custApply").hidden = $("custCancel").hidden = false;
  }catch(err){ $("custMsg").textContent = err.message || "تعذرت قراءة هذا الملف. / Could not read that file."; }
});
on("custCancel", "click", () => { custPending = null; $("custApply").hidden = $("custCancel").hidden = true; $("custMsg").textContent = "تم الإلغاء. لم تتغير قائمة العملاء. / Cancelled. The customer list was not changed."; });
on("custApply", "click", async () => {
  if (!custPending) return;
  const rows = custPending.rows; custPending = null; $("custApply").hidden = $("custCancel").hidden = true;
  setCustomers(rows, Date.now()); saveCustLocal(); showCustInfo();
  if (!CONFIG.SCRIPT_URL){ $("custMsg").textContent = "تم تحديث قائمة العملاء على هذا الجهاز. اربط Google Sheet لمشاركتها. / Customer list updated on this tablet. Connect the Google Sheet to share it."; return; }
  $("custMsg").textContent = "تم التحديث هنا. جارٍ المشاركة مع الأجهزة الأخرى… / Updated here. Sharing with the other tablets…";
  try{
    const t = await pushCustomers(rows); custStamp = t; saveCustLocal(); showCustInfo();
    $("custMsg").textContent = "تم تحديث قائمة العملاء ومشاركتها. / Customer list updated and shared.";
  }catch(err){ $("custMsg").textContent = "تم التحديث على هذا الجهاز فقط. تعذرت المشاركة: " + (err.message || err); }
});

/* Scanner focus: a barcode scanner "types" into the search box, so keep the cursor there.
   - after an item is added the cursor goes back to the box;
   - if the cursor is on a button or on empty space and a character arrives (the scanner), it is sent to the box;
   - typing in other boxes (qty, discount, customer ...) is left alone. */
if (IS_MAIN){
  const sBox = $("searchIn");
  const focusScan = () => { try{ sBox.focus({ preventScroll:true }); }catch(e){ sBox.focus(); } };
  $("pasteForm").addEventListener("submit", () => setTimeout(focusScan, 0));
  $("searchOut").addEventListener("click", () => setTimeout(focusScan, 0));
  document.addEventListener("keydown", e => {
    if (e.ctrlKey || e.altKey || e.metaKey || e.key.length !== 1) return;
    const a = document.activeElement, tag = a ? a.tagName : "";
    if (tag === "INPUT" || tag === "TEXTAREA" || tag === "SELECT" || (a && a.isContentEditable)) return;
    focusScan();                        // the character is then typed into the search box
  });
  window.addEventListener("load", focusScan);
}
if (IS_MAIN){ restore(); newQuote(); }
showItemsInfo(); pullItems(); showCustInfo(); pullCustomers();
if (IS_MAIN){ armed = true; renderPreview(); }

/* ---------- link to the Sheet ---------- */
const validSheetUrl = u => /^https:\/\/docs\.google\.com\/spreadsheets\/d\/[\w-]+/.test(u || "");
function updateSheetLink(){
  const on = !!CONFIG.SCRIPT_URL || validSheetUrl(CONFIG.SHEET_URL);
  document.querySelectorAll(".openQuote").forEach(b => { b.disabled = !on; b.title = on ? "" : "اربط Google Sheet أولاً / Connect a Google Sheet first"; });
}
/* Open the Google Sheet tab of the quotation that is open right now (saving it first if it has no tab yet). */
async function openQuotationInSheet(){
  if (!IS_MAIN){   // settings page: the open quotation lives on the main page, so use the one saved last
    restore();
    const ids = Object.keys(state.saved).sort((x, y) => ((state.saved[y].quote || {}).savedAt || 0) - ((state.saved[x].quote || {}).savedAt || 0));
    const last = ids.length && state.saved[ids[0]];
    if (last && validSheetUrl(last.tabUrl)) { window.open(last.tabUrl, "_blank", "noopener"); return; }
    if (validSheetUrl(CONFIG.SHEET_URL)) { window.open(CONFIG.SHEET_URL, "_blank", "noopener"); return; }
    return toast("افتح صفحة العروض وأضف صنفاً أولاً ليتم إنشاء ورقة العرض.\nOpen the quotations page and add an item first.");
  }
  const q = state.quote, id = q && q.id;
  const hasTab = () => state.saved[id] && validSheetUrl(state.saved[id].tabUrl);
  if (hasTab()){ window.open(state.saved[id].tabUrl, "_blank", "noopener"); return; }
  if (!CONFIG.SCRIPT_URL) return toast("اربط Google Sheet أولاً (قسم ربط Google Sheet).\nConnect a Google Sheet first.");
  if (!state.lines.length) return toast("أضف صنفاً واحداً على الأقل ليتم إنشاء ورقة العرض.\nAdd at least one item first.");
  const w = window.open("", "_blank");   // opened now, inside the click, so the browser allows it
  try{ if (w) w.document.title = "جارٍ فتح " + id + "… / Opening…"; }catch(e){}
  setSync("جارٍ إرسال " + id + " إلى Google Sheets… / Sending to Google Sheets…");
  await save({ auto:true });
  await pushChain;
  if (hasTab()){
    if (w) w.location.href = state.saved[id].tabUrl; else window.open(state.saved[id].tabUrl, "_blank", "noopener");
  } else {
    if (w) w.close();
    toast("تعذر فتح الورقة. راجع الرسالة تحت الأزرار.\nCould not open the tab.");
  }
}
document.querySelectorAll(".openQuote").forEach(b => b.addEventListener("click", openQuotationInSheet));
/* The script tells us the Sheet's address in its answers; remember it so staff can open the Sheet. */
function noteSheetUrl(out){
  if (out && validSheetUrl(out.sheetUrl) && out.sheetUrl !== CONFIG.SHEET_URL){
    CONFIG.SHEET_URL = out.sheetUrl;
    try{ localStorage.setItem("gqs-config", JSON.stringify(CONFIG)); }catch(e){}
    $("cfgSheet").value = CONFIG.SHEET_URL; updateSheetLink();
  }
}

/* ---------- connect to Google Sheets ---------- */
function showConfig(msg){
  $("cfgSheet").value = CONFIG.SHEET_URL; updateSheetLink();
  $("cfgUrl").value = CONFIG.SCRIPT_URL; $("cfgKey").value = CONFIG.SECRET;
  $("cfgState").textContent = msg || (CONFIG.SCRIPT_URL ? "محفوظ على هذا الجهاز. اضغط اختبار الاتصال. / Saved on this tablet. Press Test connection." : "غير متصل. العروض محفوظة على هذا الجهاز فقط. / Not connected. Quotations are kept on this tablet only.");
}
$("cfgSave").onclick = () => {
  CONFIG.SCRIPT_URL = $("cfgUrl").value.trim(); CONFIG.SECRET = $("cfgKey").value.trim();
  const typed = $("cfgSheet").value.trim(); CONFIG.SHEET_URL = validSheetUrl(typed) ? typed : "";
  try{ localStorage.setItem("gqs-config", JSON.stringify(CONFIG)); }catch(e){ return toast("تعذر حفظ الإعدادات في هذا المتصفح.\nThis browser would not save the settings."); }
  showConfig(); renderSaved(); toast("تم حفظ إعدادات الاتصال.\nConnection settings saved.");
};
$("cfgTest").onclick = async () => {
  if (!$("cfgUrl").value.trim()) return toast("الصق رابط تطبيق Apps Script أولاً.\nPaste the Apps Script web app URL first.");
  $("cfgSave").onclick();
  $("cfgState").textContent = "جارٍ الاختبار… / Testing…";
  try{
    const res = await fetch(CONFIG.SCRIPT_URL, { method:"POST", headers:{"Content-Type":"text/plain;charset=utf-8"},
      body:JSON.stringify({ key:CONFIG.SECRET, ping:true }) });
    const out = await res.json();
    noteSheetUrl(out);
    showConfig(!out.ok ? "رد الجدول: " + out.error + ". تحقق من الرمز السري. / Check the secret."
      : (out.features || []).includes("doc-tabs") ? "Connected. Saves will go to your Google Sheet. Script version " + out.version + "."
      : "Connected, but your Apps Script is out of date (it does not know the latest features). Paste the new Code.gs and deploy a New version.");
  }catch(err){ showConfig("تعذر الوصول إلى النص. تحقق من الرابط وأن الوصول \"أي شخص\". / Could not reach the script. Check the URL and access (Anyone)."); }
};
showConfig();
