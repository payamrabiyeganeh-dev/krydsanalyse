"use strict";
const $ = id => document.getElementById(id);
const LS = {
  get(k){ try { return JSON.parse(localStorage.getItem(k)); } catch(e){ return null; } },
  set(k,v){ try { localStorage.setItem(k, JSON.stringify(v)); } catch(e){} },
  del(k){ try { localStorage.removeItem(k); } catch(e){} }
};
let DATA = { rows: [], name: "", warnings: [] };

/* ---------- dates ---------- */
const pad = n => String(n).padStart(2,"0");
const iso = (y,m,d) => `${y}-${pad(m)}-${pad(d)}`;
const wday = s => { const [y,m,d] = s.split("-").map(Number); return new Date(Date.UTC(y,m-1,d)).getUTCDay(); };
const dkDate = s => { const [y,m,d] = s.split("-"); return `${d}.${m}.${y}`; };
const DAYS = ["søndag","mandag","tirsdag","onsdag","torsdag","fredag","lørdag"];
function addMonths(s, n){
  const [y,m,d] = s.split("-").map(Number);
  const t = (y*12 + (m-1)) + n, ny = Math.floor(t/12), nm = t%12 + 1;
  const dim = new Date(Date.UTC(ny, nm, 0)).getUTCDate();
  return iso(ny, nm, Math.min(d, dim));
}
function toISO(v){
  if (v == null || v === "") return null;
  if (typeof v === "number" && window.XLSX){ const p = XLSX.SSF.parse_date_code(v); if (p && p.y) return iso(p.y,p.m,p.d); }
  if (v instanceof Date && !isNaN(v)) return iso(v.getFullYear(), v.getMonth()+1, v.getDate());
  const s = String(v).trim(); let m;
  if ((m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/))) return iso(+m[1],+m[2],+m[3]);
  if ((m = s.match(/^(\d{1,2})[.\/-](\d{1,2})[.\/-](\d{4})/))) return iso(+m[3],+m[2],+m[1]);
  return null;
}

/* ---------- data ---------- */
const compact = rows => rows.map(r => [r.date, ...r.main, ...r.star]);
const fromCompact = arr => arr.map(a => ({ date:a[0], main:a.slice(1,6), star:a.slice(6,8) }));
function maxStar(date){ return date <= "2014-10-03" ? 8 : date <= "2022-03-18" ? 10 : 12; }
function validate(rows){
  const w = [], seen = new Set();
  rows.forEach(r => {
    const wd = wday(r.date), k = dkDate(r.date), ms = maxStar(r.date);
    if (seen.has(r.date)) w.push(`${k}: dato forekommer flere gange`); seen.add(r.date);
    if (wd !== 2 && wd !== 5) w.push(`${k}: hverken tirsdag eller fredag`);
    if (wd === 2 && r.date < "2022-03-29") w.push(`${k}: tirsdag før 29.03.2022`);
    if (r.main.some(n => !(n>=1 && n<=50)) || new Set(r.main).size !== 5) w.push(`${k}: vindertal uden for 1–50 eller gentaget`);
    if (r.star.some(n => !(n>=1 && n<=ms)) || new Set(r.star).size !== 2) w.push(`${k}: stjernetal uden for 1–${ms} eller gentaget`);
  });
  return w;
}
function parseWorkbook(wb){
  const ws = wb.Sheets[wb.SheetNames[0]];
  const aoa = XLSX.utils.sheet_to_json(ws, { header:1, raw:true, blankrows:false });
  let h = aoa.findIndex(r => r.some(c => /dato/i.test(String(c ?? ""))));
  let cDate, cMain = [], cStar = [];
  if (h >= 0){
    const hdr = aoa[h].map(c => String(c ?? "").trim());
    cDate = hdr.findIndex(c => /dato/i.test(c));
    hdr.forEach((c,i) => { if (/^vinder/i.test(c)) cMain.push(i); else if (/^stjern/i.test(c)) cStar.push(i); });
  }
  if (h < 0 || cMain.length !== 5 || cStar.length !== 2){ h = -1; cDate = 1; cMain = [2,3,4,5,6]; cStar = [7,8]; }
  const rows = []; let bad = 0;
  aoa.slice(h+1).forEach(r => {
    const date = toISO(r[cDate]);
    const main = cMain.map(i => Number(r[i])), star = cStar.map(i => Number(r[i]));
    if (!date || main.some(isNaN) || star.some(isNaN)){ if (r.some(c => c !== null && c !== "" && c !== undefined)) bad++; return; }
    rows.push({ date, main, star });
  });
  return { rows, bad };
}

/* Sources, merged in order of trust: online data (updated automatically) → indlæst fil → data i appen → manuelle.
   A later source only fills dates the earlier ones lack; disagreements are reported. */
let ONLINE = (() => { const o = LS.get("ejk.online"); return o && Array.isArray(o.draws) ? o : null; })();
let FILE = (() => { const f = LS.get("ejk.data"); return f && Array.isArray(f.rows) && f.rows.length ? f : null; })();
let MANUAL = new Map((LS.get("ejk.manual") || []).map(a => [a[0], { date:a[0], main:a.slice(1,6), star:a.slice(6,8) }]));
function saveManual(){ LS.set("ejk.manual", compact([...MANUAL.values()])); }
function rebuild(){
  const byDate = new Map(), from = new Map(), conflicts = new Set();
  const layer = (rows, tag) => rows.forEach(r => {
    const b = byDate.get(r.date);
    if (!b){ byDate.set(r.date, r); from.set(r.date, tag); }
    else if (b.main.join() !== r.main.join() || b.star.join() !== r.star.join()) conflicts.add(`${dkDate(r.date)} (${tag} vs. ${from.get(r.date)})`);
  });
  if (ONLINE) layer(fromCompact(ONLINE.draws), "online");
  if (FILE) layer(fromCompact(FILE.rows), "fil");
  layer(fromCompact(BUNDLED), "app");
  const nOnline = ONLINE ? ONLINE.draws.length : 0;
  let manual = 0;
  for (const [d, r] of MANUAL){ if (!byDate.has(d)){ byDate.set(d, r); from.set(d, "manuel"); manual++; } else { const b = byDate.get(d); if (b.main.join() !== r.main.join() || b.star.join() !== r.star.join()) conflicts.add(`${dkDate(d)} (manuel vs. ${from.get(d)})`); } }
  const extra = (FILE && FILE.extra) ? FILE.extra.slice() : [];
  if (conflicts.size) extra.push(`tallene afviger mellem kilderne for ${[...conflicts].join(", ")}; den første kilde bruges`);
  const parts = [nOnline ? "Online data" : "Data i appen"];
  if (FILE) parts.push(`fil: ${FILE.name}`);
  if (manual) parts.push(`${manual} manuel${manual>1?"le":""}`);
  applyRows([...byDate.values()], parts.join(" + "), extra);
  renderManual(); syncLine();
}
function inData(date){ return (ONLINE && ONLINE.draws.find(r => r[0] === date)) || (FILE && FILE.rows.find(r => r[0] === date)) || BUNDLED.find(r => r[0] === date); }
function applyRows(rows, name, extra){
  rows.sort((a,b) => a.date < b.date ? 1 : -1);
  const prevTop = DATA.rows.length ? DATA.rows[0].date : null;
  DATA = { rows, name, warnings: extra.concat(validate(rows)) };
  const sel = $("cutoff"), prev = sel.value;
  sel.innerHTML = rows.map(r => `<option value="${r.date}">${DAYS[wday(r.date)]} ${dkDate(r.date)}</option>`).join("");
  const saved = LS.get("ejk.settings");
  // keep a chosen older date; follow new draws when the latest one was selected
  const want = (prev && prev !== prevTop) ? prev : (!prev && saved && saved.cutoff && saved.cutoff !== saved.latest ? saved.cutoff : null);
  if (want && rows.some(r => r.date === want)) sel.value = want;
  if (rows.length){
    const first = rows[rows.length-1], last = rows[0];
    $("dataStatus").innerHTML = `<strong>${esc(name)}</strong><br>${rows.length} trækninger · ${dkDate(first.date)} – ${dkDate(last.date)}`;
  } else $("dataStatus").textContent = "Ingen trækninger indlæst.";
  const w = DATA.warnings;
  $("dataWarn").innerHTML = w.length ? `<div class="warn">${w.length} advarsel${w.length>1?"er":""} i data: ${esc(w.slice(0,4).join("; "))}${w.length>4?" …":""}</div>` : "";
  run();
}
const esc = s => String(s).replace(/[&<>"]/g, c => ({ "&":"&amp;", "<":"&lt;", ">":"&gt;", '"':"&quot;" }[c]));

$("file").addEventListener("change", e => {
  const f = e.target.files[0]; if (!f) return;
  if (!window.XLSX){ $("dataWarn").innerHTML = `<div class="warn">Excel-læseren er ikke klar endnu. Prøv igen om et øjeblik.</div>`; return; }
  const rd = new FileReader();
  rd.onload = () => {
    try {
      const wb = XLSX.read(new Uint8Array(rd.result), { type:"array" });
      const { rows, bad } = parseWorkbook(wb);
      if (!rows.length){ $("dataWarn").innerHTML = `<div class="warn">Filen indeholder ingen læsbare trækninger. Den skal have kolonnerne Dato, Vindertal_01–05 og Stjernertal_01–02.</div>`; return; }
      const extra = bad ? [`${bad} række${bad>1?"r":""} kunne ikke læses og er sprunget over`] : [];
      FILE = { name:f.name, rows:compact(rows), extra };
      LS.set("ejk.data", FILE);
      // manual draws that the file now contains are no longer needed
      rows.forEach(r => MANUAL.delete(r.date)); saveManual();
      rebuild();
      $("dataWarn").insertAdjacentHTML("afterbegin", `<div class="okmsg">${esc(f.name)} er indlæst med ${rows.length} trækninger.</div>`);
    } catch(err){ $("dataWarn").innerHTML = `<div class="warn">Filen kunne ikke læses (${esc(err.message)}). Gem den som .xlsx og prøv igen.</div>`; }
    e.target.value = "";
  };
  rd.readAsArrayBuffer(f);
});
$("reset").addEventListener("click", () => { LS.del("ejk.data"); FILE = null; rebuild(); });

/* ---------- automatic update: data.json is refreshed daily on the server ---------- */
let lastCheck = 0, checking = false, lastResult = "";
function fmtStamp(s){ try { return new Date(s).toLocaleString("da-DK", { day:"2-digit", month:"2-digit", year:"numeric", hour:"2-digit", minute:"2-digit" }); } catch(e){ return s; } }
function syncLine(){
  const el = $("syncStatus"); if (!el) return;
  const top = DATA.rows.length ? DATA.rows[0].date : null;
  const parts = [];
  if (top) parts.push(`Seneste trækning: ${DAYS[wday(top)]} ${dkDate(top)}.`);
  if (ONLINE && ONLINE.updated) parts.push(`Serveren tjekkede for nye trækninger ${fmtStamp(ONLINE.updated)}.`);
  if (lastResult) parts.push(lastResult);
  el.textContent = parts.join(" ");
  $("barSub").dataset.latest = top || "";
}
async function checkOnline(force){
  if (checking || location.protocol === "file:") return;
  if (!force && Date.now() - lastCheck < 10*60*1000) return;
  checking = true; lastCheck = Date.now(); $("checkNow").disabled = true; $("checkNow").textContent = "Søger …";
  try {
    const r = await fetch(`data.json?t=${Date.now()}`, { cache:"no-store" });
    if (!r.ok) throw new Error(`HTTP ${r.status}`);
    const doc = await r.json();
    if (!doc || !Array.isArray(doc.draws) || !doc.draws.length) throw new Error("tom fil");
    const before = DATA.rows.length ? DATA.rows[0].date : "";
    const prevCount = ONLINE ? ONLINE.draws.length : 0;
    ONLINE = { updated: doc.updated, draws: doc.draws };
    LS.set("ejk.online", ONLINE);
    // manual draws now covered by online data are no longer needed
    doc.draws.forEach(a => { const m = MANUAL.get(a[0]); if (m && m.main.join() === a.slice(1,6).join() && m.star.join() === a.slice(6,8).join()) MANUAL.delete(a[0]); });
    saveManual();
    rebuild();
    const after = DATA.rows.length ? DATA.rows[0].date : "";
    lastResult = navigator.onLine === false ? "Ingen internetforbindelse. Appen bruger de gemte data." : after > before ? `Ny trækning hentet: ${dkDate(after)}.` : (force ? "Ingen nye trækninger lige nu." : "");
  } catch(e){
    lastResult = navigator.onLine === false ? "Ingen internetforbindelse. Appen bruger de gemte data." : "Kunne ikke søge efter nye trækninger lige nu. Appen bruger de gemte data.";
  } finally {
    checking = false; $("checkNow").disabled = false; $("checkNow").textContent = "Søg efter nye trækninger"; syncLine();
  }
}
$("checkNow").addEventListener("click", () => checkOnline(true));
document.addEventListener("visibilitychange", () => { if (document.visibilityState === "visible") checkOnline(false); });
window.addEventListener("online", () => checkOnline(true));

/* ---------- manual entry ---------- */
function renderManual(){
  const items = [...MANUAL.values()].sort((a,b) => a.date < b.date ? 1 : -1);
  $("manualList").innerHTML = items.length ? items.map(r => `<div class="li"><span>${DAYS[wday(r.date)]} ${dkDate(r.date)} · ${r.main.join(" ")} + ${r.star.join(" ")}</span><button class="btn ghost small" type="button" data-del="${r.date}">Fjern</button></div>`).join("") : "";
}
$("manualList").addEventListener("click", e => {
  const d = e.target.getAttribute("data-del"); if (!d) return;
  MANUAL.delete(d); saveManual(); rebuild();
  $("eMsg").innerHTML = `<div class="okmsg">Trækningen ${dkDate(d)} er fjernet.</div>`;
});
$("eAdd").addEventListener("click", () => {
  const date = $("eDate").value;
  const main = [...$("eMain").querySelectorAll("input")].map(i => parseInt(i.value));
  const star = [...$("eStar").querySelectorAll("input")].map(i => parseInt(i.value));
  const msg = t => $("eMsg").innerHTML = `<div class="warn">${t}</div>`;
  if (!date) return msg("Vælg datoen for trækningen.");
  const wd = wday(date);
  if (wd !== 2 && wd !== 5) return msg(`${dkDate(date)} er en ${DAYS[wd]}. Eurojackpot trækkes tirsdag og fredag.`);
  if (main.some(n => !(n>=1 && n<=50)) || new Set(main).size !== 5) return msg("Skriv 5 forskellige vindertal mellem 1 og 50.");
  const ms = maxStar(date);
  if (star.some(n => !(n>=1 && n<=ms)) || new Set(star).size !== 2) return msg(`Skriv 2 forskellige stjernetal mellem 1 og ${ms}.`);
  const r = { date, main: main.sort((a,b)=>a-b), star: star.sort((a,b)=>a-b) };
  const inBase = inData(date);
  if (inBase) return msg(`${dkDate(date)} findes allerede i dine data (${inBase.slice(1,6).join(" ")} + ${inBase.slice(6,8).join(" ")}).`);
  MANUAL.set(date, r); saveManual();
  [...$("eMain").querySelectorAll("input"), ...$("eStar").querySelectorAll("input")].forEach(i => i.value = "");
  rebuild();
  $("eMsg").innerHTML = `<div class="okmsg">${DAYS[wd][0].toUpperCase()+DAYS[wd].slice(1)} ${dkDate(date)} er tilføjet.</div>`;
});

/* ---------- analysis ---------- */
function period(cutoff, months, day){
  const start = addMonths(cutoff, -months);
  const rows = DATA.rows.filter(r => r.date > start && r.date <= cutoff && (day === "all" || wday(r.date) === +day));
  return { start, cutoff, rows, months, day };
}
function counts(rows, key, N){ const c = Array(N+1).fill(0); rows.forEach(r => r[key].forEach(n => { if (n>=1 && n<=N) c[n]++; })); return c; }
function ranks(c, N, dir){
  const nums = Array.from({length:N}, (_,i) => i+1);
  nums.sort((a,b) => dir === "hi" ? c[b]-c[a] : c[a]-c[b]);
  const avg = Array(N+1), min = Array(N+1);
  for (let i = 0; i < N; ){
    let j = i; while (j+1 < N && c[nums[j+1]] === c[nums[i]]) j++;
    for (let k = i; k <= j; k++){ avg[nums[k]] = (i+j)/2 + 1; min[nums[k]] = i + 1; }
    i = j + 1;
  }
  return { avg, min };
}
function combine(pA, pB, key, N, rA, rB, method, topN){
  const cA = counts(pA.rows, key, N), cB = counts(pB.rows, key, N);
  const RA = ranks(cA, N, rA), RB = ranks(cB, N, rB);
  const items = [];
  for (let n = 1; n <= N; n++) items.push({ n, a:cA[n], b:cB[n], ra:RA.avg[n], rb:RB.avg[n], score:RA.avg[n]+RB.avg[n],
    ok: method === "sum" ? true : (RA.min[n] <= topN && RB.min[n] <= topN) });
  const dv = (x, r) => r === "hi" ? -x : x;
  items.sort((x,y) => x.score - y.score || dv(x.a,rA) - dv(y.a,rA) || dv(x.b,rB) - dv(y.b,rB) || x.n - y.n);
  items.forEach((it,i) => it.pos = i+1);
  const per = key === "main" ? 5 : 2;
  return { items, expA: pA.rows.length * per / N, expB: pB.rows.length * per / N };
}
function makeTickets(main, star, nRows){
  const m = main.items.filter(i => i.ok), s = star.items.filter(i => i.ok);
  const maxM = Math.floor(m.length/5), maxS = Math.floor(s.length/2);
  const out = [], notes = [];
  if (maxM < 1 || maxS < 1){
    if (maxM < 1) notes.push(`Kun ${m.length} vindertal opfylder begge regler. Der skal bruges mindst 5. Hæv N vindertal.`);
    if (maxS < 1) notes.push(`Kun ${s.length} stjernetal opfylder begge regler. Der skal bruges mindst 2. Hæv N stjernetal.`);
    return { out, notes };
  }
  const k = Math.min(nRows, maxM);
  if (k < nRows) notes.push(`Kun ${maxM} række${maxM>1?"r":""} kan dannes uden at gentage vindertal.`);
  if (k > maxS) notes.push(`Der er kun ${maxS} par stjernetal uden gentagelse, så stjernetallene gentages fra række ${maxS+1}.`);
  for (let i = 0; i < k; i++){
    out.push({ main: m.slice(i*5, i*5+5).map(x=>x.n).sort((a,b)=>a-b), star: s.slice((i%maxS)*2, (i%maxS)*2+2).map(x=>x.n).sort((a,b)=>a-b) });
  }
  return { out, notes };
}

/* ---------- render ---------- */
const fmt1 = x => x.toLocaleString("da-DK", { maximumFractionDigits:1, minimumFractionDigits:1 });
const dayTxt = d => d === "all" ? "alle trækninger" : d === "2" ? "kun tirsdage" : "kun fredage";
const ruleTxt = r => r === "hi" ? "hyppigste" : "mindst hyppige";
const clampInt = (v, lo, hi) => Math.max(lo, Math.min(hi, parseInt(v) || lo));
const radio = n => (document.querySelector(`input[name=${n}]:checked`) || {}).value;
function readSettings(){
  return {
    cutoff: $("cutoff").value,
    mA: clampInt($("mA").value, 1, 240), dA: $("dA").value, rA: radio("rA"),
    mB: clampInt($("mB").value, 1, 240), dB: $("dB").value, rB: radio("rB"),
    meth: radio("meth"), nMain: clampInt($("nMain").value, 5, 50), nStar: clampInt($("nStar").value, 2, 12),
    nRows: clampInt($("nRows").value, 1, 10)
  };
}
function applySettings(s){
  ["mA","mB","nMain","nStar","nRows","dA","dB"].forEach(k => { if (s[k] != null) $(k).value = s[k]; });
  [["rA",s.rA],["rB",s.rB],["meth",s.meth]].forEach(([n,v]) => { const el = document.querySelector(`input[name=${n}][value="${v}"]`); if (el) el.checked = true; });
}
function table(res, used, s){
  const hA = `A <span style="font-weight:400">(${s.rA==="hi"?"hyp.":"min."})</span>`, hB = `B <span style="font-weight:400">(${s.rB==="hi"?"hyp.":"min."})</span>`;
  return `<table><thead><tr><th>Plads</th><th>Tal</th><th>${hA}</th><th>${hB}</th><th>Pl. A</th><th>Pl. B</th><th>Sum</th></tr></thead><tbody>` +
    res.items.map(it => `<tr class="${used.has(it.n)?"used":""} ${it.ok?"":"out"}"><td>${it.pos}</td><td class="num">${it.n}</td><td>${it.a}</td><td>${it.b}</td><td>${fmt1(it.ra)}</td><td>${fmt1(it.rb)}</td><td>${fmt1(it.score)}</td></tr>`).join("") +
    `</tbody></table>`;
}
function run(){
  const s = readSettings();
  $("topBox").hidden = s.meth !== "top";
  $("methHint").textContent = s.meth === "sum"
    ? "Alle tal rangeres efter summen af deres placering i A og B. Række 1 får de 5 bedste vindertal og de 2 bedste stjernetal, række 2 de næste osv."
    : "Kun tal, der ligger i top N i både A og B, kan komme med. Inden for dem afgør den samlede placering rækkefølgen.";
  LS.set("ejk.settings", { ...s, latest: DATA.rows.length ? DATA.rows[0].date : null });
  if (!DATA.rows.length || !s.cutoff){ $("tickets").innerHTML = `<div class="empty">Indlæs data under fanen Data.</div>`; return; }
  $("barSub").textContent = `til ${dkDate(s.cutoff)}`;
  const pA = period(s.cutoff, s.mA, s.dA), pB = period(s.cutoff, s.mB, s.dB);
  $("sumA").textContent = `${s.mA} mdr · ${ruleTxt(s.rA)}`; $("sumB").textContent = `${s.mB} mdr · ${ruleTxt(s.rB)}`;
  $("ruleSummary").textContent = `A ${s.mA} mdr ${ruleTxt(s.rA)} × B ${s.mB} mdr ${ruleTxt(s.rB)}`;
  const main = combine(pA, pB, "main", 50, s.rA, s.rB, s.meth, s.nMain);
  const star = combine(pA, pB, "star", 12, s.rA, s.rB, s.meth, s.nStar);
  const pc = (L, p, r) => {
    const f = p.rows.length ? p.rows[p.rows.length-1].date : null, l = p.rows.length ? p.rows[0].date : null;
    return `<div class="pcard"><div class="t">Periode ${L} · ${ruleTxt(r)} · ${dayTxt(p.day)}</div>
      <div class="m">Efter ${dkDate(p.start)} til og med ${dkDate(p.cutoff)}</div>
      <div class="m">${p.rows.length} trækninger${f ? ` · ${dkDate(f)} – ${dkDate(l)}` : ""}</div>
      <div class="m">Forventet pr. tal: ${fmt1(L==="A"?main.expA:main.expB)} (vindertal) · ${fmt1(L==="A"?star.expA:star.expB)} (stjernetal)</div></div>`;
  };
  $("periods").innerHTML = pc("A", pA, s.rA) + pc("B", pB, s.rB);
  const pw = [];
  if (!pA.rows.length || !pB.rows.length) pw.push("En af perioderne indeholder ingen trækninger. Forlæng perioden eller vælg en senere dato.");
  if ([pA, pB].some(p => p.rows.some(r => r.date < "2022-03-25"))) pw.push("En periode går længere tilbage end 25.03.2022, hvor stjernetal 11 og 12 ikke fandtes (og 9–10 ikke før 10.10.2014). De gamle trækninger trækker disse stjernetal ned.");
  $("periodWarn").innerHTML = pw.map(x => `<div class="warn">${x}</div>`).join("");
  const { out, notes } = makeTickets(main, star, s.nRows);
  const usedM = new Set(out.flatMap(t => t.main)), usedS = new Set(out.flatMap(t => t.star));
  $("tickets").innerHTML = out.length ? out.map((t,i) => `<div class="ticket"><span class="lbl">Række ${i+1}</span>
      <div class="balls">${t.main.map(n => `<span class="ball">${n}</span>`).join("")}</div><span class="sep" aria-hidden="true"></span>
      <div class="balls">${t.star.map(n => `<span class="ball s"><i>${n}</i></span>`).join("")}</div></div>`).join("")
    : `<div class="empty">Ingen rækker kan dannes med de valgte regler.</div>`;
  $("ticketNote").innerHTML = notes.map(x => `<div class="warn">${x}</div>`).join("");
  $("tMain").innerHTML = table(main, usedM, s);
  $("tStar").innerHTML = table(star, usedS, s);
}

/* ---------- tabs ---------- */
document.querySelectorAll(".tabbar button").forEach(b => b.addEventListener("click", () => {
  document.querySelectorAll(".tabbar button").forEach(x => x.setAttribute("aria-selected", x === b ? "true" : "false"));
  document.querySelectorAll(".view").forEach(v => v.classList.toggle("on", v.id === b.dataset.v));
  window.scrollTo(0, 0);
}));

/* ---------- install help ---------- */
const standalone = window.matchMedia("(display-mode: standalone)").matches || window.navigator.standalone === true;
const isIOS = /iphone|ipad|ipod/i.test(navigator.userAgent) || (navigator.platform === "MacIntel" && navigator.maxTouchPoints > 1);
let deferredPrompt = null;
if (!standalone){
  if (isIOS){ $("installCard").classList.add("show"); $("installText").textContent = "Åbn siden i Safari, tryk på Del-knappen (firkant med pil op) og vælg 'Føj til hjemmeskærm'."; }
  window.addEventListener("beforeinstallprompt", e => {
    e.preventDefault(); deferredPrompt = e;
    $("installCard").classList.add("show"); $("installText").textContent = "Installer appen, så den ligger på din hjemmeskærm og virker uden internet."; $("installBtn").hidden = false;
  });
  $("installBtn").addEventListener("click", async () => {
    if (!deferredPrompt) return; deferredPrompt.prompt();
    const r = await deferredPrompt.userChoice; deferredPrompt = null;
    if (r.outcome === "accepted") $("installCard").classList.remove("show");
  });
}
if ("serviceWorker" in navigator && location.protocol !== "file:") navigator.serviceWorker.register("sw.js").catch(() => {});

/* ---------- start ---------- */
applySettings(LS.get("ejk.settings") || { rA:"hi", rB:"lo", meth:"sum", dA:"5", dB:"all", mA:20, mB:10, nRows:3 });
if (!radio("rA")) document.querySelector('input[name=rA][value="hi"]').checked = true;
if (!radio("rB")) document.querySelector('input[name=rB][value="lo"]').checked = true;
if (!radio("meth")) document.querySelector('input[name=meth][value="sum"]').checked = true;
document.querySelectorAll("#v-settings input, #v-settings select").forEach(el => el.addEventListener(el.type === "radio" ? "change" : "input", run));
$("eDate").value = new Date().toISOString().slice(0,10);
rebuild();
checkOnline(true);
