"use strict";

const ARES = "https://ares.gov.cz/ekonomicke-subjekty-v-be/rest/ekonomicke-subjekty/";
const $ = (id) => document.getElementById(id);

// IČO: 8 číslic, vážený součet 8..2 prvních 7 číslic, kontrolní = (11 - sum % 11) % 10
function isValidIco(ico) {
  if (!/^\d{8}$/.test(ico)) return false;
  const d = [...ico].map(Number);
  const sum = d.slice(0, 7).reduce((s, n, i) => s + n * (8 - i), 0);
  return (11 - (sum % 11)) % 10 === d[7];
}

const normalize = (s) => s.replace(/\s+/g, "").padStart(8, "0");

// Vrací Map ico -> zdroj nálezu (seřazeno podle priority)
function extractIcos(text) {
  const found = new Map();
  const add = (raw, src) => {
    const ico = normalize(raw);
    if (isValidIco(ico) && !found.has(ico)) found.set(ico, src);
  };
  // 1) označené: "IČ: 27074358", "IČO 270 74 358", "IC:"
  for (const m of text.matchAll(/(?<![A-Za-z])I[ČC]O?(?![A-Za-z])\s*[:.]?\s*((?:\d\s?){5,7}\d)\b/gi)) add(m[1], "označeno jako IČ/IČO");
  // 2) DIČ právnické osoby: CZ + 8 číslic -> IČO
  for (const m of text.matchAll(/\b(?:DI[ČC]\s*[:.]?\s*)?CZ\s?(\d{8})\b/gi)) add(m[1], "z DIČ");
  // 3) holé osmimístné číslo, které projde kontrolním součtem
  for (const m of text.matchAll(/(?<![\d+\/-])(\d{8})(?![\d\/-])/g)) add(m[1], "osmimístné číslo (kontrolní součet OK)");
  return found;
}

async function lookup(ico) {
  const r = await fetch(ARES + ico, { headers: { Accept: "application/json" } });
  if (r.status === 404) throw new Error("Subjekt v ARES nenalezen");
  if (!r.ok) throw new Error("ARES vrátil HTTP " + r.status);
  return r.json();
}

const esc = (s) => String(s ?? "").replace(/[&<>"]/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]));
const row = (k, v) => (v ? `<dt>${k}</dt><dd><span class="copy" title="Kliknutím zkopírovat">${esc(v)}</span></dd>` : "");
const fmtDate = (d) => (d ? new Date(d).toLocaleDateString("cs-CZ") : "");

function renderCard(ico, src, data, err) {
  const el = document.createElement("section");
  el.className = "card" + (err ? " err" : "");
  if (err) {
    el.innerHTML = `<h3>IČO ${esc(ico)}</h3><div class="src">${esc(src)}</div><p>${esc(err.message)}</p>`;
  } else {
    const nace = (data.czNace || []).slice(0, 3).join(", ");
    el.innerHTML = `
      <h3>${esc(data.obchodniJmeno)}</h3>
      <div class="src">${esc(src)}</div>
      <dl>
        ${row("IČO", data.ico)}
        ${row("DIČ", data.dic)}
        ${row("Sídlo", data.sidlo?.textovaAdresa)}
        ${row("Vznik", fmtDate(data.datumVzniku))}
        ${row("Zánik", fmtDate(data.datumZaniku))}
        ${row("CZ-NACE", nace)}
      </dl>
      <div class="links">
        <a href="https://ares.gov.cz/ekonomicke-subjekty?ico=${ico}" target="_blank">ARES</a>
        <a href="https://or.justice.cz/ias/ui/rejstrik-$firma?ico=${ico}" target="_blank">Obchodní rejstřík</a>
      </div>`;
  }
  $("results").appendChild(el);
}

async function processIcos(found) {
  $("results").innerHTML = "";
  if (!found.size) { $("status").textContent = "V e-mailu jsem nenašel žádné platné IČO."; return; }
  $("status").textContent = `Nalezeno ${found.size}× IČO, dotazuji ARES…`;
  await Promise.all([...found].map(async ([ico, src]) => {
    try { renderCard(ico, src, await lookup(ico)); }
    catch (e) { renderCard(ico, src, null, e); }
  }));
  $("status").textContent = `Nalezeno ${found.size}× IČO.`;
}

function scanCurrentItem() {
  const item = Office.context.mailbox.item;
  if (!item) { $("status").textContent = "Není otevřena žádná zpráva."; return; }
  $("status").textContent = "Čtu tělo e-mailu…";
  item.body.getAsync(Office.CoercionType.Text, (res) => {
    if (res.status !== Office.AsyncResultStatus.Succeeded) {
      $("status").textContent = "Nepodařilo se načíst tělo: " + res.error.message;
      return;
    }
    processIcos(extractIcos((item.subject || "") + "\n" + res.value));
  });
}

document.addEventListener("click", (e) => {
  if (e.target.classList.contains("copy")) navigator.clipboard?.writeText(e.target.textContent);
});

$("manual").addEventListener("submit", (e) => {
  e.preventDefault();
  const ico = normalize($("manualIco").value.replace(/\D/g, ""));
  if (!isValidIco(ico)) { $("status").textContent = `${ico} není platné IČO (kontrolní součet).`; return; }
  processIcos(new Map([[ico, "zadáno ručně"]]));
});

Office.onReady(() => {
  scanCurrentItem();
  // připnutý panel: přepočítat při přepnutí na jinou zprávu
  Office.context.mailbox.addHandlerAsync?.(Office.EventType.ItemChanged, scanCurrentItem);
});
