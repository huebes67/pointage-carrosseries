/* Pointage Carrosseries — lit et modifie la feuille "Commande" du fichier Excel.
   Colonnes : A Date · B Faite ("F") · C Modèle · D Couleur · E Client / note · F Observations */
(function () {
  "use strict";

  // ---------- Réglages ----------
  const LS = {
    get(k, d) { try { const v = localStorage.getItem(k); return v === null ? d : JSON.parse(v); } catch { return d; } },
    set(k, v) { try { localStorage.setItem(k, JSON.stringify(v)); } catch { /* stockage plein ou bloqué */ } },
    del(k) { try { localStorage.removeItem(k); } catch { /* ignore */ } },
  };
  const base = window.APP_CONFIG || {};
  const cfg = () => ({
    scriptUrl: LS.get("pc_scriptUrl", "") || base.SCRIPT_URL || "",
    key: LS.get("pc_key", ""), // code d'accès : jamais dans le code publié, saisi sur chaque téléphone
    delai: Number(LS.get("pc_delai", base.DELAI_JOURS || 10)) || 10,
  });

  // ---------- Outils ----------
  const $ = (s, el = document) => el.querySelector(s);
  const $$ = (s, el = document) => [...el.querySelectorAll(s)];
  const esc = (s) => String(s ?? "").replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
  const clean = (s) => String(s ?? "").replace(/\s+/g, " ").trim();
  const fold = (s) => clean(s).toLowerCase().normalize("NFD").replace(/[̀-ͯ]/g, "");
  const DAY = 86400000;

  const serialToDate = (n) => { const d = new Date(Math.round((n - 25569) * DAY)); return new Date(d.getUTCFullYear(), d.getUTCMonth(), d.getUTCDate()); };
  const dateToSerial = (d) => Date.UTC(d.getFullYear(), d.getMonth(), d.getDate()) / DAY + 25569;
  const today = () => { const d = new Date(); return new Date(d.getFullYear(), d.getMonth(), d.getDate()); };
  const addDays = (d, n) => new Date(d.getFullYear(), d.getMonth(), d.getDate() + n);
  const diffDays = (a, b) => Math.round((b - a) / DAY);
  const fmt = (d, o = { day: "numeric", month: "short", year: "numeric" }) => d ? d.toLocaleDateString("fr-FR", o) : "—";
  const isoDate = (d) => `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-${String(d.getDate()).padStart(2, "0")}`;

  function parseCellDate(v) {
    if (typeof v === "number" && v > 20000 && v < 80000) return serialToDate(v);
    if (typeof v === "string") {
      const m = v.trim().match(/^(\d{1,2})[/.-](\d{1,2})[/.-](\d{2,4})$/);
      if (m) return new Date(+(m[3].length === 2 ? "20" + m[3] : m[3]), +m[2] - 1, +m[1]);
    }
    return null;
  }

  // ---------- Normalisation modèles / couleurs ----------
  const MODEL_ALIASES = [[/\bgranite\b/g, "granit"], [/\bporshe\b/g, "porsche"], [/\bkkpit\b/g, "kpit"], [/\bcoxmaxx\b/g, "cox maxx"], [/\bsc ?370\b/g, "sc 370"], [/\bstpro\b/g, "st pro"], [/\b205 ?gti\b/g, "205 gti"], [/\bbig ?mama\b/g, "big mama"], [/\bchargeur\b/g, "charger"]];
  function modelKey(s) {
    let k = fold(s);
    for (const [re, to] of MODEL_ALIASES) k = k.replace(re, to);
    return k;
  }

  const COLORS = {
    Noir: "#15161a", Blanc: "#f4f4f4", Bleu: "#2f7cf6", Jaune: "#f5c518", Rouge: "#e5383b", Vert: "#2fb36d",
  };
  const COLOR_WORDS = [
    ["Noir", /^(n|net n)$|\bnoir/], ["Blanc", /^b$|\bblanc/], ["Vert", /^v$|\bvert/], ["Bleu", /\bbleu|^beu$/],
    ["Jaune", /\bjaun/], ["Rouge", /\brouge/],
  ];
  function colorName(raw) {
    const f = fold(raw);
    if (!f) return "";
    for (const [name, re] of COLOR_WORDS) if (re.test(f)) return name;
    return clean(raw);
  }
  const swatch = (name, cls = "") => {
    const c = COLORS[name];
    const bg = c || (name ? "conic-gradient(#e5383b,#f5c518,#2fb36d,#2f7cf6,#e5383b)" : "transparent");
    return `<i class="swatch ${cls}" style="--sw:${bg}"></i>`;
  };

  // ---------- État ----------
  const state = {
    rows: [],          // commandes (lignes avec un modèle)
    lastUsedRow: 1,    // dernière ligne Excel utilisée (pour ajouter à la suite)
    matrix: null,      // valeurs brutes A:F
    source: null,      // "sheet" | "file"
    fileName: "",
    syncedAt: null,
    view: "home",
    hist: { q: "", status: "all", list: [], shown: 0 },
    topRange: "365",
  };

  function buildRows(matrix) {
    const rows = [];
    let carry = null, lastUsed = 1;
    // Colonne G "Payé" seulement si son titre est présent (sinon G contient d'autres tableaux)
    const hasPay = fold((matrix[0] || [])[6]) === "paye";
    for (let i = 1; i < matrix.length; i++) {
      const a = matrix[i] || [];
      const v = (j) => (a[j] === "" || a[j] === undefined ? null : a[j]);
      const own = parseCellDate(v(0));
      if (own) carry = own;
      if ([1, 2, 3, 4, 5].some((j) => v(j) !== null && clean(v(j)) !== "")) lastUsed = i + 1;
      const model = clean(v(2));
      if (!model) continue;
      rows.push({
        r: i + 1,
        ownDate: own,
        date: own || carry,
        doneRaw: v(1),
        done: clean(v(1)) !== "",
        model,
        key: modelKey(model),
        colorRaw: clean(v(3)),
        color: colorName(v(3)),
        client: clean(v(4)),
        obs: clean(v(5)),
        pay: hasPay ? payState(v(6)) : "",
      });
    }
    // Nom affiché = variante la plus fréquente de chaque modèle
    const variants = {};
    for (const r of rows) { (variants[r.key] ||= {}); variants[r.key][r.model] = (variants[r.key][r.model] || 0) + 1; }
    const display = {};
    for (const k in variants) {
      const best = Object.entries(variants[k]).sort((x, y) => y[1] - x[1])[0][0];
      display[k] = best.charAt(0).toUpperCase() + best.slice(1);
    }
    for (const r of rows) r.name = display[r.key];
    state.rows = rows;
    state.lastUsedRow = lastUsed;
    state.matrix = matrix;
  }

  // "" = non renseigné (anciennes commandes) · "paid" · "unpaid"
  function payState(raw) {
    const f = fold(raw);
    if (!f) return "";
    if (f.startsWith("non") || f === "n") return "unpaid";
    if (f.startsWith("pay") || f === "oui" || f === "p") return "paid";
    return "";
  }
  const PAY_VALUE = { paid: "Payé", unpaid: "Non payé" };
  const payPill = (r) => r.pay ? `<span class="pay ${r.pay}">€ ${r.pay === "paid" ? "Payée" : "Non payée"}</span>` : "";

  function urgency(r) {
    if (r.done) return null;
    if (!r.date) return { cls: "unknown", age: 0, left: null };
    const d = cfg().delai;
    const age = diffDays(r.date, today());
    const left = d - age;
    const cls = left < 0 ? "late" : left <= 1 ? "hot" : left <= 4 ? "warn" : "ok";
    return { cls, age, left, due: addDays(r.date, d) };
  }

  // ---------- Source : Google Sheets (via le script Google du tableau) ----------
  const api = {
    ready() { const c = cfg(); return !!(c.scriptUrl && c.key); },
    async call(body) {
      const c = cfg();
      let res;
      try {
        // text/plain : évite la requête de vérification que Google refuse
        res = await fetch(c.scriptUrl, { method: "POST", headers: { "Content-Type": "text/plain;charset=utf-8" }, body: JSON.stringify({ ...body, key: c.key }) });
      } catch {
        throw new Error("Pas de connexion internet ?");
      }
      if (!res.ok) throw new Error("Serveur Google indisponible (" + res.status + ")");
      let j;
      try { j = await res.json(); } catch { throw new Error("Réponse inattendue : vérifie l'adresse du script"); }
      if (!j.ok) { const err = new Error(j.error || "Erreur"); err.auth = /code/i.test(j.error || ""); throw err; }
      return j;
    },
    async load() { return (await this.call({ action: "read" })).values; },
  };

  // ---------- Écriture (Google Sheets ou fichier local en mémoire) ----------
  // values = [A..F] (date en numéro de série Excel) ; null = ne pas toucher la cellule
  // rowNum = null pour une nouvelle commande ; expect = modèle attendu sur la ligne (garde-fou)
  async function writeRow(rowNum, values, expect) {
    if (state.source === "sheet") {
      const send = values.slice();
      if (send[0] !== null) { const d = serialToDate(send[0]); send[0] = [d.getFullYear(), d.getMonth() + 1, d.getDate()]; }
      const res = rowNum
        ? await api.call({ action: "update", row: rowNum, expect, values: send })
        : await api.call({ action: "add", values: send });
      rowNum = res.row;
    } else if (!rowNum) {
      rowNum = state.lastUsedRow + 1;
    }
    const m = state.matrix;
    // Le script crée la colonne "Payé" au premier usage : même chose sur la copie locale
    if (values[6] != null && fold((m[0] || [])[6]) !== "paye") m[0] = Object.assign((m[0] || []).slice(), { 6: "Payé" });
    while (m.length < rowNum) m.push([]);
    const row = m[rowNum - 1] = (m[rowNum - 1] || []).slice();
    values.forEach((v, j) => { if (v !== null) row[j] = v; });
    buildRows(m);
    saveCache();
    renderAll();
  }

  // ---------- Cache local (affichage instantané) ----------
  function saveCache() {
    if (state.source !== "sheet") return;
    LS.set("pc_cache", { matrix: state.matrix, at: state.syncedAt });
  }

  // ---------- Chargement ----------
  async function refresh(silent) {
    if (state.source !== "sheet") { if (!silent) toast("Mode fichier : rouvre le fichier pour actualiser"); return false; }
    const btn = $("#btnRefresh");
    btn.classList.add("spin");
    try {
      const m = await api.load();
      state.syncedAt = Date.now();
      buildRows(m);
      saveCache();
      hideBanner();
      renderAll();
      if (!silent) toast("À jour ✓");
      return true;
    } catch (e) {
      console.error(e);
      banner("Impossible de lire le tableau : " + e.message, true);
      if (e.auth) { LS.del("pc_key"); LS.del("pc_cache"); askCode(e.message); }
      return false;
    } finally {
      btn.classList.remove("spin");
    }
  }

  async function openLocalFile(file) {
    try {
      const buf = await file.arrayBuffer();
      const wb = XLSX.read(buf, { type: "array" });
      const name = wb.SheetNames.find((n) => fold(n) === "commande") || wb.SheetNames[0];
      const ws = wb.Sheets[name];
      const ref = XLSX.utils.decode_range(ws["!ref"] || "A1:F1");
      const m = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null, blankrows: true, range: { s: { r: 0, c: 0 }, e: { r: ref.e.r, c: 6 } } });
      state.source = "file";
      state.fileName = file.name;
      state.syncedAt = Date.now();
      buildRows(m);
      $("#welcome").hidden = true;
      banner("Mode fichier : les modifications restent sur l'appli, elles ne sont pas enregistrées dans Excel.");
      renderAll();
      toast(`${state.rows.length} commandes chargées`);
    } catch (e) {
      console.error(e);
      toast("Fichier illisible : " + e.message, true);
    }
  }

  // ---------- Rendu : Atelier ----------
  function renderHome() {
    const t = today();
    const todo = state.rows.filter((r) => !r.done).map((r) => ({ r, u: urgency(r) }));
    todo.sort((a, b) => (a.r.date ? a.r.date : Infinity) - (b.r.date ? b.r.date : Infinity) || a.r.r - b.r.r);
    const late = todo.filter((x) => x.u.cls === "late").length;
    const month = state.rows.filter((r) => r.date && r.date.getMonth() === t.getMonth() && r.date.getFullYear() === t.getFullYear()).length;
    const unpaid = state.rows.filter((r) => r.pay === "unpaid").length;

    $("#kpis").innerHTML = `
      <div class="kpi accent"><div class="v">${todo.length}</div><div class="l">À fabriquer</div></div>
      <div class="kpi ${late ? "danger" : ""}"><div class="v">${late}</div><div class="l">En retard (&gt; ${cfg().delai} j)</div></div>
      <button class="kpi ${unpaid ? "warn" : ""}" data-goto="unpaid"><div class="v">${unpaid}</div><div class="l">Non payées</div></button>
      <div class="kpi"><div class="v">${month}</div><div class="l">Commandes en ${fmt(t, { month: "long" })}</div></div>`;

    $("#todoList").innerHTML = todo.length ? todo.map(({ r, u }) => jobCard(r, u)).join("") :
      `<div class="empty"><b>Tout est fabriqué 🎉</b>Aucune commande en attente.</div>`;

    const recent = state.rows.filter((r) => r.done).slice(-6).reverse();
    $("#recentList").innerHTML = recent.map((r, i) => rowHtml(r, i === recent.length - 1)).join("");
  }

  function jobCard(r, u) {
    let n, lbl;
    if (u.cls === "unknown") { n = "?"; lbl = "sans date"; }
    else if (u.left < 0) { n = "+" + -u.left; lbl = -u.left > 1 ? "jours de retard" : "jour de retard"; }
    else if (u.left === 0) { n = "0"; lbl = "dernier jour"; }
    else { n = u.left; lbl = u.left > 1 ? "jours restants" : "jour restant"; }
    const pct = u.cls === "unknown" ? 0 : Math.min(100, Math.max(4, (u.age / cfg().delai) * 100));
    return `<button class="job ${u.cls}" data-row="${r.r}">
      <div class="job-title">${esc(r.name)}</div>
      <div class="job-meta">
        ${r.color ? `<span class="tag">${swatch(r.color)}${esc(r.color)}</span>` : ""}
        <span>${r.date ? "Cdé le " + fmt(r.date, { day: "numeric", month: "short" }) : "Date inconnue"}</span>
        ${payPill(r)}
      </div>
      ${r.client || r.obs ? `<div class="job-note">${esc([r.client, r.obs].filter(Boolean).join(" · "))}</div>` : ""}
      <div class="countdown"><div class="n">${n}</div><div class="t">${lbl}</div></div>
      <div class="progress"><i style="width:${pct}%"></i></div>
    </button>`;
  }

  function rowHtml(r, last) {
    const u = urgency(r);
    const sub = [r.color, r.client, r.obs].filter(Boolean).join(" · ");
    return `<button class="row ${last ? "last" : ""}" data-row="${r.r}" ${u ? `style="--u:var(--${u.cls === "unknown" ? "muted" : u.cls})"` : ""}>
      ${swatch(r.color)}
      <div class="row-main">
        <div class="row-title">${esc(r.name)}</div>
        <div class="row-sub">${esc(sub || fmt(r.date))}</div>
      </div>
      <div class="row-side">
        <span class="status ${r.done ? "done" : "todo"}">${r.done ? "Faite" : "À faire"}</span>
        ${payPill(r)}
      </div>
    </button>`;
  }

  // ---------- Rendu : Historique ----------
  const PAGE = 80;
  function filterHistory() {
    const q = fold(state.hist.q);
    const terms = q.split(" ").filter(Boolean);
    const st = state.hist.status;
    state.hist.list = state.rows.filter((r) => {
      if (st === "todo" && r.done) return false;
      if (st === "done" && !r.done) return false;
      if (st === "unpaid" && r.pay !== "unpaid") return false;
      if (!terms.length) return true;
      const hay = fold([r.name, r.model, r.colorRaw, r.color, r.client, r.obs, r.pay === "unpaid" ? "non payee" : r.pay === "paid" ? "payee" : "", r.date ? fmt(r.date, { day: "numeric", month: "long", year: "numeric" }) : ""].join(" "));
      return terms.every((t) => hay.includes(t));
    }).reverse();
    state.hist.shown = 0;
    $("#historyList").innerHTML = "";
    $("#resultCount").textContent = `${state.hist.list.length} commande${state.hist.list.length > 1 ? "s" : ""}`;
    moreHistory();
  }

  function moreHistory() {
    const { list, shown } = state.hist;
    if (shown >= list.length) return;
    const slice = list.slice(shown, shown + PAGE);
    const prev = list[shown - 1];
    let html = "", lastKey = prev && prev.date ? prev.date.getTime() : "x";
    slice.forEach((r, i) => {
      const key = r.date ? r.date.getTime() : "none";
      if (key !== lastKey) {
        html += `<div class="day-head">${r.date ? fmt(r.date, { weekday: "long", day: "numeric", month: "long", year: "numeric" }) : "Sans date"}</div>`;
        lastKey = key;
      }
      const next = slice[i + 1] || list[shown + PAGE];
      const isLast = !next || (next.date ? next.date.getTime() : "none") !== key;
      html += rowHtml(r, isLast);
    });
    $("#historyList").insertAdjacentHTML("beforeend", html);
    state.hist.shown += slice.length;
  }

  // ---------- Rendu : Stats ----------
  function renderStats() {
    const t = today();
    // Commandes par mois (12 mois)
    const months = [];
    for (let i = 11; i >= 0; i--) months.push(new Date(t.getFullYear(), t.getMonth() - i, 1));
    const counts = months.map((m) => state.rows.filter((r) => r.date && r.date.getFullYear() === m.getFullYear() && r.date.getMonth() === m.getMonth()).length);
    const max = Math.max(1, ...counts);
    $("#chartMonths").innerHTML = `<div class="bars">${months.map((m, i) => `
      <div class="bar"><span class="num">${counts[i]}</span>
        <div class="col ${i === 11 ? "cur" : ""}" style="height:${(counts[i] / max) * 100}%"></div>
        <span class="lab">${fmt(m, { month: "short" }).replace(".", "")}</span></div>`).join("")}</div>`;

    // Top modèles
    const from = state.topRange === "all" ? null : addDays(t, -365);
    const scope = state.rows.filter((r) => !from || (r.date && r.date > from));
    const byModel = {};
    for (const r of scope) byModel[r.name] = (byModel[r.name] || 0) + 1;
    const top = Object.entries(byModel).sort((a, b) => b[1] - a[1]).slice(0, 10);
    const tmax = top.length ? top[0][1] : 1;
    $("#chartModels").innerHTML = `<div class="hbars">${top.map(([n, c]) => `
      <div class="hbar"><span class="name">${esc(n)}</span><div class="track"><div class="fill" style="width:${(c / tmax) * 100}%"></div></div><span class="val">${c}</span></div>`).join("")}</div>`;

    // Couleurs (1 an)
    const yr = state.rows.filter((r) => r.date && r.date > addDays(t, -365) && r.color);
    const byColor = {};
    for (const r of yr) { const k = COLORS[r.color] ? r.color : "Autre"; byColor[k] = (byColor[k] || 0) + 1; }
    const cols = Object.entries(byColor).sort((a, b) => b[1] - a[1]);
    const cmax = cols.length ? cols[0][1] : 1;
    $("#chartColors").innerHTML = `<div class="hbars">${cols.map(([n, c]) => `
      <div class="hbar"><span class="name">${swatch(n === "Autre" ? "Autre" : n)}${esc(n)}</span><div class="track"><div class="fill" style="width:${(c / cmax) * 100}%;background:${COLORS[n] && n !== "Noir" ? COLORS[n] : n === "Noir" ? "#5b6275" : "var(--accent)"}"></div></div><span class="val">${c}</span></div>`).join("")}</div>`;
  }

  function renderAll() {
    $("#syncInfo").textContent = state.source === "file"
      ? `Fichier : ${state.fileName}`
      : state.syncedAt ? `Synchro ${new Date(state.syncedAt).toLocaleTimeString("fr-FR", { hour: "2-digit", minute: "2-digit" })} · ${state.rows.length} cdes` : "—";
    renderHome();
    if (state.view === "history") filterHistory();
    if (state.view === "stats") renderStats();
  }

  // ---------- Navigation ----------
  const TITLES = { home: "Atelier", history: "Historique", stats: "Statistiques" };
  function show(view) {
    state.view = view;
    $$(".view").forEach((v) => (v.hidden = v.id !== "view-" + view));
    $$(".tab").forEach((t) => t.classList.toggle("active", t.dataset.view === view));
    $("#pageTitle").textContent = TITLES[view];
    if (view === "history") filterHistory();
    if (view === "stats") renderStats();
    window.scrollTo({ top: 0 });
  }

  // ---------- Feuilles modales ----------
  function openSheet(html) {
    $("#sheetBody").innerHTML = html;
    $("#sheet").hidden = false;
    $("#sheetBackdrop").hidden = false;
    $("#sheet").scrollTop = 0;
  }
  function closeSheet() { $("#sheet").hidden = true; $("#sheetBackdrop").hidden = true; }
  const sheetOpen = () => !$("#sheet").hidden;

  // Fiche d'actions rapides : un appui = une modification enregistrée
  function openDetail(rowNum) {
    const r = state.rows.find((x) => x.r === rowNum);
    if (!r) return;
    const u = urgency(r);
    let statut = "Fabriquée ✓";
    if (u) statut = u.cls === "unknown" ? "À faire" : u.left < 0 ? `En retard de ${-u.left} j` : `À faire · ${u.left} j restant${u.left > 1 ? "s" : ""}`;
    const payTxt = r.pay === "paid" ? "Payée" : r.pay === "unpaid" ? "Non payée" : "Non renseigné";
    openSheet(`
      <div class="detail-head">
        ${swatch(r.color, "lg")}
        <div><h3>${esc(r.name)}</h3><span class="status ${r.done ? "done" : "todo"}" ${u ? `style="--u:var(--${u.cls === "unknown" ? "muted" : u.cls})"` : ""}>${esc(statut)}</span></div>
      </div>
      <div class="quick">
        <button class="qbtn ${r.done ? "on-done" : ""}" data-act="done">
          <span class="ico">${r.done ? "✓" : "🔧"}</span>${r.done ? "Fabriquée" : "Marquer fabriquée"}<small>${r.done ? "appuyer pour annuler" : "en cours"}</small>
        </button>
        <button class="qbtn ${r.pay === "paid" ? "on-paid" : r.pay === "unpaid" ? "on-unpaid" : ""}" data-act="pay">
          <span class="ico">€</span>${payTxt}<small>appuyer pour changer</small>
        </button>
      </div>
      <div class="qlabel">Couleur</div>
      <div class="qcolors">
        ${COLOR_CHOICES.map((c) => `<button data-color="${c}" class="${r.color === c ? "active" : ""}" aria-label="${c}">${swatch(c, "lg")}</button>`).join("")}
      </div>
      <div class="qlabel">Client / note</div>
      <div class="qnote">
        <textarea id="qNote" placeholder="Ex : Julien mail, payé virement…">${esc(r.client)}</textarea>
        <button data-act="note">OK</button>
      </div>
      <div class="detail-grid">
        <div class="cell"><div class="k">Commandé le</div><div class="v">${fmt(r.date)}${r.ownDate ? "" : " <span class='muted'>(ligne du dessus)</span>"}</div></div>
        <div class="cell"><div class="k">Échéance (${cfg().delai} j)</div><div class="v">${r.date ? fmt(addDays(r.date, cfg().delai)) : "—"}</div></div>
        ${r.obs ? `<div class="cell full"><div class="k">Observations</div><div class="v">${esc(r.obs)}</div></div>` : ""}
      </div>
      <button class="btn ghost" data-act="edit" style="width:100%">Tout modifier (date, modèle, observations…)</button>`);

    const save = async (values, msg, close) => {
      $$("#sheetBody button").forEach((b) => (b.disabled = true));
      try {
        await writeRow(r.r, values, r.model);
        toast(msg);
        if (close) closeSheet(); else openDetail(r.r);
      } catch (err) {
        $$("#sheetBody button").forEach((b) => (b.disabled = false));
        toast("Échec : " + err.message, true);
      }
    };
    $("#sheetBody").onclick = (e) => {
      const c = e.target.closest("[data-color]");
      if (c) {
        if (c.dataset.color !== r.color) save([null, null, null, c.dataset.color, null, null, null], `Couleur : ${c.dataset.color}`);
        return;
      }
      const act = e.target.closest("[data-act]")?.dataset.act;
      if (act === "edit") return openForm(r);
      if (act === "done") return save([null, r.done ? "" : "F", null, null, null, null, null], r.done ? "Remise en cours" : `${r.name} fabriquée ✓`, !r.done);
      if (act === "pay") return save([null, null, null, null, null, null, PAY_VALUE[r.pay === "paid" ? "unpaid" : "paid"]], r.pay === "paid" ? "Non payée" : "Payée ✓");
      if (act === "note") return save([null, null, null, null, clean($("#qNote").value), null, null], "Note enregistrée ✓");
    };
  }

  const COLOR_CHOICES = ["Noir", "Blanc", "Bleu", "Jaune", "Rouge", "Vert"];
  function openForm(r) {
    const isNew = !r;
    const topModels = (() => {
      const c = {}; const from = addDays(today(), -365);
      for (const x of state.rows) if (x.date && x.date > from) c[x.name] = (c[x.name] || 0) + 1;
      return Object.entries(c).sort((a, b) => b[1] - a[1]).map((e) => e[0]);
    })();
    const allModels = [...new Set(state.rows.map((x) => x.name))].sort((a, b) => a.localeCompare(b, "fr"));
    const curColor = r ? r.color : "Noir";
    const isOther = r && r.color && !COLOR_CHOICES.includes(r.color);
    openSheet(`
      <h3>${isNew ? "Nouvelle commande" : "Modifier la commande"}</h3>
      <p class="muted" style="margin:0 0 14px">${isNew ? `Ajoutée à la ligne ${state.lastUsedRow + 1} du fichier` : `Ligne ${r.r} du fichier`}</p>
      <form class="form" id="orderForm" autocomplete="off">
        <div class="field"><label>Date de commande</label>
          <input type="date" name="date" value="${isoDate(r ? r.date || today() : today())}" required></div>
        <div class="field"><label>Modèle</label>
          <input name="model" list="modelList" placeholder="Ex : Maxx v2" value="${esc(r ? r.model : "")}" required>
          <datalist id="modelList">${allModels.map((m) => `<option value="${esc(m)}">`).join("")}</datalist>
          ${isNew ? `<div class="suggest">${topModels.slice(0, 8).map((m) => `<button type="button" data-model="${esc(m)}">${esc(m)}</button>`).join("")}</div>` : ""}
        </div>
        <div class="field"><label>Couleur</label>
          <div class="color-pick">
            ${COLOR_CHOICES.map((c) => `<button type="button" data-color="${c}" class="${!isOther && curColor === c ? "active" : ""}">${swatch(c, "lg")}${c}</button>`).join("")}
            <button type="button" data-color="__other" class="${isOther ? "active" : ""}">${swatch("Autre", "lg")}Autre</button>
          </div>
          <input name="colorOther" placeholder="Couleur / précision (ex : Noir 3 mm)" value="${esc(isOther ? r.colorRaw : "")}" style="margin-top:10px" ${isOther ? "" : "hidden"}>
        </div>
        <div class="field"><label>Client / note</label>
          <input name="client" placeholder="Ex : Julien mail" value="${esc(r ? r.client : "")}"></div>
        <div class="field"><label>Observations</label>
          <textarea name="obs" placeholder="Optionnel">${esc(r ? r.obs : "")}</textarea></div>
        <div class="switch-row"><span>Payée</span><button type="button" class="switch ${!r || r.pay !== "unpaid" ? "on" : ""}" id="paySwitch" aria-label="Payée"></button></div>
        <div class="switch-row"><span>Déjà fabriquée</span><button type="button" class="switch ${r && r.done ? "on" : ""}" id="doneSwitch" aria-label="Déjà fabriquée"></button></div>
        <button class="btn primary" type="submit">${isNew ? "Ajouter la commande" : "Enregistrer"}</button>
      </form>`);

    const form = $("#orderForm");
    let color = isOther ? "__other" : curColor;
    $("#sheetBody").onclick = (e) => {
      const m = e.target.closest("[data-model]");
      if (m) { form.model.value = m.dataset.model; return; }
      const c = e.target.closest("[data-color]");
      if (c) {
        color = c.dataset.color;
        $$(".color-pick button", form).forEach((b) => b.classList.toggle("active", b === c));
        form.colorOther.hidden = color !== "__other";
        if (color === "__other") form.colorOther.focus();
        return;
      }
      if (e.target.id === "doneSwitch" || e.target.id === "paySwitch") e.target.classList.toggle("on");
    };
    form.onsubmit = async (e) => {
      e.preventDefault();
      const btn = form.querySelector("[type=submit]"); btn.disabled = true;
      const [y, mo, d] = form.date.value.split("-").map(Number);
      const date = new Date(y, mo - 1, d);
      const colorVal = color === "__other" ? clean(form.colorOther.value) : color;
      const done = $("#doneSwitch").classList.contains("on") ? "F" : "";
      const dateChanged = isNew || !r.date || r.date.getTime() !== date.getTime();
      const payOn = $("#paySwitch").classList.contains("on");
      // Ancienne commande sans info de paiement : on ne touche pas la colonne si l'interrupteur n'a pas bougé
      const pay = !isNew && !r.pay && payOn ? null : PAY_VALUE[payOn ? "paid" : "unpaid"];
      const values = [dateChanged ? dateToSerial(date) : null, done, clean(form.model.value), colorVal, clean(form.client.value), clean(form.obs.value), pay];
      try {
        await writeRow(isNew ? null : r.r, values, isNew ? null : r.model);
        closeSheet();
        toast(isNew ? "Commande ajoutée ✓" : "Modifiée ✓");
      } catch (err) { btn.disabled = false; toast("Échec : " + err.message, true); }
    };
  }

  function openSettings() {
    const c = cfg();
    openSheet(`
      <h3>Réglages</h3>
      <p class="muted" style="margin:0 0 14px">${state.source === "sheet" ? "Connecté au tableau Google Sheets" : state.source === "file" ? `Fichier local : ${esc(state.fileName)}` : "Aucune source"}</p>
      <form class="form" id="setForm" autocomplete="off">
        <div class="field"><label>Délai de fabrication (jours)</label><input name="delai" type="number" min="1" max="90" value="${c.delai}"></div>
        <div class="field"><label>Adresse du script Google</label><input name="url" placeholder="https://script.google.com/macros/s/…/exec" value="${esc(c.scriptUrl)}"></div>
        <button class="btn primary" type="submit">Enregistrer</button>
      </form>
      <div class="btns two" style="margin-top:12px">
        <label class="btn ghost file-btn">Ouvrir un fichier<input type="file" id="fileInput2" accept=".xlsx,.xlsm,.xls"></label>
        ${c.key ? `<button class="btn ghost" id="btnForget">Oublier le code</button>` : `<button class="btn ghost" id="btnCode2">Code d'accès</button>`}
      </div>`);
    const f = $("#setForm");
    f.onsubmit = async (e) => {
      e.preventDefault();
      LS.set("pc_delai", Math.max(1, Number(f.delai.value) || 10));
      LS.set("pc_scriptUrl", clean(f.url.value));
      closeSheet();
      renderAll();
      toast("Réglages enregistrés");
      if (state.source === "sheet") refresh(true);
    };
    $("#fileInput2").onchange = (e) => { if (e.target.files[0]) { closeSheet(); openLocalFile(e.target.files[0]); } };
    const fg = $("#btnForget");
    if (fg) fg.onclick = () => { LS.del("pc_key"); LS.del("pc_cache"); location.reload(); };
    const cd = $("#btnCode2"); if (cd) cd.onclick = () => askCode();
  }

  // Saisie du code d'accès (une fois par téléphone)
  function askCode(err) {
    const c = cfg();
    $("#welcome").hidden = true;
    openSheet(`
      <h3>Code d'accès</h3>
      <p class="muted" style="margin:0 0 14px">Saisis le code une seule fois : il reste enregistré sur ce téléphone.</p>
      ${err ? `<div class="banner error" style="margin:0 0 14px">${esc(err)}</div>` : ""}
      <form class="form" id="codeForm" autocomplete="off">
        ${base.SCRIPT_URL ? "" : `
        <div class="field"><label>Adresse du script Google</label><input name="url" placeholder="https://script.google.com/macros/s/…/exec" value="${esc(c.scriptUrl)}" required></div>`}
        <div class="field"><label>Code d'accès</label><input name="key" type="password" autocapitalize="off" autocorrect="off" spellcheck="false" required></div>
        <button class="btn primary" type="submit">Se connecter</button>
      </form>`);
    const f = $("#codeForm");
    f.onsubmit = async (e) => {
      e.preventDefault();
      const btn = f.querySelector("[type=submit]"); btn.disabled = true; btn.textContent = "Connexion…";
      if (f.url) LS.set("pc_scriptUrl", clean(f.url.value));
      LS.set("pc_key", f.key.value.trim());
      state.source = "sheet";
      closeSheet();
      if (await refresh(true)) toast(`Connecté ✓ ${state.rows.length} commandes`);
    };
    setTimeout(() => f.key.focus(), 250);
  }

  // ---------- Messages ----------
  let toastTimer;
  function toast(msg, err) {
    const t = $("#toast");
    t.textContent = msg; t.className = "toast" + (err ? " error" : ""); t.hidden = false;
    clearTimeout(toastTimer); toastTimer = setTimeout(() => (t.hidden = true), err ? 4500 : 2200);
  }
  function banner(msg, err) { const b = $("#banner"); b.textContent = msg; b.className = "banner" + (err ? " error" : ""); b.hidden = false; }
  function hideBanner() { $("#banner").hidden = true; }

  // ---------- Événements ----------
  $$(".tab").forEach((t) => (t.onclick = () => show(t.dataset.view)));
  $("#fab").onclick = () => {
    if (!state.source) { $("#welcome").hidden = false; return; }
    openForm(null);
  };
  $("#btnRefresh").onclick = () => refresh(false);
  $("#btnSettings").onclick = openSettings;
  $("#sheetBackdrop").onclick = closeSheet;
  $("#main").addEventListener("click", (e) => {
    const go = e.target.closest("[data-goto]");
    if (go) {
      state.hist.status = go.dataset.goto;
      $$("#statusChips .chip").forEach((c) => c.classList.toggle("active", c.dataset.status === go.dataset.goto));
      return show("history");
    }
    const el = e.target.closest("[data-row]");
    if (el) openDetail(Number(el.dataset.row));
  });
  let searchTimer;
  $("#searchInput").oninput = (e) => { clearTimeout(searchTimer); searchTimer = setTimeout(() => { state.hist.q = e.target.value; filterHistory(); }, 150); };
  $("#statusChips").onclick = (e) => {
    const b = e.target.closest("[data-status]"); if (!b) return;
    $$("#statusChips .chip").forEach((c) => c.classList.toggle("active", c === b));
    state.hist.status = b.dataset.status; filterHistory();
  };
  $("#topRange").onclick = (e) => {
    const b = e.target.closest("[data-range]"); if (!b) return;
    $$("#topRange button").forEach((c) => c.classList.toggle("active", c === b));
    state.topRange = b.dataset.range; renderStats();
  };
  new IntersectionObserver((ents) => { if (ents[0].isIntersecting && state.view === "history") moreHistory(); }, { rootMargin: "600px" })
    .observe($("#historySentinel"));
  $("#btnCode").onclick = () => askCode();
  $("#fileInput").onchange = (e) => { if (e.target.files[0]) openLocalFile(e.target.files[0]); };
  document.addEventListener("visibilitychange", () => {
    if (!document.hidden && state.source === "sheet" && Date.now() - (state.syncedAt || 0) > 60000) refresh(true);
  });

  // Bouton retour Android (Capacitor)
  const CapApp = window.Capacitor && window.Capacitor.Plugins && window.Capacitor.Plugins.App;
  if (CapApp) {
    CapApp.addListener("backButton", () => {
      if (sheetOpen()) closeSheet();
      else if (state.view !== "home") show("home");
      else CapApp.exitApp();
    });
  }

  // ---------- Démarrage ----------
  (function start() {
    if (api.ready()) {
      state.source = "sheet";
      const cache = LS.get("pc_cache", null);
      if (cache && cache.matrix) { state.syncedAt = cache.at; buildRows(cache.matrix); }
      renderAll();
      refresh(true);
      return;
    }
    renderAll();
    $("#welcome").hidden = false;
  })();

  // Web-app installable : cache de l'interface (inutile dans l'APK)
  const isNative = window.Capacitor && window.Capacitor.isNativePlatform && window.Capacitor.isNativePlatform();
  if ("serviceWorker" in navigator && !isNative && location.protocol === "https:") {
    navigator.serviceWorker.register("sw.js", { updateViaCache: "none" }).then((reg) => {
      // Cherche une nouvelle version à chaque retour dans l'appli
      document.addEventListener("visibilitychange", () => { if (!document.hidden) reg.update().catch(() => {}); });
    }).catch((e) => console.warn("SW", e));
    // Nouvelle version installée : on recharge une fois pour l'afficher
    let reloaded = false;
    const hadController = !!navigator.serviceWorker.controller;
    navigator.serviceWorker.addEventListener("controllerchange", () => {
      if (hadController && !reloaded && !sheetOpen()) { reloaded = true; location.reload(); }
    });
  }

  // Point d'entrée de test (navigateur de développement uniquement)
  window.__pc = { openLocalFile, state };
})();
