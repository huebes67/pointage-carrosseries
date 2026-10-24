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
  // Regroupement des différentes écritures d'un même modèle (le tableau n'est pas modifié)
  const MODEL_ALIASES = [
    [/\bshort ?course\b/g, "sc"], [/\bgranite\b/g, "granit"], [/\bporshe\b/g, "porsche"], [/\bkkpit\b/g, "kpit"],
    [/\bcoxmaxx\b/g, "cox maxx"], [/\bsc ?370\b/g, "sc 370"], [/\bste? ?pro\b/g, "st pro"], [/\bbig ?mama\b/g, "big mama"],
    [/\bbig ?rock\b/g, "big rock"], [/\b(205|250) ?gti\b/g, "205 gti"], [/\bx ?monster\b/g, "x monster"],
    [/\berevo ?2 ?pro\b/g, "erevo 2 pro"], [/\berevo 2\.0\b/g, "erevo 2"], [/\b(mojave|mojito) ?4 ?s\b/g, "$1 4s"],
    [/\bpick ?up\b/g, "pick up"], [/\bhoos\b/g, "hoss"], [/^gorgon$/, "gorgonzola"], [/\bcox gold ?[mn]\b/g, "cox gold"], [/^maxx$/, "maxx v1"],
    [/\baggressif\b/g, "agressif"], [/\btekkno\b/g, "tekno"], [/^chargeur$/, "dodge chargeur"], [/\bempattement\b/g, ""],
    [/\bnew\b/g, "nouveau"], [/\bl ancien model\b/g, "ancien"], [/^chevy 80\b/, "chevy 1980"], [/\bxmax\b/g, "xmaxx"],
    [/\btm\b/g, "team magic"], [/\beb 6\b/g, "eb6"], [/^eb6$/, "team magic eb6"], [/\bcox 1\/0\b/g, "cox 1/10"],
  ];
  const MODEL_STOPWORDS = new Set(["du", "de", "la", "le", "et", "l", "d"]);
  // Nom affiché imposé pour certains groupes (sinon : l'écriture la plus fréquente)
  const MODEL_DISPLAY = {
    "chargeur dodge": "Dodge chargeur", "eb6 magic team": "Team magic EB6",
    "1/8 ancien truggy": "Truggy 1/8 ancien", "1/10 porsche": "Porsche 1/10",
    "maxx v1": "Maxx v1", "cox gold": "Cox gold",
  };
  function baseKey(s) {
    let k = fold(s);
    for (const [re, to] of MODEL_ALIASES) k = k.replace(re, to);
    // l'ordre des mots ne compte pas : "chargeur dodge" = "dodge chargeur"
    return k.split(/\s+/).filter((w) => w && !MODEL_STOPWORDS.has(w)).sort().join(" ");
  }

  // Regroupements faits à la main dans l'appli : gardés sur cet appareil, le tableau n'est jamais modifié.
  // pc_groups = [{ name: "Nom final", members: [{ key: clé de base, label: nom affiché }] }]
  let userGroupOf = {}, userGroupName = {};
  function loadUserGroups() {
    userGroupOf = {}; userGroupName = {};
    for (const g of LS.get("pc_groups", [])) {
      userGroupName["ug:" + fold(g.name)] = g.name;
      userGroupOf[baseKey(g.name)] = g; // un moule enregistré sous le nom du regroupement s'y rattache aussi
      for (const m of g.members) userGroupOf[m.key] = g;
    }
  }
  loadUserGroups();
  // Nom écrit dans le tableau : l'écriture réelle la plus utilisée du modèle, jamais un nom de regroupement inventé
  function realName(name) {
    const k = modelKey(name);
    if (!k.startsWith("ug:")) return name;
    const c = {};
    for (const r of state.rows) if (r.key === k) c[r.model] = (c[r.model] || 0) + 1;
    for (const m of state.moules) if (m.key === k) c[m.model] = (c[m.model] || 0) + 0.5;
    const best = Object.entries(c).sort((a, b) => b[1] - a[1])[0];
    return best ? best[0] : name;
  }
  function modelKey(s) {
    const k = baseKey(s);
    const g = userGroupOf[k];
    return g ? "ug:" + fold(g.name) : k;
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

  // Bande en haut de la tuile, de la couleur de la carrosserie
  const colorBand = (name) => {
    if (!name) return "";
    const c = COLORS[name];
    const light = name === "Blanc" || name === "Jaune";
    const bg = c || "#5b6275"; // couleur non standard : gris uni avec son nom
    return `<div class="cband${name === "Noir" ? " noir" : ""}" style="--c:${bg};--t:${light ? "#15161a" : "#fff"}">${esc(name)}</div>`;
  };

  // Icônes au trait, même style que la barre de navigation
  const ICONS = {
    check: '<path d="M5 12.5l4.5 4.5L19 7.5"/>',
    undo: '<path d="M9 14L4 9l5-5"/><path d="M4 9h10.5a5.5 5.5 0 0 1 0 11H11"/>',
    trash: '<path d="M4 7h16M10 11v6M14 11v6M6 7l1 13h10l1-13M9 7V4h6v3"/>',
    tool: '<path d="M14.7 6.3a4 4 0 0 0-5.4 5.4L3.5 17.5l3 3 5.8-5.8a4 4 0 0 0 5.4-5.4l-2.6 2.6-2.4-.6-.6-2.4z"/>',
    euro: '<path d="M17.5 6.5a7 7 0 1 0 0 11M4 10h9M4 14h9"/>',
    moon: '<path d="M20 14.5A8 8 0 1 1 9.5 4a6.5 6.5 0 0 0 10.5 10.5z"/>',
    sun: '<circle cx="12" cy="12" r="4"/><path d="M12 2.5v2M12 19.5v2M4.6 4.6l1.4 1.4M18 18l1.4 1.4M2.5 12h2M19.5 12h2M4.6 19.4L6 18M18 6l1.4-1.4"/>',
    auto: '<circle cx="12" cy="12" r="8.5"/><path d="M12 3.5v17"/>',
    alert: '<path d="M12 3.5L21.5 20h-19z"/><path d="M12 10v4.5M12 17.2v.3"/>',
    clock: '<circle cx="12" cy="12" r="8.5"/><path d="M12 7.5V12l3 2"/>',
    mold: '<path d="M12 3l8 4.5v9L12 21l-8-4.5v-9z"/><path d="M4 7.5L12 12l8-4.5M12 12v9"/>',
    swap: '<path d="M7 4L3 8l4 4M3 8h14M17 20l4-4-4-4M21 16H7"/>',
  };
  const icon = (name, cls = "") => `<svg class="ic ${cls}" viewBox="0 0 24 24" aria-hidden="true">${ICONS[name]}</svg>`;

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
    topCount: "10",
    moules: [],        // moules (feuille "Moules")
    mmatrix: null,     // valeurs brutes de la feuille Moules
    mouleAlert: {},    // modèle -> moule cassé ou en séchage
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
      display[k] = userGroupName[k] || MODEL_DISPLAY[k] || best.charAt(0).toUpperCase() + best.slice(1);
    }
    for (const r of rows) r.name = display[r.key];
    state.rows = rows;
    state.lastUsedRow = lastUsed;
    state.matrix = matrix;
  }

  // ---------- Moules : A Modèle · B État · C Note · D Lieu · E Échelle ----------
  const MOULE_HEAD = ["Modèle", "État", "Note", "Lieu", "Échelle"];
  const ECHELLES = ["1/5", "1/7", "1/8", "1/10", "1/16"];
  // Échelle écrite dans le nom du modèle ("Cox baja 1/10", "Truggy 1/8"…)
  function scaleFromName(name) {
    const m = String(name || "").match(/\b1\s*\/\s*(5|7|8|10|16)\b/);
    return m ? "1/" + m[1] : "";
  }
  const LIEUX = ["Chez Jean", "À l'atelier"];
  const MOULE_LABEL = { ok: "OK", casse: "Cassé", sechage: "Séchage" };
  function mouleState(raw) {
    const f = fold(raw);
    if (f.startsWith("cass")) return "casse";
    if (f.startsWith("sech")) return "sechage";
    return "ok";
  }
  function buildMoules(m) {
    if (!m || !m.length) m = [MOULE_HEAD.slice()];
    const list = [];
    for (let i = 1; i < m.length; i++) {
      const a = m[i] || [];
      const model = clean(a[0]);
      if (!model) continue;
      list.push({ r: i + 1, model, key: modelKey(model), etat: mouleState(a[1]), note: clean(a[2]), lieu: clean(a[3]), scale: clean(a[4]) });
    }
    // Pour chaque modèle : le moule le plus bloquant (cassé avant séchage)
    const alert = {};
    for (const x of list) {
      if (x.etat === "ok") continue;
      const cur = alert[x.key];
      if (!cur || (x.etat === "casse" && cur.etat !== "casse")) alert[x.key] = x;
    }
    state.moules = list;
    state.mmatrix = m;
    state.mouleAlert = alert;
  }
  // Voyant sur une commande dont le moule est cassé ou en séchage
  function mouleFlag(key, long) {
    const x = state.mouleAlert[key];
    if (!x) return "";
    if (x.etat === "casse") return `<span class="mflag casse">${icon("alert")}Moule cassé</span>`;
    return `<span class="mflag sechage">${icon("clock")}Moule en séchage</span>`;
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
  const payPill = (r) => r.pay ? `<span class="pay ${r.pay}">${icon("euro")}${r.pay === "paid" ? "Payée" : "Non payée"}</span>` : "";

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
        const err = new Error("Pas de connexion internet ?"); err.network = true; throw err;
      }
      if (!res.ok) { const err = new Error("Serveur Google indisponible (" + res.status + ")"); err.network = true; throw err; }
      let j;
      try { j = await res.json(); } catch { throw new Error("Réponse inattendue : vérifie l'adresse du script"); }
      if (!j.ok) { const err = new Error(j.error || "Erreur"); err.auth = /code/i.test(j.error || ""); throw err; }
      return j;
    },
    async load() { return this.call({ action: "read" }); },
  };

  // ---------- File d'envoi en arrière-plan ----------
  // Chaque modification s'affiche tout de suite ; l'envoi au tableau se fait ensuite, dans l'ordre.
  // La file est gardée sur le téléphone : sans réseau, elle repart dès que la connexion revient.
  const queue = { items: LS.get("pc_queue", []), busy: false, wantRefresh: false };
  function enqueue(body) {
    queue.items.push(body);
    LS.set("pc_queue", queue.items);
    renderSync();
    runQueue();
  }
  let retryTimer;
  async function runQueue() {
    if (queue.busy || state.source !== "sheet" || !queue.items.length) return;
    queue.busy = true;
    $("#btnRefresh").classList.add("spin");
    let resync = false;
    try {
      while (queue.items.length) {
        const body = queue.items[0];
        try {
          const res = await api.call(body);
          // Quelqu'un d'autre a ajouté une ligne en même temps : on se recale sur le tableau
          if (body.action === "add" && res.row !== body.localRow) resync = true;
        } catch (e) {
          if (e.network) { clearTimeout(retryTimer); retryTimer = setTimeout(runQueue, 15000); return; }
          if (e.auth) { queue.items = []; LS.del("pc_queue"); LS.del("pc_key"); askCode(e.message); return; }
          toast("Modification refusée : " + e.message, true);
          resync = true;
        }
        queue.items.shift();
        LS.set("pc_queue", queue.items);
        renderSync();
      }
    } finally {
      queue.busy = false;
      $("#btnRefresh").classList.remove("spin");
      renderSync();
    }
    if (resync || queue.wantRefresh) { queue.wantRefresh = false; refresh(true); }
  }
  window.addEventListener("online", runQueue);

  function renderSync() {
    const n = queue.items.length;
    $("#syncInfo").textContent = state.source === "file"
      ? `Fichier : ${state.fileName}`
      : n ? `${n} modif${n > 1 ? "s" : ""} en attente d'envoi…`
      : state.syncedAt ? `Synchro ${new Date(state.syncedAt).toLocaleTimeString("fr-FR", { hour: "2-digit", minute: "2-digit" })} · ${state.rows.length} cdes` : "—";
  }

  // ---------- Écriture (affichage immédiat, envoi en arrière-plan) ----------
  // values = [A..G] (date en numéro de série Excel) ; null = ne pas toucher la cellule
  // rowNum = null pour une nouvelle commande ; expect = modèle attendu sur la ligne (garde-fou)
  async function writeRow(rowNum, values, expect) {
    const isNew = !rowNum;
    if (isNew) rowNum = state.lastUsedRow + 1;
    if (state.source === "sheet") {
      const send = values.slice();
      if (send[0] !== null) { const d = serialToDate(send[0]); send[0] = [d.getFullYear(), d.getMonth() + 1, d.getDate()]; }
      enqueue(isNew ? { action: "add", values: send, localRow: rowNum } : { action: "update", row: rowNum, expect, values: send });
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

  async function deleteRow(r) {
    if (state.source === "sheet") enqueue({ action: "delete", row: r.r, expect: r.model });
    // Même opération sur la copie locale : la ligne suivante garde sa date si elle en héritait
    const m = state.matrix, i = r.r - 1, next = m[i + 1];
    const own = (m[i] || [])[0];
    if (own !== null && own !== "" && own !== undefined && next && (next[0] === null || next[0] === "" || next[0] === undefined) && clean(next[2])) {
      m[i + 1] = next.slice(); m[i + 1][0] = own;
    }
    m.splice(i, 1);
    buildRows(m);
    saveCache();
    renderAll();
  }

  // values = [modèle, état, note, lieu] ; "" = vider ; null = ne pas toucher
  async function writeMoule(rowNum, values, expect) {
    const m = state.mmatrix || [MOULE_HEAD.slice()];
    const isNew = !rowNum;
    if (isNew) rowNum = (state.moules.length ? Math.max(...state.moules.map((x) => x.r)) : 1) + 1;
    if (state.source === "sheet") {
      const send = values.map((v) => (typeof v === "number" ? (() => { const d = serialToDate(v); return [d.getFullYear(), d.getMonth() + 1, d.getDate()]; })() : v));
      enqueue(isNew ? { action: "add", sheet: "moules", values: send, localRow: rowNum } : { action: "update", sheet: "moules", row: rowNum, expect, values: send });
    }
    while (m.length < rowNum) m.push([]);
    const row = m[rowNum - 1] = (m[rowNum - 1] || []).slice();
    values.forEach((v, j) => { if (v !== null) row[j] = v; });
    buildMoules(m);
    saveCache();
    renderAll();
  }

  async function deleteMoule(x) {
    if (state.source === "sheet") enqueue({ action: "delete", sheet: "moules", row: x.r, expect: x.model });
    const m = state.mmatrix;
    m.splice(x.r - 1, 1);
    buildMoules(m);
    saveCache();
    renderAll();
  }

  function confirmDelete(r) {
    openSheet(`
      <h3>Supprimer cette commande ?</h3>
      <p class="muted" style="margin:6px 0 16px">${esc(r.name)}${r.color ? " · " + esc(r.color) : ""}${r.client ? " · " + esc(r.client) : ""}<br>
      Commandée le ${fmt(r.date)} · ligne ${r.r} du tableau</p>
      <div class="btns two">
        <button class="btn ghost" data-act="cancel">Annuler</button>
        <button class="btn danger" data-act="del">Supprimer</button>
      </div>`);
    $("#sheetBody").onclick = async (e) => {
      const act = e.target.closest("[data-act]")?.dataset.act;
      if (act === "cancel") return closeSheet();
      if (act !== "del") return;
      $$("#sheetBody button").forEach((b) => (b.disabled = true));
      try {
        await deleteRow(r);
        closeSheet();
        toast("Commande supprimée");
      } catch (err) {
        $$("#sheetBody button").forEach((b) => (b.disabled = false));
        toast("Échec : " + err.message, true);
      }
    };
  }

  // ---------- Glisser (swipe) ----------
  let swOpen = null, sw = null;
  function setSwipe(w, x, anim) {
    const front = w.querySelector(".sw-front");
    front.style.transition = anim ? "transform .22s cubic-bezier(.2,.8,.2,1)" : "none";
    front.style.transform = x ? `translateX(${x}px)` : "";
    w.dataset.dir = x > 0 ? "right" : x < 0 ? "left" : "";
    w._x = x;
  }
  function closeSwipe() { if (swOpen) { setSwipe(swOpen, 0, true); swOpen = null; } }
  const mainEl = $("#main");
  mainEl.addEventListener("pointerdown", (e) => {
    if (e.pointerType === "mouse" && e.button !== 0) return;
    const w = e.target.closest(".swipe");
    if (e.target.closest(".sw-act")) return;
    if (swOpen && swOpen !== w) closeSwipe();
    if (!w) return;
    sw = { w, x0: e.clientX, y0: e.clientY, base: w._x || 0, mode: null };
  });
  mainEl.addEventListener("pointermove", (e) => {
    if (!sw) return;
    const dx = e.clientX - sw.x0, dy = e.clientY - sw.y0;
    if (!sw.mode) {
      if (Math.abs(dx) > 10 && Math.abs(dx) > Math.abs(dy) * 1.3) sw.mode = "h";
      else if (Math.abs(dy) > 10) { sw = null; return; }
      else return;
    }
    setSwipe(sw.w, Math.max(-150, Math.min(150, sw.base + dx)), false);
  });
  const endSwipe = () => {
    if (!sw) return;
    const { w, mode } = sw; sw = null;
    if (mode !== "h") return;
    w._swipedAt = Date.now();
    const x = w._x || 0;
    if (x > 60) { setSwipe(w, 120, true); swOpen = w; }
    else if (x < -60) { setSwipe(w, -140, true); swOpen = w; }
    else { setSwipe(w, 0, true); if (swOpen === w) swOpen = null; }
  };
  mainEl.addEventListener("pointerup", endSwipe);
  mainEl.addEventListener("pointercancel", endSwipe);

  // ---------- Cache local (affichage instantané) ----------
  function saveCache() {
    if (state.source !== "sheet") return;
    LS.set("pc_cache", { matrix: state.matrix, mmatrix: state.mmatrix, at: state.syncedAt });
  }

  // ---------- Chargement ----------
  async function refresh(silent) {
    if (state.source !== "sheet") { if (!silent) toast("Mode fichier : rouvre le fichier pour actualiser"); return false; }
    if (queue.items.length || queue.busy) {
      queue.wantRefresh = true;
      runQueue();
      if (!silent) toast("Envoi des modifications en cours…");
      return false;
    }
    const btn = $("#btnRefresh");
    btn.classList.add("spin");
    try {
      const res = await api.load();
      state.syncedAt = Date.now();
      buildRows(res.values);
      buildMoules(res.moules);
      saveCache();
      hideBanner();
      renderAll();
      if (!silent) toast("À jour");
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
      const mName = wb.SheetNames.find((n) => /^moules?$/.test(fold(n)));
      buildMoules(mName ? XLSX.utils.sheet_to_json(wb.Sheets[mName], { header: 1, raw: true, defval: null }) : null);
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
    // Tri choisi par l'utilisateur (sans date : toujours en bas)
    const recentFirst = LS.get("pc_todoSort", "old") === "new";
    todo.sort((a, b) => {
      if (!a.r.date || !b.r.date) return (a.r.date ? 0 : 1) - (b.r.date ? 0 : 1) || a.r.r - b.r.r;
      const d = (a.r.date - b.r.date) || (a.r.r - b.r.r);
      return recentFirst ? -d : d;
    });
    $("#todoSort").classList.toggle("recent", recentFirst);
    const late = todo.filter((x) => x.u.cls === "late").length;
    const month = state.rows.filter((r) => r.date && r.date.getMonth() === t.getMonth() && r.date.getFullYear() === t.getFullYear()).length;
    const unpaid = state.rows.filter((r) => r.pay === "unpaid").length;

    $("#kpis").innerHTML = `
      <div class="kpi accent"><div class="v">${todo.length}</div><div class="l">À fabriquer</div></div>
      <div class="kpi ${late ? "danger" : ""}"><div class="v">${late}</div><div class="l">En retard (&gt; ${cfg().delai} j)</div></div>
      <button class="kpi ${unpaid ? "warn" : ""}" data-goto="unpaid"><div class="v">${unpaid}</div><div class="l">Non payées</div></button>
      <div class="kpi"><div class="v">${month}</div><div class="l">Commandes en ${fmt(t, { month: "long" })}</div></div>`;

    $("#todoList").innerHTML = todo.length ? todo.map(({ r, u }) => jobCard(r, u)).join("") :
      `<div class="empty"><b>Tout est fabriqué</b>Aucune commande en attente.</div>`;

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
    return swipeWrap(r, "card", `<button class="job sw-front ${u.cls} ${r.color ? "has-band" : ""}" data-row="${r.r}">
      ${colorBand(r.color)}
      <div class="job-title">${esc(r.name)}</div>
      ${mouleFlag(r.key) ? `<div class="job-flag">${mouleFlag(r.key)}</div>` : ""}
      <div class="job-meta">
        <span>${r.date ? "Cdé le " + fmt(r.date, { day: "numeric", month: "short" }) : "Date inconnue"}</span>
        ${payPill(r)}
      </div>
      ${r.client || r.obs ? `<div class="job-note">${esc([r.client, r.obs].filter(Boolean).join(" · "))}</div>` : ""}
      <div class="countdown"><div class="n">${n}</div><div class="t">${lbl}</div></div>
      <div class="progress"><i style="width:${pct}%"></i></div>
    </button>`);
  }

  // Droite = fait / pas fait · gauche = supprimer
  function swipeWrap(r, cls, inner) {
    return `<div class="swipe ${cls}" data-swipe="${r.r}">
      <button class="sw-act sw-done" data-sw="done" tabindex="-1">${r.done ? icon("undo") + "Pas fait" : icon("check") + "Fait"}</button>
      <button class="sw-act sw-del" data-sw="del" tabindex="-1">${icon("trash")}Supprimer</button>
      ${inner}
    </div>`;
  }

  function rowHtml(r, last) {
    const u = urgency(r);
    const sub = [r.color, r.client, r.obs].filter(Boolean).join(" · ");
    return swipeWrap(r, last ? "last" : "", `<button class="row sw-front" data-row="${r.r}" ${u ? `style="--u:var(--${u.cls === "unknown" ? "muted" : u.cls})"` : ""}>
      ${swatch(r.color)}
      <div class="row-main">
        <div class="row-title">${esc(r.name)}</div>
        <div class="row-sub">${esc(sub || fmt(r.date))}</div>
      </div>
      <div class="row-side">
        <span class="status ${r.done ? "done" : "todo"}">${r.done ? "Faite" : "À faire"}</span>
        ${payPill(r)}
      </div>
    </button>`);
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
    const sorted = Object.entries(byModel).sort((a, b) => b[1] - a[1] || a[0].localeCompare(b[0], "fr"));
    const top = state.topCount === "all" ? sorted : sorted.slice(0, 10);
    const tmax = top.length ? top[0][1] : 1;
    $("#modelsInfo").textContent = `${sorted.length} modèles différents · ${scope.length} commandes`;
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
    renderSync();
    renderHome();
    if (state.view === "history") filterHistory();
    if (state.view === "stats") renderStats();
    if (state.view === "moules") renderMoules();
  }

  // ---------- Rendu : Moules ----------
  // Un moule par modèle des commandes (par défaut : OK, chez Jean) + ceux ajoutés à la main.
  // Seuls les moules modifiés sont écrits dans la feuille "Moules".
  const LIEU_JEAN = "Chez Jean", LIEU_ATELIER = "À l'atelier";
  // Seuls les modèles déjà commandés à la mise en place des moules en créent un automatiquement :
  // un nom tapé librement dans une nouvelle commande n'ajoute pas de moule (on l'ajoute avec le +).
  const MOULES_DEPUIS = new Date(2026, 9, 24);
  function allMoules() {
    const names = {}, auto = new Set();
    for (const r of state.rows) {
      names[r.key] = r.name;
      if (!r.date || r.date <= MOULES_DEPUIS) auto.add(r.key);
    }
    const byKey = {};
    for (const x of state.moules) byKey[x.key] = { ...x, model: names[x.key] || x.model, lieu: x.lieu || LIEU_JEAN };
    for (const k of auto) if (!byKey[k]) byKey[k] = { r: null, model: names[k], key: k, etat: "ok", note: "", lieu: LIEU_JEAN, scale: "", virtual: true };
    const list = Object.values(byKey);
    for (const x of list) x.echelle = ECHELLES.includes(x.scale) ? x.scale : scaleFromName(x.model);
    return list;
  }
  const findMoule = (key) => allMoules().find((x) => x.key === key);

  // Enregistre un changement (état, note ou lieu) ; crée la ligne si le moule n'est pas encore dans la feuille
  function saveMoule(x, patch) {
    if (x.virtual) {
      return writeMoule(null, [realName(x.model), MOULE_LABEL[patch.etat || x.etat], patch.note ?? x.note, patch.lieu ?? x.lieu, patch.scale ?? x.scale ?? ""]);
    }
    return writeMoule(x.r, [null, patch.etat ? MOULE_LABEL[patch.etat] : null, patch.note ?? null, patch.lieu ?? null, patch.scale ?? null], x.model);
  }

  function mouleCard(x, waiting) {
    const w = waiting[x.key] || 0;
    return `<div class="mcard ${x.etat}" data-moule="${esc(x.key)}" role="button" tabindex="0">
      <span class="mstate">${icon(x.etat === "casse" ? "alert" : x.etat === "sechage" ? "clock" : "check")}</span>
      <span class="mmain"><span class="mname">${esc(x.model)}${x.echelle && x.echelle !== scaleFromName(x.model) ? ` <span class="mscale">${x.echelle}</span>` : ""}</span>${x.note ? `<span class="msub">${esc(x.note)}</span>` : ""}
        ${x.etat !== "ok" || w ? `<span class="mline">${x.etat !== "ok" ? `<span class="mbadge ${x.etat}">${MOULE_LABEL[x.etat]}</span>` : ""}${w ? `<span class="mwait">${w} cde${w > 1 ? "s" : ""} en attente</span>` : ""}</span>` : ""}</span>
      <span class="mplace ${x.lieu === LIEU_ATELIER ? "atelier" : ""}">${esc(x.lieu)}</span>
    </div>`;
  }

  function renderMoules() {
    const waiting = {};
    for (const r of state.rows) if (!r.done) waiting[r.key] = (waiting[r.key] || 0) + 1;
    const all = allMoules();
    const n = (e) => all.filter((x) => x.etat === e).length;
    $("#mouleSummary").innerHTML = `
      <div class="mstat casse"><b>${n("casse")}</b>Cassé${n("casse") > 1 ? "s" : ""}</div>
      <div class="mstat sechage"><b>${n("sechage")}</b>En séchage</div>
      <div class="mstat ok"><b>${n("ok")}</b>OK</div>`;
    // Recherche : chaque mot tapé doit être le début d'un mot du modèle
    const terms = fold(state.mq || "").split(/\s+/).filter(Boolean);
    const sc = state.mscale || "all";
    $$("#mouleScale .chip").forEach((c) => {
      const v = c.dataset.scale;
      c.classList.toggle("active", v === sc);
      const n = v === "all" ? all.length : all.filter((x) => (v === "none" ? !x.echelle : x.echelle === v)).length;
      c.querySelector("b").textContent = n;
    });
    const list = all.filter((x) => {
      if (sc !== "all" && (sc === "none" ? x.echelle : x.echelle !== sc)) return false;
      if (!terms.length) return true;
      const words = fold(x.model).split(/[\s/.-]+/);
      return terms.every((t) => words.some((w) => w.startsWith(t)) || fold(x.model).replace(/\s/g, "").startsWith(t));
    });
    const order = { casse: 0, sechage: 1, ok: 2 };
    const sortFn = (a, b) => order[a.etat] - order[b.etat] || (waiting[b.key] || 0) - (waiting[a.key] || 0) || a.model.localeCompare(b.model, "fr");
    const atelier = list.filter((x) => x.lieu === LIEU_ATELIER).sort(sortFn);
    const jean = list.filter((x) => x.lieu !== LIEU_ATELIER).sort(sortFn);
    const jeanFirst = LS.get("pc_mouleOrder", "atelier") === "jean";
    $("#mouleOrder").classList.toggle("recent", jeanFirst);
    const section = (title, items) => `
      <div class="msect"><h2>${title}</h2><span class="mcount">${items.length}</span></div>
      <div class="mlist">${items.length ? items.map((x) => mouleCard(x, waiting)).join("") : `<div class="mempty">Aucun moule${terms.length ? " trouvé" : ""}</div>`}</div>`;
    const parts = [section("À l'atelier", atelier), section("Chez Jean", jean)];
    if (jeanFirst) parts.reverse();
    $("#mouleList").innerHTML = parts.join("");
  }

  // ---------- Navigation ----------
  const TITLES = { home: "Atelier", history: "Historique", stats: "Statistiques", moules: "Moules" };
  function show(view) {
    state.view = view;
    $$(".view").forEach((v) => (v.hidden = v.id !== "view-" + view));
    $$(".tab").forEach((t) => t.classList.toggle("active", t.dataset.view === view));
    $("#pageTitle").textContent = TITLES[view];
    if (view === "history") filterHistory();
    if (view === "stats") renderStats();
    if (view === "moules") renderMoules();
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
    let statut = "Fabriquée";
    if (u) statut = u.cls === "unknown" ? "À faire" : u.left < 0 ? `En retard de ${-u.left} j` : `À faire · ${u.left} j restant${u.left > 1 ? "s" : ""}`;
    const payTxt = r.pay === "paid" ? "Payée" : r.pay === "unpaid" ? "Non payée" : "Non renseigné";
    openSheet(`
      <div class="detail-head">
        ${swatch(r.color, "lg")}
        <div><h3>${esc(r.name)}</h3><span class="status ${r.done ? "done" : "todo"}" ${u ? `style="--u:var(--${u.cls === "unknown" ? "muted" : u.cls})"` : ""}>${esc(statut)}</span></div>
      </div>
      ${mouleFlag(r.key) && !r.done ? `<div class="mwarn">${mouleFlag(r.key, true)}</div>` : ""}
      <div class="quick">
        <button class="qbtn ${r.done ? "on-done" : ""}" data-act="done">
          ${icon(r.done ? "check" : "tool", "q")}${r.done ? "Fabriquée" : "Marquer fabriquée"}<small>${r.done ? "appuyer pour annuler" : "en cours"}</small>
        </button>
        <button class="qbtn ${r.pay === "paid" ? "on-paid" : r.pay === "unpaid" ? "on-unpaid" : ""}" data-act="pay">
          ${icon("euro", "q")}${payTxt}<small>appuyer pour changer</small>
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
        // Même couleur touchée une 2e fois = on enlève la couleur
        if (c.dataset.color === r.color) save([null, null, null, "", null, null, null], "Couleur enlevée");
        else save([null, null, null, c.dataset.color, null, null, null], `Couleur : ${c.dataset.color}`);
        return;
      }
      const act = e.target.closest("[data-act]")?.dataset.act;
      if (act === "edit") return openForm(r);
      if (act === "done") return save([null, r.done ? "" : "F", null, null, null, null, null], r.done ? "Remise en cours" : `${r.name} fabriquée`, !r.done);
      if (act === "pay") return save([null, null, null, null, null, null, PAY_VALUE[r.pay === "paid" ? "unpaid" : "paid"]], r.pay === "paid" ? "Non payée" : "Payée");
      if (act === "note") return save([null, null, null, null, clean($("#qNote").value), null, null], "Note enregistrée");
    };
  }

  const COLOR_CHOICES = ["Noir", "Blanc", "Bleu", "Jaune", "Rouge", "Vert"];
  // Modèles connus, les plus commandés d'abord (la dernière année compte double)
  function knownModels() {
    const c = {}; const from = addDays(today(), -365);
    for (const x of state.rows) c[x.name] = (c[x.name] || 0) + (x.date && x.date > from ? 2 : 1);
    // modèles qui n'existent encore que comme moule (ajoutés avec le +)
    const known = new Set(state.rows.map((x) => x.key));
    for (const m of state.moules) if (!known.has(m.key) && !(m.model in c)) c[m.model] = 0.5;
    return Object.entries(c).sort((a, b) => b[1] - a[1]).map(([name]) => ({ name, words: fold(name).split(/[\s/.-]+/).filter(Boolean) }));
  }
  // "pi fin" -> chaque mot tapé doit être le début d'un mot du modèle
  function suggestFrom(models, q) {
    const terms = fold(q).split(/\s+/).filter(Boolean);
    if (!terms.length) return models.slice(0, 8);
    return models.filter((m) => terms.every((t) => m.words.some((w) => w.startsWith(t)) || fold(m.name).replace(/\s/g, "").startsWith(t)))
      .filter((m) => fold(m.name) !== fold(q))
      .slice(0, 8);
  }
  const suggestHtml = (list) => list.map((m) => `<button type="button" data-model="${esc(m.name)}">${esc(m.name)}</button>`).join("");

  function openForm(r) {
    const isNew = !r;
    const models = knownModels();
    const suggestModels = (q) => suggestFrom(models, q);
    const curColor = r ? r.color : "Noir";
    const isOther = r && r.color && !COLOR_CHOICES.includes(r.color);
    openSheet(`
      <h3>${isNew ? "Nouvelle commande" : "Modifier la commande"}</h3>
      <p class="muted" style="margin:0 0 14px">${isNew ? `Ajoutée à la ligne ${state.lastUsedRow + 1} du fichier` : `Ligne ${r.r} du fichier`}</p>
      <form class="form" id="orderForm" autocomplete="off">
        <div class="field"><label>Date de commande</label>
          <input type="date" name="date" value="${isoDate(r ? r.date || today() : today())}" required></div>
        <div class="field"><label>Modèle</label>
          <input name="model" placeholder="Tape les premières lettres : pi, max, cox…" value="${esc(r ? r.model : "")}" required autocomplete="off" autocapitalize="sentences">
          <div class="suggest" id="modelSuggest">${isNew ? suggestHtml(suggestModels("")) : ""}</div>
          <div id="modelMoule">${r && !r.done ? mouleFlag(r.key, true) : ""}</div>
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
    // Prévient dès que le modèle tapé a un moule cassé ou en séchage
    const checkMoule = () => { $("#modelMoule").innerHTML = mouleFlag(modelKey(form.model.value), true); };
    form.model.addEventListener("input", () => { $("#modelSuggest").innerHTML = suggestHtml(suggestModels(form.model.value)); checkMoule(); });
    let color = isOther ? "__other" : curColor;
    $("#sheetBody").onclick = (e) => {
      const m = e.target.closest("[data-model]");
      if (m) { form.model.value = m.dataset.model; $("#modelSuggest").innerHTML = ""; checkMoule(); return; }
      const c = e.target.closest("[data-color]");
      if (c) {
        // Même couleur touchée une 2e fois = aucune couleur
        color = c.classList.contains("active") ? "" : c.dataset.color;
        $$(".color-pick button", form).forEach((b) => b.classList.toggle("active", b === c && color !== ""));
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
      const values = [dateChanged ? dateToSerial(date) : null, done, realName(clean(form.model.value)), colorVal, clean(form.client.value), clean(form.obs.value), pay];
      try {
        await writeRow(isNew ? null : r.r, values, isNew ? null : r.model);
        closeSheet();
        toast(isNew ? "Commande ajoutée" : "Modifiée");
      } catch (err) { btn.disabled = false; toast("Échec : " + err.message, true); }
    };
  }

  // Fiche d'un moule : changer l'état en un appui
  function openMoule(key) {
    const x = findMoule(key);
    if (!x) return;
    const inOrders = state.rows.some((r) => r.key === x.key);
    const w = state.rows.filter((r) => !r.done && r.key === x.key).length;
    openSheet(`
      <div class="detail-head">
        <span class="mstate big ${x.etat}">${icon(x.etat === "casse" ? "alert" : x.etat === "sechage" ? "clock" : "check")}</span>
        <div><h3>${esc(x.model)}</h3><span class="mbadge ${x.etat}">${MOULE_LABEL[x.etat]}</span>
        ${w ? `<span class="muted"> · ${w} commande${w > 1 ? "s" : ""} en attente</span>` : ""}</div>
      </div>
      <div class="qlabel">État du moule</div>
      <div class="mchoice">
        <button class="${x.etat === "ok" ? "active ok" : ""}" data-etat="ok">${icon("check")}OK</button>
        <button class="${x.etat === "casse" ? "active casse" : ""}" data-etat="casse">${icon("alert")}Cassé</button>
        <button class="${x.etat === "sechage" ? "active sechage" : ""}" data-etat="sechage">${icon("clock")}Séchage</button>
      </div>
      <div class="qlabel">Où est le moule ?</div>
      <div class="mchoice lieu">
        ${LIEUX.map((l) => `<button data-lieu="${esc(l)}" class="${x.lieu === l ? "active lieu" : ""}">${esc(l)}</button>`).join("")}
      </div>
      <div class="qlabel">Échelle${!x.scale && x.echelle ? ` <span class="muted">· trouvée dans le nom</span>` : ""}</div>
      <div class="mchoice scale">
        ${ECHELLES.map((e) => `<button data-scale="${e}" class="${x.echelle === e ? "active lieu" : ""}">${e}</button>`).join("")}
      </div>
      <div class="qlabel">Note</div>
      <div class="qnote">
        <textarea id="mNote" placeholder="Ex : fissure côté capot, 2e moule en cours…">${esc(x.note)}</textarea>
        <button data-act="note">OK</button>
      </div>
      ${x.virtual ? "" : `<button class="btn ghost danger-text" data-act="del" style="width:100%">${icon(inOrders ? "undo" : "trash")}${inOrders ? "Remettre par défaut (OK, chez Jean)" : "Supprimer ce moule"}</button>`}`);

    const save = async (p, msg) => {
      try { await saveMoule(x, p); toast(msg); openMoule(x.key); }
      catch (err) { toast("Échec : " + err.message, true); }
    };
    $("#sheetBody").onclick = async (e) => {
      const scb = e.target.closest("[data-scale]");
      if (scb) {
        // même échelle touchée une 2e fois = on l'enlève
        const scale = scb.dataset.scale === x.scale ? "" : scb.dataset.scale;
        if (!scale && !x.scale) return;
        return save({ scale }, scale ? `Échelle ${scale}` : "Échelle enlevée");
      }
      const li = e.target.closest("[data-lieu]");
      if (li) {
        if (li.dataset.lieu === x.lieu) return;
        return save({ lieu: li.dataset.lieu }, li.dataset.lieu);
      }
      const b = e.target.closest("[data-etat]");
      if (b) {
        const etat = b.dataset.etat;
        if (etat === x.etat) return;
        return save({ etat }, etat === "ok" ? "Moule OK" : etat === "casse" ? "Moule cassé" : "Moule en séchage");
      }
      const act = e.target.closest("[data-act]")?.dataset.act;
      if (act === "note") return save({ note: clean($("#mNote").value) }, "Note enregistrée");
      if (act === "del") {
        if (!e.target.closest("[data-act]").classList.contains("confirm")) {
          const btn = e.target.closest("[data-act]");
          btn.classList.add("confirm"); btn.innerHTML = `${icon("trash")}Appuie encore pour confirmer`;
          return;
        }
        await deleteMoule(x); closeSheet(); toast(inOrders ? "Moule remis par défaut" : "Moule supprimé");
      }
    };
  }

  // Ajout d'un moule
  function openMouleForm() {
    const models = knownModels();
    let etat = "ok", lieu = LIEU_JEAN, scale = "";
    openSheet(`
      <h3>Nouveau moule</h3>
      <form class="form" id="mouleForm" autocomplete="off">
        <div class="field"><label>Modèle</label>
          <input name="model" placeholder="Tape les premières lettres : pi, max, cox…" required autocomplete="off" autocapitalize="sentences">
          <div class="suggest" id="mSuggest">${suggestHtml(suggestFrom(models, ""))}</div>
        </div>
        <div class="field"><label>État</label>
          <div class="mchoice">
            <button type="button" data-etat="ok" class="active ok">${icon("check")}OK</button>
            <button type="button" data-etat="casse">${icon("alert")}Cassé</button>
            <button type="button" data-etat="sechage">${icon("clock")}Séchage</button>
          </div>
        </div>
        <div class="field"><label>Échelle</label>
          <div class="mchoice scale">${ECHELLES.map((e) => `<button type="button" data-scale="${e}">${e}</button>`).join("")}</div>
        </div>
        <div class="field"><label>Où est le moule ?</label>
          <div class="mchoice lieu">${LIEUX.map((l) => `<button type="button" data-lieu="${esc(l)}" class="${l === LIEU_JEAN ? "active lieu" : ""}">${esc(l)}</button>`).join("")}</div>
        </div>
        <div class="field"><label>Note</label><textarea name="note" placeholder="Optionnel"></textarea></div>
        <button class="btn primary" type="submit">Ajouter le moule</button>
      </form>`);
    const f = $("#mouleForm");
    f.model.addEventListener("input", () => { $("#mSuggest").innerHTML = suggestHtml(suggestFrom(models, f.model.value)); });
    $("#sheetBody").onclick = (e) => {
      const m = e.target.closest("[data-model]");
      if (m) { f.model.value = m.dataset.model; $("#mSuggest").innerHTML = ""; return; }
      const scb = e.target.closest("[data-scale]");
      if (scb) {
        scale = scb.dataset.scale === scale ? "" : scb.dataset.scale;
        $$(".mchoice.scale button", f).forEach((x) => (x.className = x === scb && scale ? "active lieu" : ""));
        return;
      }
      const li = e.target.closest("[data-lieu]");
      if (li) {
        lieu = li.dataset.lieu;
        $$(".mchoice.lieu button", f).forEach((x) => (x.className = x === li ? "active lieu" : ""));
        return;
      }
      const b = e.target.closest("[data-etat]");
      if (b) {
        etat = b.dataset.etat;
        $$(".mchoice:not(.lieu):not(.scale) button", f).forEach((x) => (x.className = x === b ? "active " + etat : ""));
      }
    };
    f.onsubmit = async (e) => {
      e.preventDefault();
      const model = clean(f.model.value);
      const same = findMoule(modelKey(model));
      if (same) { toast("Ce moule est déjà dans la liste"); return openMoule(same.key); }
      await writeMoule(null, [realName(model), MOULE_LABEL[etat], clean(f.note.value), lieu, scale || scaleFromName(model)]);
      closeSheet();
      toast("Moule ajouté");
    };
  }

  // ---------- Outil de regroupement des modèles ----------
  function applyGroups() {
    loadUserGroups();
    buildRows(state.matrix);
    buildMoules(state.mmatrix);
    saveCache();
    renderAll();
  }

  function openGroupTool() {
    const sel = new Set();
    let q = "", naming = false, pickName = "";
    const entries = () => {
      const c = {}, names = {};
      for (const r of state.rows) { c[r.key] = (c[r.key] || 0) + 1; names[r.key] = r.name; }
      return Object.keys(c).map((key) => ({ key, name: names[key], n: c[key] })).sort((a, b) => b.n - a.n || a.name.localeCompare(b.name, "fr"));
    };
    openSheet(`
      <h3>Regrouper des modèles</h3>
      <p class="muted" style="margin:4px 0 12px">Coche les modèles qui sont en fait le même, puis choisis le nom à garder.
      Le tableau n'est jamais modifié : les regroupements sont gardés sur cet appareil.</p>
      <div class="search"><svg viewBox="0 0 24 24"><circle cx="11" cy="11" r="7"/><path d="m20 20-3.5-3.5"/></svg>
        <input id="gSearch" type="search" placeholder="Chercher un modèle…" autocomplete="off"></div>
      <div id="gBar" class="gbar"></div>
      <div id="gList" class="glist"></div>
      <div id="gDone"></div>`);

    const render = () => {
      const all = entries();
      const byKey = Object.fromEntries(all.map((e) => [e.key, e]));
      // Barre : modèles cochés + bouton Regrouper, ou choix du nom final
      const chosen = [...sel].map((k) => byKey[k]).filter(Boolean);
      $("#gBar").innerHTML = naming ? `
        <div class="qlabel">Nom à garder</div>
        <div class="gnames">${chosen.map((e) => `<button type="button" class="grow ${pickName === e.name ? "on" : ""}" data-gname="${esc(e.name)}">
          <span class="gbox">${pickName === e.name ? icon("check") : ""}</span><span class="gname">${esc(e.name)}</span><span class="gn">${e.n}</span></button>`).join("")}</div>
        <div class="btns two" style="margin-top:10px">
          <button class="btn ghost" data-gact="cancel">Annuler</button>
          <button class="btn primary" data-gact="save">Regrouper</button>
        </div>` : chosen.length ? `
        <div class="gsel">${chosen.map((e) => `<span class="gchip">${esc(e.name)} <b>${e.n}</b></span>`).join("")}</div>
        <button class="btn primary" data-gact="merge" ${chosen.length < 2 ? "disabled" : ""} style="width:100%">Regrouper ${chosen.length} modèles</button>` : "";
      // Liste des modèles (recherche : début des mots)
      const terms = fold(q).split(/\s+/).filter(Boolean);
      const list = all.filter((e) => !terms.length || terms.every((t) => fold(e.name).split(/[\s/.-]+/).some((w) => w.startsWith(t))));
      $("#gList").innerHTML = naming ? "" : list.slice(0, 120).map((e) => `
        <button type="button" class="grow ${sel.has(e.key) ? "on" : ""}" data-gkey="${esc(e.key)}">
          <span class="gbox">${sel.has(e.key) ? icon("check") : ""}</span>
          <span class="gname">${esc(e.name)}${e.key.startsWith("ug:") ? ` <span class="gtag">regroupé</span>` : ""}</span>
          <span class="gn">${e.n}</span>
        </button>`).join("") + (list.length > 120 ? `<div class="mempty">Affine la recherche pour voir les autres modèles</div>` : "");
      // Regroupements existants
      const groups = LS.get("pc_groups", []);
      $("#gDone").innerHTML = naming || !groups.length ? "" : `
        <div class="qlabel" style="margin-top:16px">Mes regroupements</div>
        ${groups.map((g, i) => `<div class="gdone"><div><b>${esc(g.name)}</b><span class="msub">${g.members.map((m) => esc(m.label)).join(" · ")}</span></div>
          <button class="btn ghost" data-gundo="${i}">${icon("undo")}Défaire</button></div>`).join("")}`;
    };

    const save = (name) => {
      name = clean(name);
      if (!name) return toast("Choisis un nom", true);
      let groups = LS.get("pc_groups", []);
      const members = [];
      for (const k of sel) {
        if (k.startsWith("ug:")) {
          // un regroupement déjà fait est fusionné dans le nouveau
          const g = groups.find((x) => "ug:" + fold(x.name) === k);
          if (g) { members.push(...g.members); groups = groups.filter((x) => x !== g); }
        } else {
          const r = state.rows.find((x) => x.key === k);
          members.push({ key: k, label: r ? r.name : k });
        }
      }
      groups = groups.filter((x) => fold(x.name) !== fold(name));
      groups.push({ name, members });
      LS.set("pc_groups", groups);
      sel.clear(); naming = false;
      applyGroups();
      toast(`Regroupé sous « ${name} »`);
      render();
    };

    $("#gSearch").oninput = (e) => { q = e.target.value; render(); };
    $("#sheetBody").onclick = (e) => {
      const row = e.target.closest("[data-gkey]");
      if (row) { const k = row.dataset.gkey; sel.has(k) ? sel.delete(k) : sel.add(k); return render(); }
      const nm = e.target.closest("[data-gname]");
      if (nm) { pickName = nm.dataset.gname; return render(); }
      const un = e.target.closest("[data-gundo]");
      if (un) {
        const groups = LS.get("pc_groups", []);
        const g = groups.splice(Number(un.dataset.gundo), 1)[0];
        LS.set("pc_groups", groups);
        applyGroups(); render();
        return toast(`« ${g.name} » défait`);
      }
      const act = e.target.closest("[data-gact]")?.dataset.gact;
      if (act === "merge") {
        naming = true;
        // par défaut : le nom le plus utilisé parmi les modèles cochés
        const e0 = entries().filter((e) => sel.has(e.key))[0];
        pickName = e0 ? e0.name : "";
        return render();
      }
      if (act === "cancel") { naming = false; return render(); }
      if (act === "save") return save(pickName);
    };
    render();
  }

  function openSettings() {
    const c = cfg();
    openSheet(`
      <h3>Réglages</h3>
      <p class="muted" style="margin:0 0 14px">${state.source === "sheet" ? "Connecté au tableau Google Sheets" : state.source === "file" ? `Fichier local : ${esc(state.fileName)}` : "Aucune source"}</p>
      <form class="form" id="setForm" autocomplete="off">
        <div class="field"><label>Thème</label>
          <div class="seg full" id="themeSeg">
            ${[["dark", "moon", "Sombre"], ["light", "sun", "Clair"], ["auto", "auto", "Auto"]].map(([v, ic, l]) => `<button type="button" data-theme="${v}" class="${LS.get("pc_theme", "dark") === v ? "active" : ""}">${icon(ic)}${l}</button>`).join("")}
          </div>
        </div>
        <div class="field"><label>Délai de fabrication (jours)</label><input name="delai" type="number" min="1" max="90" value="${c.delai}"></div>
        <div class="field"><label>Adresse du script Google</label><input name="url" placeholder="https://script.google.com/macros/s/…/exec" value="${esc(c.scriptUrl)}"></div>
        <button class="btn primary" type="submit">Enregistrer</button>
      </form>
      <div class="btns two" style="margin-top:12px">
        <button class="btn ghost" id="btnGroups" type="button" style="grid-column:1/-1">Regrouper des modèles</button>
        <label class="btn ghost file-btn">Ouvrir un fichier<input type="file" id="fileInput2" accept=".xlsx,.xlsm,.xls"></label>
        ${c.key ? `<button class="btn ghost" id="btnForget">Oublier le code</button>` : `<button class="btn ghost" id="btnCode2">Code d'accès</button>`}
      </div>`);
    const f = $("#setForm");
    // Le thème s'applique tout de suite, sans attendre "Enregistrer"
    $("#themeSeg").onclick = (e) => {
      const b = e.target.closest("[data-theme]"); if (!b) return;
      LS.set("pc_theme", b.dataset.theme);
      if (window.applyTheme) window.applyTheme();
      $$("#themeSeg button").forEach((x) => x.classList.toggle("active", x === b));
    };
    f.onsubmit = async (e) => {
      e.preventDefault();
      LS.set("pc_delai", Math.max(1, Number(f.delai.value) || 10));
      LS.set("pc_scriptUrl", clean(f.url.value));
      closeSheet();
      renderAll();
      toast("Réglages enregistrés");
      if (state.source === "sheet") refresh(true);
    };
    $("#btnGroups").onclick = openGroupTool;
    $("#fileInput2").onchange = (e) => { if (e.target.files[0]) { closeSheet(); openLocalFile(e.target.files[0]); } };
    const fg = $("#btnForget");
    if (fg) fg.onclick = () => { LS.del("pc_key"); LS.del("pc_cache"); LS.del("pc_queue"); location.reload(); };
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
      if (await refresh(true)) toast(`Connecté · ${state.rows.length} commandes`);
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
    if (state.view === "moules") return openMouleForm();
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
    const act = e.target.closest(".sw-act");
    if (act) {
      const w = act.closest(".swipe");
      const r = state.rows.find((x) => x.r === Number(w.dataset.swipe));
      closeSwipe();
      if (!r) return;
      if (act.dataset.sw === "del") return confirmDelete(r);
      writeRow(r.r, [null, r.done ? "" : "F", null, null, null, null, null], r.model)
        .then(() => toast(r.done ? "Remise en cours" : `${r.name} fabriquée`))
        .catch((err) => toast("Échec : " + err.message, true));
      return;
    }
    const w = e.target.closest(".swipe");
    if (w && (Date.now() - (w._swipedAt || 0) < 400)) return; // fin d'un glisser, pas un appui
    if (w && w === swOpen) return closeSwipe();
    const mo = e.target.closest("[data-moule]");
    if (mo) return openMoule(mo.dataset.moule);
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
  let mSearchTimer;
  $("#mouleSearch").oninput = (e) => { clearTimeout(mSearchTimer); mSearchTimer = setTimeout(() => { state.mq = e.target.value; renderMoules(); }, 120); };
  $("#mouleScale").onclick = (e) => {
    const b = e.target.closest("[data-scale]"); if (!b) return;
    state.mscale = b.dataset.scale;
    renderMoules();
  };
  $("#mouleOrder").onclick = () => {
    const next = LS.get("pc_mouleOrder", "atelier") === "jean" ? "atelier" : "jean";
    LS.set("pc_mouleOrder", next);
    renderMoules();
    toast(next === "jean" ? "Chez Jean en haut" : "Atelier en haut");
  };
  $("#todoSort").onclick = () => {
    const next = LS.get("pc_todoSort", "old") === "new" ? "old" : "new";
    LS.set("pc_todoSort", next);
    renderHome();
    toast(next === "new" ? "Plus récentes en haut" : "Plus anciennes en haut");
  };
  $("#btnGroupStats").onclick = openGroupTool;
  $("#topCount").onclick = (e) => {
    const b = e.target.closest("[data-count]"); if (!b) return;
    $$("#topCount button").forEach((c) => c.classList.toggle("active", c === b));
    state.topCount = b.dataset.count; renderStats();
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
      if (cache && cache.matrix) { state.syncedAt = cache.at; buildRows(cache.matrix); buildMoules(cache.mmatrix); }
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
