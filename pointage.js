/* =====================================================================
   RSG — onglet « Pointage RSG »
   • PC (grand écran) : la feuille « Pointage journalier » (une ligne par heure,
     RSG + 3 personnes, chantier / travail effectué) à remplir comme dans Excel,
     et la feuille « Pointage mensuel » calculée toute seule.
   • Téléphone : un menu guidé — qui, quel chantier, quel travail, quelles heures.
   Données : collection « pointage », une fiche par personne et par jour
   (id = date_colonne, ex. 2026-10-28_1), avec les heures { "8": {c, t} }.
   ===================================================================== */
"use strict";
(function () {
  const COL = "pointage", NB = 4, HOURS = [...Array(24).keys()], SANS = "Sans chantier indiqué";
  const LOGO = () => window.RSG_LOGO_DEVIS || "logo-devis.png";
  const pad = n => String(n).padStart(2, "0");
  const hr = h => `${pad(h)}h – ${pad(h + 1)}h`;
  const isoToday = () => { const d = new Date(); return new Date(d.getTime() - d.getTimezoneOffset() * 6e4).toISOString().slice(0, 10); };
  const addDays = (iso, n) => { const d = new Date(iso + "T12:00:00"); d.setDate(d.getDate() + n); return new Date(d.getTime() - d.getTimezoneOffset() * 6e4).toISOString().slice(0, 10); };
  const addMonths = (ym, n) => { let [y, m] = ym.split("-").map(Number); m += n; while (m < 1) { m += 12; y--; } while (m > 12) { m -= 12; y++; } return `${y}-${pad(m)}`; };
  const cap = s => s.charAt(0).toUpperCase() + s.slice(1);
  const frLong = iso => cap(new Date(iso + "T12:00:00").toLocaleDateString("fr-FR", { weekday: "long", day: "numeric", month: "long", year: "numeric" }));
  const frShort = iso => cap(new Date(iso + "T12:00:00").toLocaleDateString("fr-FR", { weekday: "long", day: "numeric", month: "long" }));
  const frMonth = ym => cap(new Date(ym + "-15T12:00:00").toLocaleDateString("fr-FR", { month: "long", year: "numeric" }));
  const mobile = () => matchMedia("(max-width: 760px)").matches;
  const clone = o => JSON.parse(JSON.stringify(o));
  const tr = s => String(s ?? "").trim();
  const key = s => norm(tr(s));
  const filled = e => !!(e && (tr(e.c) || tr(e.t)));
  const H = n => (n ? n + " h" : "–");
  const jours = n => n + " j";
  const initials = n => (tr(n).split(/\s+/).map(w => w[0]).join("").slice(0, 2) || "?").toUpperCase();
  const relDay = iso => { const t = isoToday(); return iso === t ? "Aujourd'hui" : iso === addDays(t, -1) ? "Hier" : iso === addDays(t, 1) ? "Demain" : ""; };

  const tab = {
    id: "pointage",
    title: "Pointage RSG",
    collection: COL,
    el: null, mode: null, view: "jour", date: isoToday(), month: isoToday().slice(0, 7), built: "",
    pending: new Map(), timers: new Map(), busy: new Set(), wz: null,

    /* ---------- données ---------- */
    doc(date, s) {
      const id = `${date}_${s}`;
      return this.pending.get(id) || store.get(COL, id) || { id, date, slot: s, nom: s ? "" : "RSG", h: {} };
    },
    day(date) { return [0, 1, 2, 3].map(s => this.doc(date, s)); },
    name(d) { return d.slot === 0 ? "RSG" : tr(d.nom); },
    label(d) { return this.name(d) || `Personne ${d.slot + 1}`; },
    all() {
      const m = new Map(store.list(COL).map(d => [d.id, d]));
      for (const [id, d] of this.pending) m.set(id, d);
      return [...m.values()];
    },
    hoursOf(d) { return Object.keys(d.h || {}).map(Number).filter(h => filled(d.h[h])).sort((a, b) => a - b); },

    detail(date) {
      const docs = this.day(date), rows = new Map(), per = [0, 0, 0, 0], sans = [0, 0, 0, 0];
      docs.forEach((d, s) => {
        for (const h of this.hoursOf(d)) {
          const e = d.h[h], c = tr(e.c); per[s]++;
          if (!c) { sans[s]++; continue; }
          if (!rows.has(key(c))) rows.set(key(c), { nom: c, v: [0, 0, 0, 0] });
          rows.get(key(c)).v[s]++;
        }
      });
      return { docs, rows: [...rows.values()], per, sans, total: per.reduce((a, b) => a + b, 0) };
    },

    monthStats(ym) {
      const pers = new Map(), ch = new Map(); let total = 0;
      const add = (m, nom, date, n) => {
        const k = key(nom); if (!m.has(k)) m.set(k, { nom, h: 0, j: new Set() });
        const x = m.get(k); x.h += n; x.j.add(date);
      };
      const docs = this.all().filter(d => String(d.date || "").startsWith(ym))
        .sort((a, b) => String(a.date).localeCompare(String(b.date)) || a.slot - b.slot);
      for (const d of docs) {
        const hs = this.hoursOf(d); if (!hs.length) continue;
        add(pers, this.label(d), d.date, hs.length); total += hs.length;
        for (const h of hs) add(ch, tr(d.h[h].c) || SANS, d.date, 1);
      }
      const out = m => [...m.values()].map(x => ({ nom: x.nom, h: x.h, j: x.j.size }));
      return { pers: out(pers), ch: out(ch), total, days: new Set(docs.filter(d => this.hoursOf(d).length).map(d => d.date)).size };
    },

    personStats(ym, nom) {
      const ch = new Map(), days = []; let total = 0;
      const docs = this.all().filter(d => String(d.date || "").startsWith(ym) && key(this.label(d)) === key(nom))
        .sort((a, b) => String(a.date).localeCompare(String(b.date)));
      for (const d of docs) {
        const bl = this.blocks(d); if (!bl.length) continue;
        days.push({ date: d.date, blocks: bl });
        for (const b of bl) {
          const c = b.c || SANS, k = key(c);
          if (!ch.has(k)) ch.set(k, { nom: c, h: 0, j: new Set() });
          const x = ch.get(k); x.h += b.hours.length; x.j.add(d.date); total += b.hours.length;
        }
      }
      return { ch: [...ch.values()].map(x => ({ nom: x.nom, h: x.h, j: x.j.size })), days, total, nbJours: days.length };
    },
    // Synthèse d'un chantier : sur un mois (ym) ou depuis le début (ym = null)
    chantierStats(ym, nom) {
      const pers = new Map(), byDay = new Map(); let total = 0;
      const docs = this.all().filter(d => !ym || String(d.date || "").startsWith(ym))
        .sort((a, b) => String(a.date).localeCompare(String(b.date)) || a.slot - b.slot);
      for (const d of docs) {
        for (const b of this.blocks(d)) {
          if (key(b.c || SANS) !== key(nom)) continue;
          const pn = this.label(d), k = key(pn), n = b.hours.length;
          if (!pers.has(k)) pers.set(k, { nom: pn, h: 0, j: new Set() });
          const x = pers.get(k); x.h += n; x.j.add(d.date); total += n;
          if (!byDay.has(d.date)) byDay.set(d.date, []);
          byDay.get(d.date).push({ pers: pn, from: b.from, to: b.to, t: b.t, h: n });
        }
      }
      const days = [...byDay.entries()].sort((a, b) => a[0].localeCompare(b[0])).map(([date, rows]) => ({ date, rows }));
      return { pers: [...pers.values()].map(x => ({ nom: x.nom, h: x.h, j: x.j.size })), days, total, nbJours: days.length,
        first: days[0]?.date, last: days.at(-1)?.date };
    },
    chantierSheet(ym, nom, extra = "") {
      const c = this.chantierStats(ym, nom), fr = iso => new Date(iso + "T12:00:00").toLocaleDateString("fr-FR");
      const per = ym ? frMonth(ym) : "Depuis le début du chantier";
      return `<div class="pt-sheet pt-msheet" ${extra}>
        <div class="pt-title"><img src="${LOGO()}" alt="RSG"><div><b>SYNTHÈSE CHANTIER</b><span>${esc(nom)} · ${per}</span></div></div>
        <h2>1. Heures par personne</h2>
        <table class="pt-sum"><thead><tr><th>Personne</th><th>Heures totales</th><th>Jours</th></tr></thead>
          <tbody>${c.pers.length ? c.pers.map(x => `<tr><td>${esc(x.nom)}</td><td class="b">${H(x.h)}</td><td>${jours(x.j)}</td></tr>`).join("")
            : `<tr><td colspan="3" class="empty">Aucune heure sur ce chantier pour cette période.</td></tr>`}</tbody></table>
        <h2>2. Détail par jour</h2>
        <table class="pt-sum pt-pday"><thead><tr><th>Date</th><th>Personne</th><th>Horaires</th><th>Travail effectué</th><th>Heures</th></tr></thead>
          <tbody>${c.days.length ? c.days.map(d => d.rows.map((r, i) => `<tr class="${i ? "cont" : ""}">
              <td>${i ? "" : esc(cap(new Date(d.date + "T12:00:00").toLocaleDateString("fr-FR", { weekday: "short", day: "numeric", month: "short", ...(ym ? {} : { year: "2-digit" }) })))}</td>
              <td>${esc(r.pers)}</td><td>${pad(r.from)}h – ${pad(r.to)}h</td><td>${esc(r.t)}</td><td class="b">${r.h} h</td></tr>`).join("")).join("")
            : `<tr><td colspan="5" class="empty">—</td></tr>`}</tbody></table>
        <h2>3. Total des heures sur le chantier</h2>
        <div class="pt-total"><span>Total ${esc(nom)}${c.first ? ` <small style="font-weight:400;opacity:.8">· du ${fr(c.first)} au ${fr(c.last)}</small>` : ""}</span><b>${c.total} h · ${jours(c.nbJours)}</b></div>
      </div>`;
    },
    personSheet(ym, nom, extra = "") {
      const p = this.personStats(ym, nom);
      return `<div class="pt-sheet pt-msheet" ${extra}>
        <div class="pt-title"><img src="${LOGO()}" alt="RSG"><div><b>RÉCAPITULATIF DES HEURES</b><span>${esc(nom)} · ${frMonth(ym)}</span></div></div>
        <h2>1. Heures par chantier</h2>
        <table class="pt-sum"><thead><tr><th>Chantier</th><th>Heures totales</th><th>Jours</th></tr></thead>
          <tbody>${p.ch.length ? p.ch.map(x => `<tr><td>${esc(x.nom)}</td><td class="b">${H(x.h)}</td><td>${jours(x.j)}</td></tr>`).join("")
            : `<tr><td colspan="3" class="empty">Aucune heure pointée ce mois-ci.</td></tr>`}</tbody></table>
        <h2>2. Détail par jour</h2>
        <table class="pt-sum pt-pday"><thead><tr><th>Date</th><th>Horaires</th><th>Chantier</th><th>Travail effectué</th><th>Heures</th></tr></thead>
          <tbody>${p.days.length ? p.days.map(d => d.blocks.map((b, i) => `<tr class="${i ? "cont" : ""}">
              <td>${i ? "" : esc(cap(new Date(d.date + "T12:00:00").toLocaleDateString("fr-FR", { weekday: "short", day: "numeric", month: "short" })))}</td>
              <td>${pad(b.from)}h – ${pad(b.to)}h</td><td>${esc(b.c || SANS)}</td><td>${esc(b.t)}</td><td class="b">${b.hours.length} h</td></tr>`).join("")).join("")
            : `<tr><td colspan="5" class="empty">—</td></tr>`}</tbody></table>
        <h2>3. Total des heures du mois</h2>
        <div class="pt-total"><span>Total des heures de ${esc(nom)}</span><b>${p.total} h · ${jours(p.nbJours)}</b></div>
      </div>`;
    },

    known() {
      const docs = this.all().sort((a, b) => String(b.date).localeCompare(String(a.date)));
      const pe = new Map(), ch = new Map(), trv = new Map(), allT = new Map();
      for (const d of docs) {
        const n = this.name(d);
        if (d.slot && n && !pe.has(key(n))) pe.set(key(n), n);
        for (const h of this.hoursOf(d)) {
          const c = tr(d.h[h].c), t = tr(d.h[h].t);
          if (c && !ch.has(key(c))) ch.set(key(c), c);
          if (t) {
            if (!allT.has(key(t))) allT.set(key(t), t);
            const m = trv.get(key(c)) || new Map(); if (!m.has(key(t))) m.set(key(t), t); trv.set(key(c), m);
          }
        }
      }
      return { pers: [...pe.values()], ch: [...ch.values()], trv: c => [...(trv.get(key(c))?.values() || [])], allT: [...allT.values()] };
    },

    /* ---------- enregistrement (une écriture regroupée par personne et par jour) ---------- */
    edit(date, s, fn, now) {
      if (!canEdit()) return false;
      const d = clone(this.doc(date, s)); d.h ||= {};
      fn(d);
      for (const h of Object.keys(d.h)) if (!filled(d.h[h])) delete d.h[h];
      this.pending.set(d.id, d); this.status();
      clearTimeout(this.timers.get(d.id));
      this.timers.set(d.id, setTimeout(() => this.write(d.id), now ? 0 : 600));
      return true;
    },
    async write(id) {
      this.timers.delete(id);
      if (this.busy.has(id)) { this.timers.set(id, setTimeout(() => this.write(id), 300)); return; }
      const d = this.pending.get(id); if (!d) return;
      this.busy.add(id);
      const body = { date: d.date, slot: d.slot, nom: d.nom || "", h: d.h };
      try {
        if (store.get(COL, id)) await store.patch(COL, id, body);
        else { try { await store.create(COL, body, id); } catch { await store.patch(COL, id, body); } }
        if (this.pending.get(id) === d) this.pending.delete(id);
      } catch (e) { toast(frError(e)); }
      finally { this.busy.delete(id); this.status(); if (!this.timers.size) this.render(); }
    },
    flushAll() { for (const [id, t] of [...this.timers]) { clearTimeout(t); this.write(id); } },
    status() {
      const st = this.el?.querySelector("#ptSt"); if (!st) return;
      const wait = this.pending.size > 0;
      st.classList.toggle("pending", wait);
      st.querySelector("span").textContent = wait ? "Enregistrement…" : "Enregistré";
    },

    /* ---------- affichage ---------- */
    mount(el) {
      this.el = el; this.mode = mobile() ? "m" : "pc"; this.built = "";
      if (!this.mq) {
        this.mq = true;
        matchMedia("(max-width: 760px)").addEventListener("change", () => { if (current === this) this.mount(this.el); });
        window.addEventListener("pagehide", () => this.flushAll());
      }
      if (this.mode === "pc") this.mountPc(); else this.mountM();
    },
    render() {
      if (!this.el || current !== this) return;
      if ((mobile() ? "m" : "pc") !== this.mode) return this.mount(this.el);
      if (this.mode === "pc") this.renderPc(); else if (!this.wz) this.renderM();
    },

    /* ================= PC : feuilles comme dans Excel ================= */
    mountPc() {
      this.el.innerHTML = `<div class="pt pt-pc">
        <div class="pt-head">
          <h1>Pointage RSG</h1>
          <div class="pt-seg" role="tablist">
            <button data-v="jour">Pointage journalier</button>
            <button data-v="mois">Pointage mensuel</button>
          </div>
          <span class="pt-st" id="ptSt"><i></i><span>Enregistré</span></span>
        </div>
        <div id="ptBody"></div>
        <datalist id="ptCh"></datalist><datalist id="ptTr"></datalist><datalist id="ptPers"></datalist>
      </div>`;
      const el = this.el;
      el.querySelector(".pt-seg").onclick = e => {
        const b = e.target.closest("[data-v]"); if (!b) return;
        this.view = b.dataset.v; if (this.view === "mois") this.month = this.date.slice(0, 7);
        this.person = null; this.chantier = null; this.built = ""; this.renderPc();
      };
      if (!this.pcBound) this.bindPc(el);
      this.renderPc();
    },

    bindPc(el) {
      this.pcBound = true;
      el.addEventListener("click", e => {
        if (current !== this || this.mode !== "pc") return;
        const n = e.target.closest("[data-nav]");
        if (n) { if (this.view === "jour") this.date = addDays(this.date, +n.dataset.nav); else this.month = addMonths(this.month, +n.dataset.nav); return this.renderPc(); }
        const ml = e.target.closest("[data-mail]"); if (ml) return this.sendMail(ml.dataset.mail);
        if (e.target.closest("[data-today]")) { this.date = isoToday(); this.month = this.date.slice(0, 7); return this.renderPc(); }
        const pr = e.target.closest("[data-person]"); if (pr) { this.person = pr.dataset.person; this.chantier = null; return this.renderPc(); }
        const chx = e.target.closest("[data-chantier]"); if (chx) { this.chantier = chx.dataset.chantier; this.chScope = "mois"; this.person = null; return this.renderPc(); }
        const sc = e.target.closest("[data-scope]"); if (sc) { this.chScope = sc.dataset.scope; return this.renderPc(); }
        if (e.target.closest("[data-person-back]")) { this.person = null; this.chantier = null; return this.renderPc(); }
        const dj = e.target.closest("[data-day]");
        if (dj) { this.date = dj.dataset.day; this.view = "jour"; return this.renderPc(); }
      });
      el.addEventListener("change", e => {
        if (current !== this) return;
        if (e.target.id === "ptDate" && e.target.value) { this.date = e.target.value; this.renderPc(); }
      });
      el.addEventListener("input", e => {
        const i = e.target; if (current !== this || !i.matches(".pt-grid input")) return;
        const s = +i.dataset.s, v = i.value;
        const ok = this.edit(this.date, s, d => {
          if (i.dataset.k === "nom") d.nom = v;
          else { const h = +i.dataset.h; d.h[h] = { ...(d.h[h] || {}), [i.dataset.k]: v }; }
        });
        if (!ok) { this.fillGrid(true); return; }
        if (i.dataset.h != null) {
          const e1 = this.doc(this.date, s).h[i.dataset.h];
          el.querySelectorAll(`.pt-grid input[data-s="${s}"][data-h="${i.dataset.h}"]`).forEach(x => x.closest("td").classList.toggle("on", filled(e1)));
        }
        this.fillDetail();
      });
      el.addEventListener("keydown", e => {
        const i = e.target; if (current !== this || !i.matches(".pt-grid tbody input")) return;
        const go = dh => { const n = el.querySelector(`.pt-grid input[data-s="${i.dataset.s}"][data-k="${i.dataset.k}"][data-h="${+i.dataset.h + dh}"]`); if (n) { n.focus(); n.select(); } };
        if (e.key === "Enter") { e.preventDefault(); go(e.shiftKey ? -1 : 1); }
        else if ((e.ctrlKey || e.metaKey) && (e.key === "d" || e.key === "D")) {
          e.preventDefault();
          const up = el.querySelector(`.pt-grid input[data-s="${i.dataset.s}"][data-k="${i.dataset.k}"][data-h="${+i.dataset.h - 1}"]`);
          if (up) { i.value = up.value; i.dispatchEvent(new Event("input", { bubbles: true })); go(1); }
        }
      });
    },

    renderPc() {
      const body = this.el.querySelector("#ptBody"); if (!body) return;
      this.el.querySelectorAll(".pt-seg button").forEach(b => b.setAttribute("aria-selected", b.dataset.v === this.view));
      const k = this.known();
      const opts = a => a.slice(0, 60).map(v => `<option value="${esc(v)}"></option>`).join("");
      this.el.querySelector("#ptCh").innerHTML = opts(k.ch);
      this.el.querySelector("#ptTr").innerHTML = opts(k.allT);
      this.el.querySelector("#ptPers").innerHTML = opts(k.pers);
      this.status();

      if (this.view === "mois") { this.built = ""; body.innerHTML = this.monthPc(); return; }
      const sig = "jour|" + this.date;
      if (this.built !== sig) {
        this.built = sig;
        const rel = relDay(this.date);
        body.innerHTML = `
          <div class="pt-bar">
            <button class="btn" data-nav="-1" aria-label="Jour précédent">‹</button>
            <input type="date" id="ptDate" value="${this.date}" aria-label="Date du pointage">
            <button class="btn" data-nav="1" aria-label="Jour suivant">›</button>
            <button class="btn" data-today ${rel === "Aujourd'hui" ? "disabled" : ""}>Aujourd'hui</button>
            <button class="btn primary" data-mail="jour">Envoyer par mail</button>
            <span class="pt-hint">Cases jaunes à remplir · <kbd>Entrée</kbd> case du dessous · <kbd>Ctrl</kbd>+<kbd>D</kbd> recopier la case du dessus</span>
          </div>
          <div class="pt-sheet">
            <div class="pt-title">
              <img src="${LOGO()}" alt="RSG">
              <div><b>POINTAGE JOURNALIER</b><span>${frLong(this.date)}${rel ? ` · ${rel}` : ""}</span></div>
            </div>
            <table class="pt-grid">
              <colgroup><col class="hcol">${"<col class=\"cc\"><col class=\"tc\">".repeat(NB)}</colgroup>
              <thead>
                <tr><th rowspan="2" class="hcell">Heure</th>
                  <th colspan="2" class="p0">RSG</th>
                  ${[1, 2, 3].map(s => `<th colspan="2" class="pn"><input data-s="${s}" data-k="nom" list="ptPers" placeholder="Nom de la personne ${s + 1}" aria-label="Nom de la personne ${s + 1}" autocomplete="off"></th>`).join("")}
                </tr>
                <tr>${"<th>Chantier</th><th>Travail effectué</th>".repeat(NB)}</tr>
              </thead>
              <tbody>${HOURS.map(h => `<tr class="${h < 6 || h > 20 ? "night" : ""}"><th>${hr(h)}</th>${[0, 1, 2, 3].map(s =>
                `<td class="c"><input data-s="${s}" data-h="${h}" data-k="c" list="ptCh" autocomplete="off" aria-label="Chantier ${hr(h)}"></td><td class="t"><input data-s="${s}" data-h="${h}" data-k="t" list="ptTr" autocomplete="off" aria-label="Travail ${hr(h)}"></td>`).join("")}</tr>`).join("")}
              </tbody>
            </table>
            <table class="pt-det" id="ptDet"></table>
          </div>`;
      }
      this.fillGrid(); this.fillDetail();
    },

    fillGrid(force) {
      const docs = this.day(this.date), act = document.activeElement;
      this.el.querySelectorAll(".pt-grid input").forEach(i => {
        const d = docs[+i.dataset.s];
        const v = i.dataset.k === "nom" ? (d.nom || "") : (d.h?.[i.dataset.h]?.[i.dataset.k] || "");
        if ((i !== act || force) && i.value !== v) i.value = v;
        if (i.dataset.k !== "nom") i.closest("td").classList.toggle("on", filled(d.h?.[i.dataset.h]));
      });
    },
    fillDetail() {
      const t = this.el.querySelector("#ptDet"); if (!t) return;
      t.innerHTML = this.detailInner(this.date);
    },
    detailInner(date) {
      const { docs, rows, per, sans, total } = this.detail(date);
      const tdv = (v, s) => `<td class="${s === 0 ? "r" : ""}">${H(v)}</td>`;
      return `<thead><tr><th>Détail des heures par chantier</th>${docs.map((d, s) => `<th class="${s ? "" : "p0"}">${esc(this.label(d))}</th>`).join("")}<th>Total chantier</th></tr></thead>
        <tbody>${rows.length ? rows.map(r => `<tr><td>${esc(r.nom)}</td>${r.v.map(tdv).join("")}<td class="tot">${H(r.v.reduce((a, b) => a + b, 0))}</td></tr>`).join("")
          : `<tr><td colspan="6" class="empty">Les chantiers s'ajoutent ici dès qu'une heure est remplie.</td></tr>`}
          <tr class="sans"><td>${SANS}</td>${sans.map(v => `<td>${H(v)}</td>`).join("")}<td>${H(sans.reduce((a, b) => a + b, 0))}</td></tr>
        </tbody>
        <tfoot><tr><td>Heures totales</td>${per.map((v, s) => `<td class="${s ? "" : "p0"}">${H(v)}</td>`).join("")}<td class="big">${H(total)}</td></tr></tfoot>`;
    },

    monthSheet(ym, extra = "", click = false) {
      const m = this.monthStats(ym);
      const table = (titre, col, list, vide, pers) => `<h2>${titre}</h2>
        <table class="pt-sum"><thead><tr><th>${col}</th><th>Heures totales</th><th>Jours</th></tr></thead>
        <tbody>${list.length ? list.map(x => click
          ? `<tr class="click" ${pers ? "data-person" : "data-chantier"}="${esc(x.nom)}" title="Voir le détail de ${esc(x.nom)}"><td><span class="plink">${esc(x.nom)}</span><i class="go">Détail ›</i></td><td class="b">${H(x.h)}</td><td>${jours(x.j)}</td></tr>`
          : `<tr><td>${esc(x.nom)}</td><td class="b">${H(x.h)}</td><td>${jours(x.j)}</td></tr>`).join("")
          : `<tr><td colspan="3" class="empty">${vide}</td></tr>`}</tbody></table>`;
      return `<div class="pt-sheet pt-msheet" ${extra}>
            <div class="pt-title"><img src="${LOGO()}" alt="RSG"><div><b>POINTAGE MENSUEL</b><span>${frMonth(ym)}</span></div></div>
            ${table("1. Heures totales par personne", "Personne", m.pers, "Aucune heure pointée ce mois-ci.", true)}
            ${click && m.pers.length ? `<p class="pt-note noprint">Cliquez sur le nom d'une personne pour voir ses heures sur chaque chantier.</p>` : ""}
            ${table("2. Heures totales réalisées par chantier", "Chantier", m.ch, "Aucun chantier ce mois-ci.")}
            ${click && m.ch.length ? `<p class="pt-note noprint">Cliquez sur un chantier pour voir qui y a travaillé et combien de temps.</p>` : ""}
            <h2>3. Total des heures du mois</h2>
            <div class="pt-total"><span>Total des heures du mois</span><b>${m.total} h</b></div>
            <p class="pt-note">Un même chantier ou une même personne doit toujours être écrit de la même façon pour être additionné sur la même ligne.</p>
          </div>`;
    },
    dayHtml(date) {
      const docs = this.day(date);
      return `<div class="pt pt-pdf"><div class="pt-sheet" style="width:1120px">
        <div class="pt-title"><img src="${LOGO()}" alt="RSG"><div><b>POINTAGE JOURNALIER</b><span>${frLong(date)}</span></div></div>
        <table class="pt-grid">
          <colgroup><col class="hcol">${"<col class=\"cc\"><col class=\"tc\">".repeat(NB)}</colgroup>
          <thead><tr><th rowspan="2" class="hcell">Heure</th><th colspan="2" class="p0">RSG</th>
            ${[1, 2, 3].map(s => `<th colspan="2" class="pn">${esc(this.name(docs[s]) || "")}</th>`).join("")}</tr>
            <tr>${"<th>Chantier</th><th>Travail effectué</th>".repeat(NB)}</tr></thead>
          <tbody>${HOURS.map(h => `<tr class="${h < 6 || h > 20 ? "night" : ""}"><th>${hr(h)}</th>${docs.map(d => { const e = d.h?.[h], on = filled(e) ? " on" : "";
            return `<td class="c${on}">${esc(tr(e?.c))}</td><td class="t${on}">${esc(tr(e?.t))}</td>`; }).join("")}</tr>`).join("")}</tbody>
        </table>
        <table class="pt-det">${this.detailInner(date)}</table>
        <div class="pt-sign">Signature :</div>
      </div></div>`;
    },
    // Sauvegarde PDF (Réglages) : pour chaque mois de la période, le mensuel, le récapitulatif
    // de chaque personne et chaque journée pointée.
    exportPdf(p) {
      const k = window.devisKit, jobs = [];
      const docs = this.all().filter(d => this.hoursOf(d).length && (!p || (d.date >= p.from && d.date <= p.to)));
      const months = [...new Set(docs.map(d => d.date.slice(0, 7)))].sort();
      for (const ym of months) {
        const dir = `Pointage/${ym}/`, wrap = h => `<div class="pt pt-pdf">${h}</div>`;
        jobs.push({ path: dir + `Pointage mensuel RSG - ${ym}.pdf`, make: () => k.makePDF({}, { html: wrap(this.monthSheet(ym, 'style="width:794px"')) }) });
        for (const x of this.monthStats(ym).pers) {
          const slug = String(x.nom).replace(/[\\/:*?"<>|]+/g, "-");
          jobs.push({ path: dir + `Personnes/Récapitulatif ${slug} - ${ym}.pdf`, make: () => k.makePDF({}, { html: wrap(this.personSheet(ym, x.nom, 'style="width:794px"')) }) });
        }
        for (const x of this.monthStats(ym).ch) {
          const slug = String(x.nom).replace(/[\\/:*?"<>|]+/g, "-");
          jobs.push({ path: dir + `Chantiers/Synthèse ${slug} - ${ym}.pdf`, make: () => k.makePDF({}, { html: wrap(this.chantierSheet(ym, x.nom, 'style="width:794px"')) }) });
        }
        for (const date of [...new Set(docs.filter(d => d.date.startsWith(ym)).map(d => d.date))].sort())
          jobs.push({ path: dir + `Jours/Pointage journalier RSG - ${date}.pdf`, make: () => k.makePDF({}, { html: this.dayHtml(date), landscape: true }) });
      }
      if (!p) {                                   // période « Tout » : aussi la synthèse de chaque chantier depuis le début
        const chs = new Map(); docs.forEach(d => this.blocks(d).forEach(b => { const c = b.c || SANS; if (!chs.has(key(c))) chs.set(key(c), c); }));
        for (const c of chs.values()) {
          const slug = String(c).replace(/[\\/:*?"<>|]+/g, "-");
          jobs.push({ path: `Pointage/Chantiers depuis le début/Synthèse ${slug}.pdf`, make: () => k.makePDF({}, { html: `<div class="pt pt-pdf">${this.chantierSheet(null, c, 'style="width:794px"')}</div>` }) });
        }
      }
      return jobs;
    },
    sendMail(kind) {
      const k = window.devisKit; if (!k) return;
      this.flushAll();
      const ME = k.ENT.mail, list = a => a.length ? a.map(x => `- ${x.nom} : ${x.h} h`).join("\n") : "- aucune";
      let o;
      if (kind === "jour") {
        const { docs, rows, per, sans, total } = this.detail(this.date);
        const pers = docs.map((d, s) => ({ nom: this.label(d), h: per[s] })).filter(x => x.h);
        const ch = rows.map(r => ({ nom: r.nom, h: r.v.reduce((a, b) => a + b, 0) }));
        const ns = sans.reduce((a, b) => a + b, 0); if (ns) ch.push({ nom: SANS, h: ns });
        o = { html: this.dayHtml(this.date), landscape: true, title: `Pointage journalier ${this.date}`, name: `Pointage journalier RSG - ${this.date}.pdf`,
          what: "du pointage journalier",
          texts: { to: ME, subject: `Pointage journalier RSG – ${frLong(this.date)}`,
            body: `Bonjour,\n\nCi-joint le pointage journalier du ${frLong(this.date).toLowerCase()}.\n\nHeures par personne :\n${list(pers)}\n\nHeures par chantier :\n${list(ch)}\n\nTotal de la journée : ${total} h\n\nRSG – Réalisations Soudures Générales` } };
      } else if (kind === "chantier" && this.chantier) {
        const nom = this.chantier, ym = this.chScope === "tout" ? null : this.month, c = this.chantierStats(ym, nom);
        const slug = String(nom).replace(/[\\/:*?"<>|]+/g, "-"), per = ym ? frMonth(ym) : "depuis le début";
        o = { html: `<div class="pt pt-pdf">${this.chantierSheet(ym, nom, 'style="width:794px"')}</div>`, title: `Chantier ${nom}`,
          name: `Synthèse chantier ${slug} RSG - ${ym || "depuis le début"}.pdf`, what: `de la synthèse du chantier ${nom}`,
          texts: { to: ME, subject: `Synthèse du chantier ${nom} – ${per} – RSG`,
            body: `Bonjour,\n\nCi-joint la synthèse des heures du chantier ${nom} (${per.toLowerCase()}).\n\nHeures par personne :\n${list(c.pers)}\n\nTotal sur le chantier : ${c.total} h sur ${c.nbJours} jour${c.nbJours > 1 ? "s" : ""}.\n\nRSG – Réalisations Soudures Générales` } };
      } else if (kind === "personne" && this.person) {
        const nom = this.person, p = this.personStats(this.month, nom), slug = String(nom).replace(/[\\/:*?"<>|]+/g, "-");
        o = { html: `<div class="pt pt-pdf">${this.personSheet(this.month, nom, 'style="width:794px"')}</div>`, title: `Heures de ${nom} ${this.month}`,
          name: `Pointage ${slug} RSG - ${this.month}.pdf`, what: `du récapitulatif de ${nom}`,
          texts: { to: ME, subject: `Pointage de ${nom} – ${frMonth(this.month)} – RSG`,
            body: `Bonjour,\n\nCi-joint le récapitulatif des heures de ${nom} : ${frMonth(this.month).toLowerCase()}.\n\nHeures par chantier :\n${list(p.ch)}\n\nTotal : ${p.total} h sur ${p.nbJours} jour${p.nbJours > 1 ? "s" : ""}.\n\nRSG – Réalisations Soudures Générales` } };
      } else {
        const m = this.monthStats(this.month);
        o = { html: `<div class="pt pt-pdf">${this.monthSheet(this.month, 'style="width:794px"')}</div>`, title: `Pointage mensuel ${this.month}`, name: `Pointage mensuel RSG - ${this.month}.pdf`,
          what: "du pointage mensuel",
          texts: { to: ME, subject: `Pointage mensuel RSG – ${frMonth(this.month)}`,
            body: `Bonjour,\n\nCi-joint le pointage mensuel : ${frMonth(this.month).toLowerCase()}.\n\n1. Heures totales par personne :\n${list(m.pers)}\n\n2. Heures totales réalisées par chantier :\n${list(m.ch)}\n\n3. Total des heures du mois : ${m.total} h\n\nRSG – Réalisations Soudures Générales` } };
      }
      k.openMail({}, () => toast("Pointage prêt : il ne reste qu'à l'envoyer"), { ...o, label: "pointage", dialogTitle: "Envoyer le pointage par mail", fixedTo: ME });
    },

    monthPc() {
      const m = this.monthStats(this.month), cur = this.month === isoToday().slice(0, 7);
      const daysInMonth = new Date(+this.month.slice(0, 4), +this.month.slice(5, 7), 0).getDate();
      const dayList = [...Array(daysInMonth).keys()].map(i => `${this.month}-${pad(i + 1)}`);
      const perDay = dayList.map(d => [d, this.detail(d).total]);
      return `
        <div class="pt-bar">
          <button class="btn" data-nav="-1" aria-label="Mois précédent">‹</button>
          <b class="pt-mlabel">${frMonth(this.month)}</b>
          <button class="btn" data-nav="1" aria-label="Mois suivant">›</button>
          <button class="btn" data-today ${cur ? "disabled" : ""}>Mois en cours</button>
          ${this.chantier ? `<button class="btn" data-person-back>‹ Retour au récapitulatif du mois</button>
          <span class="pt-scope"><button data-scope="mois" class="${this.chScope !== "tout" ? "on" : ""}">Ce mois-ci</button><button data-scope="tout" class="${this.chScope === "tout" ? "on" : ""}">Depuis le début</button></span>
          <button class="btn primary" data-mail="chantier">Envoyer par mail</button>`
          : this.person ? `<button class="btn" data-person-back>‹ Retour au récapitulatif du mois</button>
          <button class="btn primary" data-mail="personne">Envoyer par mail</button>` : `<button class="btn primary" data-mail="mois">Envoyer par mail</button>`}
          <span class="pt-hint">Calculé tout seul à partir des pointages journaliers : rien à remplir ici.</span>
        </div>
        <div class="pt-month">
          ${this.chantier ? this.chantierSheet(this.chScope === "tout" ? null : this.month, this.chantier)
            : this.person ? this.personSheet(this.month, this.person) : this.monthSheet(this.month, "", true)}
          <aside class="pt-days">
            <h3>Jours du mois</h3>
            <div>${perDay.map(([d, n]) => { const wd = new Date(d + "T12:00:00").getDay();
              return `<button data-day="${d}" class="${n ? "has" : ""}${wd === 0 || wd === 6 ? " we" : ""}"><span>${cap(new Date(d + "T12:00:00").toLocaleDateString("fr-FR", { weekday: "short", day: "numeric" }))}</span><b>${n ? n + " h" : ""}</b></button>`; }).join("")}</div>
          </aside>
        </div>`;
    },

    /* ================= Téléphone : menu guidé ================= */
    mountM() {
      if (this.wz) return this.renderWz();
      this.el.innerHTML = `<div class="dv-m pt-m">
        <h1>Pointage RSG</h1>
        <div class="dv-seg pt-mseg"><button data-v="jour">Journée</button><button data-v="mois">Mois</button></div>
        <div id="ptBody"></div>
      </div>`;
      if (!this.mBound) {
        this.mBound = true;
        this.el.addEventListener("click", e => this.clickM(e));
        this.el.addEventListener("change", e => this.changeM(e));
        this.el.addEventListener("input", e => this.inputM(e));
      }
      this.renderM();
    },
    renderM() {
      const body = this.el.querySelector("#ptBody"); if (!body) return;
      this.el.querySelectorAll(".pt-mseg button[data-v]").forEach(b => b.classList.toggle("on", b.dataset.v === this.view));
      body.innerHTML = this.view === "mois" ? this.monthM() : this.dayM();
    },
    blocks(d) {
      const out = []; let cur = null;
      for (const h of this.hoursOf(d)) {
        const e = d.h[h], sig = key(e.c) + "|" + key(e.t);
        if (cur && cur.sig === sig && cur.to === h) { cur.to = h + 1; cur.hours.push(h); }
        else { cur = { sig, from: h, to: h + 1, c: tr(e.c), t: tr(e.t), hours: [h] }; out.push(cur); }
      }
      return out;
    },
    dayM() {
      const { docs, rows, per, sans, total } = this.detail(this.date), rel = relDay(this.date);
      const cards = docs.map((d, s) => {
        const bl = this.blocks(d); if (!bl.length) return "";
        return `<div class="dv-card pt-pers">
          <div class="pt-ph"><span class="av${s ? "" : " rsg"}">${s ? esc(initials(this.label(d))) : "RSG"}</span><b>${esc(this.label(d))}</b><em>${per[s]} h</em></div>
          ${bl.map(b => `<div class="pt-blkrow"><button class="pt-blk" data-blk="${s}|${b.hours.join(",")}">
            <span class="hrs">${pad(b.from)}h – ${pad(b.to)}h</span>
            <span class="what"><b>${esc(b.c || SANS)}</b>${b.t ? `<small>${esc(b.t)}</small>` : ""}</span>
            <i>${b.hours.length} h</i></button>
            <button class="pt-trash" data-del="${s}|${b.hours.join(",")}" aria-label="Supprimer ce pointage"><svg viewBox="0 0 24 24" width="20" height="20" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M4 7h16M9 7V4h6v3M6 7l1 13h10l1-13M10 11v6M14 11v6"/></svg><span>Supprimer ?</span></button></div>`).join("")}
        </div>`;
      }).join("");
      return `
        <div class="pt-mdate">
          <button data-mnav="-1" aria-label="Jour précédent">‹</button>
          <label><b>${frShort(this.date)}</b><span>${rel || new Date(this.date + "T12:00:00").getFullYear()}</span>
            <input type="date" id="ptMDate" value="${this.date}" aria-label="Choisir la date"></label>
          <button data-mnav="1" aria-label="Jour suivant">›</button>
        </div>
        <button class="btn primary big" data-new>+ Pointer des heures</button>
        ${cards || `<div class="dv-card pt-empty">Aucune heure pointée ${rel === "Aujourd'hui" ? "aujourd'hui" : "ce jour-là"}.<br>Touchez « + Pointer des heures » pour commencer.</div>`}
        ${total ? `<div class="dv-card pt-mdet"><h3>Détail des heures par chantier</h3>
          ${rows.map(r => `<div class="row"><span>${esc(r.nom)}</span><b>${H(r.v.reduce((a, b) => a + b, 0))}</b></div>`).join("")}
          ${sans.some(Boolean) ? `<div class="row sans"><span>${SANS}</span><b>${H(sans.reduce((a, b) => a + b, 0))}</b></div>` : ""}
          <div class="row tot"><span>Heures totales du jour</span><b>${total} h</b></div></div>
          <button class="btn big" data-mail="jour">Envoyer le pointage du jour par mail</button>` : ""}`;
    },
    personM() {
      const nom = this.person, p = this.personStats(this.month, nom);
      return `
        <button class="btn pt-back" data-person-back>‹ Récapitulatif du mois</button>
        <div class="pt-mdate">
          <button data-mnav="-1" aria-label="Mois précédent">‹</button>
          <label><b>${esc(nom)}</b><span>${frMonth(this.month)}</span></label>
          <button data-mnav="1" aria-label="Mois suivant">›</button>
        </div>
        <div class="dv-card pt-mdet"><h3>Heures par chantier</h3>${p.ch.length
          ? p.ch.map(x => `<div class="row"><span>${esc(x.nom)}</span><small>${jours(x.j)}</small><b>${x.h} h</b></div>`).join("")
          : `<p class="small">Aucune heure pointée ce mois-ci.</p>`}
          <div class="row tot"><span>Total de ${esc(nom)}</span><b>${p.total} h</b></div></div>
        ${p.days.length ? `<div class="dv-card pt-mdet pt-mdays"><h3>Détail par jour</h3>${p.days.map(d => `<div class="pday">
          <b>${esc(frShort(d.date))}</b>${d.blocks.map(b => `<div class="row"><span><em>${pad(b.from)}h – ${pad(b.to)}h</em> ${esc(b.c || SANS)}${b.t ? `<small> · ${esc(b.t)}</small>` : ""}</span><b>${b.hours.length} h</b></div>`).join("")}</div>`).join("")}</div>` : ""}
        <button class="btn big" data-mail="personne">Envoyer le récapitulatif de ${esc(nom)} par mail</button>`;
    },
    chantierM() {
      const nom = this.chantier, tout = this.chScope === "tout", c = this.chantierStats(tout ? null : this.month, nom);
      return `
        <button class="btn pt-back" data-person-back>‹ Récapitulatif du mois</button>
        <div class="pt-mdate">
          ${tout ? "<span></span>" : `<button data-mnav="-1" aria-label="Mois précédent">‹</button>`}
          <label><b>${esc(nom)}</b><span>${tout ? "Depuis le début" : frMonth(this.month)}</span></label>
          ${tout ? "<span></span>" : `<button data-mnav="1" aria-label="Mois suivant">›</button>`}
        </div>
        <div class="dv-seg pt-mseg"><button data-scope="mois" class="${tout ? "" : "on"}">Ce mois-ci</button><button data-scope="tout" class="${tout ? "on" : ""}">Depuis le début</button></div>
        <div class="dv-card pt-mdet"><h3>Heures par personne</h3>${c.pers.length
          ? c.pers.map(x => `<div class="row"><span>${esc(x.nom)}</span><small>${jours(x.j)}</small><b>${x.h} h</b></div>`).join("")
          : `<p class="small">Aucune heure sur ce chantier pour cette période.</p>`}
          <div class="row tot"><span>Total sur le chantier</span><b>${c.total} h</b></div></div>
        ${c.days.length ? `<div class="dv-card pt-mdet pt-mdays"><h3>Détail par jour</h3>${c.days.map(d => `<div class="pday">
          <b>${esc(frShort(d.date))}${tout ? " " + d.date.slice(0, 4) : ""}</b>${d.rows.map(r => `<div class="row"><span><em>${pad(r.from)}h – ${pad(r.to)}h</em> ${esc(r.pers)}${r.t ? `<small> · ${esc(r.t)}</small>` : ""}</span><b>${r.h} h</b></div>`).join("")}</div>`).join("")}</div>` : ""}
        <button class="btn big" data-mail="chantier">Envoyer la synthèse du chantier par mail</button>`;
    },
    monthM() {
      if (this.chantier) return this.chantierM();
      if (this.person) return this.personM();
      const m = this.monthStats(this.month);
      const list = (titre, arr, vide, pers) => `<div class="dv-card pt-mdet"><h3>${titre}</h3>${arr.length
        ? arr.map(x => `<button class="row click" ${pers ? "data-person" : "data-chantier"}="${esc(x.nom)}"><span>${esc(x.nom)}</span><small>${jours(x.j)}</small><b>${x.h} h</b><i>›</i></button>`).join("")
        : `<p class="small">${vide}</p>`}</div>`;
      return `
        <div class="pt-mdate">
          <button data-mnav="-1" aria-label="Mois précédent">‹</button>
          <label><b>${frMonth(this.month)}</b><span>Calculé tout seul</span></label>
          <button data-mnav="1" aria-label="Mois suivant">›</button>
        </div>
        ${list("1. Heures totales par personne", m.pers, "Aucune heure pointée ce mois-ci.", true)}
        ${m.pers.length ? `<p class="small pt-tip">Touchez une personne pour voir ses heures sur chaque chantier.</p>` : ""}
        ${list("2. Heures totales réalisées par chantier", m.ch, "Aucun chantier ce mois-ci.")}
        ${m.ch.length ? `<p class="small pt-tip">Touchez un chantier pour voir qui y a travaillé et combien de temps.</p>` : ""}
        <div class="dv-card pt-mdet"><h3>3. Total des heures du mois</h3><div class="row tot"><span>Total des heures du mois</span><b>${m.total} h</b></div></div>
        <button class="btn big" data-mail="mois">Envoyer le pointage du mois par mail</button>`;
    },
    clickM(e) {
      if (current !== this || this.mode !== "m") return;
      const t = e.target;
      if (this.wz) return this.clickWz(e);
      const pr = t.closest("[data-person]"); if (pr) { this.person = pr.dataset.person; this.chantier = null; window.scrollTo(0, 0); return this.renderM(); }
      const chx = t.closest("[data-chantier]"); if (chx) { this.chantier = chx.dataset.chantier; this.chScope = "mois"; this.person = null; window.scrollTo(0, 0); return this.renderM(); }
      const sc = t.closest("[data-scope]"); if (sc) { this.chScope = sc.dataset.scope; return this.renderM(); }
      if (t.closest("[data-person-back]")) { this.person = null; this.chantier = null; return this.renderM(); }
      const v = t.closest(".pt-mseg [data-v]");
      if (v) { this.person = null; this.chantier = null; this.view = v.dataset.v; if (this.view === "mois") this.month = this.date.slice(0, 7); return this.renderM(); }
      const n = t.closest("[data-mnav]");
      if (n) { if (this.view === "jour") this.date = addDays(this.date, +n.dataset.mnav); else this.month = addMonths(this.month, +n.dataset.mnav); return this.renderM(); }
      const ml = t.closest("[data-mail]"); if (ml) return this.sendMail(ml.dataset.mail);
      if (t.closest("[data-new]")) { if (!canEdit()) return; return this.openWz(); }
      const del = t.closest("[data-del]");
      if (del) {
        if (!canEdit()) return;
        if (!del.classList.contains("armed")) {
          this.el.querySelectorAll(".pt-trash.armed").forEach(x => x.classList.remove("armed"));
          del.classList.add("armed"); clearTimeout(this.disarmT);
          this.disarmT = setTimeout(() => del.classList.remove("armed"), 4000); return;
        }
        const [s, hs] = del.dataset.del.split("|"), hours = hs.split(",").map(Number);
        this.edit(this.date, +s, x => hours.forEach(h => delete x.h[h]), true);
        toast(`Pointage supprimé (${hours.length} h)`); return this.renderM();
      }
      const b = t.closest("[data-blk]");
      if (b) {
        if (!canEdit()) return;
        const [s, hs] = b.dataset.blk.split("|"), hours = hs.split(",").map(Number), d = this.doc(this.date, +s), e0 = d.h[hours[0]];
        return this.openWz({ date: this.date, nom: this.label(d), c: tr(e0.c), t: tr(e0.t), hours: new Set(hours), edit: { date: this.date, slot: +s, hours }, step: 3 });
      }
    },
    changeM(e) {
      if (current !== this || this.mode !== "m") return;
      if (e.target.id === "ptMDate" && e.target.value) { this.date = e.target.value; this.renderM(); }
      if (e.target.id === "wzDate" && e.target.value && this.wz) { this.wz.date = e.target.value; this.renderWz(); }
    },
    inputM(e) {
      if (current !== this || !this.wz) return;
      const i = e.target;
      if (i.id === "wzNom") { this.wz.nom = i.value; this.el.querySelectorAll("[data-pers]").forEach(b => b.classList.toggle("on", key(b.dataset.pers) === key(i.value))); }
      if (i.id === "wzC") { this.wz.c = i.value; this.el.querySelectorAll("[data-c]").forEach(b => b.classList.toggle("on", key(b.dataset.c) === key(i.value))); }
      if (i.id === "wzT") { this.wz.t = i.value; this.el.querySelectorAll("[data-t]").forEach(b => b.classList.toggle("on", key(b.dataset.t) === key(i.value))); }
      this.navWz();
    },

    /* ---------- assistant : qui → chantier → travail → heures ---------- */
    openWz(w) {
      this.wz = Object.assign({ step: 0, date: this.date, nom: "", c: "", t: "", hours: new Set(), edit: null }, w || {});
      this.renderWz(); window.scrollTo(0, 0);
    },
    slotFor(date, nom) {
      if (key(nom) === "rsg") return 0;
      const docs = this.day(date);
      const same = docs.find(d => d.slot && key(this.name(d)) === key(nom)); if (same) return same.slot;
      const free = docs.find(d => d.slot && !this.name(d) && !this.hoursOf(d).length);
      return free ? free.slot : -1;
    },
    busyHours() {
      const w = this.wz, s = this.slotFor(w.date, w.nom); if (s < 0) return new Map();
      const d = this.doc(w.date, s), m = new Map();
      for (const h of this.hoursOf(d)) {
        if (w.edit && w.edit.date === w.date && w.edit.slot === s && w.edit.hours.includes(h)) continue;
        m.set(h, d.h[h]);
      }
      return m;
    },
    renderWz() {
      const w = this.wz, k = this.known(), rel = relDay(w.date);
      const steps = `<div class="dv-steps">${[0, 1, 2, 3].map(i => `<i class="${i <= w.step ? "on" : ""}"></i>`).join("")}</div>`;
      let h = "";
      if (w.step === 0) {
        const t = isoToday(), y = addDays(t, -1);
        const people = ["RSG", ...k.pers.filter(p => key(p) !== "rsg")];
        const dayDocs = this.day(w.date);
        const hoursFor = n => { const d = dayDocs.find(x => key(this.label(x)) === key(n)); return d ? this.hoursOf(d).length : 0; };
        h = `<p class="dv-q">Qui a travaillé ?</p>
          <div class="dv-chips pt-when">
            <button data-date="${t}" class="${w.date === t ? "on" : ""}">Aujourd'hui</button>
            <button data-date="${y}" class="${w.date === y ? "on" : ""}">Hier</button>
            <label class="pt-other ${w.date !== t && w.date !== y ? "on" : ""}">${w.date !== t && w.date !== y ? frShort(w.date) : "Autre jour"}<input type="date" id="wzDate" value="${w.date}"></label>
          </div>
          <div class="dv-pick">${people.map(p => { const n = hoursFor(p);
            return `<button data-pers="${esc(p)}" class="${key(p) === key(w.nom) ? "on" : ""}"><span class="av${key(p) === "rsg" ? " rsg" : ""}">${key(p) === "rsg" ? "RSG" : esc(initials(p))}</span><span><b>${esc(p)}</b><span>${n ? `Déjà ${n} h pointée${n > 1 ? "s" : ""} ${rel ? rel.toLowerCase() : "ce jour-là"}` : key(p) === "rsg" ? "Vous" : "Personne déjà pointée"}</span></span></button>`; }).join("")}</div>
          <div class="field"><label for="wzNom">Autre personne</label><input id="wzNom" placeholder="Prénom ou nom" autocomplete="off" value="${people.some(p => key(p) === key(w.nom)) ? "" : esc(w.nom)}"></div>`;
      } else if (w.step === 1) {
        h = `<p class="dv-q">Sur quel chantier ?</p><p class="dv-sub">${esc(w.nom)} · ${esc(frShort(w.date))}</p>
          <div class="field"><label for="wzC">Chantier</label><input id="wzC" placeholder="ex. Dupont, Atelier…" autocomplete="off" value="${esc(w.c)}"></div>
          ${k.ch.length ? `<div class="pt-lbl">Chantiers récents</div><div class="dv-chips">${k.ch.slice(0, 10).map(c => `<button data-c="${esc(c)}" class="${key(c) === key(w.c) ? "on" : ""}">${esc(c)}</button>`).join("")}</div>` : ""}`;
      } else if (w.step === 2) {
        const sug = (w.c ? k.trv(w.c) : []).slice(0, 8);
        h = `<p class="dv-q">Qu'est-ce qui a été fait ?</p><p class="dv-sub">${esc(w.nom)} · ${esc(w.c || SANS)}</p>
          <div class="field"><label for="wzT">Travail effectué</label><textarea id="wzT" rows="3" placeholder="ex. Pose garde-corps">${esc(w.t)}</textarea></div>
          ${sug.length ? `<div class="pt-lbl">Déjà fait sur ce chantier</div><div class="dv-chips">${sug.map(t => `<button data-t="${esc(t)}" class="${key(t) === key(w.t) ? "on" : ""}">${esc(t)}</button>`).join("")}</div>` : ""}`;
      } else {
        const busy = this.busyHours();
        const pre = [["Matin", "8h – 12h", [8, 9, 10, 11]], ["Après-midi", "13h – 17h", [13, 14, 15, 16]], ["Journée", "8h – 12h + 13h – 17h", [8, 9, 10, 11, 13, 14, 15, 16]]];
        const same = a => a.length === w.hours.size && a.every(x => w.hours.has(x));
        h = `<p class="dv-q">Quelles heures ?</p><p class="dv-sub pt-hsub">${w.anchor != null
            ? `Début : <b>${pad(w.anchor)}h</b>. Touchez maintenant la <b>dernière heure travaillée</b> : tout sera pointé entre les deux.`
            : `Touchez l'heure de début, puis la dernière heure travaillée : toutes les heures entre les deux sont pointées. Touchez une heure rouge pour l'enlever.`}</p>
          <div class="pt-pre">${pre.map(([a, b, hs]) => `<button data-pre="${hs.join(",")}" class="${same(hs) ? "on" : ""}"><b>${a}</b><span>${b}</span></button>`).join("")}</div>
          <div class="pt-hgrid">${HOURS.map(x => { const bz = busy.get(x), on = w.hours.has(x);
            return `<button data-hour="${x}" class="${on ? "on" : ""}${w.anchor === x ? " anchor" : ""}${bz ? " busy" : ""}${x < 6 || x > 20 ? " night" : ""}" aria-pressed="${on}"><b>${pad(x)}h</b><span>${bz ? esc(tr(bz.c) || tr(bz.t)).slice(0, 14) : `– ${pad(x + 1)}h`}</span></button>`; }).join("")}</div>
          ${[...w.hours].some(x => busy.has(x)) ? `<p class="pt-warn">Les heures en gris étaient déjà pointées : elles seront remplacées.</p>` : ""}
          ${w.edit ? `<button class="btn danger pt-del" data-wz="del">Supprimer ces heures</button>` : ""}`;
      }
      this.el.innerHTML = `<div class="dv-m pt-wz">${steps}${h}</div>
        <div class="dv-nav"><button class="btn" data-wz="back">${w.step ? "Retour" : "Annuler"}</button><div class="tt" id="wzSum"></div><button class="btn primary" data-wz="next" id="wzNext"></button></div>`;
      this.navWz();
    },
    navWz() {
      const w = this.wz, sum = this.el.querySelector("#wzSum"), nx = this.el.querySelector("#wzNext"); if (!sum) return;
      const n = w.hours.size;
      sum.innerHTML = w.step === 3 ? `${esc(w.nom)} · ${esc(this.ranges(w.hours) || "aucune heure")}<b>${n} h</b>` : `${esc(relDay(w.date) || frShort(w.date))}<b>${esc(w.nom || "—")}</b>`;
      nx.textContent = w.step === 3 ? "Enregistrer" : "Suivant";
      nx.disabled = (w.step === 0 && !tr(w.nom)) || (w.step === 3 && !n);
    },
    ranges(set) {
      const hs = [...set].sort((a, b) => a - b), out = [];
      for (const h of hs) { const r = out.at(-1); if (r && r[1] === h) r[1] = h + 1; else out.push([h, h + 1]); }
      return out.map(([a, b]) => `${pad(a)}h–${pad(b)}h`).join(", ");
    },
    clickWz(e) {
      const w = this.wz, t = e.target;
      const d = t.closest("[data-date]"); if (d) { w.date = d.dataset.date; return this.renderWz(); }
      const p = t.closest("[data-pers]"); if (p) { w.nom = p.dataset.pers; return this.next(); }
      const c = t.closest("[data-c]"); if (c) { w.c = c.dataset.c; return this.next(); }
      const tt = t.closest("[data-t]"); if (tt) { w.t = tt.dataset.t; return this.next(); }
      const pr = t.closest("[data-pre]"); if (pr) { w.hours = new Set(pr.dataset.pre.split(",").map(Number)); w.anchor = null; return this.renderWz(); }
      const hb = t.closest("[data-hour]");
      if (hb) {
        const x = +hb.dataset.hour;
        if (w.anchor == null) {
          if (w.hours.has(x)) w.hours.delete(x);           // heure déjà pointée : on l'enlève
          else { w.hours.add(x); w.anchor = x; }           // heure de début
        } else if (x === w.anchor) { w.hours.delete(x); w.anchor = null; }
        else {                                               // dernière heure : tout l'intervalle
          for (let i = Math.min(x, w.anchor); i <= Math.max(x, w.anchor); i++) w.hours.add(i);
          w.anchor = null;
        }
        return this.renderWz();
      }
      const a = t.closest("[data-wz]"); if (!a) return;
      if (a.dataset.wz === "back") { if (w.step) { w.step--; this.renderWz(); } else this.closeWz(); }
      else if (a.dataset.wz === "next") this.next();
      else if (a.dataset.wz === "del") {
        if (!a.classList.contains("armed")) { a.classList.add("armed"); a.textContent = "Confirmer la suppression"; return; }
        this.edit(w.edit.date, w.edit.slot, x => w.edit.hours.forEach(h => delete x.h[h]), true);
        toast(`${w.edit.hours.length} h supprimée${w.edit.hours.length > 1 ? "s" : ""}`); this.closeWz();
      }
    },
    next() {
      const w = this.wz;
      if (w.step === 0) {
        w.nom = tr(w.nom); if (!w.nom) return;
        if (key(w.nom) === "rsg") w.nom = "RSG";
        if (this.slotFor(w.date, w.nom) < 0 && !(w.edit && key(this.label(this.doc(w.edit.date, w.edit.slot))) === key(w.nom)))
          return toast("Déjà 4 personnes pointées ce jour-là (RSG + 3), comme sur la feuille de pointage.");
      }
      if (w.step === 1) w.c = tr(w.c);
      if (w.step === 2) { w.t = tr(w.t); if (!w.c && !w.t) return toast("Indiquez au moins le chantier ou le travail effectué."); }
      if (w.step < 3) { w.step++; this.renderWz(); window.scrollTo(0, 0); return; }
      this.saveWz();
    },
    saveWz() {
      const w = this.wz, hours = [...w.hours].sort((a, b) => a - b);
      if (w.edit) this.edit(w.edit.date, w.edit.slot, x => w.edit.hours.forEach(h => delete x.h[h]), true);
      const s = this.slotFor(w.date, w.nom);
      if (s < 0) return toast("Déjà 4 personnes pointées ce jour-là (RSG + 3).");
      const ok = this.edit(w.date, s, x => { if (s) x.nom = w.nom; hours.forEach(h => { x.h[h] = { c: w.c, t: w.t }; }); }, true);
      if (!ok) return;
      toast(`${hours.length} h pointée${hours.length > 1 ? "s" : ""} pour ${w.nom}`);
      this.date = w.date; this.view = "jour"; this.closeWz();
    },
    closeWz() { this.wz = null; this.mountM(); window.scrollTo(0, 0); },
  };

  window.pointageTab = tab;
})();
