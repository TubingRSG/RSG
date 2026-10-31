/* =====================================================================
   RSG — Avenants, rangés sous leur devis (onglet « Création de devis »)
   Un avenant modifie un devis accepté : travaux en plus (prix positifs) ou en moins
   (prix négatifs). Il se signe « Bon pour accord » comme un devis, puis se facture
   à part (« Facturer cet avenant ») ou avec la facture de fin de chantier.
   PC : la page A4 modifiable comme dans Excel. Téléphone : lignes guidées + aperçu.
   Données : collection « avenants ».
   ===================================================================== */
"use strict";
(function () {
  const K = () => window.devisKit;
  const COL = "avenants";
  const STATUTS = { brouillon: "Brouillon", envoye: "Envoyé", accepte: "Accepté", refuse: "Refusé" };
  const NB_LIGNES = 7;
  const round2 = v => Math.round(v * 100) / 100;

  function nextNumero(iso, selfId) {
    if (!/^\d{4}-\d{2}-\d{2}$/.test(iso || "")) return "";
    const [y, m, dd] = iso.split("-"), pre = `${y}/${m}${dd}-A`;
    const self = selfId && store.get(COL, selfId);
    if (self && String(self.numero || "").startsWith(pre)) return self.numero;
    const max = store.list(COL).filter(x => x.id !== selfId && String(x.numero || "").startsWith(pre))
      .reduce((mx, x) => Math.max(mx, parseInt(String(x.numero).slice(pre.length), 10) || 0), 0);
    return pre + (max + 1);
  }
  function normalize(a) {
    const k = K(), x = Object.assign({ numero: "", date: k.isoToday(), devisId: null, devisNumero: "", devisDate: "", rang: 1,
      statut: "brouillon", clientId: null, tva: 0.2, devisHT: 0, devisTTC: 0 }, k.clone(a));
    x.client = Object.assign({ nom: "", adr1: "", adr2: "", ref: "", email: "" }, x.client || {});
    x.lignes = (x.lignes || []).map(l => ({ d: l.d || "", pu: l.pu ?? "", q: l.q ?? "" }));
    return x;
  }
  const ofDevis = id => store.list(COL).filter(a => a.devisId === id).sort((a, b) => (a.rang || 0) - (b.rang || 0));
  const ttcOf = a => round2(K().calc(normalize(a)).ttc);
  function cumul(a) {
    const prev = round2(ofDevis(a.devisId).filter(x => x.id !== a.id && (x.rang || 0) < (a.rang || 0) && x.statut !== "refuse").reduce((s, x) => s + ttcOf(x), 0));
    const cur = round2(K().calc(a).ttc);
    return { devis: round2(a.devisTTC), prev, cur, total: round2(a.devisTTC + prev + cur) };
  }
  const fileName = a => ["Avenant RSG", a.client?.nom, a.client?.ref, String(a.numero || "").replace(/\//g, "-")].map(t => K().cleanName(t)).filter(Boolean).join(" - ") + ".pdf";
  const signed = v => (K().num(v) < 0 ? "− " : "+ ") + K().eur.format(Math.abs(K().num(v)));

  /* ---------- la page A4 de l'avenant ---------- */
  function page(a, edit) {
    const k = K(), { ENT, eur, num, money, pctTxt, frDate } = k, c = a.client, t = k.calc(a), cu = cumul(a);
    const inp = (path, val, o = {}) => edit
      ? `<input class="f ${o.cls || ""}" data-f="${path}" value="${esc(val ?? "")}" ${o.type ? `type="${o.type}"` : ""} inputmode="${o.im || "text"}" placeholder="${esc(o.ph || "")}" autocomplete="off">`
      : esc(o.show ? o.show(val) : (val ?? ""));
    const rows = Math.max(NB_LIGNES, a.lignes.length + (edit ? 1 : 0));
    let body = "";
    for (let i = 0; i < rows; i++) {
      const l = a.lignes[i] || { d: "", pu: "", q: "" };
      const has = String(l.d).trim() || num(l.pu) || num(l.q);
      const tot = has && l.pu !== "" && l.q !== "" ? eur.format(num(l.pu) * num(l.q)) : "";
      body += `<tr data-row="${i}" class="${num(l.pu) < 0 ? "moins" : ""}">
        <td>${edit ? `<textarea class="f" data-f="lignes.${i}.d" rows="1" placeholder="${i === 0 && !a.lignes.length ? "Travaux en plus ou en moins…" : ""}">${esc(l.d)}</textarea>` : esc(l.d)}</td>
        <td class="n">${edit ? inp(`lignes.${i}.pu`, money(l.pu), { cls: "num-in money", im: "decimal" }) : (has && l.pu !== "" ? eur.format(num(l.pu)) : "")}</td>
        <td class="q">${edit ? inp(`lignes.${i}.q`, l.q, { cls: "c", im: "decimal" }) : (has ? esc(String(l.q).replace(".", ",")) : "")}</td>
        <td class="t" data-calc="l${i}">${tot}${edit && i < a.lignes.length ? `<button class="rowdel noprint" data-del="${i}" title="Supprimer la ligne" aria-label="Supprimer la ligne">×</button>` : ""}</td></tr>`;
    }
    const tvaSel = edit
      ? `<select class="f" data-f="tva" aria-label="Taux de TVA">${k.TVAS.map(v => `<option value="${v}" ${v === a.tva ? "selected" : ""}>${pctTxt(v)}</option>`).join("")}</select>`
      : pctTxt(a.tva);
    return `<div class="a4 av-a4 ${edit ? "" : "static"}">
      <div class="hd">
        <div class="co"><img class="logo" src="${k.LOGO}" alt="RSG">
          <div>${ENT.adresse}</div><div>Téléphone : ${ENT.tel}</div><div>Mail : ${ENT.mail}</div>
          <small>${ENT.ape}</small><small>${ENT.siret}</small><small>${ENT.tva}</small></div>
        <div>
          <div class="title"><b>AVENANT N°</b><span class="num" data-calc="num">${esc(a.numero || "—")}</span></div>
          <div class="meta"><span>En date du</span><em>${inp("date", a.date, { type: "date", show: frDate, cls: "c" })}</em></div>
          <div class="clientlbl">CLIENT</div>
          <div class="client">
            <div class="cn">${inp("client.nom", c.nom, { ph: "Nom du client" })}</div>
            <div>${inp("client.adr1", c.adr1, { ph: "Adresse" })}</div>
            <div>${inp("client.adr2", c.adr2, { ph: "Code postal et ville" })}</div>
            <div class="ref"><span>Référence :</span><b>${inp("client.ref", c.ref, { ph: "Chantier" })}</b></div>
            ${edit ? `<div class="mailrow noprint"><span>E-mail du client (non imprimé) :</span>${inp("client.email", c.email, { type: "email", ph: "client@exemple.fr" })}</div>` : ""}
          </div>
        </div>
      </div>
      <div class="ft-origin">Avenant n° ${a.rang} au devis n° ${esc(a.devisNumero)}${a.devisDate ? ` du ${frDate(a.devisDate)}` : ""}</div>
      ${edit ? `<div class="av-tip noprint">Travaux en plus : prix positif. Travaux en moins : prix négatif, par exemple -150.</div>` : ""}
      <table class="ft-table"><colgroup><col class="c1"><col class="c2"><col class="c3"><col class="c4"></colgroup>
        <thead><tr><th>Modifications (travaux en plus / en moins)</th><th>Prix unitaire HT</th><th>Quantité</th><th>Total HT</th></tr></thead>
        <tbody>${body}</tbody></table>
      ${edit ? `<div class="addrow noprint"><button data-act="addline">+ Ajouter une ligne</button></div>` : ""}
      <div class="tot">
        <div><span>Total HT de l'avenant</span><b data-calc="ht">${eur.format(t.ht)}</b></div>
        <div><span>Total TVA (${tvaSel})</span><b data-calc="tv">${eur.format(t.tv)}</b></div>
        <div class="ttc"><span>TOTAL TTC DE L'AVENANT</span><b data-calc="ttc">${eur.format(t.ttc)}</b></div>
      </div>
      ${a.tva === 0 ? `<div class="autoliq">${k.AUTOLIQ}</div>` : ""}
      <div class="av-recap">
        <div><span>Montant du devis n° ${esc(a.devisNumero)} (TTC)</span><b>${eur.format(cu.devis)}</b></div>
        ${cu.prev ? `<div><span>Avenants précédents</span><b>${signed(cu.prev)}</b></div>` : ""}
        <div><span>Présent avenant n° ${a.rang}</span><b data-calc="avcur">${signed(cu.cur)}</b></div>
        <div class="tt"><span>Nouveau montant total du chantier (TTC)</span><b data-calc="avtot">${eur.format(cu.total)}</b></div>
      </div>
      <div class="acc"><div class="lbl">ACCORD DU CLIENT</div><div class="box"><span>Signature avec la mention "Bon pour accord"</span><span>Date</span></div></div>
      <div class="foot">${ENT.paiement}<br>${ENT.tva}</div>
    </div>`;
  }

  function mailTexts(a) {
    const k = K(), c = a.client, cu = cumul(a);
    return {
      to: c.email || "",
      subject: `Avenant n° ${a.rang} (${a.numero}) – RSG${c.ref ? " – " + c.ref : ""}`,
      body: `Bonjour${c.nom ? " " + c.nom : ""},\n\nVeuillez trouver ci-joint l'avenant n° ${a.rang} à notre devis n° ${a.devisNumero}${c.ref ? ` pour « ${c.ref} »` : ""}, d'un montant de ${signed(cu.cur)} TTC.\nLe nouveau montant total du chantier est de ${k.eur.format(cu.total)} TTC.\n\nPour l'accepter, il suffit de nous le retourner signé avec la mention « Bon pour accord ».\n\nCordialement,\nRSG – Réalisations Soudures Générales\n${k.ENT.tel} – ${k.ENT.mail}`,
    };
  }
  const pdfOpts = a => ({ html: page(a, false), title: `Avenant ${a.numero}`, name: fileName(a), texts: mailTexts(a), label: "avenant",
    what: "de l'avenant", dialogTitle: "Envoyer l'avenant par mail" });

  async function create(devisId) {
    const k = K(), dv = store.get("devis", devisId); if (!dv) throw new Error("Devis introuvable.");
    const d = k.cleanLines(k.normalize(dv)), t = k.calc(d), date = k.isoToday();
    const rang = ofDevis(devisId).reduce((m, x) => Math.max(m, x.rang || 0), 0) + 1;
    return store.create(COL, {
      numero: nextNumero(date), date, devisId, devisNumero: d.numero, devisDate: d.date, rang, statut: "brouillon",
      clientId: d.clientId || null, client: { nom: d.client.nom, adr1: d.client.adr1, adr2: d.client.adr2, ref: d.client.ref, email: d.client.email },
      lignes: [], tva: d.tva, devisHT: round2(t.ht), devisTTC: round2(t.ttc),
    });
  }

  /* =================================================================
     L'avenant ouvert sous son devis (onglet « Création de devis »).
     Le devis affiche des sous-onglets : Devis | Avenant n° 1 | … | + Créer un avenant.
     ================================================================= */
  const A = {
    id: null, draft: null, dirty: false, saving: false, timer: null, pane: null, m: { screen: "view" },
    sorted() { return store.list(COL).sort((a, b) => (b.date || "").localeCompare(a.date || "") || String(b.numero).localeCompare(String(a.numero), "fr", { numeric: true })); },
    exportPdf(p) {
      return this.sorted().filter(a => !p || (String(a.date || "") >= p.from && String(a.date || "") <= p.to)).map(a => { const n = K().cleanLines(normalize(a));
        return { path: "Avenants/" + fileName(n), make: () => K().makePDF(n, pdfOpts(n)) }; });
    },
    count: devisId => ofDevis(devisId).length,
    load(id) { const a = store.get(COL, id); this.id = a ? id : null; this.draft = a ? normalize(a) : null; this.dirty = false; this.m = { screen: a && !(a.lignes || []).length ? "edit" : "view" }; return !!a; },
    async create(devisId) { const n = await create(devisId); this.load(n.id); this.m = { screen: "edit" }; toast(`Avenant n° ${n.rang} (${n.numero}) créé`); return n; },

    // sous-onglets du devis
    tabsHtml(devisId, devisNumero, activeId) {
      return `<div class="dv-subtabs" role="tablist">
        <button data-dsub="" class="${activeId ? "" : "on"}"><b>Devis</b><small>${esc(devisNumero || "")}</small></button>
        ${ofDevis(devisId).map(a => `<button data-dsub="${a.id}" class="${a.id === activeId ? "on" : ""}"><b>Avenant n° ${a.rang}</b><small>${signed(ttcOf(a))}${window.factureTab?.avenantFacture(a.id) ? " · facturé" : ""}</small></button>`).join("")}
        <button data-act="newav" class="add">+ Créer un avenant</button></div>`;
    },

    /* ----------------------------- PC ----------------------------- */
    renderPC(pane) {
      this.pane = pane || this.pane; if (!this.pane || !this.draft) return;
      const a = this.draft, facts = window.factureTab?.avenantFactures(a.id) || [];
      this.pane.innerHTML = `<div id="avPane"><div class="dv-bar">
          <button class="btn" data-dsub="">‹ Retour au devis</button>
          <button class="btn primary" data-act="pdf">Enregistrer en PDF sur le PC</button>
          <button class="btn" data-act="mail">Envoyer par mail</button>
          <select class="btn" data-act="statut" aria-label="Statut de l'avenant">${Object.entries(STATUTS).map(([k, v]) => `<option value="${k}" ${a.statut === k ? "selected" : ""}>${v}</option>`).join("")}</select>
          <button class="btn" data-act="facturer">Transformer en facture d'avenant</button>
          ${facts.length ? `<button class="btn" data-act="voirfacture">Voir ${facts.length > 1 ? `les ${facts.length} factures` : `la facture ${esc(facts[0].numero)}`}</button>` : ""}
          <button class="btn danger" data-act="del">Supprimer l'avenant</button>
          <span class="st" id="avSt"><i></i><span>Enregistré</span></span>
        </div>
        <div class="dv-hint">Avenant n° ${a.rang} au devis ${esc(a.devisNumero)} : écrivez les travaux en plus (prix positif) ou en moins (prix négatif), comme dans Excel. Tout est enregistré automatiquement.</div>
        <div class="dv-stage">${page(a, true)}</div></div>`;
      this.pane.querySelectorAll(".a4 textarea.f").forEach(K().grow);
    },
    refreshPC() {
      if (!this.id || this.dirty || this.saving || this.pane?.querySelector(".a4")?.contains(document.activeElement)) return;
      const fresh = store.get(COL, this.id);
      if (fresh && JSON.stringify(normalize(fresh)) !== JSON.stringify(this.draft)) { this.draft = normalize(fresh); this.renderPC(); }
    },
    status(txt, pending) { const s = $("#avSt"); if (!s) return; s.classList.toggle("pending", !!pending); s.lastElementChild.textContent = txt; },
    queueSave() {
      if (!canEdit()) return;
      this.dirty = true; this.status("Modifications en cours…", true);
      clearTimeout(this.timer); this.timer = setTimeout(() => this.flush(), 700);
    },
    async flush() {
      if (!this.dirty || this.saving || !this.id) return;
      this.dirty = false; this.saving = true; this.status("Enregistrement…", true);
      try { await store.patch(COL, this.id, K().cleanLines(this.draft)); this.status("Enregistré", false); }
      catch (e) { this.dirty = true; this.status("Non enregistré", true); toast(frError(e)); }
      finally { this.saving = false; if (this.dirty) this.queueSave(); }
    },
    updateCalc() {
      const k = K(), a = this.draft, t = k.calc(a), cu = cumul(a), set = (n, v) => { const e = this.pane.querySelector(`[data-calc="${n}"]`); if (e) e.textContent = v; };
      a.lignes.forEach((l, i) => {
        const has = String(l.d).trim() || k.num(l.pu) || k.num(l.q);
        const cell = this.pane.querySelector(`[data-calc="l${i}"]`); if (!cell) return;
        const btn = cell.querySelector(".rowdel"); cell.textContent = has && l.pu !== "" && l.q !== "" ? k.eur.format(k.num(l.pu) * k.num(l.q)) : ""; if (btn) cell.appendChild(btn);
        cell.closest("tr")?.classList.toggle("moins", k.num(l.pu) < 0);
      });
      set("ht", k.eur.format(t.ht)); set("tv", k.eur.format(t.tv)); set("ttc", k.eur.format(t.ttc));
      set("avcur", signed(cu.cur)); set("avtot", k.eur.format(cu.total));
      const tab = document.querySelector(`.dv-subtabs [data-dsub="${this.id}"] small`); if (tab) tab.textContent = signed(cu.cur);
      if (!this.pane.querySelector(`[data-f="lignes.${a.lignes.length - 1}.q"]`)?.value) { const qi = this.pane.querySelector(`[data-f="lignes.${a.lignes.length - 1}.q"]`); if (qi && document.activeElement !== qi) qi.value = a.lignes.at(-1)?.q ?? ""; }
    },
    onInputPC(e) {
      const f = e.target.dataset.f; if (!f || !this.draft) return;
      let v = e.target.value;
      if (f === "tva") v = +v;
      if (/^lignes\.\d+\.(pu|q)$/.test(f)) v = v === "" ? "" : v.replace(/[\s  €]/g, "").replace(",", ".").replace(/^−/, "-");
      if (/^lignes\.(\d+)\./.test(f)) {
        const i = +f.split(".")[1];
        while (this.draft.lignes.length <= i) this.draft.lignes.push({ d: "", pu: "", q: 1 });
        const needRow = i === this.draft.lignes.length - 1 && !this.pane.querySelector(`tr[data-row="${i + 1}"]`);
        K().setPath(this.draft, f, v);
        if (needRow) { const pos = K().caretOf(e.target); this.renderPC(); K().refocus(this.pane, f, pos); }
      } else K().setPath(this.draft, f, v);
      if (f === "date" && v) { this.draft.numero = nextNumero(v, this.id); const n = this.pane.querySelector('[data-calc="num"]'); if (n) n.textContent = this.draft.numero; }
      if (e.target.tagName === "TEXTAREA") K().grow(e.target);
      this.updateCalc(); this.queueSave();
    },
    onChangePC(e) {
      const t = e.target;
      if (t.dataset.f === "tva") { this.draft.tva = +t.value; this.renderPC(); this.queueSave(); }
      if (t.dataset.act === "statut") { this.draft.statut = t.value; this.queueSave(); }
    },
    // renvoie "deleted" quand l'avenant est supprimé (le devis se réaffiche)
    async onClickPC(e) {
      const g = s => e.target.closest(s);
      if (g("[data-del]")) { this.draft.lignes.splice(+g("[data-del]").dataset.del, 1); this.renderPC(); this.updateCalc(); this.queueSave(); return; }
      const act = g("[data-act]")?.dataset.act; if (!act || e.target.tagName === "SELECT") return;
      if (act === "addline") { this.draft.lignes = K().cleanLines(this.draft).lignes; this.draft.lignes.push({ d: "", pu: "", q: 1 }); this.renderPC(); K().refocus(this.pane, `lignes.${this.draft.lignes.length - 1}.d`); }
      if (act === "pdf") { await this.flush(); try { const r = await K().savePDF(this.draft, pdfOpts(K().cleanLines(this.draft))); if (r === "saved") toast("PDF enregistré dans Téléchargements : " + fileName(this.draft)); } catch (err) { if (err?.code !== "declined") toast(err.message || String(err)); } }
      if (act === "mail") { await this.flush(); const a = K().cleanLines(this.draft); K().openMail(a, () => { if (this.draft.statut === "brouillon") { this.draft.statut = "envoye"; this.queueSave(); this.renderPC(); } }, pdfOpts(a)); }
      if (act === "facturer") { await this.flush(); return this.facturer(g("[data-act]")); }
      if (act === "voirfacture") { const f = window.factureTab.avenantFacture(this.id); if (f) window.factureTab.showFacture(f.id); }
      if (act === "del") {
        const b = g("[data-act]");
        if (!b.classList.contains("armed")) { b.classList.add("armed"); b.textContent = "Confirmer la suppression"; setTimeout(() => { b.classList.remove("armed"); b.textContent = "Supprimer l'avenant"; }, 4000); return; }
        if (!canEdit()) return;
        clearTimeout(this.timer); this.dirty = false;
        await store.remove(COL, this.id).catch(err => toast(frError(err)));
        this.id = null; this.draft = null; toast("Avenant supprimé"); return "deleted";
      }
    },
    async facturer() { window.factureTab.openCreateAvenant(this.id); },

    /* -------------------------- téléphone -------------------------- */
    // rendu dans el ; le devis garde la main pour le bouton « ‹ Devis » (data-act="avback")
    renderM(el, devis) {
      const k = K(), a0 = this.id && store.get(COL, this.id); if (!a0) return false;
      const strip = this.tabsHtml(a0.devisId, devis?.numero, this.id);
      if (this.m.screen === "edit") {
        const a = this.draft && this.draft.id === this.id ? this.draft : (this.draft = normalize(a0));
        if (!a.lignes.length) a.lignes.push({ d: "", pu: "", q: 1 });
        const cu = cumul(a);
        el.innerHTML = `<div class="dv dv-m"><button class="btn dv-back" data-act="avback">‹ Retour au devis</button>${strip}
          <p class="dv-q">Qu'est-ce qui change ?</p><p class="dv-sub">Avenant n° ${a.rang} au devis ${esc(a.devisNumero)} · ${esc(a.client.nom)}</p>
          ${a.lignes.map((l, i) => { const moins = l._s === -1 || k.num(l.pu) < 0;
            return `<div class="dv-card dv-line">
              <div class="dv-seg av-sign"><button data-sign="${i}|1" class="${moins ? "" : "on"}">+ En plus</button><button data-sign="${i}|-1" class="${moins ? "on" : ""}">− En moins</button></div>
              <div class="field"><label for="al${i}">Désignation</label><textarea id="al${i}" data-am="lignes.${i}.d" rows="2">${esc(l.d)}</textarea></div>
              <div class="row"><div class="field"><label for="ap${i}">Prix unitaire HT (€)</label><input id="ap${i}" data-am="lignes.${i}.pu" inputmode="decimal" value="${esc(l.pu === "" ? "" : String(Math.abs(k.num(l.pu))).replace(".", ","))}"></div>
                <div class="field"><label>Quantité</label><div class="stepper"><button data-aqty="${i}|-1" aria-label="Moins">−</button><input data-am="lignes.${i}.q" inputmode="decimal" value="${esc(String(l.q).replace(".", ","))}" aria-label="Quantité"><button data-aqty="${i}|1" aria-label="Plus">+</button></div></div></div>
              <div class="lt"><button class="del" data-amdel="${i}">Retirer</button><span>Total HT <b data-amcalc="${i}">${k.eur.format(k.num(l.pu) * k.num(l.q))}</b></span></div></div>`; }).join("")}
          <button class="btn big" data-act="maddline">+ Ajouter une ligne</button></div>
          <div class="dv-nav"><button class="btn" data-act="avback">‹ Devis</button>
            <div class="tt">Avenant TTC<b id="amTot">${signed(cu.cur)}</b></div>
            <button class="btn primary" data-act="mview">Voir l'avenant</button></div>`;
        el.querySelectorAll(".dv-m textarea").forEach(K().grow);
        return true;
      }
      const a = normalize(a0), t = k.calc(a), cu = cumul(a), facts = window.factureTab?.avenantFactures(a.id) || [];
      el.innerHTML = `<div class="dv dv-m"><button class="btn dv-back" data-act="avback">‹ Retour au devis</button>${strip}
        <p class="dv-q">Avenant ${esc(a.numero)}</p><p class="dv-sub">${esc(a.client.nom)} · avenant n° ${a.rang} au devis ${esc(a.devisNumero)}</p>
        <div class="dv-prev" id="mPrev">${page(a, false)}</div>
        <div class="dv-card dv-sum"><div><span>Devis (TTC)</span><b>${k.eur.format(cu.devis)}</b></div>
          ${cu.prev ? `<div><span>Avenants précédents</span><b>${signed(cu.prev)}</b></div>` : ""}
          <div><span>Cet avenant (TTC)</span><b>${signed(cu.cur)}</b></div>
          <div class="ttc"><span>Nouveau total du chantier</span><b>${k.eur.format(cu.total)}</b></div></div>
        <div class="dv-card"><div class="field"><label>Statut de l'avenant</label></div>
          <div class="dv-seg dv-statut">${Object.entries(STATUTS).map(([s, v]) => `<button data-amstatut="${s}" class="st-${s} ${a.statut === s ? "on" : ""}">${v}</button>`).join("")}</div></div>
        <button class="btn primary big" data-act="mmail">Envoyer par mail</button>
        <button class="btn big" data-act="mpdf">Enregistrer le PDF</button>
        <button class="btn big" data-act="mfacturer">Transformer en facture d'avenant</button>
        ${facts.length ? `<button class="btn big" data-act="mvoirfacture">Voir ${facts.length > 1 ? `les ${facts.length} factures` : `la facture ${esc(facts[0].numero)}`}</button>` : ""}
        <div class="dv-edit"><span class="dv-sub" style="margin:0;align-self:center">Modifier :</span><button data-act="medit">Les lignes</button></div>
        <button class="btn danger" data-act="mdelete">Supprimer cet avenant</button></div>
        <div class="dv-nav"><button class="btn" data-act="avback">‹ Devis</button><div class="tt">Avenant TTC<b>${signed(t.ttc)}</b></div></div>`;
      k.fitPreview();
      return true;
    },
    async saveM(patch) {
      if (!canEdit()) return false;
      try { await store.patch(COL, this.id, patch); return true; } catch (e) { toast(frError(e)); return false; }
    },
    async saveDraftM() {
      if (!this.draft) return true;
      this.draft = K().cleanLines(this.draft);
      return this.saveM({ lignes: this.draft.lignes.map(({ d, pu, q }) => ({ d, pu, q })) });
    },
    onInputM(e) {
      const t = e.target, f = t.dataset.am; if (!f || !this.draft) return;
      let v = t.value;
      const mm = f.match(/^lignes\.(\d+)\.(pu|q|d)$/), i = mm ? +mm[1] : -1, l = this.draft.lignes[i];
      if (mm && mm[2] === "pu") { const neg = l._s === -1 || K().num(l.pu) < 0; v = v.replace(",", ".").replace(/^[-−]/, ""); v = v === "" ? "" : (neg ? "-" + v : v); }
      if (mm && mm[2] === "q") v = v.replace(",", ".");
      K().setPath(this.draft, f, v);
      if (t.tagName === "TEXTAREA") K().grow(t);
      const c = document.querySelector(`[data-amcalc="${i}"]`); if (c && l) c.textContent = K().eur.format(K().num(l.pu) * K().num(l.q));
      const tot = $("#amTot"); if (tot) tot.textContent = signed(cumul(this.draft).cur);
    },
    // renvoie "rerender" (réafficher l'avenant), "deleted" ou rien
    async onClickM(e) {
      const g = s => e.target.closest(s), k = K();
      if (g("[data-sign]")) { const [i, s] = g("[data-sign]").dataset.sign.split("|").map(Number), l = this.draft.lignes[i]; l._s = s; if (l.pu !== "") l.pu = String(s * Math.abs(k.num(l.pu))); return "rerender"; }
      if (g("[data-aqty]")) { const [i, s] = g("[data-aqty]").dataset.aqty.split("|").map(Number), l = this.draft.lignes[i]; l.q = Math.max(0, k.num(l.q) + s); return "rerender"; }
      if (g("[data-amdel]")) { this.draft.lignes.splice(+g("[data-amdel]").dataset.amdel, 1); return "rerender"; }
      if (g("[data-amstatut]")) { const s = g("[data-amstatut]").dataset.amstatut; if (await this.saveM({ statut: s })) toast("Statut : " + STATUTS[s]); return "rerender"; }
      const act = g("[data-act]")?.dataset.act; if (!act) return;
      if (act === "maddline") { this.draft.lignes.push({ d: "", pu: "", q: 1 }); return "rerender-focus"; }
      if (act === "mview") {
        if (!k.cleanLines(this.draft).lignes.some(l => String(l.d).trim())) { toast("Écrivez au moins une ligne (désignation)."); return; }
        if (!(await this.saveDraftM())) return; this.m.screen = "view"; window.scrollTo(0, 0); return "rerender";
      }
      if (act === "medit") { this.draft = normalize(store.get(COL, this.id)); this.m.screen = "edit"; window.scrollTo(0, 0); return "rerender"; }
      const a = this.id && store.get(COL, this.id) ? k.cleanLines(normalize(store.get(COL, this.id))) : null; if (!a) return;
      if (act === "mpdf") { try { const r = await k.savePDF(a, pdfOpts(a)); if (r === "saved") toast("PDF enregistré : " + fileName(a)); } catch (err) { if (err?.code !== "declined") toast(err.message || String(err)); } }
      if (act === "mmail") k.openMail(a, async () => { if (a.statut === "brouillon") await this.saveM({ statut: "envoye" }); toast("Avenant envoyé"); }, pdfOpts(a));
      if (act === "mfacturer") return this.facturer(g("[data-act]"));
      if (act === "mvoirfacture") { const f = window.factureTab.avenantFacture(a.id); if (f) window.factureTab.showFacture(f.id); }
      if (act === "mdelete") {
        const b = g("[data-act]");
        if (!b.classList.contains("armed")) { b.classList.add("armed"); b.textContent = "Confirmer la suppression"; return; }
        await store.remove(COL, this.id).catch(err => toast(frError(err)));
        this.id = null; this.draft = null; toast("Avenant supprimé"); return "deleted";
      }
    },
  };

  window.avenantsView = A;
  window.avenantsKit = { ofDevis, ttcOf, normalize, round2 };
})();
