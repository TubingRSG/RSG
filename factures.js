/* =====================================================================
   RSG — onglet « Création de facture »
   Une facture se crée à partir d'un devis :
   • facture identique au devis (toutes les lignes ; les acomptes déjà facturés
     sur ce devis sont déduits du net à payer),
   • ou facture d'acompte : un pourcentage du devis.
   PC : la page A4 de la facture, modifiable comme dans Excel. Téléphone : liste,
   aperçu, PDF, mail et statut (À envoyer / Envoyée / Payée).
   Données : collection « factures » (même base que le reste de RSG).
   ===================================================================== */
"use strict";
(function () {
  const K = () => window.devisKit;
  const COL = "factures";
  const STATUTS = { a_envoyer: "À envoyer", envoyee: "Envoyée", payee: "Payée" };
  const STCSS = { a_envoyer: "brouillon", envoyee: "envoye", payee: "accepte" };
  const NB_LIGNES = 8;
  const ECHEANCES = [["À réception", 0], ["15 jours", 15], ["30 jours", 30]];
  const round2 = v => Math.round(v * 100) / 100;
  const addDays = (iso, n) => { const d = new Date(iso + "T12:00:00"); d.setDate(d.getDate() + n); return new Date(d.getTime() - d.getTimezoneOffset() * 6e4).toISOString().slice(0, 10); };
  const pctFr = v => String(round2(v)).replace(".", ",") + " %";

  function nextNumero(iso, selfId) {
    if (!/^\d{4}-\d{2}-\d{2}$/.test(iso || "")) return "";
    const [y, m, dd] = iso.split("-"), pre = `${y}/${m}${dd}-F`;
    const self = selfId && store.get(COL, selfId);
    if (self && String(self.numero || "").startsWith(pre)) return self.numero;
    const max = store.list(COL).filter(x => x.id !== selfId && String(x.numero || "").startsWith(pre))
      .reduce((mx, x) => Math.max(mx, parseInt(String(x.numero).slice(pre.length), 10) || 0), 0);
    return pre + (max + 1);
  }
  function normalize(f) {
    const k = K(), x = Object.assign({ numero: "", date: k.isoToday(), echeance: "", type: "totale", pct: 100, devisId: null, devisNumero: "", devisDate: "",
      statut: "a_envoyer", clientId: null, tva: 0.2, deduit: 0, devisHT: 0, devisTTC: 0 }, k.clone(f));
    x.client = Object.assign({ nom: "", adr1: "", adr2: "", ref: "", email: "" }, x.client || {});
    x.lignes = (x.lignes || []).map(l => ({ d: l.d || "", pu: l.pu ?? "", q: l.q ?? "" }));
    return x;
  }
  const totals = f => { const t = K().calc(f), ded = K().num(f.deduit); return { ...t, ded, net: round2(t.ttc - ded) }; };
  const cleanName = t => K().cleanName(t);
  const fileName = f => ["Facture RSG", f.client?.nom, f.client?.ref, String(f.numero || "").replace(/\//g, "-")].map(cleanName).filter(Boolean).join(" - ") + ".pdf";
  const facturesOf = devisId => store.list(COL).filter(f => f.devisId === devisId);
  const acomptesTTC = (devisId, exceptId) => round2(facturesOf(devisId).filter(f => f.type === "acompte" && f.id !== exceptId).reduce((s, f) => s + K().calc(normalize(f)).ttc, 0));
  const isAvAc = f => f.type === "avenant" && f.avPart === "acompte";
  const acTxt = f => f.avancement ? `Avancement ${pctFr(f.avancement)}` : `Acompte ${pctFr(f.pct)}`;
  const typeTxt = f => f.type === "acompte" ? acTxt(f) : isAvAc(f) ? `${acTxt(f)} avenant n° ${f.avenantRang || ""}` : f.type === "avenant" ? `Avenant n° ${f.avenantRang || ""}` : "Facture totale";
  const AV = () => window.avenantsView, AK = () => window.avenantsKit;
  // Toutes les factures d'un même devis sont rangées ensemble (acomptes, facture totale, factures d'avenant) :
  // la liste n'en montre qu'une ligne, et la facture ouverte affiche les autres en sous-onglets.
  const ordre = f => f.type === "acompte" ? 0 : f.type === "totale" ? 1 : 2 + (f.avenantRang || 0) * 2 + (isAvAc(f) ? 0 : 1);
  const familyOf = f => !f?.devisId ? [f] : store.list(COL).filter(x => x.devisId === f.devisId)
    .sort((a, b) => ordre(a) - ordre(b) || (a.avenantRang || 0) - (b.avenantRang || 0) || (a.date || "").localeCompare(b.date || "") || String(a.numero).localeCompare(String(b.numero), "fr", { numeric: true }));
  const topOf = f => familyOf(f)[0];
  const tabLabel = x => isAvAc(x) ? `${acTxt(x)} avenant n° ${x.avenantRang}` : x.type === "avenant" ? `Facture d'avenant n° ${x.avenantRang}` : x.type === "acompte" ? (x.avancement ? `Acompte à l'avancement ${pctFr(x.avancement)}` : `Facture d'acompte ${pctFr(x.pct)}`) : "Facture totale";
  function familyTabs(f, attr) {
    const fam = familyOf(f); if (fam.length < 2) return "";
    return `<div class="dv-subtabs"><span class="dv-subtl">Devis ${esc(f.devisNumero)} :</span>${fam.map(x => `<button ${attr}="${x.id}" class="${x.id === f.id ? "on" : ""}"><b>${tabLabel(x)}</b><small>${esc(x.numero)}</small></button>`).join("")}</div>`;
  }

  /* ---------- la page A4 de la facture ---------- */
  function page(f, edit) {
    const k = K(), { ENT, eur, num, money, pctTxt, frDate } = k, c = f.client, t = totals(f);
    const inp = (path, val, o = {}) => edit
      ? `<input class="f ${o.cls || ""}" data-f="${path}" value="${esc(val ?? "")}" ${o.type ? `type="${o.type}"` : ""} inputmode="${o.im || "text"}" placeholder="${esc(o.ph || "")}" autocomplete="off">`
      : esc(o.show ? o.show(val) : (val ?? ""));
    const rows = Math.max(NB_LIGNES, f.lignes.length + (edit ? 1 : 0));
    let body = "";
    for (let i = 0; i < rows; i++) {
      const l = f.lignes[i] || { d: "", pu: "", q: "" };
      const has = String(l.d).trim() || num(l.pu) || num(l.q);
      const tot = has && l.pu !== "" && l.q !== "" ? eur.format(num(l.pu) * num(l.q)) : "";
      body += `<tr data-row="${i}">
        <td>${edit ? `<textarea class="f" data-f="lignes.${i}.d" rows="1">${esc(l.d)}</textarea>` : String(l.d).split(/\r?\n/).map(x => `<div class="dl">${esc(x) || "&nbsp;"}</div>`).join("")}</td>
        <td class="n">${edit ? inp(`lignes.${i}.pu`, money(l.pu), { cls: "num-in money", im: "decimal" }) : (has && l.pu !== "" ? eur.format(num(l.pu)) : "")}</td>
        <td class="q">${edit ? inp(`lignes.${i}.q`, l.q, { cls: "c", im: "decimal" }) : (has ? esc(String(l.q).replace(".", ",")) : "")}</td>
        <td class="t" data-calc="l${i}">${tot}${edit && i < f.lignes.length ? `<button class="rowdel noprint" data-del="${i}" title="Supprimer la ligne" aria-label="Supprimer la ligne">×</button>` : ""}</td></tr>`;
    }
    const tvaSel = edit
      ? `<select class="f" data-f="tva" aria-label="Taux de TVA">${k.TVAS.map(v => `<option value="${v}" ${v === f.tva ? "selected" : ""}>${pctTxt(v)}</option>`).join("")}</select>`
      : pctTxt(f.tva);
    const duDevis = `devis n° ${esc(f.devisNumero)}${f.devisDate ? ` du ${frDate(f.devisDate)}` : ""}`;
    const inclus = (f.avenantsInclus || []).map(x => x.rang);
    const origin = f.devisNumero
      ? (f.type === "acompte" ? (f.avancement ? `Facture d'acompte à l'avancement (travaux réalisés à ${pctFr(f.avancement)}) sur le ${duDevis}` : `Facture d'acompte de ${pctFr(f.pct)} sur le ${duDevis}`)
        : isAvAc(f) ? `Facture d'acompte ${f.avancement ? `à l'avancement (travaux réalisés à ${pctFr(f.avancement)})` : `de ${pctFr(f.pct)}`} sur l'avenant n° ${f.avenantRang} (${esc(f.avenantNumero)}) au ${duDevis}`
        : f.type === "avenant" ? `Facture de l'avenant n° ${f.avenantRang} (${esc(f.avenantNumero)}) au ${duDevis}`
        : `Facture établie d'après le ${duDevis}${inclus.length ? ` et ${inclus.length > 1 ? "ses avenants n° " + inclus.join(", ") : "son avenant n° " + inclus[0]}` : ""}`)
      : "";
    return `<div class="a4 ft-a4 ${edit ? "" : "static"}">
      <div class="hd">
        <div class="co"><img class="logo" src="${k.LOGO}" alt="RSG">
          <div>${ENT.adresse}</div><div>Téléphone : ${ENT.tel}</div><div>Mail : ${ENT.mail}</div>
          <small>${ENT.ape}</small><small>${ENT.siret}</small><small>${ENT.tva}</small></div>
        <div>
          <div class="title"><b>FACTURE N°</b><span class="num" data-calc="num" title="Calculé à partir de la date">${esc(f.numero || "—")}</span></div>
          <div class="meta"><span>En date du</span><em>${inp("date", f.date, { type: "date", show: frDate, cls: "c" })}</em></div>
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
      ${origin ? `<div class="ft-origin">${origin}</div>` : ""}
      <table class="ft-table"><colgroup><col class="c1"><col class="c2"><col class="c3"><col class="c4"></colgroup>
        <thead><tr><th>Désignation des produits et prestations</th><th>Prix unitaire HT</th><th>Quantité</th><th>Total HT</th></tr></thead>
        <tbody>${body}</tbody></table>
      ${edit ? `<div class="addrow noprint"><button data-act="addline">+ Ajouter une ligne</button></div>` : ""}
      ${isAvAc(f) && f.avTTC != null ? `<div class="ft-recap">Montant total de l'avenant : <b>${eur.format(f.avHT)} HT</b> soit <b>${eur.format(f.avTTC)} TTC</b> · ${f.avancement ? `avancement : <b>${pctFr(f.avancement)}</b>` : `acompte demandé : <b>${pctFr(f.pct)}</b>`}</div>` : ""}
      ${f.type === "acompte" && f.devisTTC ? `<div class="ft-recap">Montant total du devis : <b>${eur.format(f.devisHT)} HT</b> soit <b>${eur.format(f.devisTTC)} TTC</b> · ${f.avancement ? `avancement : <b>${pctFr(f.avancement)}</b>` : `acompte demandé : <b>${pctFr(f.pct)}</b>`}</div>` : ""}
      <div class="tot">
        <div><span>Total Hors Taxe (HT)</span><b data-calc="ht">${eur.format(t.ht)}</b></div>
        <div><span>Total TVA (${tvaSel})</span><b data-calc="tv">${eur.format(t.tv)}</b></div>
        <div class="${t.ded ? "ttc2" : "ttc"}"><span>${t.ded ? "Total TTC" : "NET À PAYER TTC"}</span><b data-calc="ttc">${eur.format(t.ttc)}</b></div>
        ${t.ded ? `<div><span>Acompte(s) déjà facturé(s)</span><b data-calc="ded">− ${eur.format(t.ded)}</b></div>
        <div class="ttc"><span>NET À PAYER</span><b data-calc="net">${eur.format(t.net)}</b></div>` : ""}
      </div>
      ${f.tva === 0 ? `<div class="autoliq">${k.AUTOLIQ}</div>` : ""}
      <div class="ft-ech"><span>Date d'échéance du règlement :</span><b>${inp("echeance", f.echeance, { type: "date", show: frDate, cls: "c" })}</b></div>
      <div class="foot">${ENT.paiement}<br>${ENT.tva}</div>
    </div>`;
  }

  function mailTexts(f) {
    const k = K(), t = totals(f), c = f.client;
    const quoi = f.type === "acompte" ? "facture d'acompte" : "facture";
    return {
      to: c.email || "",
      subject: `Facture n° ${f.numero} – RSG${c.ref ? " – " + c.ref : ""}`,
      body: `Bonjour${c.nom ? " " + c.nom : ""},\n\nVeuillez trouver ci-joint notre ${quoi} n° ${f.numero}${c.ref ? ` pour « ${c.ref} »` : ""}, d'un montant de ${k.eur.format(t.net)} TTC à régler${f.echeance ? ` avant le ${k.frDate(f.echeance)}` : ""}, par virement (coordonnées bancaires sur la facture).\n\nNous vous remercions de votre confiance.\n\nCordialement,\nRSG – Réalisations Soudures Générales\n${k.ENT.tel} – ${k.ENT.mail}`,
    };
  }
  const pdfOpts = f => ({ html: page(f, false), title: `Facture ${f.numero}`, name: fileName(f), texts: mailTexts(f), label: "facture" });

  /* ---------- acompte : pourcentage fixe ou « à l'avancement » ---------- */
  // st.mode = "pct" (pourcentage du montant) ou "avct" (travaux réalisés à X %, moins les acomptes déjà facturés)
  function acompteAmount(st, t, dejaHT) {
    const ht = st.mode === "avct" ? round2(t.ht * st.avct / 100 - dejaHT) : round2(t.ht * st.pct / 100);
    return { ht, ttc: round2(ht * (t.ht ? t.ttc / t.ht : 1)), pct: t.ht ? round2(ht / t.ht * 100) : 0 };
  }
  function acompteBlock(st, ac, t, dejaHT) {
    const k = K(), m = acompteAmount(st, t, dejaHT), avct = st.mode === "avct";
    return `<div class="ft-pct">
      <div class="dv-chips">${ac.map(([lab, p], i) => `<button data-fpct="${i}" class="${!avct && st.idx === i ? "on" : ""}">${esc(lab)} · ${p} %</button>`).join("")}
        <button data-favct class="${avct ? "on" : ""}">Acompte à l'avancement</button></div>
      <div class="ft-pctrow"><label for="ftPctIn">${avct ? "Avancement des travaux" : "Pourcentage"}</label><input id="ftPctIn" inputmode="decimal" value="${String(avct ? st.avct : st.pct).replace(".", ",")}"><span>%</span></div>
      <div class="ft-amt" id="ftAmtBox">${acompteTxt(st, t, dejaHT)}</div></div>`;
  }
  function acompteTxt(st, t, dejaHT) {
    const k = K(), m = acompteAmount(st, t, dejaHT);
    if (st.mode !== "avct") return `Montant de l'acompte : <b>${k.eur.format(m.ht)} HT · ${k.eur.format(m.ttc)} TTC</b>`;
    const fait = round2(t.ht * st.avct / 100);
    return `Travaux réalisés à ${pctFr(st.avct)} : ${k.eur.format(fait)} HT${dejaHT ? ` − acomptes déjà facturés ${k.eur.format(dejaHT)} HT` : ""}<br>
      ${m.ht > 0 ? `À facturer : <b>${k.eur.format(m.ht)} HT · ${k.eur.format(m.ttc)} TTC</b>` : `<b class="pt-warn">Rien à facturer : les acomptes déjà facturés couvrent cet avancement.</b>`}`;
  }
  function acompteLine(st, label, sur, ref, m) {
    return st.mode === "avct"
      ? `Acompte à l'avancement (travaux réalisés à ${pctFr(st.avct)}) sur ${sur}${ref ? ` – ${ref}` : ""}`
      : `${label ? label + " (" + pctFr(st.pct) + ")" : "Acompte de " + pctFr(st.pct)} sur ${sur}${ref ? ` – ${ref}` : ""}`;
  }

  /* ---------- création à partir d'un devis ---------- */
  function fromDevis(dv, type, pct, label, avs = []) {
    const k = K(), d = k.cleanLines(k.normalize(dv)), t = k.calc(d), date = k.isoToday(), ref = d.client.ref;
    const f = {
      numero: nextNumero(date), date, echeance: addDays(date, 30), type, pct: type === "acompte" ? pct : 100,
      devisId: dv.id, devisNumero: d.numero, devisDate: d.date, statut: "a_envoyer", clientId: d.clientId || null,
      client: { nom: d.client.nom, adr1: d.client.adr1, adr2: d.client.adr2, ref, email: d.client.email },
      tva: d.tva, devisHT: round2(t.ht), devisTTC: round2(t.ttc), deduit: 0,
    };
    if (type === "acompte") {
      f.lignes = [{ d: `${label ? label + " (" + pctFr(pct) + ")" : "Acompte de " + pctFr(pct)} sur le devis n° ${d.numero} du ${k.frDate(d.date)}${ref ? ` – ${ref}` : ""}`, pu: round2(t.ht * pct / 100), q: 1 }];
    } else {
      f.lignes = d.lignes.map(l => ({ ...l }));
      for (const a of avs) {
        const n = k.cleanLines(AK().normalize(a));
        n.lignes.forEach(l => f.lignes.push({ d: `Avenant n° ${n.rang} – ${l.d}`, pu: l.pu, q: l.q }));   // (plus proposé : les avenants se facturent à part)
      }
      if (avs.length) f.avenantsInclus = avs.map(a => ({ id: a.id, rang: a.rang, numero: a.numero }));
      f.deduit = acomptesTTC(dv.id);
    }
    return f;
  }

  /* =================================================================
     L'onglet
     ================================================================= */
  const tab = {
    id: "factures",
    title: "Création de facture",
    collections: ["factures", "avenants", "devis", "clients"],
    el: null, mode: null, sub: "factures",
    selId: null, draft: null, dirty: false, saving: false, timer: null, q: "",
    m: { screen: "list", id: null },

    // Sauvegarde PDF (Réglages) : un PDF par facture de la période
    exportPdf(p) {
      return this.sorted().filter(f => !p || (String(f.date || "") >= p.from && String(f.date || "") <= p.to)).map(f => { const n = K().cleanLines(normalize(f));
        return { path: "Factures/" + fileName(n), make: () => K().makePDF(n, pdfOpts(n)) }; }).concat(AV() ? AV().exportPdf(p) : []);
    },
    // avenants : facture liée, ouverture d'une facture, facturation d'un avenant
    avenantFacture(avId) { return this.avenantFactures(avId)[0]; },
    avenantFactures(avId) { return store.list(COL).filter(f => f.avenantId === avId).sort((a, b) => ordre(a) - ordre(b) || (a.date || "").localeCompare(b.date || "")); },
    showFacture(id) {
      this.pickMode = false; this.selId = id; this.draft = normalize(store.get(COL, id)); this.m = { screen: "view", id };
      if (current === this) this.mount(this.el); else showTab("factures");
      window.scrollTo(0, 0);
    },
    // part = "totale" (facture identique à l'avenant, acomptes déjà facturés déduits) ou "acompte" (pct %)
    async createFromAvenant(a, part = "totale", pct = 100, label = "", st = null) {
      if (!canEdit()) return;
      const k = K(), date = k.isoToday(), t = k.calc(a);
      const f = { numero: nextNumero(date), date, echeance: addDays(date, 30), type: "avenant", avPart: part, pct: part === "acompte" ? pct : 100,
        avenantId: a.id, avenantNumero: a.numero, avenantRang: a.rang, avHT: round2(t.ht), avTTC: round2(t.ttc),
        devisId: a.devisId, devisNumero: a.devisNumero, devisDate: a.devisDate, statut: "a_envoyer", clientId: a.clientId || null,
        client: { ...a.client }, tva: a.tva, devisHT: a.devisHT, devisTTC: a.devisTTC, deduit: 0 };
      if (part === "acompte") {
        const dejaHT = round2(this.avenantFactures(a.id).filter(isAvAc).reduce((s, x) => s + K().calc(normalize(x)).ht, 0));
        const S = st || { mode: "pct", pct }, m = acompteAmount(S, t, dejaHT);
        if (S.mode === "avct") { f.avancement = S.avct; f.pct = m.pct; }
        f.lignes = [{ d: acompteLine(S, label, `l'avenant n° ${a.rang} (${a.numero}) au devis n° ${a.devisNumero}`, a.client.ref, m), pu: m.ht, q: 1 }];
      } else {
        f.lignes = a.lignes.map(({ d, pu, q }) => ({ d, pu, q }));
        f.deduit = round2(this.avenantFactures(a.id).filter(isAvAc).reduce((s, x) => s + K().calc(normalize(x)).ttc, 0));
      }
      try {
        const n = await store.create(COL, f);
        if (a.statut !== "accepte") await store.patch("avenants", a.id, { statut: "accepte" }).catch(() => {});
        toast(`Facture ${f.numero} créée${part === "acompte" ? ` (acompte ${pctFr(pct)})` : ""} pour l'avenant n° ${a.rang}`);
        this.showFacture(n.id);
      } catch (e) { toast(frError(e)); }
    },
    // fenêtre « Transformer l'avenant en facture » : identique ou acompte
    openCreateAvenant(avId) {
      const k = K(), a0 = store.get("avenants", avId); if (!a0) return toast("Avenant introuvable.");
      if (!canEdit()) return;
      const a = k.cleanLines(AK().normalize(a0)), t = k.calc(a), dv = store.get("devis", a.devisId);
      if (!a.lignes.length) return toast("L'avenant n'a encore aucune ligne.");
      const deja = this.avenantFactures(avId), acTTC = round2(deja.filter(isAvAc).reduce((s, f) => s + k.calc(normalize(f)).ttc, 0));
      const dejaTTC = round2(deja.reduce((s, f) => s + totals(normalize(f)).net, 0)), reste = round2(t.ttc - dejaTTC);
      const ac = dv ? k.ACOMPTES.map((lab, i) => [lab.replace(/\s*:$/, ""), k.num(k.normalize(dv).acomptes[i])]).filter(([, p]) => p > 0) : [];
      let dlg = $("#ftNew");
      if (!dlg) { document.body.insertAdjacentHTML("beforeend", `<dialog id="ftNew"><div class="dlg" id="ftNewBody"></div></dialog>`); dlg = $("#ftNew"); }
      const avHT = round2(deja.filter(isAvAc).reduce((s, f) => s + k.calc(normalize(f)).ht, 0));
      const st = { type: "totale", pct: ac[0]?.[1] || 30, idx: ac.length ? 0 : -1, mode: "pct", avct: 50 };
      const draw = () => {
        $("#ftNewBody").innerHTML = `
          <h3>Facturer l'avenant n° ${a.rang} (${esc(a.numero)})</h3>
          <div class="small">${esc(a.client.nom || "Client")}${a.client.ref ? " · " + esc(a.client.ref) : ""} · devis ${esc(a.devisNumero)} · avenant <b>${k.eur.format(t.ttc)} TTC</b></div>
          ${a.statut !== "accepte" ? `<p class="pt-warn">L'avenant n'est pas encore marqué « Accepté » : il le sera à la création de la facture.</p>` : ""}
          ${deja.length ? `<div class="ft-deja"><b>Déjà facturé sur cet avenant</b>${deja.map(f => `<div><span>${esc(f.numero)} · ${esc(typeTxt(f))}</span><b>${k.eur.format(totals(normalize(f)).net)}</b></div>`).join("")}
            <div class="reste"><span>Reste à facturer</span><b>${k.eur.format(Math.max(0, reste))}</b></div></div>` : ""}
          <div class="ft-choice">
            <button data-ftype="totale" class="${st.type === "totale" ? "on" : ""}"><b>Facture identique à l'avenant</b>
              <span>Toutes les lignes de l'avenant, ${k.eur.format(t.ttc)} TTC.${acTTC ? ` Les acomptes déjà facturés (${k.eur.format(acTTC)}) sont déduits : net à payer ${k.eur.format(round2(t.ttc - acTTC))}.` : ""}</span></button>
            <button data-ftype="acompte" class="${st.type === "acompte" ? "on" : ""}"><b>Facture d'acompte</b><span>Un pourcentage de l'avenant.</span></button>
          </div>
          ${st.type === "acompte" ? acompteBlock(st, ac, t, avHT) : ""}
          <div class="err" id="ftErr"></div>
          <div class="foot"><span class="spacer"></span><button class="btn" data-fnew="cancel">Annuler</button><button class="btn primary" data-fnew="ok">Créer la facture</button></div>`;
      };
      draw();
      $("#ftNewBody").onclick = async ev => {
        const g = s => ev.target.closest(s);
        if (g("[data-ftype]")) { st.type = g("[data-ftype]").dataset.ftype; return draw(); }
        if (g("[data-fpct]")) { st.mode = "pct"; st.idx = +g("[data-fpct]").dataset.fpct; st.pct = ac[st.idx][1]; return draw(); }
        if (g("[data-favct]")) { st.mode = "avct"; st.idx = -1; return draw(); }
        const b = g("[data-fnew]"); if (!b) return;
        if (b.dataset.fnew === "cancel") return dlg.close();
        if (st.type === "acompte") {
          const v = st.mode === "avct" ? st.avct : st.pct;
          if (!(v > 0 && v <= 100)) { $("#ftErr").textContent = "Indiquez un pourcentage entre 1 et 100."; return; }
          if (acompteAmount(st, t, avHT).ht <= 0) { $("#ftErr").textContent = "Rien à facturer pour cet avancement."; return; }
        }
        b.disabled = true; dlg.close();
        await this.createFromAvenant(a, st.type, st.pct, st.type === "acompte" && st.mode === "pct" && st.idx >= 0 ? ac[st.idx][0] : "", st);
      };
      $("#ftNewBody").oninput = ev => {
        if (ev.target.id !== "ftPctIn") return;
        if (st.mode === "avct") st.avct = k.num(ev.target.value); else { st.pct = k.num(ev.target.value); st.idx = -1; $("#ftNewBody").querySelectorAll("[data-fpct]").forEach(x => x.classList.remove("on")); }
        const el = $("#ftAmtBox"); if (el) el.innerHTML = acompteTxt(st, t, avHT);
      };
      if (!dlg.open) dlg.showModal();
    },

    mount(el) {
      this.el = el; this.mode = K().mobile() ? "m" : "pc";
      if (!this.bound) {
        this.bound = true;
        el.addEventListener("click", e => current === this && this.route(e, "onClick"));
        el.addEventListener("input", e => current === this && this.route(e, "onInput"));
        el.addEventListener("change", e => current === this && this.route(e, "onChange"));
        el.addEventListener("focusin", e => { if (current === this && e.target.classList?.contains("money")) { const v = K().num(e.target.value); e.target.value = e.target.value === "" ? "" : String(v).replace(".", ","); e.target.select(); } });
        el.addEventListener("focusout", e => { if (current === this && e.target.classList?.contains("money")) e.target.value = K().money(e.target.value.replace(",", ".")); });
        matchMedia("(max-width: 760px)").addEventListener("change", () => { if (current === this) this.mount(this.el); });
      }
      if (this.mode === "pc") this.mountPC(); else this.renderM();
    },
    // aiguillage : sous-onglet Factures ou Avenants
    route(e, fn) { return this[fn](e); },
    render() {
      if (!this.el || current !== this) return;
      if (this.mode === "pc") {
        this.renderList();
        const inside = this.el.querySelector(".a4")?.contains(document.activeElement);
        if (!this.selId && !this.draft) { const first = this.sorted()[0]; if (first) { this.selId = first.id; this.draft = normalize(first); this.renderList(); this.renderPage(); } return; }
        if (this.selId && !this.dirty && !this.saving && !inside) {
          const fresh = store.get(COL, this.selId);
          if (fresh && (!this.draft || JSON.stringify(K().cleanLines(normalize(fresh))) !== JSON.stringify(K().cleanLines(this.draft)))) { this.draft = normalize(fresh); this.renderPage(); }
          if (!fresh) { this.selId = null; this.draft = null; this.renderPage(); }
        }
      } else if (this.m.screen !== "view") this.renderM();
    },
    sorted() {
      return store.list(COL).sort((a, b) => (b.date || "").localeCompare(a.date || "") || String(b.numero).localeCompare(String(a.numero), "fr", { numeric: true }));
    },

    /* ----------------------------- PC ----------------------------- */
    mountPC() {
      this.el.innerHTML = `<div class="dv ft"><div class="dv-pc">
        <aside class="dv-side"><div class="top">
          <button class="btn primary" data-act="pick">+ Facturer un devis</button>
          <input id="ftSearch" placeholder="Rechercher (client, n°, chantier)" value="${esc(this.q)}" aria-label="Rechercher une facture">
        </div><div class="dv-list" id="ftList"></div></aside>
        <section class="dv-main" id="ftMain"></section>
      </div></div>`;
      if (!this.selId) this.selId = this.sorted()[0]?.id || null;
      this.draft = this.selId && store.get(COL, this.selId) ? normalize(store.get(COL, this.selId)) : null;
      this.renderList(); this.renderPage();
    },
    renderList() {
      const list = $("#ftList"); if (!list) return;
      const q = norm(this.q), eur = K().eur;
      const items = this.sorted().filter(f => topOf(f).id === f.id).filter(f => !q || familyOf(f).some(x => norm([x.numero, x.client?.nom, x.client?.ref, x.devisNumero].join(" ")).includes(q)));
      const selTop = this.selId && store.get(COL, this.selId) ? topOf(store.get(COL, this.selId)).id : null;
      list.innerHTML = items.length ? items.map(f => `<button class="dv-item ${f.id === selTop ? "on" : ""}" data-open="${f.id}">
          <b>${esc(f.numero)}</b><em>${eur.format(totals(normalize(f)).net)}</em>
          <span>${esc(f.client?.nom || "Client")}${f.client?.ref ? " · " + esc(f.client.ref) : ""}</span><span class="statut ${STCSS[f.statut] || "brouillon"}">${STATUTS[f.statut] || "À envoyer"}</span>
          <span class="ft-type">${esc(typeTxt(f))}${f.devisNumero ? " · devis " + esc(f.devisNumero) : ""}${familyOf(f).length > 1 ? ` · <b class="dv-avn">+ ${familyOf(f).length - 1} autre${familyOf(f).length > 2 ? "s" : ""} facture${familyOf(f).length > 2 ? "s" : ""} de ce devis</b>` : ""}</span></button>`).join("")
        : `<div class="dv-empty">${store.list(COL).length ? "Aucune facture ne correspond." : "Aucune facture pour l'instant. Cliquez « + Facturer un devis », ou « Facturer » sur un devis."}</div>`;
    },
    renderPage() {
      const y = window.scrollY; this._renderPage(); if (window.scrollY !== y) window.scrollTo(0, y);
    },
    _renderPage() {
      const main = $("#ftMain"); if (!main) return;
      if (this.pickMode || !this.draft) { main.innerHTML = this.pickHtml(); return; }
      const f = this.draft;
      main.innerHTML = familyTabs(f, "data-open") + `<div class="dv-bar">
          <button class="btn primary" data-act="pdf">Enregistrer en PDF sur le PC</button>
          <button class="btn" data-act="mail">Envoyer par mail</button>
          <select class="btn" data-act="statut" aria-label="Statut de la facture">${Object.entries(STATUTS).map(([k, v]) => `<option value="${k}" ${f.statut === k ? "selected" : ""}>${v}</option>`).join("")}</select>
          ${f.devisId && store.get("devis", f.devisId) ? `<button class="btn" data-act="devis">Voir le devis</button>` : ""}
          <button class="btn danger" data-act="del">Supprimer</button>
          <span class="st" id="ftSt"><i></i><span>save</span></span>
        </div>
        <div class="dv-hint">${esc(typeTxt(f))}${f.devisNumero ? ` du devis n° ${esc(f.devisNumero)}` : ""}. Vous pouvez corriger les cases comme dans Excel : tout est enregistré automatiquement.</div>
        <div class="dv-stage">${page(f, true)}</div>`;
      this.el.querySelectorAll(".a4 textarea.f").forEach(K().grow);
    },
    pickHtml() {
      const k = K(), list = store.list("devis").sort((a, b) => (b.date || "").localeCompare(a.date || ""));
      return `<div class="dv-card ft-pick"><h2>Quel devis facturer ?</h2>
        <p class="dv-hint">Choisissez le devis : vous pourrez le facturer à l'identique ou facturer un acompte.</p>
        ${list.length ? `<div class="ft-dlist">${list.map(d => this.devisRow(d)).join("")}</div>` : `<p class="dv-hint">Aucun devis pour l'instant : créez-le d'abord dans l'onglet « Création de devis ».</p>`}
        ${this.draft ? `<button class="btn" data-act="unpick">Annuler</button>` : ""}</div>`;
    },
    devisRow(d) {
      const k = K(), n = k.normalize(d), t = k.calc(n), fs = facturesOf(d.id);
      return `<button class="dv-mitem" data-devis="${d.id}"><b>${esc(n.client.nom || "Client à renseigner")}</b><em>${k.eur.format(t.ttc)}</em>
        <span>${esc(n.numero)}${n.client.ref ? " · " + esc(n.client.ref) : ""}</span>
        <span class="ft-done">${fs.length ? `${fs.length} facture${fs.length > 1 ? "s" : ""}` : ""}</span></button>`;
    },
    status(txt, pending) { const s = $("#ftSt"); if (!s) return; s.classList.toggle("pending", !!pending); s.lastElementChild.textContent = pending ? "modif" : (/Non/.test(txt) ? "erreur" : "save"); },
    queueSave() {
      if (!canEdit()) return;
      this.dirty = true; this.status("Enregistrement…", true);
      clearTimeout(this.timer); this.timer = setTimeout(() => this.flush(), 700);
    },
    async flush() {
      if (!this.dirty || this.saving || !this.selId) return;
      this.dirty = false; this.saving = true; this.status("Enregistrement…", true);
      try { await store.patch(COL, this.selId, K().cleanLines(this.draft)); this.status("Enregistré", false); }
      catch (e) { this.dirty = true; this.status("Non enregistré", true); toast(frError(e)); }
      finally { this.saving = false; if (this.dirty) this.queueSave(); else this.renderList(); }
    },
    updateCalc() {
      const k = K(), f = this.draft, t = totals(f), set = (n, v) => { const e = this.el.querySelector(`[data-calc="${n}"]`); if (e) e.textContent = v; };
      f.lignes.forEach((l, i) => {
        const has = String(l.d).trim() || k.num(l.pu) || k.num(l.q);
        const cell = this.el.querySelector(`[data-calc="l${i}"]`); if (!cell) return;
        const btn = cell.querySelector(".rowdel"); cell.textContent = has && l.pu !== "" && l.q !== "" ? k.eur.format(k.num(l.pu) * k.num(l.q)) : ""; if (btn) cell.appendChild(btn);
      });
      set("ht", k.eur.format(t.ht)); set("tv", k.eur.format(t.tv)); set("ttc", k.eur.format(t.ttc)); set("net", k.eur.format(t.net));
    },
    onInput(e) {
      if (e.target.id === "ftSearch") { this.q = e.target.value; this.renderList(); return; }
      if (this.mode === "m") return;
      const f = e.target.dataset.f; if (!f || !this.draft) return;
      let v = e.target.value;
      if (f === "tva") v = +v;
      if (/^lignes\.\d+\.(pu|q)$/.test(f)) v = v === "" ? "" : v.replace(/[\s  €]/g, "").replace(",", ".");
      if (/^lignes\.(\d+)\./.test(f)) {
        const i = +f.split(".")[1];
        while (this.draft.lignes.length <= i) this.draft.lignes.push({ d: "", pu: "", q: "" });
        const needRow = i === this.draft.lignes.length - 1 && !this.el.querySelector(`tr[data-row="${i + 1}"]`);
        K().setPath(this.draft, f, v);
        if (needRow) { const pos = K().caretOf(e.target); this.renderPage(); K().refocus(this.el, f, pos); }
      } else K().setPath(this.draft, f, v);
      if (f === "date" && v) { this.draft.numero = nextNumero(v, this.selId); const n = this.el.querySelector('[data-calc="num"]'); if (n) n.textContent = this.draft.numero; }
      if (e.target.tagName === "TEXTAREA") K().grow(e.target);
      this.updateCalc(); this.queueSave();
    },
    onChange(e) {
      const t = e.target;
      if (this.mode === "m") {
        if (t.id === "fmEch" && t.value) this.saveM({ echeance: t.value });
        return;
      }
      if (t.dataset.f === "tva") { this.draft.tva = +t.value; this.renderPage(); this.queueSave(); if (this.draft.tva === 0) toast("Mention ajoutée : " + K().AUTOLIQ); }
      if (t.dataset.act === "statut") { this.draft.statut = t.value; this.queueSave(); }
    },
    async onClick(e) {
      const g = s => e.target.closest(s);
      if (g("[data-devis]")) return this.openCreate(g("[data-devis]").dataset.devis);
      if (this.mode === "m") return this.onClickM(e, g);
      if (g("[data-open]")) { await this.flush(); this.pickMode = false; this.selId = g("[data-open]").dataset.open; this.draft = normalize(store.get(COL, this.selId)); this.renderList(); this.renderPage(); return; }
      if (g("[data-del]")) { this.draft.lignes.splice(+g("[data-del]").dataset.del, 1); this.renderPage(); this.updateCalc(); this.queueSave(); return; }
      const a = g("[data-act]")?.dataset.act; if (!a || e.target.tagName === "SELECT") return;
      if (a === "pick") { if (!canEdit()) return; await this.flush(); this.pickMode = true; return this.renderPage(); }
      if (a === "unpick") { this.pickMode = false; return this.renderPage(); }
      if (a === "addline") { this.draft.lignes = K().cleanLines(this.draft).lignes; this.draft.lignes.push({ d: "", pu: "", q: 1 }); this.renderPage(); K().refocus(this.el, `lignes.${this.draft.lignes.length - 1}.d`); }
      if (a === "pdf") { await this.flush(); try { const r = await K().savePDF(this.draft, pdfOpts(K().cleanLines(this.draft))); if (r === "saved") toast("PDF enregistré dans Téléchargements : " + fileName(this.draft)); } catch (err) { if (err?.code !== "declined") toast(err.message || String(err)); } }
      if (a === "mail") { await this.flush(); const f = K().cleanLines(this.draft); K().openMail(f, () => { if (this.draft.statut === "a_envoyer") { this.draft.statut = "envoyee"; this.queueSave(); this.renderPage(); } }, pdfOpts(f)); }
      if (a === "devis") { await this.flush(); const dv = window.devisTab; dv.selId = this.draft.devisId; dv.draft = null; dv.m = { ...dv.m, screen: "list" }; showTab("devis"); }
      if (a === "del") {
        const b = g("[data-act]");
        if (!b.classList.contains("armed")) { b.classList.add("armed"); b.textContent = "Confirmer la suppression"; setTimeout(() => { b.classList.remove("armed"); b.textContent = "Supprimer"; }, 4000); return; }
        if (!canEdit()) return;
        clearTimeout(this.timer); this.dirty = false;
        await store.remove(COL, this.selId).catch(err => toast(frError(err)));
        this.selId = this.sorted()[0]?.id || null; this.draft = this.selId ? normalize(store.get(COL, this.selId)) : null;
        this.renderList(); this.renderPage(); toast("Facture supprimée");
      }
    },

    /* ---------------- fenêtre « Facturer ce devis » (PC et téléphone) ---------------- */
    openCreate(devisId) {
      const k = K(), dv = store.get("devis", devisId); if (!dv) return toast("Devis introuvable.");
      if (!canEdit()) return;
      const d = k.normalize(dv), t = k.calc(d), deja = facturesOf(devisId).sort((a, b) => (a.date || "").localeCompare(b.date || ""));
      const dejaTTC = round2(deja.reduce((s, f) => s + totals(normalize(f)).net, 0));
      const acTTC = acomptesTTC(devisId), hasTotale = deja.some(f => f.type === "totale");
      const avAcc = AK().ofDevis(devisId).filter(a => a.statut === "accepte");
      const avDispo = avAcc.filter(a => !this.avenantFacture(a.id));
      const avTTC = round2(avDispo.reduce((s, a) => s + AK().ttcOf(a), 0));
      const marche = round2(t.ttc + avAcc.reduce((s, a) => s + AK().ttcOf(a), 0));
      const reste = Math.max(0, round2(marche - dejaTTC));
      const ac = k.ACOMPTES.map((lab, i) => [lab.replace(/\s*:$/, ""), k.num(d.acomptes[i])]).filter(([, p]) => p > 0);
      let dlg = $("#ftNew");
      if (!dlg) { document.body.insertAdjacentHTML("beforeend", `<dialog id="ftNew"><div class="dlg" id="ftNewBody"></div></dialog>`); dlg = $("#ftNew"); }
      const st = { type: "totale", pct: ac[0]?.[1] || 30, idx: ac.length ? 0 : -1, av: false, mode: "pct", avct: 50 };
      const dvHT = round2(deja.filter(f => f.type === "acompte").reduce((s, f) => s + k.calc(normalize(f)).ht, 0));
      const draw = () => {
        const amtHT = round2(t.ht * st.pct / 100), amtTTC = round2(t.ttc * st.pct / 100);
        $("#ftNewBody").innerHTML = `
          <h3>Facturer le devis n° ${esc(d.numero)}</h3>
          <div class="small">${esc(d.client.nom || "Client")}${d.client.ref ? " · " + esc(d.client.ref) : ""} · total du devis <b>${k.eur.format(t.ttc)} TTC</b>${avAcc.length ? ` · avec avenants acceptés <b>${k.eur.format(marche)} TTC</b>` : ""}</div>
          ${deja.length ? `<div class="ft-deja"><b>Déjà facturé sur ce devis</b>${deja.map(f => `<div><span>${esc(f.numero)} · ${esc(typeTxt(f))}</span><b>${k.eur.format(totals(normalize(f)).net)}</b></div>`).join("")}
            <div class="reste"><span>Reste à facturer</span><b>${k.eur.format(reste)}</b></div></div>` : ""}
          <div class="ft-choice">
            <button data-ftype="totale" class="${st.type === "totale" ? "on" : ""}"><b>Facture identique au devis</b>
              <span>Toutes les lignes du devis${st.av && avDispo.length ? " et des avenants" : ""}, ${k.eur.format(round2(t.ttc + (st.av ? avTTC : 0)))} TTC.${acTTC ? ` Les acomptes déjà facturés (${k.eur.format(acTTC)}) sont déduits : net à payer ${k.eur.format(round2(t.ttc + (st.av ? avTTC : 0) - acTTC))}.` : ""}</span></button>
            <button data-ftype="acompte" class="${st.type === "acompte" ? "on" : ""}"><b>Facture d'acompte</b><span>Un pourcentage du devis.</span></button>
          </div>
          ${st.type === "acompte" ? acompteBlock(st, ac, t, dvHT) : ""}
          ${st.type === "totale" && hasTotale ? `<p class="pt-warn">Attention : ce devis a déjà une facture totale.</p>` : ""}
          <div class="err" id="ftErr"></div>
          <div class="foot"><span class="spacer"></span><button class="btn" data-fnew="cancel">Annuler</button><button class="btn primary" data-fnew="ok">Créer la facture</button></div>`;
      };
      draw();
      $("#ftNewBody").onclick = async ev => {
        const g = s => ev.target.closest(s);
        if (ev.target.id === "ftAvInc") { st.av = ev.target.checked; return draw(); }
        if (g("[data-ftype]")) { st.type = g("[data-ftype]").dataset.ftype; return draw(); }
        if (g("[data-fpct]")) { st.mode = "pct"; st.idx = +g("[data-fpct]").dataset.fpct; st.pct = ac[st.idx][1]; return draw(); }
        if (g("[data-favct]")) { st.mode = "avct"; st.idx = -1; return draw(); }
        const b = g("[data-fnew]"); if (!b) return;
        if (b.dataset.fnew === "cancel") return dlg.close();
        if (st.type === "acompte") {
          const v = st.mode === "avct" ? st.avct : st.pct;
          if (!(v > 0 && v <= 100)) { $("#ftErr").textContent = "Indiquez un pourcentage entre 1 et 100."; return; }
          if (acompteAmount(st, t, dvHT).ht <= 0) { $("#ftErr").textContent = "Rien à facturer pour cet avancement."; return; }
        }
        b.disabled = true;
        try {
          const f = fromDevis(dv, st.type, st.pct, st.mode === "pct" && st.idx >= 0 ? ac[st.idx][0] : "", st.type === "totale" && st.av ? avDispo : []);
          if (st.type === "acompte" && st.mode === "avct") {
            const m = acompteAmount(st, t, dvHT); f.avancement = st.avct; f.pct = m.pct;
            f.lignes = [{ d: acompteLine(st, "", `le devis n° ${d.numero} du ${k.frDate(d.date)}`, d.client.ref, m), pu: m.ht, q: 1 }];
          }
          const n = await store.create(COL, f);
          if (["brouillon", "envoye"].includes(dv.statut || "brouillon")) await store.patch("devis", dv.id, { statut: "accepte" }).catch(() => {});
          dlg.close();
          toast(`Facture ${f.numero} créée${st.type === "acompte" ? ` (acompte ${pctFr(st.pct)})` : ""}`);
          this.showFacture(n.id);
        } catch (e) { $("#ftErr").textContent = frError(e); b.disabled = false; }
      };
      $("#ftNewBody").oninput = ev => {
        if (ev.target.id !== "ftPctIn") return;
        if (st.mode === "avct") st.avct = k.num(ev.target.value); else { st.pct = k.num(ev.target.value); st.idx = -1; $("#ftNewBody").querySelectorAll("[data-fpct]").forEach(x => x.classList.remove("on")); }
        const el = $("#ftAmtBox"); if (el) el.innerHTML = acompteTxt(st, t, dvHT);
      };
      if (!dlg.open) dlg.showModal();
    },

    /* -------------------------- téléphone -------------------------- */
    renderM() {
      const k = K(), m = this.m;
      if (m.screen === "pick") {
        const list = store.list("devis").sort((a, b) => (b.date || "").localeCompare(a.date || ""));
        this.el.innerHTML = `<div class="dv dv-m"><p class="dv-q">Quel devis facturer ?</p>
          <p class="dv-sub">Touchez le devis : vous choisirez ensuite facture identique ou acompte.</p>
          ${list.length ? list.map(d => this.devisRow(d)).join("") : `<div class="dv-card dv-sub" style="margin:0">Aucun devis : créez-le d'abord dans « Création de devis ».</div>`}</div>
          <div class="dv-nav"><button class="btn" data-act="mlist">Liste</button><div class="tt"></div></div>`;
        return;
      }
      if (m.screen === "view" && store.get(COL, m.id)) {
        const f = normalize(store.get(COL, m.id)), t = totals(f);
        const ech = f.echeance ? Math.round((new Date(f.echeance) - new Date(f.date)) / 864e5) : null;
        this.el.innerHTML = `<div class="dv dv-m">
          <button class="btn dv-back" data-act="mlist">‹ Liste des factures</button>${familyTabs(f, "data-mopen")}<p class="dv-q">Facture ${esc(f.numero)}</p><p class="dv-sub">${esc(f.client.nom)} · ${esc(typeTxt(f))}${f.devisNumero ? " du devis " + esc(f.devisNumero) : ""}</p>
          <div class="dv-prev" id="mPrev">${page(f, false)}</div>
          <div class="dv-card dv-sum"><div><span>Total HT</span><b>${k.eur.format(t.ht)}</b></div><div><span>TVA ${k.pctTxt(f.tva)}</span><b>${k.eur.format(t.tv)}</b></div>
            ${t.ded ? `<div><span>Total TTC</span><b>${k.eur.format(t.ttc)}</b></div><div><span>Acomptes déjà facturés</span><b>− ${k.eur.format(t.ded)}</b></div>` : ""}
            <div class="ttc"><span>Net à payer</span><b>${k.eur.format(t.net)}</b></div></div>
          <div class="dv-card"><div class="field"><label>Échéance du règlement</label></div>
            <div class="dv-chips">${ECHEANCES.map(([n, j]) => `<button data-mech="${j}" class="${ech === j ? "on" : ""}">${n}</button>`).join("")}</div>
            <div class="field" style="margin-top:8px"><input id="fmEch" type="date" value="${esc(f.echeance)}" aria-label="Date d'échéance"></div></div>
          <div class="dv-card"><div class="field"><label>Statut de la facture</label></div>
            <div class="dv-seg ft-seg">${Object.entries(STATUTS).map(([s, v]) => `<button data-mstatut="${s}" class="st-${STCSS[s]} ${f.statut === s ? "on" : ""}">${v}</button>`).join("")}</div></div>
          <button class="btn primary big" data-act="mmail">Envoyer par mail</button>
          <button class="btn big" data-act="mpdf">Enregistrer le PDF</button>
          <button class="btn danger" data-act="mdelete">Supprimer cette facture</button></div>
          <div class="dv-nav"><button class="btn" data-act="mlist">Liste</button><div class="tt">Net à payer<b>${k.eur.format(t.net)}</b></div></div>`;
        k.fitPreview();
        return;
      }
      m.screen = "list";
      const items = this.sorted();
      this.el.innerHTML = `<div class="dv dv-m">
        <h1>Création de facture</h1>
        <button class="btn primary big" data-act="mpick">+ Facturer un devis</button>
        ${items.length ? `<div class="dv-sub" style="margin:4px 0 0">Factures récentes</div>` + items.filter(f => topOf(f).id === f.id).map(f => `<button class="dv-mitem" data-mopen="${f.id}"><b>${esc(f.client?.nom || "Client")}</b><em>${k.eur.format(totals(normalize(f)).net)}</em>
            <span>${esc(f.numero)} · ${esc(typeTxt(f))}${familyOf(f).length > 1 ? ` · + ${familyOf(f).length - 1} autre${familyOf(f).length > 2 ? "s" : ""}` : ""}</span><span class="statut ${STCSS[f.statut] || "brouillon"}">${STATUTS[f.statut] || "À envoyer"}</span></button>`).join("")
          : `<div class="dv-card dv-sub" style="margin:0">Aucune facture pour l'instant. Touchez « Facturer un devis », ou « Facturer ce devis » depuis un devis.</div>`}</div>`;
    },
    async saveM(patch) {
      if (!canEdit()) return false;
      try { await store.patch(COL, this.m.id, patch); this.renderM(); return true; } catch (e) { toast(frError(e)); return false; }
    },
    async onClickM(e, g) {
      const m = this.m, k = K();
      if (g("[data-mopen]")) { m.id = g("[data-mopen]").dataset.mopen; m.screen = "view"; window.scrollTo(0, 0); return this.renderM(); }
      if (g("[data-mech]")) { const f = store.get(COL, m.id); return this.saveM({ echeance: addDays(f.date, +g("[data-mech]").dataset.mech) }); }
      if (g("[data-mstatut]")) { const s = g("[data-mstatut]").dataset.mstatut; if (await this.saveM({ statut: s })) toast("Statut : " + STATUTS[s]); return; }
      const a = g("[data-act]")?.dataset.act; if (!a) return;
      if (a === "mpick") { if (!canEdit()) return; m.screen = "pick"; window.scrollTo(0, 0); return this.renderM(); }
      if (a === "mlist") { m.screen = "list"; return this.renderM(); }
      const f = m.id && store.get(COL, m.id) ? k.cleanLines(normalize(store.get(COL, m.id))) : null; if (!f) return;
      if (a === "mpdf") { try { const r = await k.savePDF(f, pdfOpts(f)); if (r === "saved") toast("PDF enregistré : " + fileName(f)); } catch (err) { if (err?.code !== "declined") toast(err.message || String(err)); } }
      if (a === "mmail") k.openMail(f, async () => { if (f.statut === "a_envoyer") await this.saveM({ statut: "envoyee" }); toast("Facture envoyée"); }, pdfOpts(f));
      if (a === "mdelete") {
        const b = g("[data-act]");
        if (!b.classList.contains("armed")) { b.classList.add("armed"); b.textContent = "Confirmer la suppression"; return; }
        await store.remove(COL, m.id).catch(err => toast(frError(err)));
        m.screen = "list"; m.id = null; toast("Facture supprimée"); this.renderM();
      }
    },
  };

  window.factureTab = tab;
})();
