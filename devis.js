/* =====================================================================
   RSG — onglet « Création de devis »
   • PC (grand écran) : la page A4 du devis Excel, à remplir directement dans les cases.
   • Téléphone : un assistant qui pose les questions une par une et remplit le devis tout seul.
   Données : collections « devis » et « clients » (même base que le reste de RSG).
   ===================================================================== */
"use strict";
(function () {
  const ENT = {
    adresse: "Siège social : 13 rue Pasteur, 54580 MOINEVILLE.",
    tel: "+33 7 50 00 25 43",
    mail: "rsg.courrier@gmail.com",
    ape: "E.I. - Code APE : 2420Z",
    siret: "N° SIRET : 792 868 135 00023",
    tva: "N° TVA Intracommunautaire : FR 02792868135",
    paiement: "Conditions de paiement : par virement _ Coordonnées bancaires : IBAN FR76 1610 6000 2996 0306 0579 686 / BIC AGRIFRPP861",
  };
  const LOGO = window.RSG_LOGO_DEVIS || "logo-devis.png";
  const ACOMPTES = ["Acompte à la signature :", "Acompte Début de travaux :", "Acompte 3/4 des travaux :", "Solde à réception des travaux :"];
  const PRESETS = [["20 / 30 / 30 / 20", [20, 30, 30, 20]], ["30 / 40 / 0 / 30", [30, 40, 0, 30]], ["50 / 0 / 0 / 50", [50, 0, 0, 50]], ["Tout à la fin", [0, 0, 0, 100]]];
  const TVAS = [0.2, 0.1, 0.055, 0];
  const VALIDITES = ["15 jours", "1 mois", "2 mois", "3 mois"];
  const STATUTS = { brouillon: "Brouillon", envoye: "Envoyé", accepte: "Accepté", refuse: "Refusé" };
  const NB_LIGNES = 9;
  const AUTOLIQ = "TVA autoliquidation à la charge du preneur";

  const eur = new Intl.NumberFormat("fr-FR", { style: "currency", currency: "EUR" });
  const pctTxt = v => (Math.round(v * 1000) / 10).toLocaleString("fr-FR", { minimumFractionDigits: v * 100 % 1 ? 1 : 2 }) + " %";
  const money = v => v === "" || v == null ? "" : num(v).toLocaleString("fr-FR", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  const num = v => { const n = parseFloat(String(v ?? "").replace(/\s/g, "").replace(",", ".")); return isFinite(n) ? n : 0; };
  const isoToday = () => { const d = new Date(); return new Date(d.getTime() - d.getTimezoneOffset() * 6e4).toISOString().slice(0, 10); };
  const frDate = iso => iso ? new Date(iso + "T12:00:00").toLocaleDateString("fr-FR") : "";
  const mobile = () => matchMedia("(max-width: 760px)").matches;
  const clone = o => JSON.parse(JSON.stringify(o));
  const initials = n => (String(n || "?").replace(/^(m\.|mme|m|monsieur|madame)\s+/i, "").trim().split(/\s+/).map(w => w[0]).join("").slice(0, 2) || "?").toUpperCase();

  function calc(d) {
    const ht = (d.lignes || []).reduce((s, l) => s + num(l.pu) * num(l.q), 0);
    const tv = ht * (d.tva ?? 0.2);
    return { ht, tv, ttc: ht + tv };
  }
  function nextNumero(iso, selfId) {
    if (!/^\d{4}-\d{2}-\d{2}$/.test(iso || "")) return "";
    const [y, m, dd] = iso.split("-"), pre = `${y}/${m}${dd}-D`;
    const self = selfId && store.get("devis", selfId);
    if (self && String(self.numero || "").startsWith(pre)) return self.numero;     // même jour : le devis garde son numéro
    const max = store.list("devis").filter(x => x.id !== selfId && String(x.numero || "").startsWith(pre))
      .reduce((mx, x) => Math.max(mx, parseInt(String(x.numero).slice(pre.length), 10) || 0), 0);
    return pre + (max + 1);
  }
  function blank() {
    const date = isoToday();
    return { numero: nextNumero(date), date, validite: "1 mois", statut: "brouillon", clientId: null,
      client: { nom: "", adr1: "", adr2: "", ref: "", liv1: "", liv2: "", email: "" },
      lignes: [], tva: 0.2, acomptes: [20, 30, 30, 20] };
  }
  function normalize(d) {
    const x = Object.assign(blank(), clone(d));
    x.client = Object.assign(blank().client, x.client || {});
    x.lignes = (x.lignes || []).map(l => ({ d: l.d || "", pu: l.pu ?? "", q: l.q ?? "" }));
    x.acomptes = (x.acomptes || [20, 30, 30, 20]).slice(0, 4);
    return x;
  }
  const cleanLines = d => ({ ...d, lignes: d.lignes.filter(l => String(l.d).trim() || num(l.pu) || num(l.q)) });
  function setPath(o, path, v) {
    const k = path.split("."); let t = o;
    for (let i = 0; i < k.length - 1; i++) {
      const key = /^\d+$/.test(k[i]) ? +k[i] : k[i];
      if (t[key] == null) t[key] = /^\d+$/.test(k[i + 1]) ? [] : {};
      t = t[key];
    }
    const last = /^\d+$/.test(k.at(-1)) ? +k.at(-1) : k.at(-1);
    t[last] = v;
  }
  const cleanName = t => String(t || "").replace(/[\\/:*?"<>|]+/g, "-").replace(/\s+/g, " ").trim();
  const fileName = d => ["Devis RSG", d.client?.nom, d.client?.ref, String(d.numero || "").replace(/\//g, "-")]
    .map(cleanName).filter(Boolean).join(" - ") + ".pdf";

  /* ---------- la page A4 (statique ou éditable) ---------- */
  function page(d, edit) {
    const c = d.client, t = calc(d);
    const inp = (path, val, o = {}) => edit
      ? `<input class="f ${o.cls || ""}" data-f="${path}" value="${esc(val ?? "")}" ${o.type ? `type="${o.type}"` : ""} inputmode="${o.im || "text"}" placeholder="${esc(o.ph || "")}" ${o.list ? `list="${o.list}"` : ""} autocomplete="off">`
      : esc(o.show ? o.show(val) : (val ?? ""));
    const rows = Math.max(NB_LIGNES, d.lignes.length + (edit ? 1 : 0));
    let body = "";
    for (let i = 0; i < rows; i++) {
      const l = d.lignes[i] || { d: "", pu: "", q: "" };
      const has = String(l.d).trim() || num(l.pu) || num(l.q);
      const tot = has && l.pu !== "" && l.q !== "" ? eur.format(num(l.pu) * num(l.q)) : "";
      body += `<tr data-row="${i}">
        <td>${edit ? `<textarea class="f" data-f="lignes.${i}.d" rows="1" placeholder="${i === 0 && !d.lignes.length ? "Désignation du produit ou de la prestation…" : ""}">${esc(l.d)}</textarea>` : esc(l.d)}</td>
        <td class="n">${edit ? inp(`lignes.${i}.pu`, money(l.pu), { cls: "num-in money", im: "decimal" }) : (has && l.pu !== "" ? eur.format(num(l.pu)) : "")}</td>
        <td class="q">${edit ? inp(`lignes.${i}.q`, l.q, { cls: "c", im: "decimal" }) : (has ? esc(String(l.q).replace(".", ",")) : "")}</td>
        <td class="t" data-calc="l${i}">${tot}${edit && i < d.lignes.length ? `<button class="rowdel noprint" data-del="${i}" title="Supprimer la ligne" aria-label="Supprimer la ligne">×</button>` : ""}</td></tr>`;
    }
    const tvaSel = edit
      ? `<select class="f" data-f="tva" aria-label="Taux de TVA">${TVAS.map(v => `<option value="${v}" ${v === d.tva ? "selected" : ""}>${pctTxt(v)}</option>`).join("")}</select>`
      : pctTxt(d.tva);
    const acc = ACOMPTES.map((lab, i) => `<span>${lab}</span><b data-calc="a${i}">${eur.format(t.ttc * num(d.acomptes[i]) / 100)}</b>
        <i>${edit ? inp(`acomptes.${i}`, d.acomptes[i], { cls: "c", im: "numeric" }) : num(d.acomptes[i]) + " %"}</i>`).join("");
    const sumPct = d.acomptes.reduce((s, v) => s + num(v), 0);
    return `<div class="a4 ${edit ? "" : "static"}">
      <div class="hd">
        <div class="co"><img class="logo" src="${LOGO}" alt="RSG">
          <div>${ENT.adresse}</div><div>Téléphone : ${ENT.tel}</div><div>Mail : ${ENT.mail}</div>
          <small>${ENT.ape}</small><small>${ENT.siret}</small><small>${ENT.tva}</small></div>
        <div>
          <div class="title"><b>DEVIS N°</b><span class="num" data-calc="num" title="Calculé à partir de la date">${esc(d.numero || "—")}</span></div>
          <div class="meta"><span>En date du</span><em>${inp("date", d.date, { type: "date", show: frDate, cls: "c" })}</em>
            <span>Durée de validité :</span><em>${inp("validite", d.validite, { cls: "c", list: "dvValid" })}</em></div>
          <div class="clientlbl">CLIENT</div>
          <div class="client">
            <div class="cn">${inp("client.nom", c.nom, { ph: "Nom du client", list: "dvClients" })}</div>
            <div>${inp("client.adr1", c.adr1, { ph: "Adresse" })}</div>
            <div>${inp("client.adr2", c.adr2, { ph: "Code postal et ville" })}</div>
            <div class="ref"><span>Référence :</span><b>${inp("client.ref", c.ref, { ph: "Chantier" })}</b></div>
            ${edit ? `<div class="mailrow noprint"><span>E-mail du client (non imprimé) :</span>${inp("client.email", c.email, { type: "email", ph: "client@exemple.fr" })}</div>` : ""}
          </div>
        </div>
      </div>
      <table><colgroup><col class="c1"><col class="c2"><col class="c3"><col class="c4"></colgroup>
        <thead><tr><th>Désignation des produits et prestations</th><th>Prix unitaire HT</th><th>Quantité</th><th>Total HT</th></tr></thead>
        <tbody>${body}</tbody></table>
      ${edit ? `<div class="addrow noprint"><button data-act="addline">+ Ajouter une ligne</button></div>` : ""}
      <div class="tot">
        <div><span>Total Hors Taxe (HT)</span><b data-calc="ht">${eur.format(t.ht)}</b></div>
        <div><span>Total TVA (${tvaSel})</span><b data-calc="tv">${eur.format(t.tv)}</b></div>
        <div class="ttc"><span>TOTAL TTC</span><b data-calc="ttc">${eur.format(t.ttc)}</b></div>
      </div>
      ${d.tva === 0 ? `<div class="autoliq">${AUTOLIQ}</div>` : ""}
      <div class="cond">
        <span class="lbl">Termes de réalisation :</span><b style="border:0"></b><i></i>
        ${acc}
        <span></span><b class="sum" data-calc="as">${eur.format(t.ttc * sumPct / 100)}</b><i data-calc="ap" style="${Math.round(sumPct) !== 100 ? "color:#c00000;font-weight:700" : ""}">${sumPct} %</i>
      </div>
      <div class="acc"><div class="lbl">ACCORD DU CLIENT</div><div class="box"><span>Signature avec la mention "Bon pour accord"</span><span>Date</span></div></div>
      <div class="foot">${ENT.paiement}<br>${ENT.tva}</div>
    </div>`;
  }

  /* ---------- PDF et mail (communs PC / téléphone) ---------- */
  // o (facultatif) : { html, title, name, texts, label } pour réutiliser PDF et mail ailleurs (factures)
  async function makePDF(d, o = {}) {
    if (!window.html2canvas || !window.jspdf) throw new Error("Le module PDF n'est pas chargé : ouvrez RSG une fois avec Internet.");
    const holder = document.createElement("div"); holder.className = "pdf-stage"; holder.innerHTML = o.html || page(d, false);
    document.body.appendChild(holder);
    const el = holder.firstElementChild; el.style.boxShadow = "none";
    await Promise.all([...el.querySelectorAll("img")].map(i => i.complete ? 0 : new Promise(r => { i.onload = i.onerror = r; })));
    try {
      const canvas = await window.html2canvas(el, { scale: 2, backgroundColor: "#ffffff", useCORS: true });
      const L = !!o.landscape, PW = L ? 297 : 210, PH = L ? 210 : 297;
      const pdf = new window.jspdf.jsPDF({ unit: "mm", format: "a4", orientation: L ? "landscape" : "portrait" });
      let w = PW - 20, h = w * canvas.height / canvas.width;
      if (h > PH - 20) { h = PH - 20; w = h * canvas.width / canvas.height; }
      pdf.setProperties({ title: o.title || `Devis ${d.numero}`, author: "RSG – Réalisations Soudures Générales" });
      pdf.addImage(canvas.toDataURL("image/jpeg", 0.92), "JPEG", (PW - w) / 2, (PH - h) / 2, w, h);
      return pdf.output("blob");
    } finally { holder.remove(); }
  }
  async function savePDF(d, o = {}) {
    const name = o.name || fileName(d), blob = await makePDF(d, o);
    if (window.RSG_SAVE_FILE) return window.RSG_SAVE_FILE(name, blob);            // simulation
    const a = document.createElement("a"); a.href = URL.createObjectURL(blob); a.download = name;
    document.body.appendChild(a); a.click(); a.remove(); setTimeout(() => URL.revokeObjectURL(a.href), 8000);
    return "saved";
  }

  function mailTexts(d) {
    const t = calc(d), c = d.client;
    return {
      to: c.email || "",
      subject: `Devis n° ${d.numero} – RSG${c.ref ? " – " + c.ref : ""}`,
      body: `Bonjour${c.nom ? " " + c.nom : ""},\n\nVeuillez trouver ci-joint notre devis n° ${d.numero}${c.ref ? ` pour « ${c.ref} »` : ""}, d'un montant de ${eur.format(t.ttc)} TTC.\n\nCe devis est valable ${d.validite}. Pour l'accepter, il suffit de nous le retourner signé avec la mention « Bon pour accord ».\n\nNous restons à votre disposition pour toute question.\n\nCordialement,\nRSG – Réalisations Soudures Générales\n${ENT.tel} – ${ENT.mail}`,
    };
  }
  function ensureMailDialog() {
    if ($("#dvMail")) return;
    document.body.insertAdjacentHTML("beforeend", `<dialog id="dvMail"><form class="dlg" method="dialog">
      <h3>Envoyer le devis par mail</h3>
      <div class="from">De : <b>${ENT.mail}</b></div>
      <div class="field"><label for="dvTo">À</label><input id="dvTo" type="email" placeholder="client@exemple.fr"></div>
      <div class="cc"><span>Copie cachée (obligatoire)</span><b>${ENT.mail}</b></div>
      <div class="field"><label for="dvSub">Objet</label><input id="dvSub"></div>
      <div class="field"><label for="dvBody">Message</label><textarea id="dvBody"></textarea></div>
      <div class="att"><i>PDF</i><span id="dvAtt"></span></div>
      <ol class="steps" id="dvSteps"></ol>
      <div class="err" id="dvErr"></div>
      <div class="foot"><span class="spacer"></span><button type="button" class="btn" data-close>Annuler</button><button type="button" class="btn primary" id="dvSend">Envoyer</button></div>
    </form></dialog>`);
    $("#dvMail [data-close]").onclick = () => $("#dvMail").close();
  }
  function openMail(d, after, o = {}) {
    ensureMailDialog();
    const m = o.texts || mailTexts(d), phone = mobile(), label = o.label || "devis", fname = o.name || fileName(d);
    const fixed = o.fixedTo || "";                 // destinataire imposé (pointages : uniquement l'adresse RSG)
    const du = o.what || (label === "devis" ? "du devis" : "de la facture");
    $("#dvMail h3").textContent = o.dialogTitle || `Envoyer ${label === "devis" ? "le devis" : "la facture"} par mail`;
    $("#dvTo").readOnly = !!fixed; $("#dvMail .cc").hidden = !!fixed;
    $("#dvTo").value = fixed || m.to; $("#dvSub").value = m.subject; $("#dvBody").value = m.body; $("#dvAtt").textContent = fname;
    $("#dvErr").textContent = "";
    // copie obligatoire vers l'adresse RSG, en copie cachée (le client ne la voit pas)
    $("#dvSteps").innerHTML = (phone
      ? "<li>Le PDF " + du + " est enregistré dans le dossier <b>Téléchargements</b> du téléphone.</li><li>Gmail s'ouvre avec le destinataire" + (fixed ? "" : ", la copie à " + ENT.mail) + ", l'objet et le message déjà remplis.</li><li>Touchez 📎 → <b>Joindre un fichier</b> → Téléchargements, choisissez le PDF, puis Envoyer.</li>"
      : "<li>Le PDF " + du + " est enregistré sur ce PC (dossier Téléchargements).</li><li>Gmail s'ouvre sur le compte <b>" + ENT.mail + "</b>, avec le destinataire" + (fixed ? "" : ", la copie") + ", l'objet et le message déjà remplis.</li><li>Glissez le PDF dans le message, puis cliquez Envoyer.</li>");
    $("#dvSend").textContent = "Préparer dans Gmail";
    $("#dvSend").onclick = async () => {
      const to = fixed || $("#dvTo").value.trim(), su = $("#dvSub").value, body = $("#dvBody").value;
      if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(to)) { $("#dvErr").textContent = "Indiquez l'adresse e-mail du client."; $("#dvTo").focus(); return; }
      const btn = $("#dvSend"); btn.disabled = true; btn.textContent = "Préparation du PDF…";
      try {
        if (window.RSG_SIM_MAIL) { await window.RSG_SIM_MAIL({ from: ENT.mail, to, bcc: fixed ? "" : ENT.mail, subject: su, body, file: fname, blob: await makePDF(d, o) }); }
        else if (phone) {
          await savePDF(d, o);
          location.href = "mailto:" + encodeURIComponent(to) + "?" + (fixed ? "" : "bcc=" + encodeURIComponent(ENT.mail) + "&") +
            "subject=" + encodeURIComponent(su) + "&body=" + encodeURIComponent(body);
        } else {
          await savePDF(d, o);
          const url = "https://mail.google.com/mail/?view=cm&fs=1&authuser=" + encodeURIComponent(ENT.mail) +
            "&to=" + encodeURIComponent(to) + (fixed ? "" : "&bcc=" + encodeURIComponent(ENT.mail)) +
            "&su=" + encodeURIComponent(su) + "&body=" + encodeURIComponent(body);
          window.open(url, "_blank", "noopener");
        }
        $("#dvMail").close();
        after && after(to);
      } catch (e) { if (e?.name !== "AbortError") $("#dvErr").textContent = e.message || String(e); }
      finally { btn.disabled = false; btn.textContent = "Préparer dans Gmail"; }
    };
    $("#dvMail").showModal();
  }

  /* ---------- enregistrement (clients + devis) ---------- */
  async function saveClientOf(d) {
    const c = d.client; if (!c.nom.trim()) return d;
    const clients = store.list("clients");
    let found = d.clientId && clients.find(x => x.id === d.clientId);
    if (!found) found = clients.find(x => x.nom.trim().toLowerCase() === c.nom.trim().toLowerCase());
    const data = { nom: c.nom.trim(), adr1: c.adr1, adr2: c.adr2, email: c.email };
    if (!found) { const n = await store.create("clients", data); d.clientId = n.id; }
    else {
      d.clientId = found.id;
      if (found.adr1 !== c.adr1 || found.adr2 !== c.adr2 || (c.email && found.email !== c.email)) await store.patch("clients", found.id, data);
    }
    return d;
  }
  async function persist(id, draft) {
    let d = cleanLines(clone(draft));
    d = await saveClientOf(d);
    draft.clientId = d.clientId;
    if (!id) return (await store.create("devis", d)).id;
    await store.patch("devis", id, d);
    return id;
  }

  /* ---------- clients enregistrés : liste + suppression ---------- */
  function openClients(after) {
    let dlg = $("#dvClientsDlg");
    if (!dlg) {
      document.body.insertAdjacentHTML("beforeend", `<dialog id="dvClientsDlg"><div class="dlg"><h3>Clients enregistrés</h3>
        <p class="small">Supprimer un client l'enlève seulement de la liste des propositions. Les devis et factures déjà faits pour lui ne changent pas.</p>
        <div class="cl-list" id="clList"></div>
        <div class="foot"><span class="spacer"></span><button class="btn primary" data-close>Fermer</button></div></div></dialog>`);
      dlg = $("#dvClientsDlg");
      dlg.querySelector("[data-close]").onclick = () => dlg.close();
      dlg.addEventListener("close", () => after && after());
      $("#clList").onclick = async e => {
        const b = e.target.closest("[data-cldel]"); if (!b) return;
        if (!canEdit()) return;
        if (!b.classList.contains("armed")) { b.classList.add("armed"); b.textContent = "Confirmer"; setTimeout(() => { b.classList.remove("armed"); b.textContent = "Supprimer"; }, 4000); return; }
        const c = store.get("clients", b.dataset.cldel);
        try { await store.remove("clients", b.dataset.cldel); toast(`Client « ${c?.nom || ""} » supprimé`); draw(); } catch (err) { toast(frError(err)); }
      };
    }
    const draw = () => {
      const list = store.list("clients").sort((a, b) => a.nom.localeCompare(b.nom, "fr"));
      const nb = id => store.list("devis").filter(d => d.clientId === id).length;
      $("#clList").innerHTML = list.length ? list.map(c => `<div class="cl-row"><span class="av">${esc(initials(c.nom))}</span>
          <div><b>${esc(c.nom)}</b><span>${esc([c.adr1, c.adr2].filter(Boolean).join(", ") || "Adresse non renseignée")}${nb(c.id) ? ` · ${nb(c.id)} devis` : ""}</span></div>
          <button class="btn danger" data-cldel="${c.id}">Supprimer</button></div>`).join("")
        : `<p class="small">Aucun client enregistré.</p>`;
    };
    draw(); dlg.showModal();
  }

  /* =================================================================
     L'onglet
     ================================================================= */
  const inPeriod = (iso, p) => !p || (String(iso || "") >= p.from && String(iso || "") <= p.to);
  const tab = {
    id: "devis",
    title: "Création de devis",
    collections: ["devis", "clients"],
    el: null, mode: null,
    // PC
    selId: null, draft: null, dirty: false, saving: false, timer: null, q: "",
    // téléphone
    m: { screen: "list", step: 1, id: null, draft: null, newClient: false, cq: "" },

    // Sauvegarde PDF (Réglages) : un PDF par devis de la période
    exportPdf(p) {
      return this.sortedList().filter(d => inPeriod(d.date, p)).map(d => { const n = cleanLines(normalize(d));
        return { path: "Devis/" + fileName(n), make: () => makePDF(n) }; });
    },
    mount(el) {
      this.el = el; this.mode = mobile() ? "m" : "pc";
      if (!this.bound) {
        this.bound = true;
        el.addEventListener("click", e => current === this && this.onClick(e));
        el.addEventListener("input", e => current === this && this.onInput(e));
        el.addEventListener("change", e => current === this && this.onChange(e));
        el.addEventListener("keydown", e => current === this && this.onKey(e));
        el.addEventListener("focusin", e => { if (e.target.classList?.contains("money")) { const v = num(e.target.value); e.target.value = e.target.value === "" ? "" : String(v).replace(".", ","); e.target.select(); } });
        el.addEventListener("focusout", e => { if (e.target.classList?.contains("money")) e.target.value = money(e.target.value.replace(",", ".")); });
        matchMedia("(max-width: 760px)").addEventListener("change", () => { if (current === this) this.mount(this.el); });
      }
      if (this.mode === "pc") this.mountPC(); else this.renderM();
    },
    render() {
      if (!this.el || current !== this) return;
      if (this.mode === "pc") {
        this.renderList();
        if (this.av) { window.avenantsView.refreshPC(); return; }
        if (!this.selId && !this.draft && !this.dirty) {          // premier chargement : on ouvre le devis le plus récent
          const first = this.sortedList()[0];
          if (first) { this.selId = first.id; this.draft = normalize(first); this.renderList(); this.renderPage(); return; }
        }
        const inside = this.el.querySelector(".a4")?.contains(document.activeElement);
        if (this.selId && !this.dirty && !this.saving && !inside) {
          const fresh = store.get("devis", this.selId);
          if (fresh && (!this.draft || JSON.stringify(cleanLines(normalize(fresh))) !== JSON.stringify(cleanLines(this.draft)))) { this.draft = normalize(fresh); this.renderPage(); }
        }
        if (this.selId && !store.get("devis", this.selId) && !this.saving) { this.selId = null; this.draft = null; this.renderPage(); }
      } else if (this.m.screen === "list") this.renderM();
    },

    /* ----------------------------- PC ----------------------------- */
    mountPC() {
      this.el.innerHTML = `<div class="dv"><div class="dv-pc">
        <aside class="dv-side"><div class="top">
          <button class="btn primary" data-act="new">+ Nouveau devis</button>
          <button class="btn" data-act="clients">Clients enregistrés</button>
          <input id="dvSearch" placeholder="Rechercher (client, n°, chantier)" value="${esc(this.q)}" aria-label="Rechercher un devis">
        </div><div class="dv-list" id="dvList"></div></aside>
        <section class="dv-main" id="dvMain"></section>
      </div>
      <datalist id="dvClients"></datalist><datalist id="dvValid">${VALIDITES.map(v => `<option value="${v}">`).join("")}</datalist></div>`;
      if (!this.selId) this.selId = this.sortedList()[0]?.id || null;
      if (this.selId && store.get("devis", this.selId)) this.draft = normalize(store.get("devis", this.selId));
      this.renderList(); this.renderPage();
    },
    sortedList() {
      return store.list("devis").sort((a, b) => (b.date || "").localeCompare(a.date || "") || String(b.numero).localeCompare(String(a.numero), "fr", { numeric: true }));
    },
    renderList() {
      const list = $("#dvList"); if (!list) return;
      const q = norm(this.q);
      const items = this.sortedList().filter(d => !q || norm([d.numero, d.client?.nom, d.client?.ref].join(" ")).includes(q));
      list.innerHTML = items.length ? items.map(d => {
        const t = calc(normalize(d));
        return `<button class="dv-item ${d.id === this.selId ? "on" : ""}" data-open="${d.id}">
          <b>${esc(d.numero)}</b><em>${eur.format(t.ttc)}</em>
          <span>${esc(d.client?.nom || "Client à renseigner")}${d.client?.ref ? " · " + esc(d.client.ref) : ""}</span><span class="statut ${d.statut || "brouillon"}">${STATUTS[d.statut] || "Brouillon"}</span>
          ${window.avenantsView?.count(d.id) ? `<span class="dv-avn">+ ${window.avenantsView.count(d.id)} avenant${window.avenantsView.count(d.id) > 1 ? "s" : ""}</span>` : ""}</button>`;
      }).join("") : `<div class="dv-empty">${store.list("devis").length ? "Aucun devis ne correspond." : "Aucun devis pour l'instant. Cliquez « + Nouveau devis » pour commencer."}</div>`;
      const dl = $("#dvClients"); if (dl) dl.innerHTML = store.list("clients").map(c => `<option value="${esc(c.nom)}">${esc([c.adr2].filter(Boolean).join(" "))}</option>`).join("");
    },
    renderPage() {
      const y = window.scrollY; this._renderPage(); if (window.scrollY !== y) window.scrollTo(0, y);   // la page ne saute pas
    },
    _renderPage() {
      const main = $("#dvMain"); if (!main) return;
      if (!this.draft) {
        main.innerHTML = `<div class="dv-card"><b>Aucun devis ouvert.</b><p class="dv-hint">Créez un devis avec « + Nouveau devis », ou choisissez-en un dans la liste.</p></div>`;
        return;
      }
      const d = this.draft, AV = window.avenantsView;
      const strip = this.selId && AV ? AV.tabsHtml(this.selId, d.numero, this.av) : "";
      if (this.av && AV && AV.id === this.av) { main.innerHTML = strip + `<div id="avHost"></div>`; AV.renderPC($("#avHost")); return; }
      this.av = null;
      main.innerHTML = strip + `<div class="dv-bar">
          <button class="btn primary" data-act="pdf">Enregistrer en PDF sur le PC</button>
          <button class="btn" data-act="mail">Envoyer par mail</button>
          <button class="btn" data-act="facturer" title="Transformer ce devis en facture">Facturer</button>
          <button class="btn" data-act="dup">Dupliquer</button>
          <select class="btn" data-act="statut" aria-label="Statut du devis">${Object.entries(STATUTS).map(([k, v]) => `<option value="${k}" ${d.statut === k ? "selected" : ""}>${v}</option>`).join("")}</select>
          <button class="btn danger" data-act="del">Supprimer</button>
          <span class="st" id="dvSt"><i></i><span>save</span></span>
        </div>
        <div class="dv-hint">Remplissez directement les cases du devis, comme dans Excel (Tab pour passer à la case suivante). Tout est enregistré automatiquement.</div>
        <div class="dv-stage">${page(d, true)}</div>`;
      this.el.querySelectorAll(".a4 textarea.f").forEach(grow);
    },
    status(txt, pending) {
      const s = $("#dvSt"); if (!s) return;
      s.classList.toggle("pending", !!pending); s.lastElementChild.textContent = pending ? "modif" : (/Non/.test(txt) ? "erreur" : "save");
    },
    queueSave() {
      if (!canEdit()) return;
      this.dirty = true; this.status("Enregistrement…", true);
      clearTimeout(this.timer); this.timer = setTimeout(() => this.flush(), 700);
    },
    async flush() {
      if (!this.dirty || this.saving) return;
      this.dirty = false; this.saving = true; this.status("Enregistrement…", true);
      try {
        const wasNew = !this.selId;
        this.selId = await persist(this.selId, this.draft);
        this.status("Enregistré", false);
        if (wasNew) this.renderList();
      } catch (e) { this.dirty = true; this.status("Non enregistré", true); toast(frError(e)); }
      finally { this.saving = false; if (this.dirty) this.queueSave(); else this.renderList(); }
    },
    updateCalc() {
      const d = this.draft, t = calc(d), set = (k, v) => { const e = this.el.querySelector(`[data-calc="${k}"]`); if (e) e.textContent = v; };
      d.lignes.forEach((l, i) => {
        const has = String(l.d).trim() || num(l.pu) || num(l.q);
        const cell = this.el.querySelector(`[data-calc="l${i}"]`); if (!cell) return;
        const txt = has && l.pu !== "" && l.q !== "" ? eur.format(num(l.pu) * num(l.q)) : "";
        const btn = cell.querySelector(".rowdel"); cell.textContent = txt; if (btn) cell.appendChild(btn);
      });
      set("ht", eur.format(t.ht)); set("tv", eur.format(t.tv)); set("ttc", eur.format(t.ttc));
      d.acomptes.forEach((p, i) => set("a" + i, eur.format(t.ttc * num(p) / 100)));
      const sp = d.acomptes.reduce((s, v) => s + num(v), 0);
      set("as", eur.format(t.ttc * sp / 100));
      const ap = this.el.querySelector('[data-calc="ap"]'); if (ap) { ap.textContent = sp + " %"; ap.style.cssText = Math.round(sp) !== 100 ? "color:#c00000;font-weight:700" : ""; }
    },
    fillFromClient(name) {
      const c = store.list("clients").find(x => x.nom.trim().toLowerCase() === name.trim().toLowerCase());
      if (!c) { this.draft.clientId = null; return; }
      Object.assign(this.draft.client, { nom: c.nom, adr1: c.adr1 || "", adr2: c.adr2 || "", email: c.email || "" });
      this.draft.clientId = c.id;
      for (const k of ["adr1", "adr2", "email"]) { const i = this.el.querySelector(`[data-f="client.${k}"]`); if (i) i.value = this.draft.client[k]; }
      toast(`Client reconnu : adresse remplie automatiquement`);
    },

    onInput(e) {
      if (e.target.id === "dvSearch") { this.q = e.target.value; this.renderList(); return; }
      if (this.mode === "m") return this.onInputM(e);
      if (this.av && e.target.closest("#avPane")) return window.avenantsView.onInputPC(e);
      const f = e.target.dataset.f; if (!f || !this.draft) return;
      let v = e.target.value;
      if (f === "tva") v = +v;
      if (/^lignes\.\d+\.(pu|q)$|^acomptes\.\d$/.test(f)) v = v === "" ? "" : v.replace(/[\s\u202f\u00a0€]/g, "").replace(",", ".");
      if (/^lignes\.(\d+)\./.test(f)) {
        const i = +f.split(".")[1];
        while (this.draft.lignes.length <= i) this.draft.lignes.push({ d: "", pu: "", q: "" });
        const needRow = i === this.draft.lignes.length - 1 && !this.el.querySelector(`tr[data-row="${i + 1}"]`);
        setPath(this.draft, f, v);
        if (needRow) { const pos = caretOf(e.target); this.renderPage(); refocus(this.el, f, pos); }
      } else setPath(this.draft, f, v);
      if (f === "date" && v) {
        this.draft.numero = nextNumero(v, this.selId);
        const n = this.el.querySelector('[data-calc="num"]'); if (n) n.textContent = this.draft.numero;
      }
      if (e.target.tagName === "TEXTAREA") grow(e.target);
      this.updateCalc(); this.queueSave();
    },
    onChange(e) {
      if (this.mode === "m") return this.onChangeM(e);
      if (this.av && e.target.closest("#avPane")) return window.avenantsView.onChangePC(e);
      const t = e.target;
      if (t.dataset.f === "client.nom") { this.fillFromClient(t.value); this.queueSave(); }
      if (t.dataset.f === "tva") { this.draft.tva = +t.value; this.renderPage(); this.queueSave(); if (this.draft.tva === 0) toast("Mention ajoutée : " + AUTOLIQ); }
      if (t.dataset.act === "statut") { this.draft.statut = t.value; this.queueSave(); }
    },
    onKey(e) {
      if (this.mode !== "pc") return;
      const f = e.target.dataset?.f || "";
      const m = f.match(/^lignes\.(\d+)\.q$/);
      if (e.key === "Enter" && m) {              // Entrée sur la quantité : ligne suivante, comme dans un tableur
        e.preventDefault();
        const next = this.el.querySelector(`[data-f="lignes.${+m[1] + 1}.d"]`); next && next.focus();
      }
    },
    async onClick(e) {
      if (this.mode === "m") return this.onClickM(e);
      const AV = window.avenantsView;
      const ds = e.target.closest("[data-dsub]");
      if (ds) { await this.flush(); await AV.flush(); this.av = ds.dataset.dsub || null; if (this.av && !AV.load(this.av)) this.av = null; this.renderPage(); return; }
      if (e.target.closest('[data-act="newav"]')) {
        if (!canEdit() || !this.selId) return;
        await this.flush(); await AV.flush();
        try { const n = await AV.create(this.selId); this.av = n.id; this.renderPage(); this.renderList(); refocus(this.el, "lignes.0.d"); } catch (err) { toast(frError(err)); }
        return;
      }
      if (this.av && e.target.closest("#avPane")) { const r = await AV.onClickPC(e); if (r === "deleted") { this.av = null; this.renderPage(); this.renderList(); } return; }
      const o = e.target.closest("[data-open]");
      if (o) { await this.flush(); await AV.flush(); this.av = null; this.selId = o.dataset.open; this.draft = normalize(store.get("devis", this.selId)); this.renderList(); this.renderPage(); return; }
      const del = e.target.closest("[data-del]");
      if (del) { this.draft.lignes.splice(+del.dataset.del, 1); this.renderPage(); this.updateCalc(); this.queueSave(); return; }
      const a = e.target.closest("[data-act]")?.dataset.act; if (!a || e.target.tagName === "SELECT") return;
      if (a === "new") {
        if (!canEdit()) return;
        await this.flush(); this.av = null; this.selId = null; this.draft = normalize(blank()); this.renderList(); this.renderPage();
        this.el.querySelector('[data-f="client.nom"]')?.focus(); this.dirty = true; this.flush();
      }
      if (a === "clients") return openClients(() => this.renderList());
      if (a === "addline") { this.draft.lignes = cleanLines(this.draft).lignes; this.draft.lignes.push({ d: "", pu: "", q: 1 }); this.renderPage(); refocus(this.el, `lignes.${this.draft.lignes.length - 1}.d`); }
      if (a === "pdf") { await this.flush(); try { const r = await savePDF(cleanLines(this.draft)); if (r === "saved") toast("PDF enregistré dans Téléchargements : " + fileName(this.draft)); } catch (err) { if (err?.code !== "declined") toast(err.message || String(err)); } }
      if (a === "mail") { await this.flush(); openMail(cleanLines(this.draft), () => { if (this.draft.statut === "brouillon") { this.draft.statut = "envoye"; this.queueSave(); this.renderPage(); } }); }
      if (a === "facturer") { if (!canEdit()) return; await this.flush(); if (this.selId) window.factureTab?.openCreate(this.selId); return; }
      if (a === "dup") {
        if (!canEdit()) return;
        await this.flush(); const copy = normalize(this.draft); copy.date = isoToday(); copy.numero = nextNumero(copy.date); copy.statut = "brouillon";
        this.selId = null; this.draft = copy; this.dirty = true; await this.flush(); this.renderPage(); toast("Devis dupliqué : " + copy.numero);
      }
      if (a === "del") {
        const b = e.target.closest("[data-act]");
        if (!b.classList.contains("armed")) { b.classList.add("armed"); b.textContent = "Confirmer la suppression"; setTimeout(() => { b.classList.remove("armed"); b.textContent = "Supprimer"; }, 4000); return; }
        if (!canEdit()) return;
        const id = this.selId; clearTimeout(this.timer); this.dirty = false;
        if (id) await store.remove("devis", id).catch(err => toast(frError(err)));
        this.selId = this.sortedList()[0]?.id || null; this.draft = this.selId ? normalize(store.get("devis", this.selId)) : null;
        this.renderList(); this.renderPage(); toast("Devis supprimé");
      }
    },

    /* -------------------------- téléphone -------------------------- */
    renderM() {
      const m = this.m;
      if (m.screen === "av") {
        if (window.avenantsView.renderM(this.el, m.draft)) return;
        m.screen = "step"; m.step = 5;
      }
      if (m.screen === "list") {
        const items = this.sortedList();
        this.el.innerHTML = `<div class="dv dv-m">
          <h1>Création de devis</h1>
          <button class="btn primary big" data-act="mnew">+ Nouveau devis</button>
          <button class="btn big" data-act="mclients">Clients enregistrés</button>
          ${items.length ? `<div class="dv-sub" style="margin:4px 0 0">Devis récents</div>` + items.map(d => { const t = calc(normalize(d));
            return `<button class="dv-mitem" data-mopen="${d.id}"><b>${esc(d.client?.nom || "Client à renseigner")}</b><em>${eur.format(t.ttc)}</em>
              <span>${esc(d.numero)}${d.client?.ref ? " · " + esc(d.client.ref) : ""}</span><span class="statut ${d.statut || "brouillon"}">${STATUTS[d.statut] || "Brouillon"}</span></button>`; }).join("")
            : `<div class="dv-card dv-sub" style="margin:0">Aucun devis pour l'instant. Touchez « Nouveau devis » : RSG vous pose les questions une par une.</div>`}
        </div>`;
        return;
      }
      const d = m.draft, t = calc(d), st = m.step;
      const steps = `<div class="dv-steps">${[1, 2, 3, 4, 5].map(i => `<i class="${i <= st ? "on" : ""}"></i>`).join("")}</div>`;
      let h = "";
      if (st === 1) {
        const q = norm(m.cq), clients = store.list("clients").filter(c => !q || norm(c.nom + " " + (c.adr2 || "")).includes(q)).sort((a, b) => a.nom.localeCompare(b.nom, "fr"));
        h = `<p class="dv-q">Qui est le client ?</p><p class="dv-sub">Choisissez-le : son adresse se remplit toute seule.</p>
          ${m.newClient ? `<div class="dv-card" style="display:flex;flex-direction:column;gap:10px">
              <b>Nouveau client</b>
              <div class="field"><label for="mcNom">Nom</label><input id="mcNom" data-m="client.nom" value="${esc(d.client.nom)}" placeholder="M. Jean Dupont" autocomplete="name"></div>
              <div class="field"><label for="mcA1">Adresse</label><input id="mcA1" data-m="client.adr1" value="${esc(d.client.adr1)}" placeholder="12 rue des Forges" autocomplete="street-address"></div>
              <div class="field"><label for="mcA2">Code postal et ville</label><input id="mcA2" data-m="client.adr2" value="${esc(d.client.adr2)}" placeholder="54000 Nancy"></div>
              <div class="field"><label for="mcMail">E-mail (pour lui envoyer le devis)</label><input id="mcMail" type="email" inputmode="email" data-m="client.email" value="${esc(d.client.email)}" placeholder="client@exemple.fr"></div>
              <button class="btn" data-act="mclist">← Choisir un client existant</button></div>`
          : `<div class="field"><input id="mcq" data-mq value="${esc(m.cq)}" placeholder="Rechercher un client…" aria-label="Rechercher un client"></div>
            <div class="dv-pick">${clients.map(c => `<button data-client="${c.id}" class="${d.clientId === c.id ? "on" : ""}"><span class="av">${esc(initials(c.nom))}</span><div><b>${esc(c.nom)}</b><span>${esc([c.adr1, c.adr2].filter(Boolean).join(", ") || "Adresse à compléter")}</span></div></button>`).join("")
              || `<div class="dv-sub" style="margin:0">Aucun client trouvé.</div>`}</div>
            <button class="btn big" data-act="mnewclient">+ Nouveau client</button>
            <button class="btn" data-act="mclients">Gérer les clients enregistrés</button>`}`;
      } else if (st === 2) {
        h = `<p class="dv-q">Quel est le chantier ?</p><p class="dv-sub">Pour ${esc(d.client.nom || "le client")}.</p>
          <div class="field"><label for="mDate">Date du devis</label><input id="mDate" type="date" data-m="date" value="${esc(d.date)}"></div>
          <p class="dv-sub" style="margin:-4px 0 0">Numéro du devis : <b id="mNum">${esc(d.numero || "—")}</b></p>
          <div class="field"><label for="mRef">Référence du chantier</label><input id="mRef" data-m="client.ref" value="${esc(d.client.ref)}" placeholder="Écrivez le chantier"></div>
          <div><div class="field"><label>Validité du devis</label></div><div class="dv-chips">${VALIDITES.map(v => `<button data-valid="${v}" class="${d.validite === v ? "on" : ""}">${v}</button>`).join("")}</div></div>
          `;
      } else if (st === 3) {
        h = `<p class="dv-q">Qu'est-ce que tu proposes ?</p><p class="dv-sub">Écris chaque ligne du devis : désignation, prix et quantité.</p>
          ${d.lignes.length ? d.lignes.map((l, i) => `<div class="dv-card dv-line">
              <div class="field"><label for="ml${i}">Désignation</label><textarea id="ml${i}" data-m="lignes.${i}.d" rows="2">${esc(l.d)}</textarea></div>
              <div class="row"><div class="field"><label for="mp${i}">Prix unitaire HT (€)</label><input id="mp${i}" data-m="lignes.${i}.pu" inputmode="decimal" value="${esc(String(l.pu).replace(".", ","))}"></div>
                <div class="field"><label>Quantité</label><div class="stepper"><button data-qty="${i}|-1" aria-label="Moins">−</button><input data-m="lignes.${i}.q" inputmode="decimal" value="${esc(String(l.q).replace(".", ","))}" aria-label="Quantité"><button data-qty="${i}|1" aria-label="Plus">+</button></div></div></div>
              <div class="lt"><button class="del" data-mdel="${i}">Retirer</button><span>Total HT <b data-mcalc="${i}">${eur.format(num(l.pu) * num(l.q))}</b></span></div></div>`).join("")
          : ""}
          <button class="btn big" data-act="maddline">+ Ajouter une ligne</button>`;
      } else if (st === 4) {
        h = `<p class="dv-q">Les conditions</p>
          <div><div class="field"><label>TVA</label></div><div class="dv-seg">${TVAS.map(v => `<button data-tva="${v}" class="${d.tva === v ? "on" : ""}">${pctTxt(v).replace(",00", "")}</button>`).join("")}</div>
            ${d.tva === 0 ? `<p class="dv-sub" style="margin:8px 0 0">La mention « ${AUTOLIQ} » est ajoutée au devis.</p>` : ""}</div>
          <div><div class="field"><label>Acomptes (signature / début / 3/4 / solde)</label></div><div class="dv-chips">${PRESETS.map(([n, v], i) => `<button data-preset="${i}" class="${String(v) === String(d.acomptes.map(num)) ? "on" : ""}">${n}</button>`).join("")}</div></div>
          <div class="dv-card dv-sum">${ACOMPTES.map((lab, i) => `<div><span>${lab.replace(" :", "")} (${num(d.acomptes[i])} %)</span><b>${eur.format(t.ttc * num(d.acomptes[i]) / 100)}</b></div>`).join("")}</div>`;
      } else {
        h = `${m.id && window.avenantsView ? window.avenantsView.tabsHtml(m.id, d.numero, null) : ""}<p class="dv-q">Votre devis est prêt</p><p class="dv-sub">${esc(d.numero)} · ${esc(d.client.nom || "")}</p>
          <div class="dv-prev" id="mPrev">${page(d, false)}</div>
          <div class="dv-card dv-sum"><div><span>Total HT</span><b>${eur.format(t.ht)}</b></div><div><span>TVA ${pctTxt(d.tva)}</span><b>${eur.format(t.tv)}</b></div><div class="ttc"><span>Total TTC</span><b>${eur.format(t.ttc)}</b></div></div>
          <div class="dv-card"><div class="field"><label>Statut du devis</label></div>
            <div class="dv-seg dv-statut">${Object.entries(STATUTS).map(([k, v]) => `<button data-mstatut="${k}" class="st-${k} ${(d.statut || "brouillon") === k ? "on" : ""}" aria-pressed="${(d.statut || "brouillon") === k}">${v}</button>`).join("")}</div></div>
          <button class="btn primary big" data-act="mmail">Envoyer par mail</button>
          <button class="btn big" data-act="mfacturer">Facturer ce devis</button>
          <button class="btn big" data-act="mpdf">Enregistrer le PDF</button>
          <div class="dv-edit"><span class="dv-sub" style="margin:0;align-self:center">Modifier :</span>${["Client", "Chantier", "Prestations", "Conditions"].map((n, i) => `<button data-goto="${i + 1}">${n}</button>`).join("")}</div>
          <button class="btn danger" data-act="mdelete">Supprimer ce devis</button>`;
      }
      this.el.innerHTML = `<div class="dv dv-m"><button class="btn dv-back" data-act="mtolist">‹ Liste des devis</button>${steps}${h}</div>
        <div class="dv-nav"><button class="btn" data-act="mback">${st === 1 || st === 5 ? "Liste" : "Retour"}</button>
          <div class="tt">Total TTC<b id="mTot">${eur.format(t.ttc)}</b></div>
          ${st < 5 ? `<button class="btn primary" data-act="mnext">${st === 4 ? "Voir le devis" : "Suivant"}</button>` : ""}</div>`;
      if (st === 5) fitPreview();
      this.el.querySelectorAll(".dv-m textarea").forEach(grow);
    },
    async saveM() {
      const m = this.m; if (!canEdit()) return false;
      try { m.id = await persist(m.id, m.draft); return true; }
      catch (e) { toast(frError(e)); return false; }
    },
    onInputM(e) {
      const t = e.target, m = this.m;
      if (m.screen === "av") return window.avenantsView.onInputM(e);
      if (t.dataset.mq !== undefined) { m.cq = t.value; const pos = caretOf(t); this.renderM(); const i = $("#mcq"); i && (i.focus(), i.setSelectionRange(pos, pos)); return; }
      const f = t.dataset.m; if (!f) return;
      let v = t.value; if (/\.(pu|q)$/.test(f)) v = v.replace(",", ".");
      setPath(m.draft, f, v);
      if (f === "client.nom") m.draft.clientId = null;
      if (f === "date" && v) { m.draft.numero = nextNumero(v, m.id); const n = $("#mNum"); if (n) n.textContent = m.draft.numero; }
      if (t.tagName === "TEXTAREA") grow(t);
      const mm = f.match(/^lignes\.(\d+)\./);
      if (mm) { const l = m.draft.lignes[+mm[1]], c = this.el.querySelector(`[data-mcalc="${mm[1]}"]`); if (c) c.textContent = eur.format(num(l.pu) * num(l.q)); }
      $("#mTot").textContent = eur.format(calc(m.draft).ttc);
    },
    onChangeM(e) {
      if (e.target.dataset.act === "mliv") { this.m.liv = e.target.checked; if (!e.target.checked) { this.m.draft.client.liv1 = this.m.draft.client.liv2 = ""; } this.renderM(); }
    },
    async onClickM(e) {
      const m = this.m, g = s => e.target.closest(s), AV = window.avenantsView;
      if (g("[data-dsub]")) {
        if (m.screen === "av" && AV.m.screen === "edit") await AV.saveDraftM();
        const v = g("[data-dsub]").dataset.dsub;
        if (v && AV.load(v)) m.screen = "av"; else { m.screen = "step"; m.step = 5; }
        window.scrollTo(0, 0); return this.renderM();
      }
      if (g('[data-act="newav"]')) {
        if (!canEdit()) return;
        if (m.screen === "av" && AV.m.screen === "edit") await AV.saveDraftM();
        if (m.screen !== "av" && !(await this.saveM())) return;
        try { await AV.create(m.id); m.screen = "av"; window.scrollTo(0, 0); this.renderM(); } catch (err) { toast(frError(err)); }
        return;
      }
      if (m.screen === "av") {
        if (g('[data-act="avback"]')) { if (AV.m.screen === "edit") await AV.saveDraftM(); m.screen = "step"; m.step = 5; window.scrollTo(0, 0); return this.renderM(); }
        const r = await AV.onClickM(e);
        if (r === "deleted") { m.screen = "step"; m.step = 5; return this.renderM(); }
        if (r === "rerender") return this.renderM();
        if (r === "rerender-focus") { this.renderM(); const t = this.el.querySelectorAll(".dv-line textarea"); const last = t[t.length - 1]; last && (last.scrollIntoView({ block: "center" }), last.focus()); }
        return;
      }
      if (g("[data-mopen]")) { m.id = g("[data-mopen]").dataset.mopen; m.draft = normalize(store.get("devis", m.id)); m.screen = "step"; m.step = 5; return this.renderM(); }
      if (g("[data-client]")) {
        const c = store.get("clients", g("[data-client]").dataset.client);
        Object.assign(m.draft.client, { nom: c.nom, adr1: c.adr1 || "", adr2: c.adr2 || "", email: c.email || "" }); m.draft.clientId = c.id;
        toast(`${c.nom} : adresse remplie automatiquement`); m.step = 2; await this.saveM(); return this.renderM();
      }
      if (g("[data-valid]")) { m.draft.validite = g("[data-valid]").dataset.valid; return this.renderM(); }
      if (g("[data-qty]")) { const [i, s] = g("[data-qty]").dataset.qty.split("|").map(Number); const l = m.draft.lignes[i]; l.q = Math.max(0, num(l.q) + s); return this.renderM(); }
      if (g("[data-mdel]")) { m.draft.lignes.splice(+g("[data-mdel]").dataset.mdel, 1); return this.renderM(); }
      if (g("[data-tva]")) { m.draft.tva = +g("[data-tva]").dataset.tva; return this.renderM(); }
      if (g("[data-preset]")) { m.draft.acomptes = PRESETS[+g("[data-preset]").dataset.preset][1].slice(); return this.renderM(); }
      if (g("[data-mstatut]")) {
        if (!canEdit()) return;
        const k = g("[data-mstatut]").dataset.mstatut; if (m.draft.statut === k) return;
        m.draft.statut = k; if (await this.saveM()) { toast("Statut : " + STATUTS[k]); this.renderM(); }
        return;
      }
      if (g("[data-goto]")) { m.step = +g("[data-goto]").dataset.goto; return this.renderM(); }
      const a = g("[data-act]")?.dataset.act; if (!a) return;
      if (a === "mnew") { if (!canEdit()) return; m.draft = normalize(blank()); m.id = null; m.screen = "step"; m.step = 1; m.newClient = false; m.cq = ""; m.liv = false; return this.renderM(); }
      if (a === "mclients") return openClients(() => this.renderM());
      if (a === "mnewclient") { m.newClient = true; m.draft.clientId = null; m.draft.client = Object.assign(m.draft.client, { nom: "", adr1: "", adr2: "", email: "" }); this.renderM(); $("#mcNom")?.focus(); return; }
      if (a === "mclist") { m.newClient = false; return this.renderM(); }
      if (a === "maddline") { m.draft.lignes.push({ d: "", pu: "", q: 1 }); this.renderM(); const k = m.draft.lignes.length - 1, t = $("#ml" + k); t && (t.scrollIntoView({ block: "center" }), t.focus()); return; }
      if (a === "mtolist") { if (m.draft && (m.id || m.draft.client.nom)) await this.saveM(); m.screen = "list"; window.scrollTo(0, 0); return this.renderM(); }
      if (a === "mback") {
        if (m.step === 1 || m.step === 5) { if (m.draft && (m.id || m.draft.client.nom)) await this.saveM(); m.screen = "list"; return this.renderM(); }
        m.step--; return this.renderM();
      }
      if (a === "mnext") {
        if (m.step === 1 && !m.draft.client.nom.trim()) { toast("Choisissez ou créez le client."); return; }
        if (m.step === 2 && !m.draft.date) { toast("Indiquez la date du devis."); return; }
        if (m.step === 3 && !cleanLines(m.draft).lignes.some(l => String(l.d).trim())) { toast("Écrivez au moins une ligne (désignation)."); return; }
        m.draft = cleanLines(m.draft);
        if (!(await this.saveM())) return;
        m.step++; window.scrollTo(0, 0);
        if (m.step === 3 && !m.draft.lignes.length) m.draft.lignes.push({ d: "", pu: "", q: 1 });   // une première ligne vide, prête à remplir
        return this.renderM();
      }
      if (a === "mfacturer") { if (!canEdit()) return; if (await this.saveM()) window.factureTab?.openCreate(m.id); return; }
      if (a === "mpdf") { try { const r = await savePDF(m.draft); if (r === "saved") toast("PDF enregistré : " + fileName(m.draft)); } catch (err) { if (err?.code !== "declined") toast(err.message || String(err)); } }
      if (a === "mmail") openMail(m.draft, async () => { if (m.draft.statut === "brouillon") { m.draft.statut = "envoye"; await this.saveM(); } toast("Devis envoyé"); });
      if (a === "mdelete") {
        const b = g("[data-act]");
        if (!b.classList.contains("armed")) { b.classList.add("armed"); b.textContent = "Confirmer la suppression"; return; }
        if (m.id) await store.remove("devis", m.id).catch(err => toast(frError(err)));
        m.screen = "list"; m.id = null; toast("Devis supprimé"); this.renderM();
      }
    },
  };

  function grow(t) { t.style.height = "auto"; t.style.height = t.scrollHeight + "px"; }
  function caretOf(el) { try { return el.selectionStart; } catch { return null; } }
  function refocus(root, f, pos) {
    const el = root.querySelector(`[data-f="${f}"]`); if (!el) return;
    el.focus(); if (pos != null) try { el.setSelectionRange(pos, pos); } catch {}
  }
  function fitPreview() {
    const box = $("#mPrev"); if (!box) return;
    const a4 = box.firstElementChild, s = box.clientWidth / 796;
    a4.style.transform = `scale(${s})`; box.style.height = (a4.offsetHeight * s) + "px";
  }
  addEventListener("resize", () => fitPreview());

  window.devisTab = tab;
  // outils partagés avec l'onglet « Création de facture »
  window.devisKit = { ENT, LOGO, AUTOLIQ, TVAS, ACOMPTES, eur, pctTxt, money, num, isoToday, frDate, mobile, clone, initials,
    calc, normalize, cleanLines, setPath, cleanName, makePDF, savePDF, openMail, saveClientOf, grow, caretOf, refocus, fitPreview };
})();
