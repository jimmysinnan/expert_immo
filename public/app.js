// ─────────────────────────────────────────────────────────────────────────────
// ExpertIA — app.js
// Logique frontend : navigation, formulaire, appels API, rendu rapport
// ─────────────────────────────────────────────────────────────────────────────

// ── CADASTRE AUTO ────────────────────────────────────────────────────────────
async function fetchCadastre() {
  const adresse = document.getElementById('adresse_bien').value.trim();
  if (!adresse) return alert('Saisissez d\'abord l\'adresse du bien.');
  const btn = document.getElementById('btn-cadastre');
  btn.textContent = '⏳ Recherche…'; btn.disabled = true;
  try {
    // 1. Geocoder l'adresse → coords + code INSEE
    const geo = await fetch(`https://api-adresse.data.gouv.fr/search/?q=${encodeURIComponent(adresse)}&limit=1`).then(r => r.json());
    const feat = geo.features?.[0];
    if (!feat) throw new Error('Adresse non reconnue par le géocodeur');
    const [lon, lat] = feat.geometry.coordinates;

    // 2. Récupérer les parcelles via apicarto IGN (France métro + DOM : 971/972/973/974/976)
    const cadastre = await fetch(
      `https://apicarto.ign.fr/api/cadastre/parcelle?lon=${lon}&lat=${lat}`
    ).then(r => { if (!r.ok) throw new Error(`API cadastre : ${r.status}`); return r.json(); });

    const parcels = cadastre.features ?? [];
    if (!parcels.length) throw new Error('Aucune parcelle trouvée à cette adresse');

    // 3. Formater les références : "AB 0042, AB 0043"
    const refs = [...new Set(parcels.map(p => {
      const { section, numero } = p.properties;
      return `${section.trim()} ${String(numero).trim()}`;
    }))].join(', ');

    document.getElementById('refs_cadastrales').value = refs;
  } catch (e) {
    alert('Cadastre : ' + e.message);
  } finally {
    btn.textContent = '⊕ Cadastre'; btn.disabled = false;
  }
}

// ── HELPER BASE64 ────────────────────────────────────────────────────────────
function fileToBase64(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => {
      const base64 = reader.result.split(',')[1];
      resolve({ data: base64, mimeType: file.type || 'image/jpeg', name: file.name });
    };
    reader.onerror = reject;
    reader.readAsDataURL(file);
  });
}

// ── STATE ────────────────────────────────────────────────────────────────────
const state = {
  currentStep: 0,
  refDoc: null,           // File — rapport de référence
  photos: {               // FileList par catégorie
    terrain: null,
    ext: null,
    int: null,
    desordres: []         // Array de File (multi-blocs désordres)
  },
  photos64: {             // Photos converties en base64 pour l'export DOCX
    terrain: [],
    ext: [],
    int: [],
    desordres: []
  },
  logo: null,             // Logo extrait du DOCX de référence { data, mimeType }
  ficheLogo: null,        // Logo uploadé manuellement pour la fiche dossier
  chapter1: '',
  marche_immobilier: '',
  style: null,
  photoResults: {},
  reportMarkdown: '',
  sections: null,         // JSON sections JALTA pour export DOCX
  formData: {}
};

const STEPS = 6;

// ── NAVIGATION PAGES ─────────────────────────────────────────────────────────
function showPage(id) {
  document.querySelectorAll('.page').forEach(p => p.classList.remove('active'));
  document.querySelectorAll('.nav-tab').forEach(t => t.classList.remove('active'));
  document.getElementById('page-' + id).classList.add('active');
  document.getElementById('tab-' + id).classList.add('active');
  window.scrollTo(0, 0);
}

// ── STEPPER ───────────────────────────────────────────────────────────────────
function canGoStep(i) {
  if (i <= state.currentStep) goStep(i);
}

function goStep(i) {
  for (let j = 0; j < STEPS; j++) {
    const sec = document.getElementById('sec-' + j);
    if (sec) sec.style.display = j === i ? 'block' : 'none';
    const s = document.getElementById('s' + j);
    const sc = document.getElementById('sc' + j);
    if (!s || !sc) continue;
    s.className = 'step' + (j === i ? ' active' : j < i ? ' done' : '');
    sc.textContent = j < i ? '✓' : 'ABCDEF'[j];
  }
  state.currentStep = i;
  window.scrollTo(0, 0);
}

function nextStep() {
  if (state.currentStep < STEPS - 1) goStep(state.currentStep + 1);
}
function prevStep() {
  if (state.currentStep > 0) goStep(state.currentStep - 1);
}

// ── UPLOAD RAPPORT DE RÉFÉRENCE ───────────────────────────────────────────────
function handleRefDoc(input) {
  if (!input.files?.[0]) return;
  state.refDoc = input.files[0];
  const bar = document.getElementById('doc-upload-bar');
  const status = document.getElementById('doc-status');
  bar.style.borderColor = 'var(--green)';
  status.style.display = 'block';
  status.textContent = '✓ ' + input.files[0].name;
  toast('Rapport de référence chargé — le style sera extrait lors de la génération');
}

// ── UPLOAD PHOTOS ─────────────────────────────────────────────────────────────
function handlePhotos(cat, input) {
  if (!input.files?.length) return;
  state.photos[cat] = input.files;
  const uz = document.getElementById('uz-' + cat);
  const count = document.getElementById('count-' + cat);
  if (uz) uz.classList.add('has-files');
  if (count) {
    count.style.display = 'block';
    count.textContent = input.files.length + ' photo' + (input.files.length > 1 ? 's' : '') + ' sélectionnée' + (input.files.length > 1 ? 's' : '');
  }
}

// ── DÉSORDRES DYNAMIQUES ──────────────────────────────────────────────────────
let _desordreCount = 0;

function desordreBlockHTML(n) {
  return `
  <div class="card desordre-block" data-d="${n}">
    <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:8px">
      <div class="desordre-num">— Désordre ${n} —</div>
      <button class="btn btn-glass btn-sm" style="color:#c0392b;padding:2px 10px;font-size:13px" onclick="removeDesordre(this)" title="Supprimer ce désordre">✕</button>
    </div>
    <div class="grid2">
      <div class="field">
        <div class="field-top"><label>Localisation</label></div>
        <input type="text" data-field="loc" placeholder="Ex : Façade Nord — partie basse, Plafond cuisine RDC…">
      </div>
      <div class="field">
        <div class="field-top"><label>Niveau de gravité</label></div>
        <select data-field="grav">
          <option value="">— Sélectionner —</option>
          <option value="Esthétique">Esthétique</option>
          <option value="Fonctionnel">Fonctionnel</option>
          <option value="Structurel">Structurel</option>
        </select>
      </div>
    </div>
    <div class="field">
      <div class="field-top"><label>Nature et description précise</label></div>
      <textarea data-field="nat" style="min-height:60px" placeholder="Décrivez précisément le désordre observé…"></textarea>
    </div>
    <div class="field">
      <div class="field-top"><label>Origine probable</label></div>
      <input type="text" data-field="orig" placeholder="Ex : Tassement différentiel, infiltration, vétusté…">
    </div>
    <div class="field">
      <div class="field-top"><label>Photo(s) du désordre</label></div>
      <div class="upload-zone" onclick="this.querySelector('input').click()" style="padding:12px">
        <label>
          <div class="upload-label" style="font-size:12px">Photos désordre ${n} — Cliquer pour choisir</div>
          <div class="upload-hint">1 à 3 photos</div>
          <div class="upload-count" style="display:none"></div>
          <input type="file" multiple accept="image/*" onchange="handleDesordrePhotoBlock(this)">
        </label>
      </div>
    </div>
  </div>`;
}

function initDesordres() {
  _desordreCount = 0;
  const container = document.getElementById('desordres-container');
  if (!container) return;
  container.innerHTML = '';
  addDesordre();
  // Bouton +
  const addBtn = document.createElement('div');
  addBtn.style.cssText = 'text-align:center;margin:12px 0';
  addBtn.innerHTML = `<button class="btn btn-navy" onclick="addDesordre()" style="gap:6px">+ Ajouter un désordre</button>`;
  container.after(addBtn);
}

function addDesordre() {
  _desordreCount++;
  const container = document.getElementById('desordres-container');
  if (!container) return;
  const div = document.createElement('div');
  div.innerHTML = desordreBlockHTML(_desordreCount);
  container.appendChild(div.firstElementChild);
}

function removeDesordre(btn) {
  const block = btn.closest('.desordre-block');
  if (document.querySelectorAll('.desordre-block').length <= 1) {
    block.querySelectorAll('input[type="text"], textarea').forEach(el => el.value = '');
    block.querySelectorAll('select').forEach(el => el.selectedIndex = 0);
    return;
  }
  block.remove();
}

function handleDesordrePhotoBlock(input) {
  if (!input.files?.length) return;
  const files = Array.from(input.files);
  state.photos.desordres.push(...files);
  const uz = input.closest('.upload-zone');
  const count = uz?.querySelector('.upload-count');
  if (uz) uz.classList.add('has-files');
  if (count) { count.style.display = 'block'; count.textContent = files.length + ' photo(s)'; }
}

function handleDesordrePhoto(i, input) {
  if (!input.files?.length) return;
  const files = Array.from(input.files);
  // Stocker les photos désordres dans state.photos.desordres (tableau de File)
  state.photos.desordres.push(...files);
  const uz = document.getElementById('uz-d' + i);
  const count = document.getElementById('count-d' + i);
  if (uz) uz.classList.add('has-files');
  if (count) {
    count.style.display = 'block';
    count.textContent = files.length + ' photo(s)';
  }
}

// ── INIT SURFACES (15 lignes) ─────────────────────────────────────────────────
function initSurfaces() {
  const tbody = document.getElementById('surf-body');
  if (!tbody) return;
  const types = [
    'Séjour / Salon','Cuisine','Salle à manger','Chambre','Bureau',
    'Salle de bain','Salle d\'eau','WC','Couloir / Dégagement',
    'Dressing','Cave','Cellier / Buanderie','Garage','Véranda','Autre'
  ];
  const niveaux = ['Rez-de-chaussée','Niveau 1 (R+1)','Niveau 2 (R+2)','Sous-sol','Combles','Annexe'];
  tbody.innerHTML = '';
  for (let i = 1; i <= 15; i++) {
    tbody.innerHTML += `<tr>
      <td><select id="surf_type_${i}">
        <option value="">— Type —</option>
        ${types.map(t => `<option>${t}</option>`).join('')}
      </select></td>
      <td><input type="text" id="surf_prec_${i}" placeholder="Ex : Chambre 1, WC RDC…"></td>
      <td><select id="surf_niv_${i}">
        ${niveaux.map(n => `<option>${n}</option>`).join('')}
      </select></td>
      <td><input type="number" id="surf_m2_${i}" placeholder="m²" min="0" step="0.5" style="text-align:right;max-width:80px"></td>
    </tr>`;
  }
}

// ── COLLECTE DES DONNÉES DU FORMULAIRE ────────────────────────────────────────
function collectFormData() {
  const get = id => (document.getElementById(id)?.value || '').trim();
  const checks = name => Array.from(document.querySelectorAll(`input[name="${name}"]:checked`)).map(c => c.value).join(', ');

  // Désordres (blocs dynamiques)
  let desordresText = '';
  let dIdx = 0;
  document.querySelectorAll('.desordre-block').forEach(block => {
    const loc = (block.querySelector('[data-field="loc"]')?.value || '').trim();
    const nat = (block.querySelector('[data-field="nat"]')?.value || '').trim();
    const grav = (block.querySelector('[data-field="grav"]')?.value || '').trim();
    const orig = (block.querySelector('[data-field="orig"]')?.value || '').trim();
    if (!loc && !nat) return;
    dIdx++;
    desordresText += `**Désordre ${dIdx} — ${loc || '[localisation à compléter]'}**\n`;
    desordresText += `Nature : ${nat || '[à compléter]'}\n`;
    desordresText += `Gravité : ${grav || '[à compléter]'}\n`;
    desordresText += `Origine probable : ${orig || '[à compléter]'}\n\n`;
  });

  // Surfaces
  let surfacesText = '';
  const surfacesArray = [];
  for (let i = 1; i <= 15; i++) {
    const type = get(`surf_type_${i}`);
    const m2 = get(`surf_m2_${i}`);
    if (!type && !m2) continue;
    const prec = get(`surf_prec_${i}`);
    const niv = get(`surf_niv_${i}`);
    surfacesText += `| ${type}${prec ? ' — ' + prec : ''} | ${niv} | ${m2} |\n`;
    surfacesArray.push({ type, prec, niveau: niv, m2 });
  }

  return {
    ref_dossier: get('ref_dossier') || 'EXP-2025-XXX',
    date_visite: get('date_visite') || new Date().toLocaleDateString('fr-FR'),
    type_mission: get('type_mission'),
    donneur_ordre: get('donneur_ordre'),
    nom_donneur_ordre: get('nom_donneur_ordre'),
    adresse_bien: get('adresse_bien'),
    refs_cadastrales: get('refs_cadastrales'),
    regime_juridique: get('regime_juridique'),
    dpe_classe: get('dpe_classe'),
    ges_classe: get('ges_classe'),
    type_bien: get('type_bien'),
    annee_construction: get('annee_construction'),
    nb_niveaux: get('nb_niveaux'),
    sous_sol: get('sous_sol'),
    annexes: checks('annexes'),
    notes_bien: get('notes_bien'),
    // Appartement / Immeuble
    appart_etage: get('appart_etage'),
    appart_type_pieces: get('appart_type_pieces'),
    copro_type: get('copro_type'),
    copro_nb_batiments: get('copro_nb_batiments'),
    copro_composition: get('copro_composition'),
    copro_tantiemes: get('copro_tantiemes'),
    superficie_terrain: get('superficie_terrain'),
    forme_terrain: get('forme_terrain') === 'Autre' ? (get('forme_terrain_autre') || 'Autre') : get('forme_terrain'),
    topographie: get('topographie'),
    orientation: get('orientation'),
    acces_terrain: get('acces_terrain'),
    clotures: get('clotures'),
    reseaux: checks('reseaux'),
    contraintes: checks('contraintes') || 'Aucune contrainte connue',
    zonage_plu: get('zonage_plu'),
    notes_terrain: get('notes_terrain'),
    type_construction: get('type_construction'),
    materiau_toiture: get('materiau_toiture'),
    forme_toiture: get('forme_toiture'),
    etat_toiture: get('etat_toiture'),
    materiau_facades: get('materiau_facades'),
    etat_facades: get('etat_facades'),
    menuiseries_ext: get('menuiseries_ext'),
    chauffage: get('chauffage'),
    etat_electrique: get('etat_electrique'),
    etat_plomberie: get('etat_plomberie'),
    sols_interieurs: get('sols_interieurs'),
    revetements_murs: get('revetements_murs'),
    notes_bati: get('notes_bati'),
    desordres: desordresText,
    surfaces: surfacesText,
    surfaces_array: surfacesArray,
    situation_locative: get('situation_locative'),
    assainissement: get('assainissement'),
    documents_fournis: Array.from(document.querySelectorAll('input[name="documents_fournis"]:checked')).map(c => c.value).join(', '),
  };
}

// ── GÉNÉRATION PRINCIPALE ─────────────────────────────────────────────────────
async function startGeneration() {
  const formData = collectFormData();
  state.formData = formData;

  if (!formData.adresse_bien) {
    toast('Veuillez saisir l\'adresse du bien (Section A)');
    goStep(0);
    return;
  }

  showPage('gen');
  resetPipeline();

  try {
    // ── ÉTAPE 0 : Analyse données
    setStep(0, 'active');
    await sleep(400);
    setStep(0, 'done');

    // ── ÉTAPE 1 : Chapitre 1 géographique
    setStep(1, 'active');
    updateDetail(1, `Recherche pour : ${formData.adresse_bien}`);
    try {
      const r1 = await fetchJSON('/api/chapter1', { adresse: formData.adresse_bien });
      state.chapter1 = r1.text || '';
      state.marche_immobilier = r1.marche_immobilier || '';
      updateDetail(1, 'Chapitre 1 généré ✓');
      setStep(1, 'done');
    } catch (e) {
      setStep(1, 'done'); // Non bloquant
      state.chapter1 = '[Chapitre 1 — données géographiques non disponibles — À COMPLÉTER PAR L\'EXPERT]';
      updateDetail(1, 'Recherche web non disponible — à compléter manuellement');
    }

    // ── ÉTAPE 2 : Extraction style + logo
    setStep(2, 'active');
    if (state.refDoc) {
      updateDetail(2, `Analyse de ${state.refDoc.name}`);
      try {
        const fd2 = new FormData();
        fd2.append('document', state.refDoc);
        const r2 = await fetch('/api/extract-style', { method: 'POST', body: fd2 });
        const j2 = await r2.json();
        state.style = j2.style || null;
        state.logo = j2.logo || null;
        updateDetail(2, state.logo ? 'Style extrait + logo récupéré ✓' : 'Style extrait ✓');
      } catch (e) {
        state.style = null;
        state.logo = null;
        updateDetail(2, 'Extraction style échouée — style générique utilisé');
      }
    } else {
      updateDetail(2, 'Aucun rapport de référence — style professionnel standard');
      await sleep(600);
    }
    setStep(2, 'done');

    // ── ÉTAPE 3 : Analyse photos + conversion base64
    setStep(3, 'active');
    const hasPhotos = state.photos.terrain || state.photos.ext || state.photos.int || state.photos.desordres.length;
    if (hasPhotos) {
      updateDetail(3, 'Analyse des photos en cours…');
      try {
        const fd3 = new FormData();
        if (state.photos.terrain) Array.from(state.photos.terrain).forEach(f => fd3.append('terrain', f));
        if (state.photos.ext) Array.from(state.photos.ext).forEach(f => fd3.append('ext', f));
        if (state.photos.int) Array.from(state.photos.int).forEach(f => fd3.append('int', f));
        state.photos.desordres.forEach(f => fd3.append('desordres', f));
        const r3 = await fetch('/api/analyze-photos', { method: 'POST', body: fd3 });
        state.photoResults = await r3.json();

        // Conversion des photos en base64 pour l'export DOCX
        const p64 = state.photos64;
        if (state.photos.terrain) p64.terrain = await Promise.all(Array.from(state.photos.terrain).map(fileToBase64));
        if (state.photos.ext)     p64.ext     = await Promise.all(Array.from(state.photos.ext).map(fileToBase64));
        if (state.photos.int)     p64.int     = await Promise.all(Array.from(state.photos.int).map(fileToBase64));
        if (state.photos.desordres.length) p64.desordres = await Promise.all(state.photos.desordres.map(fileToBase64));

        const nbCats = Object.values(state.photoResults).filter(Boolean).length;
        updateDetail(3, `${nbCats} lot(s) de photos analysés ✓`);
      } catch (e) {
        state.photoResults = {};
        updateDetail(3, 'Analyse photos échouée — descriptions textuelles utilisées');
      }
    } else {
      state.photoResults = {};
      updateDetail(3, 'Aucune photo fournie — descriptions textuelles uniquement');
      await sleep(400);
    }
    setStep(3, 'done');

    // ── ÉTAPE 4 : Génération principale
    setStep(4, 'active');
    updateDetail(4, 'Rédaction des 3 chapitres en cours…');
    const payload = {
      formData,
      chapter1: state.chapter1,
      style: state.style,
      photos: {
        terrain: state.photoResults.terrain || null,
        ext: state.photoResults.ext || null,
        int: state.photoResults.int || null,
        desordres: state.photoResults.desordres || null
      },
      desordres: formData.desordres,
      surfaces: formData.surfaces
    };
    const r4 = await fetchJSON('/api/generate', payload);
    state.reportMarkdown = r4.report || '';
    // Injecter marche_immobilier depuis chapter1 dans les sections
    if (r4.sections) {
      r4.sections.marche_immobilier = r4.sections.marche_immobilier || state.marche_immobilier || '';
    }
    state.sections = r4.sections || null;
    updateDetail(4, 'Rapport rédigé ✓');
    setStep(4, 'done');

    // ── ÉTAPE 5 : Finalisation
    setStep(5, 'active');
    updateDetail(5, 'Rapport prêt — export .docx disponible');
    await sleep(500);
    setStep(5, 'done');

    // Affichage résultat
    const pdMeta = document.getElementById('pd-meta');
    if (pdMeta) pdMeta.textContent = `Dossier ${formData.ref_dossier} · ${formData.adresse_bien} · Rapport prêt pour validation.`;
    document.getElementById('pipeline-done').style.display = 'block';
    renderReport();

  } catch (err) {
    console.error(err);
    toast('Erreur : ' + (err.message || 'Erreur serveur. Vérifiez votre clé API dans .env'));
  }
}

// ── HELPERS PIPELINE ──────────────────────────────────────────────────────────
function resetPipeline() {
  for (let i = 0; i < 6; i++) {
    const ps = document.getElementById('ps' + i);
    const pst = document.getElementById('pst' + i);
    if (ps) ps.className = 'pipe-step waiting';
    if (pst) pst.innerHTML = '○';
  }
  const done = document.getElementById('pipeline-done');
  if (done) done.style.display = 'none';
}

function setStep(i, status) {
  const ps = document.getElementById('ps' + i);
  const pst = document.getElementById('pst' + i);
  if (!ps || !pst) return;
  ps.className = 'pipe-step ' + status;
  if (status === 'active') pst.innerHTML = '<div class="spinner"></div>';
  else if (status === 'done') pst.innerHTML = '✓';
  else if (status === 'error') pst.innerHTML = '✕';
  else pst.innerHTML = '○';
}

function updateDetail(i, text) {
  const pd = document.getElementById('pd' + i);
  if (pd) pd.textContent = text;
}

// ── RENDU DU RAPPORT ──────────────────────────────────────────────────────────
function renderReport() {
  const container = document.getElementById('rapport-content');
  if (!container || !state.reportMarkdown) return;
  container.innerHTML = markdownToHtml(state.reportMarkdown);
}

function markdownToHtml(md) {
  let html = '';
  const lines = md.split('\n');
  let inTable = false;
  let tableRows = [];

  const flushTable = () => {
    if (!tableRows.length) return;
    const validRows = tableRows.filter(r => !r.match(/^\|[\s-|]+\|$/));
    html += '<div style="overflow-x:auto"><table class="surf-r-table">';
    validRows.forEach((row, ri) => {
      const cells = row.split('|').filter((_, i, a) => i > 0 && i < a.length - 1).map(c => c.trim());
      html += '<tr' + (ri === 0 ? '' : ri === validRows.length - 1 && cells[0].includes('**') ? ' class="total"' : '') + '>';
      cells.forEach(cell => {
        const tag = ri === 0 ? 'th' : 'td';
        html += `<${tag}>${inlineHtml(cell)}</${tag}>`;
      });
      html += '</tr>';
    });
    html += '</table></div>';
    tableRows = [];
    inTable = false;
  };

  for (let i = 0; i < lines.length; i++) {
    const line = lines[i];

    if (line.startsWith('|')) {
      inTable = true;
      tableRows.push(line);
      continue;
    }
    if (inTable) flushTable();

    if (line.startsWith('# ')) {
      html += `<div class="rapport-header-card">
        <div class="rh-cabinet">Cabinet d'Expertise Immobilière</div>
        <div class="rh-title">${line.slice(2).trim()}</div>`;
    } else if (line.startsWith('## ') && line.includes('Pré-rapport')) {
      html += `<div class="rh-sub">${line.slice(3).trim()}</div>`;
      // Construire les métas depuis formData
      const fd = state.formData;
      html += `<div class="rh-meta">
        <div class="rm"><div class="rm-label">Référence</div><div class="rm-val">${fd.ref_dossier || '—'}</div></div>
        <div class="rm"><div class="rm-label">Date de visite</div><div class="rm-val">${fd.date_visite || '—'}</div></div>
        <div class="rm"><div class="rm-label">Mission</div><div class="rm-val">${fd.type_mission || '—'}</div></div>
        <div class="rm"><div class="rm-label">Bien</div><div class="rm-val">${fd.type_bien || '—'}</div></div>
        <div class="rm"><div class="rm-label">Adresse</div><div class="rm-val">${fd.adresse_bien || '—'}</div></div>
        <div class="rm"><div class="rm-label">Donneur d'ordre</div><div class="rm-val">${fd.nom_donneur_ordre || fd.donneur_ordre || '—'}</div></div>
      </div>`;
    } else if (line.match(/^\*Le présent document/)) {
      html += `<div class="rh-mention">${line.replace(/\*/g, '')}</div></div>`;
    } else if (line === '---') {
      // skip separators
    } else if (line.startsWith('## CHAPITRE') || line.startsWith('## Chapitre')) {
      html += `</div><div class="rapport-body"><div class="ch-header"><span class="ch-n">CH</span> ${inlineHtml(line.slice(3))}</div>`;
    } else if (line.startsWith('### ')) {
      html += `<div class="sub-h">${inlineHtml(line.slice(4))}</div>`;
    } else if (line.startsWith('**Désordre')) {
      // Bloc désordre
      const loc = line.replace(/\*\*/g, '').trim();
      const gravLine = lines[i + 2] || '';
      const gravVal = gravLine.includes('Structurel') ? 'stru' : gravLine.includes('Esthétique') ? 'esth' : 'fonc';
      const gravLabel = gravLine.replace('Gravité :', '').trim();
      html += `<div class="disorder-card ${gravVal === 'stru' ? 'stru' : ''}">
        <div class="disorder-head">
          <span class="grav-badge grav-${gravVal}">${gravLabel || 'Fonctionnel'}</span>
          <div class="disorder-title">${loc}</div>
        </div>`;
    } else if (line.startsWith('Nature :') || line.startsWith('Gravité :') || line.startsWith('Observation :') || line.startsWith('Origine probable :')) {
      const [key, ...rest] = line.split(':');
      html += `<div class="disorder-row"><strong>${key} :</strong>${inlineHtml(rest.join(':'))}</div>`;
      if (line.startsWith('Origine probable :')) html += '</div>'; // ferme disorder-card
    } else if (line.trim() === '') {
      html += '';
    } else if (line.trim()) {
      html += `<p class="rp">${inlineHtml(line)}</p>`;
    }
  }

  if (inTable) flushTable();
  // Fermer le dernier rapport-body ouvert
  html += `<div class="rapport-warning">⚠ Les surfaces ci-dessus sont issues des données saisies par l'expert. Elles devront être confirmées par mesurage avant tout acte.</div></div>`;

  // Boutons d'action
  html += `<div class="rapport-actions">
    <button class="btn btn-glass" onclick="showPage('form');canGoStep(0)">← Nouveau dossier</button>
    <button class="btn btn-navy" onclick="downloadDocx()">⬇ Télécharger .docx</button>
  </div>`;

  return html;
}

function inlineHtml(text) {
  return text
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/\*\*([^*]+)\*\*/g, '<strong>$1</strong>')
    .replace(/\*([^*]+)\*/g, '<em>$1</em>')
    .replace(/\[À COMPLÉTER PAR L'EXPERT\]/g, '<span style="color:var(--amber);font-weight:600">[À COMPLÉTER PAR L\'EXPERT]</span>')
    .replace(/à l'examen visuel/gi, '<span class="data-badge db-ia">à l\'examen visuel</span>')
    .replace(/\(source\s+([^)]+)\)/gi, '<span class="data-badge db-source">$1</span>')
    .replace(/\(INSEE[^)]*\)/gi, '<span class="data-badge db-source">INSEE</span>')
    .replace(/\(DVF[^)]*\)/gi, '<span class="data-badge db-source">DVF</span>');
}

// ── TOGGLE MENTIONS IA ────────────────────────────────────────────────────────
let aiVisible = true;
function toggleAI() {
  aiVisible = !aiVisible;
  const btn = document.getElementById('toggle-ai-btn');
  const content = document.getElementById('rapport-content');
  btn.classList.toggle('hidden-ai', !aiVisible);
  content.classList.toggle('hide-ai', !aiVisible);
  btn.title = aiVisible ? 'Masquer les mentions IA' : 'Afficher les mentions IA';
}

// ── EXPORT DOCX ───────────────────────────────────────────────────────────────
async function downloadDocx() {
  if (!state.reportMarkdown) {
    toast('Aucun rapport généré — lancez d\'abord la génération');
    return;
  }
  const btn = document.getElementById('btn-docx');
  if (btn) { btn.textContent = '⏳ Export…'; btn.disabled = true; }
  try {
    const res = await fetch('/api/export-docx', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        report: state.reportMarkdown,
        sections: state.sections,
        formData: state.formData,
        refDossier: state.formData.ref_dossier || 'PreRapport',
        photos64: state.photos64,
        logo: state.logo,
        photoResults: state.photoResults
      })
    });
    if (!res.ok) throw new Error('Export échoué');
    const blob = await res.blob();
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    const cd = res.headers.get('Content-Disposition') || '';
    const fn = cd.match(/filename="([^"]+)"/)?.[1] || 'PreRapport.docx';
    a.download = fn;
    a.click();
    URL.revokeObjectURL(url);
    toast('✓ ' + fn + ' téléchargé');
  } catch (e) {
    toast('Erreur export .docx : ' + e.message);
  } finally {
    if (btn) { btn.textContent = '⬇ Télécharger .docx'; btn.disabled = false; }
  }
}

// ── UTILITAIRES ───────────────────────────────────────────────────────────────
async function fetchJSON(url, body) {
  const res = await fetch(url, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(body)
  });
  if (!res.ok) {
    const err = await res.json().catch(() => ({}));
    throw new Error(err.error || `HTTP ${res.status}`);
  }
  return res.json();
}

function sleep(ms) {
  return new Promise(r => setTimeout(r, ms));
}

let toastTimer;
function toast(msg, duration = 3500) {
  const t = document.getElementById('toast');
  const m = document.getElementById('toast-msg');
  if (!t || !m) return;
  m.textContent = msg;
  t.classList.add('show');
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => t.classList.remove('show'), duration);
}

// ── FICHE DOSSIER ─────────────────────────────────────────────────────────────

function toggleCopropriete(select) {
  const bloc = document.getElementById('bloc-copropriete');
  const blocAppart = document.getElementById('bloc-appart-only');
  if (!bloc) return;
  const val = select.value;
  const showBloc = val === 'Appartement' || val === 'Immeuble';
  bloc.style.display = showBloc ? 'block' : 'none';
  if (blocAppart) blocAppart.style.display = val === 'Appartement' ? 'contents' : 'none';
}

function toggleAutreForme(select) {
  const autre = document.getElementById('forme_terrain_autre');
  if (!autre) return;
  autre.style.display = select.value === 'Autre' ? 'block' : 'none';
  if (select.value !== 'Autre') autre.value = '';
}

function handleLogoUpload(input) {
  if (!input.files?.[0]) return;
  fileToBase64(input.files[0]).then(b64 => {
    state.ficheLogo = b64;
    const status = document.getElementById('logo-status');
    if (status) { status.style.display = 'block'; status.textContent = '✓ ' + input.files[0].name; }
    toast('Logo chargé');
  });
}

function addWorkingDays(date, days) {
  let count = 0;
  const d = new Date(date);
  while (count < days) {
    d.setDate(d.getDate() + 1);
    const day = d.getDay();
    if (day !== 0 && day !== 6) count++;
  }
  return d;
}

function updateDateButoir() {
  const visiteVal = document.getElementById('f_visite')?.value;
  if (!visiteVal) return;
  const butoir = addWorkingDays(new Date(visiteVal), 15);
  const el = document.getElementById('f_date_butoir');
  if (el) el.value = butoir.toISOString().slice(0, 10);
}

function amountToWordsFr(n) {
  const ones = ['', 'UN', 'DEUX', 'TROIS', 'QUATRE', 'CINQ', 'SIX', 'SEPT', 'HUIT', 'NEUF',
    'DIX', 'ONZE', 'DOUZE', 'TREIZE', 'QUATORZE', 'QUINZE', 'SEIZE', 'DIX-SEPT', 'DIX-HUIT', 'DIX-NEUF'];

  function tens(x) {
    if (x < 20) return ones[x];
    const t = Math.floor(x / 10), u = x % 10;
    if (t === 2) return u ? 'VINGT-' + ones[u] : 'VINGT';
    if (t === 3) return u ? 'TRENTE-' + ones[u] : 'TRENTE';
    if (t === 4) return u ? 'QUARANTE-' + ones[u] : 'QUARANTE';
    if (t === 5) return u ? 'CINQUANTE-' + ones[u] : 'CINQUANTE';
    if (t === 6) return u ? 'SOIXANTE-' + ones[u] : 'SOIXANTE';
    if (t === 7) return 'SOIXANTE-' + ones[10 + u];
    if (t === 8) return u ? 'QUATRE-VINGT-' + ones[u] : 'QUATRE-VINGTS';
    if (t === 9) return 'QUATRE-VINGT-' + ones[10 + u];
    return '';
  }

  function hundreds(x) {
    if (x < 100) return tens(x);
    const h = Math.floor(x / 100), r = x % 100;
    const pre = h === 1 ? 'CENT' : ones[h] + ' CENT';
    return r ? pre + ' ' + tens(r) : (h > 1 ? pre + 'S' : pre);
  }

  if (!n) return 'ZÉRO';
  if (n < 1000) return hundreds(n);
  const k = Math.floor(n / 1000), r = n % 1000;
  const pre = k === 1 ? 'MILLE' : hundreds(k) + ' MILLE';
  return r ? pre + ' ' + hundreds(r) : pre;
}

function updateHonorairesLettres() {
  const val = parseFloat(document.getElementById('f_honoraires')?.value || '0');
  const el = document.getElementById('f_honoraires_lettres');
  if (el) el.value = val ? amountToWordsFr(Math.round(val)) : '';
}

async function generateFiche() {
  const get = id => (document.getElementById(id)?.value || '').trim();
  const suiviPar = document.querySelector('input[name="f_suivi_par"]:checked')?.value || 'RV';

  const data = {
    num_dossier: get('f_num_dossier') || 'EXP-2026-XXX',
    suivi_par: suiviPar,
    ordonnateur: get('f_ordonnateur'),
    email_ord: get('f_email'),
    objet: get('f_objet') || 'VALEUR VENALE',
    lieu: get('f_lieu'),
    description_bien: get('f_description_bien'),
    date_visite: get('f_visite'),
    date_butoir: get('f_date_butoir'),
    logo: state.ficheLogo || state.logo || null
  };

  const btn = document.getElementById('btn-fiche');
  if (btn) { btn.textContent = '⏳ Génération…'; btn.disabled = true; }

  try {
    const res = await fetch('/api/generate-fiche', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(data)
    });
    if (!res.ok) throw new Error('Erreur serveur');
    const blob = await res.blob();
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    const cd = res.headers.get('Content-Disposition') || '';
    const fn = cd.match(/filename="([^"]+)"/)?.[1] || 'Fiche.docx';
    a.download = fn;
    a.click();
    URL.revokeObjectURL(url);
    toast('✓ ' + fn + ' téléchargé');
  } catch (e) {
    toast('Erreur fiche : ' + e.message);
  } finally {
    if (btn) { btn.textContent = 'Générer la fiche dossier'; btn.disabled = false; }
  }
}

async function generateDevis() {
  const get = id => (document.getElementById(id)?.value || '').trim();
  const suiviPar = document.querySelector('input[name="f_suivi_par"]:checked')?.value || 'RV';

  const data = {
    num_dossier: get('f_num_dossier') || 'EXP-2026-XXX',
    suivi_par: suiviPar,
    ordonnateur: get('f_ordonnateur'),
    email_ord: get('f_email'),
    objet: get('f_objet') || 'VALEUR VENALE',
    lieu: get('f_lieu'),
    description_bien: get('f_description_bien'),
    honoraires_ttc: get('f_honoraires') || '0',
    date_visite: get('f_visite'),
    date_butoir: get('f_date_butoir'),
    logo: state.ficheLogo || state.logo || null
  };

  const btn = document.getElementById('btn-devis');
  if (btn) { btn.textContent = '⏳ Génération…'; btn.disabled = true; }

  try {
    const res = await fetch('/api/generate-devis', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(data)
    });
    if (!res.ok) throw new Error('Erreur serveur');
    const blob = await res.blob();
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    const cd = res.headers.get('Content-Disposition') || '';
    const fn = cd.match(/filename="([^"]+)"/)?.[1] || 'Devis.docx';
    a.download = fn;
    a.click();
    URL.revokeObjectURL(url);
    toast('✓ ' + fn + ' téléchargé');
  } catch (e) {
    toast('Erreur devis : ' + e.message);
  } finally {
    if (btn) { btn.textContent = '✦ Générer le devis .docx'; btn.disabled = false; }
  }
}

// ── INIT ──────────────────────────────────────────────────────────────────────
document.addEventListener('DOMContentLoaded', () => {
  showPage('fiche');
  goStep(0);
  initDesordres();
  initSurfaces();

  // Date par défaut = aujourd'hui pour formulaire et fiche
  const today = new Date().toISOString().slice(0, 10);
  const dv = document.getElementById('date_visite');
  if (dv) dv.value = today;
  const fv = document.getElementById('f_visite');
  if (fv) { fv.value = today; updateDateButoir(); }
});
