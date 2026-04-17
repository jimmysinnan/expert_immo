# ExpertIA — Améliorations rapport DOCX — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Corriger 5 problèmes dans le pipeline de génération : extraction de logo depuis le DOCX de référence, simplification du contexte géographique, exploitation effective des données cadastre/PLU, intégration des photos dans les sections DOCX appropriées, et amélioration du style extrait depuis l'exemple.

**Architecture:**
- `server.js` : toutes les modifications backend (prompts, routes, génération DOCX)
- `public/app.js` : stocker les photos en base64 pendant l'analyse et les transmettre à l'export

**Tech Stack:** Node.js + Express, `docx` v8 (ImageRun), `jszip` (extraction logo DOCX), `mammoth`, Claude claude-sonnet-4-6

---

## Fichiers modifiés

| Fichier | Modifications |
|---|---|
| `server.js` | Imports JSZip + ImageRun, logo extraction, prompts géo/PLU/cadastre, intégration photos DOCX |
| `public/app.js` | Stockage base64 photos pendant analyse, envoi logo+photos à export-docx |
| `package.json` | Ajout jszip comme dépendance directe |

---

## Task 1 : Installer jszip + ajouter ImageRun aux imports

**Files:**
- Modify: `package.json`
- Modify: `server.js:1-11`

- [ ] **Step 1 : Installer jszip**

```bash
cd c:\Users\jimmy\Projets\expert-immo-app
npm install jszip
```

Vérifier que `package.json` contient `"jszip": "^3.x.x"`.

- [ ] **Step 2 : Ajouter JSZip et ImageRun aux imports dans server.js**

Remplacer les lignes 1-11 de `server.js` :

```javascript
require('dotenv').config();
const express = require('express');
const multer  = require('multer');
const Anthropic = require('@anthropic-ai/sdk');
const mammoth  = require('mammoth');
const JSZip   = require('jszip');
const {
  Document, Packer, Paragraph, TextRun, HeadingLevel,
  Table, TableRow, TableCell, WidthType, AlignmentType,
  BorderStyle, Header, Footer, ShadingType, PageNumber,
  NumberFormat, convertInchesToTwip, ImageRun
} = require('docx');
```

- [ ] **Step 3 : Commit**

```bash
git add package.json package-lock.json server.js
git commit -m "chore: add jszip dependency + ImageRun import"
```

---

## Task 2 : Simplifier le prompt géographique (chapitre 1)

**Files:**
- Modify: `server.js` — fonction `buildChapter1Prompt()` (lignes ~75-95)

**Problème actuel :** 4 sous-sections demandées dont analyse de marché DVF/prix m² — trop détaillé et hors périmètre du chapitre géographique.

**Objectif :** 2-3 paragraphes narratifs sobres, style JALTA, sans données de marché.

- [ ] **Step 1 : Remplacer `buildChapter1Prompt()`**

Remplacer la fonction `buildChapter1Prompt` dans `server.js` par :

```javascript
function buildChapter1Prompt(adresse) {
  return `Tu es un expert immobilier certifié. Rédige la section "SITUATION GÉOGRAPHIQUE" d'un rapport d'expertise JALTA pour le bien situé à :

${adresse}

Rédige en 2 à 3 paragraphes dans le style sobre et factuel du Cabinet JALTA :

**Paragraphe 1 — La commune**
Situer la commune : département, caractère (résidentiel, touristique, économique), dynamisme local — 3 à 4 lignes. Exemple d'entrée : "La commune de... est située dans le département de... Elle se caractérise par..."

**Paragraphe 2 — Situation du bien dans la commune**
Décrire l'environnement immédiat du bien : quartier ou secteur, tissu bâti (pavillonnaire, mixte...), standing, desserte — 3 à 4 lignes. Exemple d'entrée : "Le bien objet de la présente expertise est situé dans le secteur..."

**Paragraphe 3 (optionnel) — Accessibilité**
Axes routiers principaux, transports — 2 lignes maximum. Uniquement si pertinent.

RÈGLES ABSOLUES :
- Maximum 200 mots au total
- Aucune donnée de prix, aucune statistique de marché, aucune référence DVF
- Style impersonnel, troisième personne, indicatif présent
- Si une donnée est inconnue, ne pas l'inventer — l'omettre
- Retourner uniquement le texte, sans titres ni marqueurs markdown`;
}
```

- [ ] **Step 2 : Redémarrer le serveur et tester**

```bash
npm run dev
```

Tester manuellement : lancer une génération avec une adresse en Martinique, vérifier que la section géographique est plus courte et sans données de marché.

- [ ] **Step 3 : Commit**

```bash
git add server.js
git commit -m "fix: simplifier prompt géographique - supprimer données marché DVF"
```

---

## Task 3 : Exploiter références cadastrales et PLU dans les prompts

**Files:**
- Modify: `server.js` — fonction `buildMainPrompt()` (lignes ~133-214)

**Problème actuel :** `refs_cadastrales` et `zonage_plu` sont dans le payload mais les instructions de génération des sections `situation_juridique` et `situation_urbanistique` n'imposent pas explicitement leur intégration.

- [ ] **Step 1 : Modifier les instructions dans le JSON schema de `buildMainPrompt()`**

Dans la fonction `buildMainPrompt`, remplacer les lignes des clés `situation_urbanistique` et `situation_juridique` dans la partie `GÉNÈRE UN JSON` :

**Avant (situation_urbanistique) :**
```
"situation_urbanistique": "Texte SITUATION URBANISTIQUE : zonage PLU, règles d'urbanisme, certificat d'urbanisme, servitudes d'utilité publique, assainissement — style JALTA 3 à 5 phrases.",
```

**Après :**
```javascript
`"situation_urbanistique": "Texte SITUATION URBANISTIQUE — INTÉGRER OBLIGATOIREMENT le zonage PLU '${formData.zonage_plu || '[zonage non renseigné]'}' dans la première phrase. Exemple : 'Au regard du Plan Local d'Urbanisme, le bien est classé en zone ${formData.zonage_plu || '[à compléter]'}...'. Décrire les règles d'urbanisme applicables à cette zone, les possibilités de construction, l'assainissement (${formData.assainissement || '[à compléter]'}), les servitudes connues — 3 à 5 phrases style JALTA.",`
```

**Avant (situation_juridique) :**
```
"situation_juridique": "Texte SITUATION JURIDIQUE : régime juridique du bien (pleine propriété, copropriété...), références cadastrales, superficie cadastrale, mentions hypothécaires si connues — 3 à 5 phrases.",
```

**Après :**
```javascript
`"situation_juridique": "Texte SITUATION JURIDIQUE — INTÉGRER OBLIGATOIREMENT la référence cadastrale '${formData.refs_cadastrales || '[référence à compléter]'}' dans le texte. Exemple : 'Le bien est cadastré sous la référence ${formData.refs_cadastrales || '[à compléter]'}...'. Mentionner le régime juridique (${formData.regime_juridique || '[à compléter]'}), la superficie du terrain (${formData.superficie_terrain || '[à compléter]'} m²), les mentions hypothécaires si connues — 3 à 5 phrases style JALTA.",`
```

Note : toute la zone `GÉNÈRE UN JSON` utilise déjà un template literal, donc ces interpolations `${formData.xxx}` fonctionnent directement.

- [ ] **Step 2 : Vérifier que `buildMainPrompt` est bien un template literal**

S'assurer que la fonction débute par `return \`` (backtick) et se termine par `\``. Si ce n'est pas le cas, c'est une chaîne ordinaire et les interpolations ne fonctionneront pas — convertir en template literal.

Vérifier en lisant la ligne ~141 de `server.js`.

- [ ] **Step 3 : Commit**

```bash
git add server.js
git commit -m "fix: intégration obligatoire cadastre et PLU dans les sections juridique/urbanistique"
```

---

## Task 4 : Améliorer l'extraction de style — exemples de contenu par section

**Files:**
- Modify: `server.js` — `buildStylePrompt()` (lignes ~97-113) et route `/api/extract-style`

**Problème actuel :** Le style extrait est uniquement méta (ton, formules-clés) mais pas des exemples de contenu réel par section. L'IA génère donc dans un style générique plutôt que de reproduire les tournures exactes du rapport de référence.

- [ ] **Step 1 : Étendre `buildStylePrompt()` pour extraire des exemples de contenu**

Remplacer la fonction `buildStylePrompt` :

```javascript
function buildStylePrompt(docText) {
  return `Analyse ce rapport d'expertise immobilière de référence et extrais en JSON :
{
  "ton_general": "description du niveau de langue (ex: professionnel, technique, sobre)",
  "formules_introduction": ["liste des formules récurrentes d'introduction de section"],
  "formules_conclusion": ["formules de conclusion"],
  "formules_conditionnelles": ["formules utilisées pour observations visuelles : 'à l'examen visuel', 'semble présenter', etc."],
  "vocabulaire_technique": ["termes techniques caractéristiques du rapport"],
  "exemple_situation_geographique": "Copier mot-pour-mot 3 à 5 phrases caractéristiques de la section situation géographique/localisation du rapport. Si absente : null.",
  "exemple_situation_urbanistique": "Copier mot-pour-mot 2 à 4 phrases de la section urbanistique/PLU. Si absente : null.",
  "exemple_situation_juridique": "Copier mot-pour-mot 2 à 4 phrases de la section juridique/cadastrale. Si absente : null.",
  "exemple_description_terrain": "Copier mot-pour-mot 4 à 6 phrases caractéristiques du chapitre terrain. Si absente : null.",
  "exemple_description_bati": "Copier mot-pour-mot 4 à 6 phrases caractéristiques du chapitre bâti/construction. Si absente : null.",
  "style_desordres": "Copier mot-pour-mot 2 à 3 phrases sur la manière dont les désordres sont présentés. Si absente : null."
}
Retourner UNIQUEMENT le JSON valide, sans texte avant ni après.

RAPPORT DE RÉFÉRENCE :
${docText}`;
}
```

- [ ] **Step 2 : Injecter les exemples dans `buildMainPrompt()`**

Dans `buildMainPrompt`, juste après la ligne `const { formData, chapter1, style, photos, desordres, surfaces } = data;`, ajouter :

```javascript
  // Parser le style extrait (JSON string → objet)
  let styleObj = null;
  if (style) {
    try { styleObj = JSON.parse(style); } catch(e) { styleObj = null; }
  }

  // Construire le bloc d'exemples de contenu
  const exemples = styleObj ? `
=== EXEMPLES DE CONTENU À REPRODUIRE (extraits du rapport de référence) ===
${styleObj.exemple_situation_geographique ? `SITUATION GÉOGRAPHIQUE — écrire dans ce style :\n"${styleObj.exemple_situation_geographique}"` : ''}
${styleObj.exemple_situation_urbanistique ? `SITUATION URBANISTIQUE — écrire dans ce style :\n"${styleObj.exemple_situation_urbanistique}"` : ''}
${styleObj.exemple_situation_juridique ? `SITUATION JURIDIQUE — écrire dans ce style :\n"${styleObj.exemple_situation_juridique}"` : ''}
${styleObj.exemple_description_terrain ? `TERRAIN — écrire dans ce style :\n"${styleObj.exemple_description_terrain}"` : ''}
${styleObj.exemple_description_bati ? `CONSTRUCTION — écrire dans ce style :\n"${styleObj.exemple_description_bati}"` : ''}
${styleObj.style_desordres ? `DÉSORDRES — écrire dans ce style :\n"${styleObj.style_desordres}"` : ''}
` : '';
```

- [ ] **Step 3 : Ajouter `${exemples}` dans le template de `buildMainPrompt`**

Dans le template return de `buildMainPrompt`, ajouter `${exemples}` juste après la section `=== STYLE DE L'EXPERT ===` :

```javascript
return `=== STYLE DE L'EXPERT (reproduire exactement) ===
${style || 'Style JALTA professionnel — français technique immobilier Antilles — TEGOVA 6e édition.'}

${exemples}
=== DONNÉES DU DOSSIER ===
...
```

- [ ] **Step 4 : Commit**

```bash
git add server.js
git commit -m "feat: enrichir extraction de style avec exemples de contenu par section"
```

---

## Task 5 : Extraction du logo depuis le DOCX de référence

**Files:**
- Modify: `server.js` — route `POST /api/extract-style`

**Principe :** Un fichier DOCX est un ZIP. Les images sont dans `word/media/`. On extrait la première image trouvée (logo du cabinet en général en haut de page).

- [ ] **Step 1 : Ajouter la fonction `extractLogoFromDocx`**

Ajouter cette fonction dans `server.js` juste avant la route `/api/extract-style` :

```javascript
async function extractLogoFromDocx(buffer) {
  try {
    const zip = await JSZip.loadAsync(buffer);
    // Les images sont dans word/media/
    const mediaFiles = Object.keys(zip.files).filter(name =>
      name.startsWith('word/media/') && /\.(png|jpg|jpeg|gif|webp)$/i.test(name)
    );
    if (!mediaFiles.length) return null;
    // Prendre la première image (généralement le logo en en-tête)
    const firstMedia = mediaFiles[0];
    const imgBuffer = await zip.files[firstMedia].async('nodebuffer');
    const ext = firstMedia.split('.').pop().toLowerCase();
    const mimeMap = { png: 'image/png', jpg: 'image/jpeg', jpeg: 'image/jpeg', gif: 'image/gif', webp: 'image/webp' };
    return {
      data: imgBuffer.toString('base64'),
      mimeType: mimeMap[ext] || 'image/png',
      filename: firstMedia
    };
  } catch (e) {
    console.warn('[extractLogo] Échec extraction logo:', e.message);
    return null;
  }
}
```

- [ ] **Step 2 : Modifier la route `/api/extract-style` pour retourner aussi le logo**

Dans la route `POST /api/extract-style`, après avoir extrait le texte avec mammoth, ajouter l'extraction du logo :

```javascript
app.post('/api/extract-style', upload.single('document'), async (req, res) => {
  try {
    if (!req.file) return res.json({ style: null, logo: null });

    let docText = '';
    let logo = null;

    if (req.file.originalname.endsWith('.docx')) {
      const result = await mammoth.extractRawText({ buffer: req.file.buffer });
      docText = result.value.slice(0, 8000);
      // Extraire le logo en parallèle de l'analyse de style
      logo = await extractLogoFromDocx(req.file.buffer);
    } else if (req.file.originalname.endsWith('.pdf') || req.file.mimetype === 'application/pdf') {
      const response = await client.messages.create({
        model: process.env.CLAUDE_MODEL || 'claude-sonnet-4-6',
        max_tokens: 1500,
        temperature: 0,
        messages: [{
          role: 'user',
          content: [
            { type: 'document', source: { type: 'base64', media_type: 'application/pdf', data: req.file.buffer.toString('base64') } },
            { type: 'text', text: buildStylePrompt('') }
          ]
        }]
      });
      return res.json({ style: response.content[0].text, logo: null });
    }

    const response = await client.messages.create({
      model: process.env.CLAUDE_MODEL || 'claude-sonnet-4-6',
      max_tokens: 1500,
      temperature: 0,
      messages: [{ role: 'user', content: buildStylePrompt(docText) }]
    });

    res.json({ style: response.content[0].text, logo });
  } catch (err) {
    console.error('[extract-style]', err.message);
    res.status(500).json({ error: err.message });
  }
});
```

- [ ] **Step 3 : Commit**

```bash
git add server.js
git commit -m "feat: extraction logo depuis DOCX de référence via JSZip"
```

---

## Task 6 : Frontend — stocker photos en base64 + logo en state

**Files:**
- Modify: `public/app.js` — `state` object + étape 2 (extract-style) + étape 3 (analyze-photos)

**Problème actuel :** Les photos sont analysées (texte JSON extrait) mais leurs buffers raw ne sont jamais stockés pour l'export DOCX. Le logo extrait n'est pas stocké non plus.

- [ ] **Step 1 : Étendre `state` pour stocker les photos base64 et le logo**

Dans `app.js`, modifier l'objet `state` (ligne ~42) pour ajouter :

```javascript
const state = {
  currentStep: 0,
  refDoc: null,
  photos: {
    terrain: null,
    ext: null,
    int: null,
    desordres: []
  },
  photos64: {          // ← NOUVEAU : photos en base64 pour export DOCX
    terrain: [],
    ext: [],
    int: [],
    desordres: []
  },
  logo: null,          // ← NOUVEAU : { data: base64, mimeType: 'image/png' }
  chapter1: '',
  style: null,
  photoResults: {},
  reportMarkdown: '',
  sections: null,
  formData: {}
};
```

- [ ] **Step 2 : Stocker le logo lors de l'extraction de style (étape 2)**

Dans `startGeneration()`, modifier le bloc étape 2 (autour de la ligne ~340) pour récupérer le logo :

```javascript
    // ── ÉTAPE 2 : Extraction style
    setStep(2, 'active');
    if (state.refDoc) {
      updateDetail(2, `Analyse de ${state.refDoc.name}`);
      try {
        const fd2 = new FormData();
        fd2.append('document', state.refDoc);
        const r2 = await fetch('/api/extract-style', { method: 'POST', body: fd2 });
        const j2 = await r2.json();
        state.style = j2.style || null;
        state.logo = j2.logo || null;  // ← NOUVEAU
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
```

- [ ] **Step 3 : Convertir les photos en base64 lors de l'analyse (étape 3)**

Dans `startGeneration()`, modifier le bloc étape 3 pour stocker les photos base64 en plus de l'analyse. Ajouter une fonction helper `fileToBase64` au début de `app.js` (avant le bloc state) :

```javascript
// ── HELPER BASE64 ─────────────────────────────────────────────────────────────
function fileToBase64(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => {
      // reader.result = "data:image/jpeg;base64,/9j/..."
      const base64 = reader.result.split(',')[1];
      resolve({ data: base64, mimeType: file.type || 'image/jpeg', name: file.name });
    };
    reader.onerror = reject;
    reader.readAsDataURL(file);
  });
}
```

Puis dans le bloc étape 3, juste APRÈS que `state.photoResults = await r3.json()` soit résolu, ajouter la conversion base64 :

```javascript
        state.photoResults = await r3.json();

        // ← NOUVEAU : convertir les photos en base64 pour l'export DOCX
        const p64 = state.photos64;
        if (state.photos.terrain) {
          p64.terrain = await Promise.all(Array.from(state.photos.terrain).map(fileToBase64));
        }
        if (state.photos.ext) {
          p64.ext = await Promise.all(Array.from(state.photos.ext).map(fileToBase64));
        }
        if (state.photos.int) {
          p64.int = await Promise.all(Array.from(state.photos.int).map(fileToBase64));
        }
        if (state.photos.desordres.length) {
          p64.desordres = await Promise.all(state.photos.desordres.map(fileToBase64));
        }

        const nbCats = Object.values(state.photoResults).filter(Boolean).length;
        updateDetail(3, `${nbCats} lot(s) de photos analysés ✓`);
```

- [ ] **Step 4 : Modifier `downloadDocx()` pour envoyer photos + logo**

Dans la fonction `downloadDocx()` (autour de la ligne ~584), modifier le body de la requête :

```javascript
    const res = await fetch('/api/export-docx', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        report: state.reportMarkdown,
        sections: state.sections,
        formData: state.formData,
        refDossier: state.formData.ref_dossier || 'PreRapport',
        photos64: state.photos64,    // ← NOUVEAU
        logo: state.logo             // ← NOUVEAU
      })
    });
```

- [ ] **Step 5 : Commit**

```bash
git add public/app.js
git commit -m "feat: stocker photos base64 + logo en state pour export DOCX"
```

---

## Task 7 : Backend — intégrer photos et logo dans `generateJaltaDocx`

**Files:**
- Modify: `server.js` — route `/api/export-docx`, fonction `generateJaltaDocx()`, fonctions de build des sections

**Principe :**
- Photos terrain → dans `buildDescriptionSection()` après le texte terrain
- Photos extérieur → dans `buildDescriptionSection()` après le texte bâti (extérieur)
- Photos intérieur → dans `buildDescriptionSection()` après le texte bâti (intérieur) — ou combinées avec ext
- Photos désordres → dans `buildDescriptionSection()` après l'état des lieux
- Section "Photographies" : ne plus générer de placeholders (les vraies photos sont dans les sections)
- Logo → dans l'en-tête DOCX et sur la page de garde

- [ ] **Step 1 : Ajouter la fonction helper `buildImageRun`**

Ajouter dans `server.js` juste avant `buildCoverPage` :

```javascript
// Crée un ImageRun docx depuis une photo base64
// opts.width et opts.height en pixels (PointsToTwip calcule automatiquement)
function buildImageRun(photo64, opts = {}) {
  if (!photo64 || !photo64.data) return null;
  try {
    const buf = Buffer.from(photo64.data, 'base64');
    return new ImageRun({
      data: buf,
      transformation: {
        width: opts.width || 450,
        height: opts.height || 300,
      },
      type: (photo64.mimeType || 'image/jpeg').replace('image/', ''),
    });
  } catch (e) {
    console.warn('[buildImageRun] Erreur:', e.message);
    return null;
  }
}

// Génère un tableau de 2 photos côte à côte (ou 1 pleine largeur si impaire)
function buildPhotoParagraphs(photos64Array, caption = '') {
  if (!photos64Array || !photos64Array.length) return [];
  const items = [];

  for (let i = 0; i < photos64Array.length; i += 2) {
    const left = buildImageRun(photos64Array[i], { width: 210, height: 155 });
    const right = i + 1 < photos64Array.length
      ? buildImageRun(photos64Array[i + 1], { width: 210, height: 155 })
      : null;

    if (!left) continue;

    if (right) {
      // 2 photos côte à côte
      items.push(new Table({
        width: { size: 100, type: WidthType.PERCENTAGE },
        borders: noBorders(),
        rows: [new TableRow({
          children: [
            new TableCell({
              borders: noBorders(),
              width: { size: 48, type: WidthType.PERCENTAGE },
              margins: { right: 100 },
              children: [new Paragraph({ children: [left], alignment: AlignmentType.CENTER })]
            }),
            new TableCell({
              borders: noBorders(),
              width: { size: 48, type: WidthType.PERCENTAGE },
              margins: { left: 100 },
              children: [new Paragraph({ children: [right], alignment: AlignmentType.CENTER })]
            }),
          ]
        })]
      }));
    } else {
      // 1 photo pleine largeur
      items.push(new Paragraph({
        children: [buildImageRun(photos64Array[i], { width: 450, height: 300 })],
        alignment: AlignmentType.CENTER
      }));
    }
    items.push(spacer(80));
  }

  if (caption) {
    items.push(new Paragraph({
      children: [new TextRun({ text: caption, size: 17, italics: true, color: C.DARK, font: 'Times New Roman' })],
      alignment: AlignmentType.CENTER,
      spacing: { before: 40, after: 120 }
    }));
  }

  return items;
}
```

- [ ] **Step 2 : Modifier `generateJaltaDocx()` pour recevoir photos64 et logo**

Modifier la signature et le début de `generateJaltaDocx` :

```javascript
async function generateJaltaDocx(sections, formData, photos64 = {}, logo = null) {
  const fd = formData || {};
  const p64 = photos64 || {};
```

- [ ] **Step 3 : Modifier `buildCoverPage()` pour inclure le logo**

Modifier la signature de `buildCoverPage(formData, logo)` et ajouter le logo en haut de la page de garde, juste avant le bandeau titre :

```javascript
function buildCoverPage(formData, logo) {
  const fd = formData || {};
  const items = [
    pageBreak(),
  ];

  // Logo du cabinet si disponible
  if (logo && logo.data) {
    const logoRun = buildImageRun(logo, { width: 180, height: 70 });
    if (logoRun) {
      items.push(new Paragraph({
        children: [logoRun],
        alignment: AlignmentType.LEFT,
        spacing: { before: 100, after: 200 }
      }));
    }
  }

  items.push(
    // Bandeau titre principal
    new Table({
      width: { size: 100, type: WidthType.PERCENTAGE },
      borders: noBorders(),
      rows: [new TableRow({
        children: [shadedCell(C.NAVY, [
          // ... garder le contenu existant du bandeau ...
```

Note : conserver le reste de `buildCoverPage` identique (bandeau titre, tableau dossier, clause confidentialité).

- [ ] **Step 4 : Modifier `buildDescriptionSection()` pour intégrer les photos**

Modifier la signature : `function buildDescriptionSection(sections, formData, p64 = {})`.

Dans le return array de `buildDescriptionSection`, remplacer les `imagePlaceholder` par les vraies photos ou un placeholder si vide :

**Pour la section terrain** (après `...splitParagraphs(sections.description_terrain...)`), remplacer :
```javascript
    imagePlaceholder('[à rajouter par l\'expert] — Plan de masse / Plan du terrain'),
```
par :
```javascript
    ...(p64.terrain && p64.terrain.length
      ? buildPhotoParagraphs(p64.terrain, 'Vues du terrain — photos prises lors de la visite')
      : [imagePlaceholder('[à rajouter par l\'expert] — Photos du terrain')]),
```

**Pour la section bâti** (après `...splitParagraphs(sections.description_bati...)`), ajouter les photos extérieures ET intérieures :
```javascript
    ...(p64.ext && p64.ext.length
      ? buildPhotoParagraphs(p64.ext, 'Vues extérieures')
      : [imagePlaceholder('[à rajouter par l\'expert] — Photos extérieures')]),
    spacer(100),
    ...(p64.int && p64.int.length
      ? buildPhotoParagraphs(p64.int, 'Vues intérieures')
      : [imagePlaceholder('[à rajouter par l\'expert] — Photos intérieures')]),
```

**Pour la section état des lieux** (après `...splitParagraphs(sections.desordres_texte...)`), ajouter :
```javascript
    ...(p64.desordres && p64.desordres.length
      ? [spacer(80), ...buildPhotoParagraphs(p64.desordres, 'Désordres constatés lors de la visite')]
      : []),
```

- [ ] **Step 5 : Modifier l'appel à `buildDescriptionSection` dans `generateJaltaDocx`**

```javascript
    // IV — Description (terrain + bâti + surfaces + désordres)
    ...buildDescriptionSection(sections, fd, p64),
```

- [ ] **Step 6 : Modifier `buildPhotosSection()` — photos non catégorisées uniquement**

Simplifier `buildPhotosSection` pour ne plus afficher de placeholders (les photos sont désormais dans les sections) :

```javascript
function buildPhotosSection() {
  return [
    pageBreak(),
    navyBanner('PHOTOGRAPHIES COMPLÉMENTAIRES'),
    spacer(120),
    new Paragraph({
      children: [new TextRun({
        text: 'Les photographies illustrant chaque section descriptive sont intégrées directement dans les chapitres correspondants du présent rapport.',
        size: 19, font: 'Times New Roman', italics: true, color: C.DARK
      })],
      alignment: AlignmentType.JUSTIFIED,
      spacing: { before: 60, after: 120 }
    }),
    imagePlaceholder('[à rajouter par l\'expert] — Photographies complémentaires ou de contexte'),
  ];
}
```

- [ ] **Step 7 : Intégrer le logo dans l'en-tête DOCX**

Dans `generateJaltaDocx`, modifier la section `headers.default` pour ajouter le logo à gauche si disponible :

```javascript
      headers: {
        default: new Header({
          children: [
            new Table({
              width: { size: 100, type: WidthType.PERCENTAGE },
              borders: noBorders(),
              rows: [new TableRow({
                children: [
                  new TableCell({
                    borders: noBorders(),
                    width: { size: logo ? 20 : 0, type: WidthType.PERCENTAGE },
                    children: logo && logo.data ? [new Paragraph({
                      children: [buildImageRun(logo, { width: 100, height: 40 })],
                      alignment: AlignmentType.LEFT
                    })] : [new Paragraph({ children: [] })]
                  }),
                  new TableCell({
                    borders: noBorders(),
                    width: { size: logo ? 50 : 70, type: WidthType.PERCENTAGE },
                    children: [new Paragraph({
                      children: [new TextRun({ text: 'RAPPORT D\'EXPERTISE IMMOBILIÈRE — CONFIDENTIEL', size: 16, color: C.NAVY, font: 'Times New Roman' })],
                      border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: C.NAVY } }
                    })]
                  }),
                  new TableCell({
                    borders: noBorders(),
                    width: { size: 30, type: WidthType.PERCENTAGE },
                    children: [new Paragraph({
                      children: [new TextRun({ text: fd.ref_dossier || '', size: 16, color: C.NAVY, font: 'Times New Roman', bold: true })],
                      alignment: AlignmentType.RIGHT,
                      border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: C.NAVY } }
                    })]
                  })
                ]
              })]
            })
          ]
        })
      },
```

- [ ] **Step 8 : Modifier la route `/api/export-docx` pour passer photos64 et logo**

```javascript
app.post('/api/export-docx', async (req, res) => {
  try {
    const { report, sections, formData, refDossier, photos64, logo } = req.body;
    let buffer;

    if (sections && formData) {
      buffer = await generateJaltaDocx(sections, formData, photos64 || {}, logo || null);
    } else {
      buffer = await generateDocx(report || '');
    }

    const filename = `${refDossier || 'PreRapport'}_${new Date().toISOString().slice(0,10)}.docx`;
    res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    res.send(buffer);
  } catch (err) {
    console.error('[export-docx]', err.message);
    res.status(500).json({ error: err.message });
  }
});
```

- [ ] **Step 9 : Mettre à jour l'appel à `buildCoverPage` dans `generateJaltaDocx`**

```javascript
    // Page de couverture
    ...buildCoverPage(fd, logo),
```

- [ ] **Step 10 : Commit**

```bash
git add server.js
git commit -m "feat: intégrer photos réelles et logo dans le DOCX par section"
```

---

## Task 8 : Test end-to-end + vérifications

- [ ] **Step 1 : Lancer le serveur**

```bash
cd c:\Users\jimmy\Projets\expert-immo-app
npm run dev
```

Vérifier qu'aucune erreur ne s'affiche au démarrage.

- [ ] **Step 2 : Tester l'extraction de style + logo**

- Uploader un fichier `.docx` de référence
- Vérifier dans les logs serveur que l'extraction se passe sans erreur
- Vérifier côté frontend (DevTools console) que `state.logo` est non null après l'étape 2
- Vérifier que `state.style` est un JSON valide avec les nouvelles clés `exemple_*`

- [ ] **Step 3 : Tester la génération**

- Remplir un dossier complet avec : adresse Martinique, références cadastrales (ex: "AB 0042"), zonage PLU (ex: "UM")
- Lancer la génération
- Vérifier que dans le rapport HTML :
  - La section SITUATION GÉOGRAPHIQUE est ≤ 200 mots et sans données DVF/prix
  - La section SITUATION JURIDIQUE mentionne "AB 0042"
  - La section SITUATION URBANISTIQUE mentionne "UM"

- [ ] **Step 4 : Tester l'export DOCX avec photos**

- Uploader des photos dans les catégories terrain, extérieur, intérieur
- Lancer la génération puis l'export
- Ouvrir le DOCX et vérifier :
  - Le logo JALTA apparaît en haut de la page de garde et dans l'en-tête
  - Les photos terrain sont dans la section "LE TERRAIN D'ASSIETTE"
  - Les photos ext/int sont dans la section "LA CONSTRUCTION"
  - La section "PHOTOGRAPHIES COMPLÉMENTAIRES" ne contient plus les 6 placeholders

- [ ] **Step 5 : Commit final**

```bash
git add -A
git commit -m "test: vérifications end-to-end rapport avec photos et logo"
```

---

## Self-Review

### Couverture spec

| Remarque | Task couvrant |
|---|---|
| Rapport d'exemple comme template de contenu | Task 4 (extraction exemples par section) |
| Logos, en-tête, page de garde | Task 5 (extraction logo) + Task 7 (intégration DOCX) |
| Géographie trop détaillée | Task 2 (simplification prompt) |
| Référence cadastrale non exploitée | Task 3 (prompts explicites) |
| Zone PLU non exploitée | Task 3 (prompts explicites) |
| Images dans les bonnes sections | Task 6 (frontend base64) + Task 7 (backend ImageRun) |

### Points d'attention

- **Taille des requêtes** : si l'utilisateur uploade beaucoup de photos haute résolution, le body JSON de `/api/export-docx` peut devenir très lourd (plusieurs Mo). Si la limite express `50mb` est trop juste, augmenter à `100mb`.
- **Type image pour ImageRun** : le champ `type` dans `ImageRun` attend `'jpeg'` pas `'image/jpeg'`. La fonction `buildImageRun` fait bien `.replace('image/', '')` pour extraire le type court.
- **Logo premier image vs logo réel** : la première image d'un DOCX n'est pas toujours le logo — ça peut être une photo. Si le logo n'est pas trouvé, le fallback est gracieux (pas de logo, pas d'erreur).
