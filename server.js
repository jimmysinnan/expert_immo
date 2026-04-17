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

// ── Validation clé API au démarrage ──────────────────────────────────────────
const API_KEY = process.env.ANTHROPIC_API_KEY;
if (!API_KEY || API_KEY === 'sk-ant-api-xxx' || !API_KEY.startsWith('sk-ant-')) {
  console.error('\n╔══════════════════════════════════════════════════════╗');
  console.error('║  ERREUR : Clé API Anthropic manquante ou invalide    ║');
  console.error('║                                                      ║');
  console.error('║  1. Ouvrir le fichier .env à la racine du projet     ║');
  console.error('║  2. Remplacer sk-ant-api-xxx par votre vraie clé     ║');
  console.error('║     → console.anthropic.com → API Keys               ║');
  console.error('║  3. Relancer : npm start                             ║');
  console.error('╚══════════════════════════════════════════════════════╝\n');
  process.exit(1);
}

const app    = express();
const upload = multer({ storage: multer.memoryStorage(), limits: { fileSize: 25 * 1024 * 1024 } });
const client = new Anthropic({ apiKey: API_KEY });

app.use(express.json({ limit: '50mb' }));
app.use(express.static('public'));

// ─────────────────────────────────────────────────────────────────────────────
// COULEURS JALTA
// ─────────────────────────────────────────────────────────────────────────────
const C = {
  NAVY:    '1F3864',
  GREEN:   '92D050',
  NAVY_L:  'D9E2F3',
  GRAY:    'F2F2F2',
  GRAY_MED:'BFBFBF',
  WHITE:   'FFFFFF',
  BLACK:   '000000',
  DARK:    '404040',
  AMBER:   'C47A00',
};

// ─────────────────────────────────────────────────────────────────────────────
// PROMPTS
// ─────────────────────────────────────────────────────────────────────────────

const SYSTEM_PROMPT = `Tu es un expert immobilier certifié TEGOVA (6e édition) et Charte de l'Expertise Immobilière (5e édition), spécialisé dans la rédaction de rapports d'expertise conformes aux standards du Cabinet JALTA en Martinique.

Tu rédiges pour le compte d'un expert immobilier et dois reproduire EXACTEMENT son style, son ton et son vocabulaire — tels qu'ils apparaissent dans les rapports de référence de sa base de connaissance.

STYLE ET VOCABULAIRE JALTA :
- Formules d'entrée : "Au jour de notre visite...", "Nous avons notamment relevé...", "Il s'agit d'un bâtiment en dur..."
- Observations visuelles : toujours au conditionnel ou avec "à l'examen visuel", "semble", "paraît"
- Style : impersonnel, troisième personne, indicatif présent pour les faits constatés
- Vocabulaire : "le bien objet de la présente expertise", "au sens de la Charte", "TEGOVA 6e édition", "valeur vénale", "critères de pondération"
- Mentions manquantes : [à rajouter par l'expert] (jamais inventer)
- Norme de surface : Loi Boutin pour la surface habitable, avec coefficient de pondération JALTA

Ta mission : générer les sections narratives du rapport JALTA en JSON structuré, à partir des données saisies et des analyses de photos.

RÈGLES ABSOLUES :
1. Reproduire EXACTEMENT le style JALTA — jamais un style générique
2. Ne JAMAIS générer de valeurs vénales, prix, estimations ou calculs de valeur
3. Si une information manque → [à rajouter par l'expert], jamais inventer
4. Toute observation issue des photos → conditionnel ou "à l'examen visuel"
5. Langue : français professionnel, vocabulaire technique immobilier Antilles
6. Retourner UNIQUEMENT le JSON valide demandé, sans aucun texte avant ni après`;

function buildChapter1Prompt(adresse) {
  return `Tu es un expert immobilier certifié. Rédige la section "SITUATION GÉOGRAPHIQUE" d'un rapport d'expertise JALTA pour le bien situé à :

${adresse}

Rédige en 2 à 3 paragraphes dans le style sobre et factuel du Cabinet JALTA :

**Paragraphe 1 — La commune**
Situer la commune : département, caractère général (résidentiel, touristique, économique), dynamisme local — 3 à 4 lignes. Entrée type : "La commune de... est située dans le département de... Elle se caractérise par..."

**Paragraphe 2 — Situation du bien dans la commune**
Décrire l'environnement immédiat du bien : quartier ou secteur, tissu bâti (pavillonnaire, mixte...), standing, desserte de proximité — 3 à 4 lignes. Entrée type : "Le bien objet de la présente expertise est situé dans le secteur..."

**Paragraphe 3 — Accessibilité (optionnel)**
Axes routiers principaux, transports — 2 lignes maximum. Uniquement si l'information est pertinente et vérifiable.

RÈGLES ABSOLUES :
- Maximum 200 mots au total
- Aucune donnée de prix, aucune statistique de marché, aucune référence DVF
- Style impersonnel, troisième personne, indicatif présent
- Si une donnée est inconnue, ne pas l'inventer — l'omettre
- Retourner uniquement le texte, sans titres ni marqueurs markdown`;
}

function buildStylePrompt(docText) {
  return `Analyse ce rapport d'expertise immobilière de référence et extrais en JSON :
{
  "ton_general": "description du niveau de langue (ex: professionnel, technique, sobre)",
  "formules_introduction": ["liste des formules récurrentes d'introduction de section"],
  "formules_conclusion": ["formules de conclusion"],
  "formules_conditionnelles": ["formules utilisées pour observations visuelles : 'à l'examen visuel', 'semble présenter', etc."],
  "vocabulaire_technique": ["termes techniques caractéristiques du rapport"],
  "exemple_situation_geographique": "Copier mot-pour-mot 3 à 5 phrases caractéristiques de la section situation géographique ou localisation du rapport. Si absente : null.",
  "exemple_situation_urbanistique": "Copier mot-pour-mot 2 à 4 phrases de la section urbanistique ou PLU. Si absente : null.",
  "exemple_situation_juridique": "Copier mot-pour-mot 2 à 4 phrases de la section juridique ou cadastrale. Si absente : null.",
  "exemple_description_terrain": "Copier mot-pour-mot 4 à 6 phrases caractéristiques du chapitre terrain. Si absente : null.",
  "exemple_description_bati": "Copier mot-pour-mot 4 à 6 phrases caractéristiques du chapitre bâti ou construction. Si absente : null.",
  "style_desordres": "Copier mot-pour-mot 2 à 3 phrases sur la manière dont les désordres sont présentés. Si absente : null."
}
Retourner UNIQUEMENT le JSON valide, sans texte avant ni après.

RAPPORT DE RÉFÉRENCE :
${docText}`;
}

function buildPhotoPrompt(category) {
  const prompts = {
    terrain: `Expert immobilier — analyse ces photos du terrain. JSON :
{"forme_observee":"","topographie_observee":"","acces_observe":"","clotures_observees":"","vegetation":"","environnement_immediat":"","observations":""}
Conditionnel ou "à l'examen visuel". JSON uniquement.`,
    ext: `Expert immobilier — analyse ces photos extérieures. JSON :
{"toiture_materiau":"","toiture_etat":"","toiture_obs":"","facades_materiau":"","facades_etat":"","fissures":"","menuiseries_materiau":"","menuiseries_etat":"","abords":"","desordres_ext":[]}
Conditionnel ou "à l'examen visuel". JSON uniquement.`,
    int: `Expert immobilier — analyse ces photos intérieures. JSON :
{"sols_type":"","sols_etat":"","murs_revetement":"","murs_etat":"","plafonds":"","chauffage_obs":"","electricite_obs":"","humidite":"","desordres_int":[]}
Conditionnel ou "à l'examen visuel". JSON uniquement.`,
    desordres: `Expert immobilier — analyse ce(s) désordre(s). JSON :
{"desordres":[{"nature":"","localisation":"","gravite":"esthetique|fonctionnel|structurel","description":"","origine":"","urgence":"immediate|a_prevoir|surveillance"}]}
Conditionnel obligatoire. JSON uniquement.`
  };
  return prompts[category] || prompts.ext;
}

function buildMainPrompt(data) {
  const { formData, chapter1, style, photos, desordres, surfaces } = data;

  const surfacesArr = formData.surfaces_array || [];
  const surfacesText = surfacesArr.length
    ? surfacesArr.map(s => `${s.type}${s.prec ? ' — ' + s.prec : ''} | ${s.niveau} | ${s.m2} m²`).join('\n')
    : (surfaces || '[à rajouter par l\'expert]');

  // Parser le style extrait (JSON string → objet) pour les exemples de contenu
  let styleObj = null;
  if (style) {
    try { styleObj = JSON.parse(style); } catch(e) { styleObj = null; }
  }

  // Bloc d'exemples de contenu issus du rapport de référence
  const exemples = styleObj ? `
=== EXEMPLES DE CONTENU À REPRODUIRE (extraits mot-pour-mot du rapport de référence) ===
${styleObj.exemple_situation_geographique ? `SITUATION GÉOGRAPHIQUE — reproduire ce style :\n"${styleObj.exemple_situation_geographique}"` : ''}
${styleObj.exemple_situation_urbanistique ? `SITUATION URBANISTIQUE — reproduire ce style :\n"${styleObj.exemple_situation_urbanistique}"` : ''}
${styleObj.exemple_situation_juridique ? `SITUATION JURIDIQUE — reproduire ce style :\n"${styleObj.exemple_situation_juridique}"` : ''}
${styleObj.exemple_description_terrain ? `TERRAIN — reproduire ce style :\n"${styleObj.exemple_description_terrain}"` : ''}
${styleObj.exemple_description_bati ? `CONSTRUCTION — reproduire ce style :\n"${styleObj.exemple_description_bati}"` : ''}
${styleObj.style_desordres ? `DÉSORDRES — reproduire ce style :\n"${styleObj.style_desordres}"` : ''}
` : '';

  return `=== STYLE DE L'EXPERT (reproduire exactement) ===
${style ? (styleObj ? `Ton : ${styleObj.ton_general || 'professionnel'}\nFormules d'introduction : ${(styleObj.formules_introduction || []).slice(0,3).join(' / ')}\nVocabulaire clé : ${(styleObj.vocabulaire_technique || []).slice(0,5).join(', ')}` : style) : 'Style JALTA professionnel — français technique immobilier Antilles — TEGOVA 6e édition.'}
${exemples}

=== DONNÉES DU DOSSIER ===
Référence : ${formData.ref_dossier}
Date de visite : ${formData.date_visite}
Type de mission : ${formData.type_mission}
Donneur d'ordre : ${formData.nom_donneur_ordre} (${formData.donneur_ordre})
Adresse : ${formData.adresse_bien}
Références cadastrales : ${formData.refs_cadastrales || '[à rajouter par l\'expert]'}
Régime juridique : ${formData.regime_juridique}
Situation locative : ${formData.situation_locative || '[à rajouter par l\'expert]'}
DPE : Classe ${formData.dpe_classe || 'NC'} — GES : Classe ${formData.ges_classe || 'NC'}
Type de bien : ${formData.type_bien}
Année de construction : ${formData.annee_construction || '[à rajouter par l\'expert]'}
Niveaux : ${formData.nb_niveaux || '[à rajouter par l\'expert]'}

=== TERRAIN ===
Superficie : ${formData.superficie_terrain} m²
Forme : ${formData.forme_terrain}
Topographie : ${formData.topographie}
Orientation : ${formData.orientation}
Accès : ${formData.acces_terrain}
Clôtures : ${formData.clotures}
Réseaux : ${formData.reseaux}
Assainissement : ${formData.assainissement || '[à rajouter par l\'expert]'}
Zonage PLU : ${formData.zonage_plu || '[à rajouter par l\'expert]'}
Contraintes : ${formData.contraintes}
Notes expert : ${formData.notes_terrain || ''}
Observations photos terrain : ${photos.terrain || 'Aucune photo fournie — [à rajouter par l\'expert]'}

=== BÂTI ===
Structure : ${formData.type_construction}
Toiture : ${formData.materiau_toiture} — ${formData.forme_toiture} — État : ${formData.etat_toiture}
Façades : ${formData.materiau_facades} — État : ${formData.etat_facades}
Menuiseries ext : ${formData.menuiseries_ext}
Sols intérieurs : ${formData.sols_interieurs || '[à rajouter par l\'expert]'}
Chauffage : ${formData.chauffage}
Électricité : ${formData.etat_electrique}
Plomberie : ${formData.etat_plomberie}
Notes expert : ${formData.notes_bati || ''}
Observations photos extérieures : ${photos.ext || 'Aucune photo fournie — [à rajouter par l\'expert]'}
Observations photos intérieures : ${photos.int || 'Aucune photo fournie — [à rajouter par l\'expert]'}

=== DÉSORDRES CONSTATÉS ===
${desordres || 'Aucun désordre renseigné.'}
Observations photos désordres : ${photos.desordres || 'Aucune photo fournie'}

=== SURFACES ===
${surfacesText}

=== SECTION SITUATION GÉOGRAPHIQUE (déjà rédigée — intégrer telle quelle) ===
${chapter1}

---

GÉNÈRE UN JSON avec exactement ces clés (UNIQUEMENT le JSON, sans markdown ni texte avant/après) :

{
  "resume_mission": "Texte introductif de la mission en 2-3 phrases style JALTA : objet de la mission, référence, donneur d'ordre, date de visite.",
  "cadre_evaluation": "Texte du cadre de l'évaluation : normes TEGOVA et Charte appliquées, conditions et limites de la mission, absence de sondages destructifs, observations visuelles au conditionnel — 4 à 6 phrases.",
  "objectif_evaluation": "Texte de l'objectif de l'évaluation : nature de la mission (vénale, locative, etc.), finalité (vente, garantie, fiscalité...) — 2 à 4 phrases style JALTA.",
  "situation_geographique": "Texte complet SITUATION GÉOGRAPHIQUE — intégrer la section géographique déjà rédigée telle quelle.",
  "situation_urbanistique": "Texte SITUATION URBANISTIQUE — INTÉGRER OBLIGATOIREMENT le zonage PLU '${formData.zonage_plu || '[zonage non renseigné]'}' dans la première phrase. Exemple d'ouverture : 'Au regard du Plan Local d'Urbanisme en vigueur, le bien est classé en zone ${formData.zonage_plu || '[à compléter]'}...'. Décrire les règles d'urbanisme applicables à cette zone, les possibilités de construction, l'assainissement (${formData.assainissement || '[à compléter]'}), les servitudes connues — 3 à 5 phrases style JALTA.",
  "situation_juridique": "Texte SITUATION JURIDIQUE — INTÉGRER OBLIGATOIREMENT la référence cadastrale '${formData.refs_cadastrales || '[référence à compléter]'}' dans le texte. Exemple d'ouverture : 'Le bien est cadastré sous la référence ${formData.refs_cadastrales || '[à compléter]'}...'. Mentionner le régime juridique (${formData.regime_juridique || '[à compléter]'}), la superficie du terrain (${formData.superficie_terrain || '[à compléter]'} m²), les mentions hypothécaires si connues — 3 à 5 phrases style JALTA.",
  "situation_locative_text": "Texte SITUATION LOCATIVE : si libre d'occupation ou occupé, conditions de l'occupation, incidence sur la valeur — 2 à 4 phrases. Si libre : le préciser clairement.",
  "description_terrain": "Texte section LE TERRAIN D'ASSIETTE — au moins 150 mots — style JALTA : 'Le terrain objet de la présente expertise...', surface, forme, accès, clôtures, réseaux, PLU, contraintes.",
  "description_bati": "Texte section LA CONSTRUCTION (extérieur et intérieur) — au moins 200 mots — style JALTA : 'Il s'agit d'un bâtiment en dur...', structure, toiture, façades, menuiseries, intérieur, équipements, DPE.",
  "desordres_texte": "Texte section ÉTAT DES LIEUX — liste tous les désordres en style JALTA avec conditionnel — si aucun : 'Au jour de notre visite, aucun désordre significatif n'a été constaté.'",
  "jugement_favorable": ["Point favorable 1", "Point favorable 2", "..."],
  "jugement_defavorable": ["Point défavorable 1", "Point défavorable 2", "..."],
  "elements_jugement_intro": "Phrase d'introduction des éléments de jugement style JALTA.",
  "conclusion": "Texte de conclusion du rapport — 3-5 phrases — style JALTA : synthèse de la mission, rappel normes TEGOVA, mention que la valeur vénale sera arrêtée par l'expert signataire."
}`;
}

// ─────────────────────────────────────────────────────────────────────────────
// ROUTES
// ─────────────────────────────────────────────────────────────────────────────

// POST /api/chapter1
app.post('/api/chapter1', async (req, res) => {
  try {
    const { adresse } = req.body;
    if (!adresse) return res.status(400).json({ error: 'Adresse manquante' });

    const response = await client.messages.create({
      model: process.env.CLAUDE_MODEL || 'claude-sonnet-4-6',
      max_tokens: 2000,
      temperature: 0,
      tools: [{ type: 'web_search_20250305', name: 'web_search', max_uses: 5 }],
      messages: [{ role: 'user', content: buildChapter1Prompt(adresse) }]
    });

    const text = response.content
      .filter(b => b.type === 'text')
      .map(b => b.text)
      .join('\n')
      .trim();

    res.json({ text });
  } catch (err) {
    console.error('[chapter1]', err.message);
    res.status(500).json({ error: err.message });
  }
});

// Extrait la première image (logo) d'un DOCX via JSZip
async function extractLogoFromDocx(buffer) {
  try {
    const zip = await JSZip.loadAsync(buffer);
    const mediaFiles = Object.keys(zip.files).filter(name =>
      name.startsWith('word/media/') && /\.(png|jpg|jpeg|gif|webp)$/i.test(name)
    );
    if (!mediaFiles.length) return null;
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

// POST /api/extract-style
app.post('/api/extract-style', upload.single('document'), async (req, res) => {
  try {
    if (!req.file) return res.json({ style: null, logo: null });

    let docText = '';
    let logo = null;

    if (req.file.originalname.endsWith('.docx')) {
      const [mammothResult, extractedLogo] = await Promise.all([
        mammoth.extractRawText({ buffer: req.file.buffer }),
        extractLogoFromDocx(req.file.buffer)
      ]);
      docText = mammothResult.value.slice(0, 8000);
      logo = extractedLogo;
    } else if (req.file.originalname.endsWith('.pdf') || req.file.mimetype === 'application/pdf') {
      const response = await client.messages.create({
        model: process.env.CLAUDE_MODEL || 'claude-sonnet-4-6',
        max_tokens: 3000,
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

// POST /api/analyze-photos
app.post('/api/analyze-photos', upload.fields([
  { name: 'terrain', maxCount: 10 },
  { name: 'ext', maxCount: 15 },
  { name: 'int', maxCount: 20 },
  { name: 'desordres', maxCount: 15 }
]), async (req, res) => {
  try {
    const results = {};
    const categories = ['terrain', 'ext', 'int', 'desordres'];

    for (const cat of categories) {
      const files = req.files?.[cat];
      if (!files || files.length === 0) { results[cat] = null; continue; }

      const imageBlocks = files.map(f => ({
        type: 'image',
        source: { type: 'base64', media_type: f.mimetype, data: f.buffer.toString('base64') }
      }));

      const response = await client.messages.create({
        model: process.env.CLAUDE_MODEL || 'claude-sonnet-4-6',
        max_tokens: 1000,
        temperature: 0,
        messages: [{
          role: 'user',
          content: [...imageBlocks, { type: 'text', text: buildPhotoPrompt(cat) }]
        }]
      });
      results[cat] = response.content[0].text;
    }

    res.json(results);
  } catch (err) {
    console.error('[analyze-photos]', err.message);
    res.status(500).json({ error: err.message });
  }
});

// POST /api/generate
app.post('/api/generate', async (req, res) => {
  try {
    const response = await client.messages.create({
      model: process.env.CLAUDE_MODEL || 'claude-sonnet-4-6',
      max_tokens: 8000,
      temperature: 0.2,
      system: SYSTEM_PROMPT,
      messages: [{ role: 'user', content: buildMainPrompt(req.body) }]
    });

    const raw = response.content[0].text.trim();

    // Extraire le JSON (au cas où Claude ajouterait du texte autour)
    const jsonMatch = raw.match(/\{[\s\S]+\}/);
    let sections = null;
    let report = raw;

    if (jsonMatch) {
      try {
        sections = JSON.parse(jsonMatch[0]);
        // Générer un rapport markdown lisible pour l'aperçu
        report = buildMarkdownFromSections(sections, req.body.formData);
      } catch (e) {
        console.warn('[generate] JSON parse failed, using raw text');
      }
    }

    res.json({ report, sections });
  } catch (err) {
    console.error('[generate]', err.message);
    res.status(500).json({ error: err.message });
  }
});

// POST /api/export-docx
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

// ─────────────────────────────────────────────────────────────────────────────
// MARKDOWN PREVIEW (aperçu HTML)
// ─────────────────────────────────────────────────────────────────────────────

function buildMarkdownFromSections(sections, formData) {
  const fd = formData || {};
  return `# RAPPORT D'EXPERTISE IMMOBILIÈRE
## Pré-rapport soumis à validation

**Référence dossier :** ${fd.ref_dossier || '[à rajouter par l\'expert]'}
**Date de visite :** ${fd.date_visite || '[à rajouter par l\'expert]'}
**Adresse :** ${fd.adresse_bien || '[à rajouter par l\'expert]'}
**Nature de la mission :** ${fd.type_mission || '[à rajouter par l\'expert]'}
**Donneur d'ordre :** ${fd.nom_donneur_ordre || ''} (${fd.donneur_ordre || ''})

*Le présent document constitue un pré-rapport préparatoire. Les valeurs vénales et conclusions définitives feront l'objet d'une analyse complémentaire par l'expert signataire.*

---

## I/ RÉSUMÉ DE LA MISSION

${sections.resume_mission || '[à rajouter par l\'expert]'}

---

## II/ EXPERTISE DÉTAILLÉE

### 1 CADRE DE L'ÉVALUATION
${sections.cadre_evaluation || '[à rajouter par l\'expert]'}

### 2 OBJECTIF DE L'ÉVALUATION
${sections.objectif_evaluation || '[à rajouter par l\'expert]'}

---

## III/ SITUATION

### 1 SITUATION GÉOGRAPHIQUE
${sections.situation_geographique || '[à rajouter par l\'expert]'}

### 2 SITUATION URBANISTIQUE
${sections.situation_urbanistique || '[à rajouter par l\'expert]'}

### 3 SITUATION JURIDIQUE
${sections.situation_juridique || '[à rajouter par l\'expert]'}

### 4 SITUATION LOCATIVE
${sections.situation_locative_text || '[à rajouter par l\'expert]'}

---

## IV/ DESCRIPTION DU BIEN

### LE TERRAIN D'ASSIETTE
${sections.description_terrain || '[à rajouter par l\'expert]'}

### LA CONSTRUCTION
${sections.description_bati || '[à rajouter par l\'expert]'}

### ÉTAT DES LIEUX
${sections.desordres_texte || '[à rajouter par l\'expert]'}

---

## V/ ÉLÉMENTS DE JUGEMENT

${sections.elements_jugement_intro || ''}

**Éléments favorables :**
${(sections.jugement_favorable || []).map(p => '- ' + p).join('\n') || '[à rajouter par l\'expert]'}

**Éléments défavorables :**
${(sections.jugement_defavorable || []).map(p => '- ' + p).join('\n') || '[à rajouter par l\'expert]'}

---

## VI/ ÉVALUATION

[à rajouter par l'expert]

---

## CONCLUSION

${sections.conclusion || '[à rajouter par l\'expert]'}`;
}

// ─────────────────────────────────────────────────────────────────────────────
// GÉNÉRATION DOCX JALTA
// ─────────────────────────────────────────────────────────────────────────────

// Helpers
function noBorders() {
  const none = { style: BorderStyle.NONE, size: 0, color: C.WHITE };
  return { top: none, bottom: none, left: none, right: none, insideH: none, insideV: none };
}

function cellBorder(color = C.GRAY_MED) {
  const b = { style: BorderStyle.SINGLE, size: 4, color };
  return { top: b, bottom: b, left: b, right: b };
}

function shadedCell(fill, children, opts = {}) {
  return new TableCell({
    shading: { fill, type: ShadingType.CLEAR, color: C.WHITE },
    borders: opts.borders || noBorders(),
    width: opts.width,
    columnSpan: opts.columnSpan,
    verticalAlign: opts.verticalAlign || 'center',
    margins: opts.margins || { top: 80, bottom: 80, left: 120, right: 120 },
    children
  });
}

function navyBanner(text, fontSize = 22) {
  return new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    borders: noBorders(),
    rows: [new TableRow({
      children: [shadedCell(C.NAVY, [
        new Paragraph({
          children: [new TextRun({ text, bold: true, color: C.WHITE, size: fontSize, font: 'Times New Roman' })],
          alignment: AlignmentType.LEFT,
          spacing: { before: 60, after: 60 }
        })
      ], { borders: noBorders() })]
    })]
  });
}

function subBanner(text) {
  return new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    borders: noBorders(),
    rows: [new TableRow({
      children: [shadedCell(C.NAVY_L, [
        new Paragraph({
          children: [new TextRun({ text, bold: true, color: C.NAVY, size: 20, font: 'Times New Roman' })],
          alignment: AlignmentType.LEFT,
          spacing: { before: 40, after: 40 }
        })
      ], { borders: noBorders() })]
    })]
  });
}

function bodyPara(text, opts = {}) {
  const runs = [];
  // Traiter les [à rajouter par l'expert] en orange
  const parts = text.split(/(\[à rajouter par l'expert\])/gi);
  for (const part of parts) {
    if (part.toLowerCase() === "[à rajouter par l'expert]") {
      runs.push(new TextRun({ text: part, color: C.AMBER, size: opts.size || 20, font: 'Times New Roman', bold: true }));
    } else if (part) {
      runs.push(new TextRun({ text: part, size: opts.size || 20, font: 'Times New Roman', color: opts.color || C.DARK }));
    }
  }
  return new Paragraph({
    children: runs.length ? runs : [new TextRun({ text, size: opts.size || 20, font: 'Times New Roman', color: opts.color || C.DARK })],
    alignment: opts.alignment || AlignmentType.JUSTIFIED,
    spacing: { before: opts.before || 60, after: opts.after || 100, line: opts.line || 276 }
  });
}

function spacer(pts = 200) {
  return new Paragraph({ text: '', spacing: { before: 0, after: pts } });
}

function pageBreak() {
  return new Paragraph({ pageBreakBefore: true, children: [new TextRun('')] });
}

function imagePlaceholder(label) {
  return new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    borders: noBorders(),
    rows: [new TableRow({
      children: [new TableCell({
        shading: { fill: C.GRAY, type: ShadingType.CLEAR },
        borders: { top: { style: BorderStyle.SINGLE, size: 4, color: C.GRAY_MED }, bottom: { style: BorderStyle.SINGLE, size: 4, color: C.GRAY_MED }, left: { style: BorderStyle.SINGLE, size: 4, color: C.GRAY_MED }, right: { style: BorderStyle.SINGLE, size: 4, color: C.GRAY_MED } },
        margins: { top: 400, bottom: 400, left: 200, right: 200 },
        children: [new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [
            new TextRun({ text: label, color: C.AMBER, size: 20, bold: true, font: 'Times New Roman' })
          ]
        })]
      })]
    })]
  });
}

// ── HELPERS IMAGES ───────────────────────────────────────────────────────────

function buildImageRun(photo64, opts = {}) {
  if (!photo64 || !photo64.data) return null;
  try {
    const buf = Buffer.from(photo64.data, 'base64');
    const imgType = (photo64.mimeType || 'image/jpeg').replace('image/', '');
    return new ImageRun({
      data: buf,
      transformation: { width: opts.width || 450, height: opts.height || 300 },
      type: imgType === 'jpg' ? 'jpeg' : imgType,
    });
  } catch (e) {
    console.warn('[buildImageRun] Erreur:', e.message);
    return null;
  }
}

function buildPhotoParagraphs(photos64Array, caption = '') {
  if (!photos64Array || !photos64Array.length) return [];
  const items = [];

  for (let i = 0; i < photos64Array.length; i += 2) {
    const left = buildImageRun(photos64Array[i], { width: 215, height: 160 });
    const right = i + 1 < photos64Array.length
      ? buildImageRun(photos64Array[i + 1], { width: 215, height: 160 })
      : null;

    if (!left) continue;

    if (right) {
      items.push(new Table({
        width: { size: 100, type: WidthType.PERCENTAGE },
        borders: noBorders(),
        rows: [new TableRow({
          children: [
            new TableCell({
              borders: noBorders(),
              width: { size: 49, type: WidthType.PERCENTAGE },
              margins: { top: 60, bottom: 60, right: 100, left: 0 },
              children: [new Paragraph({ children: [left], alignment: AlignmentType.CENTER })]
            }),
            new TableCell({
              borders: noBorders(),
              width: { size: 49, type: WidthType.PERCENTAGE },
              margins: { top: 60, bottom: 60, left: 100, right: 0 },
              children: [new Paragraph({ children: [right], alignment: AlignmentType.CENTER })]
            }),
          ]
        })]
      }));
    } else {
      const solo = buildImageRun(photos64Array[i], { width: 450, height: 300 });
      if (solo) items.push(new Paragraph({ children: [solo], alignment: AlignmentType.CENTER, spacing: { before: 60, after: 60 } }));
    }
    items.push(spacer(60));
  }

  if (caption && items.length) {
    items.push(new Paragraph({
      children: [new TextRun({ text: caption, size: 17, italics: true, color: C.DARK, font: 'Times New Roman' })],
      alignment: AlignmentType.CENTER,
      spacing: { before: 40, after: 120 }
    }));
  }

  return items;
}

// ─────────────────────────────────────────────────────────────────────────────

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
          new Paragraph({
            alignment: AlignmentType.CENTER,
            spacing: { before: 200, after: 100 },
            children: [new TextRun({ text: 'RAPPORT D\'EXPERTISE IMMOBILIÈRE', bold: true, color: C.WHITE, size: 36, font: 'Times New Roman' })]
          }),
          new Paragraph({
            alignment: AlignmentType.CENTER,
            spacing: { before: 60, after: 200 },
            children: [new TextRun({ text: 'Pré-rapport soumis à validation de l\'expert signataire', color: C.NAVY_L, size: 22, font: 'Times New Roman', italics: true })]
          })
        ], { borders: noBorders() })]
      })]
    }),
    spacer(300),

    // Tableau identité dossier
    new Table({
      width: { size: 100, type: WidthType.PERCENTAGE },
      borders: noBorders(),
      rows: [
        buildCoverRow('Référence dossier', fd.ref_dossier || '[à rajouter par l\'expert]'),
        buildCoverRow('Date de visite', fd.date_visite || '[à rajouter par l\'expert]'),
        buildCoverRow('Adresse du bien', fd.adresse_bien || '[à rajouter par l\'expert]'),
        buildCoverRow('Type de bien', fd.type_bien || '[à rajouter par l\'expert]'),
        buildCoverRow('Nature de la mission', fd.type_mission || '[à rajouter par l\'expert]'),
        buildCoverRow('Donneur d\'ordre', `${fd.nom_donneur_ordre || ''} — ${fd.donneur_ordre || ''}`.replace(/^ — | — $/, '')),
        buildCoverRow('Régime juridique', fd.regime_juridique || '[à rajouter par l\'expert]'),
        buildCoverRow('Références cadastrales', fd.refs_cadastrales || '[à rajouter par l\'expert]'),
      ]
    }),

    spacer(400),
    imagePlaceholder('[à rajouter par l\'expert] — Photo de couverture du bien'),
    spacer(300),

    // Clause confidentialité
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 100, after: 60 },
      children: [new TextRun({ text: 'CLAUSE DE CONFIDENTIALITÉ', bold: true, size: 18, font: 'Times New Roman', color: C.NAVY })]
    }),
    new Paragraph({
      alignment: AlignmentType.JUSTIFIED,
      spacing: { before: 60, after: 60 },
      children: [new TextRun({
        text: 'Le présent rapport est établi à la demande et à l\'usage exclusif du donneur d\'ordre. Il ne peut être communiqué à des tiers sans l\'accord écrit de l\'expert signataire. Toute reproduction partielle ou totale est interdite. Ce document constitue un pré-rapport préparatoire — les valeurs vénales et conclusions définitives feront l\'objet d\'une validation complémentaire par l\'expert signataire conformément à la Charte de l\'Expertise Immobilière (5e édition) et au référentiel TEGOVA (6e édition).',
        size: 17, font: 'Times New Roman', color: C.DARK, italics: true
      })]
    })
  );
  return items;
}

function buildCoverRow(label, value) {
  return new TableRow({
    children: [
      shadedCell(C.NAVY_L, [
        new Paragraph({ children: [new TextRun({ text: label, bold: true, size: 20, font: 'Times New Roman', color: C.NAVY })], spacing: { before: 40, after: 40 } })
      ], { width: { size: 35, type: WidthType.PERCENTAGE }, borders: { bottom: { style: BorderStyle.SINGLE, size: 2, color: C.WHITE } } }),
      shadedCell(C.GRAY, [
        new Paragraph({ children: [new TextRun({ text: value, size: 20, font: 'Times New Roman', color: value.includes('[à rajouter') ? C.AMBER : C.DARK })], spacing: { before: 40, after: 40 } })
      ], { width: { size: 65, type: WidthType.PERCENTAGE }, borders: { bottom: { style: BorderStyle.SINGLE, size: 2, color: C.WHITE } } })
    ]
  });
}

function buildSommaire() {
  const entries = [
    ['I/', 'Résumé de la mission'],
    ['II/', 'Expertise détaillée'],
    ['III/', 'Situation'],
    ['IV/', 'Description du bien'],
    ['V/', 'Éléments de jugement'],
    ['VI/', 'Évaluation'],
    ['', 'Conclusion'],
    ['', 'Photographies'],
    ['', 'Annexes'],
    ['', 'Glossaire'],
  ];

  return [
    pageBreak(),
    navyBanner('SOMMAIRE'),
    spacer(150),
    ...entries.map(([num, titre]) =>
      new Table({
        width: { size: 100, type: WidthType.PERCENTAGE },
        borders: noBorders(),
        rows: [new TableRow({
          children: [
            new TableCell({
              width: { size: 12, type: WidthType.PERCENTAGE },
              borders: noBorders(),
              children: [new Paragraph({ children: [new TextRun({ text: num, bold: !!num, size: 20, font: 'Times New Roman', color: C.NAVY })], spacing: { before: 60, after: 60 } })]
            }),
            new TableCell({
              width: { size: 75, type: WidthType.PERCENTAGE },
              borders: { bottom: { style: BorderStyle.DOTTED, size: 2, color: C.GRAY_MED } },
              children: [new Paragraph({ children: [new TextRun({ text: titre, bold: !!num, size: 20, font: 'Times New Roman', color: num ? C.NAVY : C.DARK })], spacing: { before: 60, after: 60 } })]
            }),
            new TableCell({
              width: { size: 13, type: WidthType.PERCENTAGE },
              borders: noBorders(),
              children: [new Paragraph({ children: [new TextRun({ text: '[à rajouter par l\'expert]', size: 18, font: 'Times New Roman', color: C.AMBER })], alignment: AlignmentType.RIGHT, spacing: { before: 60, after: 60 } })]
            })
          ]
        })]
      })
    ),
  ];
}

function buildResumeSection(sections, formData) {
  const fd = formData || {};
  const rows = [
    ['REQUÉRANT', `${fd.nom_donneur_ordre || ''} — ${fd.donneur_ordre || ''}`.replace(/^ — | — $/, '') || '[à rajouter par l\'expert]'],
    ['ADRESSE DU BIEN', fd.adresse_bien || '[à rajouter par l\'expert]'],
    ['RÉFÉRENCE CADASTRALE', fd.refs_cadastrales || '[à rajouter par l\'expert]'],
    ['TYPE D\'ACTIF', fd.type_bien || '[à rajouter par l\'expert]'],
    ['DATE DE L\'ÉVALUATION', fd.date_visite || '[à rajouter par l\'expert]'],
    ['ASSAINISSEMENT', fd.assainissement || '[à rajouter par l\'expert]'],
    ['OBJECTIF DE L\'ÉVALUATION', fd.type_mission || '[à rajouter par l\'expert]'],
    ['SITUATION URBANISTIQUE', fd.zonage_plu || '[à rajouter par l\'expert]'],
    ['MÉTHODE D\'ÉVALUATION', 'Méthode par comparaison directe'],
    ['VALEUR VÉNALE RETENUE', '[à rajouter par l\'expert]'],
  ];

  return [
    pageBreak(),
    navyBanner('I/ RÉSUMÉ'),
    spacer(120),
    bodyPara(sections.resume_mission || '[à rajouter par l\'expert]'),
    spacer(150),
    new Table({
      width: { size: 100, type: WidthType.PERCENTAGE },
      borders: noBorders(),
      rows: rows.map(([label, val], i) => {
        const isValeur = label === 'VALEUR VÉNALE RETENUE';
        return new TableRow({
          children: [
            shadedCell(isValeur ? C.GREEN : (i % 2 === 0 ? C.NAVY_L : C.GRAY), [
              new Paragraph({ children: [new TextRun({ text: label, bold: true, size: 19, font: 'Times New Roman', color: isValeur ? C.NAVY : C.NAVY })], spacing: { before: 50, after: 50 } })
            ], { width: { size: 40, type: WidthType.PERCENTAGE } }),
            shadedCell(isValeur ? C.GREEN : C.WHITE, [
              new Paragraph({ children: [new TextRun({ text: val, size: 19, font: 'Times New Roman', color: val.includes('[à rajouter') ? C.AMBER : (isValeur ? C.NAVY : C.DARK), bold: isValeur })], spacing: { before: 50, after: 50 } })
            ], { width: { size: 60, type: WidthType.PERCENTAGE } })
          ]
        });
      })
    }),
  ];
}

function buildExpertiseDetaillee(sections, formData) {
  const fd = formData || {};
  const s = sections || {};
  return [
    pageBreak(),
    navyBanner('II/ EXPERTISE DÉTAILLÉE'),
    spacer(120),
    subBanner('1   CADRE DE L\'ÉVALUATION'),
    spacer(80),
    ...splitParagraphs(s.cadre_evaluation || 'L\'expertise a été conduite conformément aux dispositions de la Charte de l\'Expertise Immobilière (5e édition) et du référentiel TEGOVA (6e édition). La présente expertise repose sur l\'examen visuel du bien lors de la visite, les documents fournis par le donneur d\'ordre, et les données de marché disponibles au jour de la mission. Il est précisé que l\'expert n\'a pas réalisé de sondages destructifs ni de diagnostics techniques spécialisés. Les observations relatives aux éléments non accessibles sont formulées au conditionnel.'),
    spacer(120),
    subBanner('2   OBJECTIF DE L\'ÉVALUATION'),
    spacer(80),
    ...splitParagraphs(s.objectif_evaluation || fd.type_mission || '[à rajouter par l\'expert]'),
    spacer(120),
    subBanner('3   DATE DE L\'ÉVALUATION'),
    spacer(80),
    bodyPara(`La présente évaluation a été réalisée à la date du ${fd.date_visite || '[à rajouter par l\'expert]'}.`),
    spacer(120),
    subBanner('4   VISITE ET DOCUMENTS MIS À DISPOSITION'),
    spacer(80),
    imagePlaceholder('[à rajouter par l\'expert] — Liste des pièces et documents consultés'),
    spacer(120),
    subBanner('5   CLAUSE DE CONFIDENTIALITÉ'),
    spacer(80),
    bodyPara('Le présent rapport est établi à la demande et à l\'usage exclusif du donneur d\'ordre. Il ne peut être communiqué à des tiers sans l\'accord écrit de l\'expert signataire. Toute reproduction partielle ou totale est interdite sans autorisation préalable.'),
  ];
}

function buildSituationSection(sections) {
  const s = sections || {};
  return [
    pageBreak(),
    navyBanner('III/ SITUATION'),
    spacer(120),
    subBanner('1   SITUATION GÉOGRAPHIQUE'),
    spacer(80),
    ...splitParagraphs(s.situation_geographique || '[à rajouter par l\'expert]'),
    spacer(100),
    imagePlaceholder('[à rajouter par l\'expert] — Plan de situation / Carte de localisation'),
    spacer(150),
    subBanner('2   SITUATION URBANISTIQUE'),
    spacer(80),
    ...splitParagraphs(s.situation_urbanistique || '[à rajouter par l\'expert]'),
    spacer(150),
    subBanner('3   SITUATION JURIDIQUE'),
    spacer(80),
    ...splitParagraphs(s.situation_juridique || '[à rajouter par l\'expert]'),
    spacer(100),
    imagePlaceholder('[à rajouter par l\'expert] — Plan cadastral'),
    spacer(150),
    subBanner('4   SITUATION LOCATIVE'),
    spacer(80),
    ...splitParagraphs(s.situation_locative_text || '[à rajouter par l\'expert]'),
  ];
}

function buildDescriptionSection(sections, formData, p64 = {}) {
  const fd = formData || {};
  const surfacesArr = fd.surfaces_array || [];

  // Calculer surfaces par catégorie pour pondération
  const totalHab = surfacesArr.filter(s => !s.type?.toLowerCase().includes('garage') && !s.type?.toLowerCase().includes('cave') && !s.type?.toLowerCase().includes('cellier') && !s.type?.toLowerCase().includes('veranda'))
    .reduce((sum, s) => sum + (parseFloat(s.m2) || 0), 0);
  const totalAnnexes = surfacesArr.filter(s => s.type?.toLowerCase().includes('garage') || s.type?.toLowerCase().includes('cave') || s.type?.toLowerCase().includes('cellier') || s.type?.toLowerCase().includes('veranda'))
    .reduce((sum, s) => sum + (parseFloat(s.m2) || 0), 0);

  const surfaceRows = surfacesArr.length > 0
    ? surfacesArr.map((s, i) =>
        new TableRow({
          children: [
            new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: `${s.type || ''}${s.prec ? ' — ' + s.prec : ''}`, size: 19, font: 'Times New Roman' })], spacing: { before: 40, after: 40 } })], borders: cellBorder() }),
            new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: s.niveau || '', size: 19, font: 'Times New Roman' })], spacing: { before: 40, after: 40 } })], borders: cellBorder() }),
            new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: s.m2 ? `${s.m2} m²` : '[à compléter]', size: 19, font: 'Times New Roman' })], alignment: AlignmentType.RIGHT, spacing: { before: 40, after: 40 } })], borders: cellBorder() }),
            new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: '1,00', size: 19, font: 'Times New Roman' })], alignment: AlignmentType.RIGHT, spacing: { before: 40, after: 40 } })], borders: cellBorder() }),
            new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: s.m2 ? `${parseFloat(s.m2).toFixed(2)} m²` : '[à compléter]', size: 19, font: 'Times New Roman' })], alignment: AlignmentType.RIGHT, spacing: { before: 40, after: 40 } })], borders: cellBorder() }),
          ]
        })
      )
    : [new TableRow({
        children: [new TableCell({ columnSpan: 5, children: [new Paragraph({ children: [new TextRun({ text: '[à rajouter par l\'expert] — Tableau des surfaces', size: 19, font: 'Times New Roman', color: C.AMBER })], alignment: AlignmentType.CENTER, spacing: { before: 80, after: 80 } })], borders: cellBorder() })]
      })];

  const surfaceTable = new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    borders: noBorders(),
    rows: [
      // En-tête
      new TableRow({
        children: [
          shadedCell(C.NAVY, [new Paragraph({ children: [new TextRun({ text: 'Désignation', bold: true, size: 19, font: 'Times New Roman', color: C.WHITE })], spacing: { before: 40, after: 40 } })], { width: { size: 35, type: WidthType.PERCENTAGE } }),
          shadedCell(C.NAVY, [new Paragraph({ children: [new TextRun({ text: 'Niveau', bold: true, size: 19, font: 'Times New Roman', color: C.WHITE })], spacing: { before: 40, after: 40 } })], { width: { size: 20, type: WidthType.PERCENTAGE } }),
          shadedCell(C.NAVY, [new Paragraph({ children: [new TextRun({ text: 'Surface', bold: true, size: 19, font: 'Times New Roman', color: C.WHITE })], alignment: AlignmentType.RIGHT, spacing: { before: 40, after: 40 } })], { width: { size: 15, type: WidthType.PERCENTAGE } }),
          shadedCell(C.NAVY, [new Paragraph({ children: [new TextRun({ text: 'Coeff.', bold: true, size: 19, font: 'Times New Roman', color: C.WHITE })], alignment: AlignmentType.RIGHT, spacing: { before: 40, after: 40 } })], { width: { size: 15, type: WidthType.PERCENTAGE } }),
          shadedCell(C.NAVY, [new Paragraph({ children: [new TextRun({ text: 'Surface pond.', bold: true, size: 19, font: 'Times New Roman', color: C.WHITE })], alignment: AlignmentType.RIGHT, spacing: { before: 40, after: 40 } })], { width: { size: 15, type: WidthType.PERCENTAGE } }),
        ]
      }),
      ...surfaceRows,
      // Ligne totale
      new TableRow({
        children: [
          shadedCell(C.NAVY_L, [new Paragraph({ children: [new TextRun({ text: 'Surface habitable totale (Loi Boutin)', bold: true, size: 19, font: 'Times New Roman', color: C.NAVY })], spacing: { before: 60, after: 60 } })], { columnSpan: 3 }),
          shadedCell(C.NAVY_L, [new Paragraph({ children: [new TextRun({ text: totalHab > 0 ? `${totalHab.toFixed(2)} m²` : '[à rajouter par l\'expert]', bold: true, size: 19, font: 'Times New Roman', color: C.NAVY })], alignment: AlignmentType.RIGHT, spacing: { before: 60, after: 60 } })], { columnSpan: 2 }),
        ]
      }),
      new TableRow({
        children: [
          shadedCell(C.GRAY, [new Paragraph({ children: [new TextRun({ text: 'Surfaces annexes (garages, caves, vérandas...)', size: 19, font: 'Times New Roman' })], spacing: { before: 40, after: 40 } })], { columnSpan: 3 }),
          shadedCell(C.GRAY, [new Paragraph({ children: [new TextRun({ text: totalAnnexes > 0 ? `${totalAnnexes.toFixed(2)} m²` : '[à rajouter par l\'expert]', size: 19, font: 'Times New Roman' })], alignment: AlignmentType.RIGHT, spacing: { before: 40, after: 40 } })], { columnSpan: 2 }),
        ]
      }),
    ]
  });

  return [
    pageBreak(),
    navyBanner('IV/ DESCRIPTION DU BIEN'),
    spacer(120),
    subBanner('LE TERRAIN D\'ASSIETTE'),
    spacer(80),
    ...splitParagraphs(sections.description_terrain || '[à rajouter par l\'expert]'),
    spacer(100),
    ...(p64.terrain && p64.terrain.length
      ? buildPhotoParagraphs(p64.terrain, 'Vues du terrain — photos prises lors de la visite')
      : [imagePlaceholder('[à rajouter par l\'expert] — Photos du terrain / Plan de masse')]),
    spacer(150),
    subBanner('LA CONSTRUCTION'),
    spacer(80),
    ...splitParagraphs(sections.description_bati || '[à rajouter par l\'expert]'),
    spacer(100),
    ...(p64.ext && p64.ext.length
      ? buildPhotoParagraphs(p64.ext, 'Vues extérieures')
      : [imagePlaceholder('[à rajouter par l\'expert] — Photos extérieures')]),
    spacer(80),
    ...(p64.int && p64.int.length
      ? buildPhotoParagraphs(p64.int, 'Vues intérieures')
      : [imagePlaceholder('[à rajouter par l\'expert] — Photos intérieures')]),
    spacer(150),
    subBanner('SURFACES'),
    spacer(80),
    surfaceTable,
    spacer(150),
    subBanner('ÉTAT DES LIEUX'),
    spacer(80),
    ...splitParagraphs(sections.desordres_texte || 'Au jour de notre visite, aucun désordre significatif n\'a été constaté.'),
    ...(p64.desordres && p64.desordres.length
      ? [spacer(80), ...buildPhotoParagraphs(p64.desordres, 'Désordres constatés lors de la visite')]
      : []),
  ];
}

function buildJugementSection(sections) {
  const favorable = sections.jugement_favorable || [];
  const defavorable = sections.jugement_defavorable || [];

  const maxRows = Math.max(favorable.length, defavorable.length, 1);
  const rows = [];
  for (let i = 0; i < maxRows; i++) {
    rows.push(new TableRow({
      children: [
        new TableCell({
          shading: { fill: C.WHITE, type: ShadingType.CLEAR },
          borders: cellBorder(C.GRAY_MED),
          margins: { top: 60, bottom: 60, left: 120, right: 120 },
          children: [new Paragraph({
            children: [new TextRun({ text: favorable[i] || '', size: 19, font: 'Times New Roman' })],
            spacing: { before: 40, after: 40 }
          })]
        }),
        new TableCell({
          shading: { fill: C.WHITE, type: ShadingType.CLEAR },
          borders: cellBorder(C.GRAY_MED),
          margins: { top: 60, bottom: 60, left: 120, right: 120 },
          children: [new Paragraph({
            children: [new TextRun({ text: defavorable[i] || '', size: 19, font: 'Times New Roman' })],
            spacing: { before: 40, after: 40 }
          })]
        }),
      ]
    }));
  }

  return [
    pageBreak(),
    navyBanner('V/ ÉLÉMENTS DE JUGEMENT'),
    spacer(120),
    bodyPara(sections.elements_jugement_intro || 'L\'appréciation du bien objet de la présente expertise repose sur l\'analyse des éléments favorables et défavorables suivants :'),
    spacer(100),
    new Table({
      width: { size: 100, type: WidthType.PERCENTAGE },
      borders: noBorders(),
      rows: [
        new TableRow({
          children: [
            shadedCell(C.NAVY, [new Paragraph({ children: [new TextRun({ text: 'FAVORABLE', bold: true, color: C.WHITE, size: 20, font: 'Times New Roman' })], alignment: AlignmentType.CENTER, spacing: { before: 60, after: 60 } })], { borders: noBorders() }),
            shadedCell(C.NAVY, [new Paragraph({ children: [new TextRun({ text: 'DÉFAVORABLE', bold: true, color: C.WHITE, size: 20, font: 'Times New Roman' })], alignment: AlignmentType.CENTER, spacing: { before: 60, after: 60 } })], { borders: noBorders() }),
          ]
        }),
        ...rows
      ]
    }),
  ];
}

function buildEvaluationSection(formData) {
  return [
    pageBreak(),
    navyBanner('VI/ ÉVALUATION'),
    spacer(120),
    subBanner('Méthode d\'évaluation'),
    spacer(80),
    bodyPara('L\'évaluation du bien objet de la présente expertise est réalisée par comparaison directe avec des références de marché récentes (méthode comparative). Les termes de comparaison ont été sélectionnés dans la même zone géographique, pour des biens de nature et caractéristiques similaires.'),
    bodyPara('[à rajouter par l\'expert] — Développement de la méthode et des termes de comparaison retenus.'),
    spacer(150),
    subBanner('Références de marché'),
    spacer(80),
    imagePlaceholder('[à rajouter par l\'expert] — Tableau des termes de comparaison'),
    spacer(150),
    subBanner('Calcul de la valeur vénale'),
    spacer(80),
    imagePlaceholder('[à rajouter par l\'expert] — Tableau de calcul et justification de la valeur'),
    spacer(200),
    // Encadré valeur vénale
    new Table({
      width: { size: 70, type: WidthType.PERCENTAGE },
      borders: noBorders(),
      rows: [new TableRow({
        children: [shadedCell(C.GREEN, [
          new Paragraph({
            alignment: AlignmentType.CENTER,
            spacing: { before: 100, after: 40 },
            children: [new TextRun({ text: 'VALEUR VÉNALE RETENUE', bold: true, color: C.NAVY, size: 22, font: 'Times New Roman' })]
          }),
          new Paragraph({
            alignment: AlignmentType.CENTER,
            spacing: { before: 40, after: 40 },
            children: [new TextRun({ text: '[à rajouter par l\'expert]', bold: true, color: C.NAVY, size: 36, font: 'Times New Roman' })]
          }),
          new Paragraph({
            alignment: AlignmentType.CENTER,
            spacing: { before: 20, after: 100 },
            children: [new TextRun({ text: 'Valeur en euros — hors honoraires — hors droits de mutation', color: C.NAVY, size: 18, font: 'Times New Roman', italics: true })]
          })
        ], { borders: noBorders() })]
      })]
    }),
  ];
}

function buildConclusionSection(sections) {
  return [
    pageBreak(),
    navyBanner('CONCLUSION'),
    spacer(120),
    ...splitParagraphs(sections.conclusion || '[à rajouter par l\'expert]'),
    spacer(200),
    // Signature
    new Table({
      width: { size: 50, type: WidthType.PERCENTAGE },
      borders: noBorders(),
      rows: [new TableRow({
        children: [new TableCell({
          borders: { top: { style: BorderStyle.SINGLE, size: 4, color: C.NAVY } },
          margins: { top: 120, bottom: 40, left: 0, right: 0 },
          children: [
            new Paragraph({ children: [new TextRun({ text: 'L\'Expert signataire', bold: true, size: 20, font: 'Times New Roman', color: C.NAVY })], spacing: { before: 60, after: 40 } }),
            new Paragraph({ children: [new TextRun({ text: '[à rajouter par l\'expert] — Nom, qualité, signature, cachet', size: 18, font: 'Times New Roman', color: C.AMBER })], spacing: { before: 40, after: 40 } }),
          ]
        })]
      })]
    }),
  ];
}

function buildPhotosSection() {
  return [
    pageBreak(),
    navyBanner('PHOTOGRAPHIES COMPLÉMENTAIRES'),
    spacer(120),
    new Paragraph({
      children: [new TextRun({
        text: 'Les photographies illustrant chaque section descriptive sont intégrées directement dans les chapitres correspondants du présent rapport (terrain, construction, état des lieux).',
        size: 19, font: 'Times New Roman', italics: true, color: C.DARK
      })],
      alignment: AlignmentType.JUSTIFIED,
      spacing: { before: 60, after: 120 }
    }),
    imagePlaceholder('[à rajouter par l\'expert] — Photographies complémentaires ou de contexte'),
  ];
}

function buildAnnexesSection() {
  return [
    pageBreak(),
    navyBanner('ANNEXES'),
    spacer(120),
    imagePlaceholder('[à rajouter par l\'expert] — Titre de propriété'),
    spacer(100),
    imagePlaceholder('[à rajouter par l\'expert] — Extrait cadastral'),
    spacer(100),
    imagePlaceholder('[à rajouter par l\'expert] — Certificat d\'urbanisme'),
    spacer(100),
    imagePlaceholder('[à rajouter par l\'expert] — DPE et Diagnostics techniques'),
    spacer(100),
    imagePlaceholder('[à rajouter par l\'expert] — Tout autre document utile à l\'expertise'),
  ];
}

function buildGlossaireSection() {
  const termes = [
    ['Valeur vénale', 'Prix auquel un bien pourrait être vendu dans des conditions normales de marché, entre un vendeur et un acquéreur consentants, disposant d\'une information complète et agissant librement.'],
    ['Surface habitable (Loi Boutin)', 'Surface de plancher construite, après déduction des surfaces occupées par les murs, cloisons, marches et cages d\'escalier, gaines, embrasures de portes et fenêtres, parties de locaux d\'une hauteur inférieure à 1,80 m.'],
    ['TEGOVA', 'The European Group of Valuers\' Associations — organisation européenne des associations d\'experts immobiliers. Le référentiel TEGOVA (6e édition) définit les standards européens de l\'expertise immobilière.'],
    ['Charte de l\'Expertise', 'Charte de l\'Expertise en Évaluation Immobilière (5e édition) — document de référence français définissant les règles déontologiques et méthodologiques applicables aux experts immobiliers.'],
    ['Coefficient de pondération', 'Coefficient appliqué à la surface brute d\'un espace selon sa nature et son usage, permettant de calculer une surface pondérée pour les besoins de l\'évaluation.'],
    ['Terme de comparaison', 'Bien similaire vendu ou mis en vente récemment, retenu comme référence pour calibrer la valeur du bien expertisé par la méthode comparative.'],
    ['Servitude', 'Charge imposée sur un immeuble (fonds servant) pour l\'utilité d\'un autre immeuble (fonds dominant) appartenant à un propriétaire différent (servitude de passage, de vue, etc.).'],
    ['PLU', 'Plan Local d\'Urbanisme — document d\'urbanisme définissant les règles d\'utilisation des sols sur le territoire communal (zonage, coefficients d\'emprise, hauteurs, destinations autorisées).'],
  ];

  return [
    pageBreak(),
    navyBanner('GLOSSAIRE'),
    spacer(120),
    new Table({
      width: { size: 100, type: WidthType.PERCENTAGE },
      borders: noBorders(),
      rows: termes.map(([terme, def], i) =>
        new TableRow({
          children: [
            shadedCell(i % 2 === 0 ? C.NAVY_L : C.GRAY, [
              new Paragraph({ children: [new TextRun({ text: terme, bold: true, size: 19, font: 'Times New Roman', color: C.NAVY })], spacing: { before: 60, after: 60 } })
            ], { width: { size: 28, type: WidthType.PERCENTAGE } }),
            shadedCell(C.WHITE, [
              new Paragraph({ children: [new TextRun({ text: def, size: 18, font: 'Times New Roman', color: C.DARK })], spacing: { before: 60, after: 60 }, alignment: AlignmentType.JUSTIFIED })
            ], { width: { size: 72, type: WidthType.PERCENTAGE } }),
          ]
        })
      )
    }),
  ];
}

// Découpe un texte multiligne en paragraphes docx
function splitParagraphs(text) {
  if (!text) return [bodyPara('[à rajouter par l\'expert]')];
  return text.split(/\n+/).filter(l => l.trim()).map(line => bodyPara(line));
}

async function generateJaltaDocx(sections, formData, photos64 = {}, logo = null) {
  const fd = formData || {};
  const p64 = photos64 || {};

  const children = [
    // Page de couverture (avec logo si disponible)
    ...buildCoverPage(fd, logo),
    // Sommaire
    ...buildSommaire(),
    // I — Résumé
    ...buildResumeSection(sections, fd),
    // II — Expertise détaillée
    ...buildExpertiseDetaillee(sections, fd),
    // III — Situation
    ...buildSituationSection(sections),
    // IV — Description (terrain + bâti + surfaces + désordres + photos intégrées)
    ...buildDescriptionSection(sections, fd, p64),
    // V — Éléments de jugement
    ...buildJugementSection(sections),
    // VI — Évaluation
    ...buildEvaluationSection(fd),
    // Conclusion
    ...buildConclusionSection(sections),
    // Photographies complémentaires
    ...buildPhotosSection(),
    // Annexes
    ...buildAnnexesSection(),
    // Glossaire
    ...buildGlossaireSection(),
  ];

  const doc = new Document({
    creator: 'ExpertIA — Cabinet JALTA',
    title: `Rapport d'expertise — ${fd.ref_dossier || 'Dossier'}`,
    description: `Rapport d'expertise immobilière — ${fd.adresse_bien || ''}`,
    sections: [{
      properties: {
        page: {
          margin: { top: convertInchesToTwip(1), right: convertInchesToTwip(0.9), bottom: convertInchesToTwip(1), left: convertInchesToTwip(0.9) }
        },
        pageNumberStart: 1,
        pageNumberFormatType: NumberFormat.DECIMAL,
      },
      headers: {
        default: new Header({
          children: [
            new Table({
              width: { size: 100, type: WidthType.PERCENTAGE },
              borders: noBorders(),
              rows: [new TableRow({
                children: [
                  // Logo du cabinet (si disponible)
                  ...(logo && logo.data ? [new TableCell({
                    borders: noBorders(),
                    width: { size: 15, type: WidthType.PERCENTAGE },
                    margins: { top: 0, bottom: 0, left: 0, right: 100 },
                    children: (() => {
                      const logoRun = buildImageRun(logo, { width: 100, height: 38 });
                      return logoRun
                        ? [new Paragraph({ children: [logoRun], alignment: AlignmentType.LEFT })]
                        : [new Paragraph({ children: [] })];
                    })()
                  })] : []),
                  new TableCell({
                    borders: noBorders(),
                    width: { size: logo && logo.data ? 60 : 75, type: WidthType.PERCENTAGE },
                    children: [new Paragraph({
                      children: [new TextRun({ text: 'RAPPORT D\'EXPERTISE IMMOBILIÈRE — CONFIDENTIEL', size: 16, color: C.NAVY, font: 'Times New Roman' })],
                      border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: C.NAVY } }
                    })]
                  }),
                  new TableCell({
                    borders: noBorders(),
                    width: { size: 25, type: WidthType.PERCENTAGE },
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
      footers: {
        default: new Footer({
          children: [
            new Paragraph({
              alignment: AlignmentType.CENTER,
              border: { top: { style: BorderStyle.SINGLE, size: 4, color: C.GRAY_MED } },
              spacing: { before: 60 },
              children: [
                new TextRun({ text: 'Page ', size: 16, color: C.DARK, font: 'Times New Roman' }),
                new TextRun({ children: [PageNumber.CURRENT], size: 16, color: C.DARK, font: 'Times New Roman' }),
                new TextRun({ text: ' — Document confidentiel — ', size: 16, color: C.DARK, font: 'Times New Roman' }),
                new TextRun({ text: fd.adresse_bien || '', size: 16, color: C.NAVY, font: 'Times New Roman', bold: true }),
              ]
            })
          ]
        })
      },
      children
    }]
  });

  return Packer.toBuffer(doc);
}

// ─────────────────────────────────────────────────────────────────────────────
// GÉNÉRATION DOCX MARKDOWN (fallback)
// ─────────────────────────────────────────────────────────────────────────────

async function generateDocx(markdown) {
  const children = parseMarkdown(markdown);
  const doc = new Document({
    styles: {
      paragraphStyles: [{
        id: 'expertTitle',
        name: 'Expert Title',
        basedOn: 'Normal',
        run: { font: 'Times New Roman', size: 36, bold: true, color: '1a2f4e' },
        paragraph: { spacing: { after: 200 }, alignment: AlignmentType.CENTER }
      }]
    },
    sections: [{
      properties: {
        page: { margin: { top: 1440, right: 1134, bottom: 1440, left: 1134 } }
      },
      headers: {
        default: new Header({
          children: [new Paragraph({
            children: [new TextRun({ text: 'RAPPORT D\'EXPERTISE IMMOBILIÈRE — CONFIDENTIEL', size: 16, color: '6b6457' })],
            border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: '9a7c38' } }
          })]
        })
      },
      children
    }]
  });
  return Packer.toBuffer(doc);
}

function parseMarkdown(md) {
  const lines = md.split('\n');
  const result = [];
  let tableBuffer = [];
  let inTable = false;

  const flushTable = () => {
    if (tableBuffer.length === 0) return;
    try {
      const rows = tableBuffer
        .filter(r => !r.match(/^\|[-\s|]+\|$/))
        .map(r => r.split('|').filter((_, i, a) => i > 0 && i < a.length - 1).map(c => c.trim()));
      if (rows.length === 0) { tableBuffer = []; inTable = false; return; }
      const tableRows = rows.map((cells, ri) =>
        new TableRow({
          children: cells.map(cell =>
            new TableCell({
              children: [new Paragraph({ children: [new TextRun({ text: cell.replace(/\*\*/g, ''), bold: ri === 0 || cell.startsWith('**'), size: 18 })] })],
              shading: ri === 0 ? { fill: 'E8E4DC', type: ShadingType.CLEAR } : undefined
            })
          )
        })
      );
      result.push(new Table({ rows: tableRows, width: { size: 100, type: WidthType.PERCENTAGE } }));
    } catch {}
    tableBuffer = [];
    inTable = false;
  };

  for (const raw of lines) {
    const line = raw.trimEnd();
    if (line.startsWith('|')) { inTable = true; tableBuffer.push(line); continue; }
    if (inTable) flushTable();
    if (line.startsWith('# ')) {
      result.push(new Paragraph({ text: line.slice(2).trim(), heading: HeadingLevel.TITLE, alignment: AlignmentType.CENTER }));
    } else if (line.startsWith('## ')) {
      result.push(new Paragraph({ text: line.slice(3).trim(), heading: HeadingLevel.HEADING_1 }));
    } else if (line.startsWith('### ')) {
      result.push(new Paragraph({ text: line.slice(4).trim(), heading: HeadingLevel.HEADING_2 }));
    } else if (line === '---') {
      result.push(new Paragraph({ text: '', spacing: { after: 120 } }));
    } else if (line.trim() === '') {
      result.push(new Paragraph({ text: '' }));
    } else {
      result.push(new Paragraph({ children: parseInline(line), spacing: { after: 80, line: 276 } }));
    }
  }
  if (inTable) flushTable();
  return result;
}

function parseInline(text) {
  const runs = [];
  const pattern = /(\*\*[^*]+\*\*|\*[^*]+\*)/g;
  let lastIndex = 0, match;
  while ((match = pattern.exec(text)) !== null) {
    if (match.index > lastIndex) runs.push(new TextRun({ text: text.slice(lastIndex, match.index), size: 22 }));
    const inner = match[0];
    if (inner.startsWith('**')) runs.push(new TextRun({ text: inner.slice(2, -2), bold: true, size: 22 }));
    else runs.push(new TextRun({ text: inner.slice(1, -1), italics: true, size: 22, color: '6b6457' }));
    lastIndex = match.index + match[0].length;
  }
  if (lastIndex < text.length) runs.push(new TextRun({ text: text.slice(lastIndex), size: 22 }));
  return runs.length ? runs : [new TextRun({ text, size: 22 })];
}

// ─────────────────────────────────────────────────────────────────────────────
const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`ExpertIA → http://localhost:${PORT}`));
