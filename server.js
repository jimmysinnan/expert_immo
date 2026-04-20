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
  return `Tu es un expert immobilier certifié JALTA en Martinique. Rédige deux sections distinctes pour un rapport d'expertise du bien situé à :

${adresse}

---
SECTION 1 — SITUATION GÉOGRAPHIQUE (clé JSON : "situation_geographique")
Rédige 2 à 3 paragraphes sobres et factuels :

Paragraphe 1 — La commune : situer dans la Martinique (Collectivité Territoriale de Martinique — CTM — depuis 2015, île française des Antilles). Mentionner le rang de la commune par population si connu (avec chiffre approximatif), sa situation géographique dans l'île (nord/sud/est/ouest, distance de Fort-de-France), son caractère général. Ne jamais écrire "département" ni "collectivité d'outre-mer régie par l'article 73".

Paragraphe 2 — Localisation du bien : quartier ou secteur, tissu bâti, desserte de proximité (commerces, services). Aucune référence à des routes spécifiques si l'information n'est pas certaine.

Paragraphe 3 — Accessibilité (seulement si certaine) : axes routiers connus, 2 lignes max. Si incertain, omettre.

SECTION 2 — ENVIRONNEMENT ÉCONOMIQUE (clé JSON : "marche_immobilier")
Rédige une analyse du marché immobilier local de la commune et du secteur concerné.
- Maximum 8 lignes
- UNIQUEMENT descriptif — aucun chiffre, aucun prix au m²
- Style factuel : tendances de la demande, profil des acquéreurs, attractivité du secteur, dynamiques récentes
- Entrée type : "Le marché immobilier de la commune de... se caractérise par..."

RÈGLES ABSOLUES :
- Style impersonnel, troisième personne, indicatif présent
- Si une donnée est inconnue, l'omettre — ne jamais inventer
- Retourner UNIQUEMENT un JSON valide avec deux clés : {"situation_geographique": "...", "marche_immobilier": "..."}`;
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
${(formData.type_bien === 'Appartement' || formData.type_bien === 'Immeuble') ? `
=== COPROPRIÉTÉ ===
${formData.appart_etage ? `Étage : ${formData.appart_etage}` : ''}
${formData.appart_type_pieces ? `Type / Pièces : ${formData.appart_type_pieces}` : ''}
Type de copropriété : ${formData.copro_type || '[à compléter]'}
Nombre de bâtiments : ${formData.copro_nb_batiments || '[à compléter]'}
Composition des bâtiments : ${formData.copro_composition || '[à compléter]'}
Tantièmes parties communes générales détenus : ${formData.copro_tantiemes || '[à compléter]'}
` : ''}

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

=== SECTION SITUATION GÉOGRAPHIQUE UNIQUEMENT (déjà rédigée — copier telle quelle, SANS le marché immobilier) ===
${chapter1}

IMPORTANT : Le texte ci-dessus contient deux blocs. Le bloc SITUATION GÉOGRAPHIQUE va dans la clé "situation_geographique". Le bloc MARCHÉ IMMOBILIER / ENVIRONNEMENT ÉCONOMIQUE va UNIQUEMENT dans la clé "marche_immobilier". Ne jamais mélanger les deux.

---

GÉNÈRE UN JSON avec exactement ces clés (UNIQUEMENT le JSON, sans markdown ni texte avant/après) :

{
  "resume_mission": "UNE SEULE phrase courte de synthèse (max 20 mots) rappelant l'objet et la localisation du bien — sans 'À la requête de', sans répéter les données du tableau.",
  "cadre_evaluation": "Commencer OBLIGATOIREMENT par le paragraphe de mission rédigé ainsi : 'À la requête de [nom donneur ordre], Nous, CABINET JALTA, avons reçu mission de déterminer la Valeur Vénale de [type et usage du bien], situé [adresse], référencé sous le dossier [réf dossier]. Après avoir visité les lieux le [date visite], en présence de notre mandant(e), relevé leur état, recueilli les renseignements nécessaires, nous avons établi le présent rapport.' Puis enchaîner avec le cadre normatif : normes TEGOVA et Charte appliquées, conditions et limites de la mission, absence de sondages destructifs, observations visuelles au conditionnel — 5 à 7 phrases au total. NE PAS mentionner le PPR. NE PAS inclure de définition de la valeur vénale.",
  "objectif_evaluation": "Texte de l'objectif de l'évaluation : nature de la mission (vénale, locative, etc.), finalité (vente, garantie, fiscalité...) — 2 à 4 phrases style JALTA. NE PAS inclure la définition 'soit le prix auquel ce bien pourrait raisonnablement être cédé...' — cette définition figure au glossaire.",
  "situation_geographique": "Copier EXACTEMENT le bloc SITUATION GÉOGRAPHIQUE fourni ci-dessus, sans aucune modification ni ajout. NE PAS inclure d'analyse de marché immobilier, de prix au m², de transactions DVF, ni d'environnement économique — ces éléments figurent UNIQUEMENT dans la clé 'marche_immobilier'.",
  "situation_urbanistique": "Texte SITUATION URBANISTIQUE — INTÉGRER OBLIGATOIREMENT le zonage PLU '${formData.zonage_plu || '[zonage non renseigné]'}' dans la première phrase. Exemple : 'Au regard du Plan Local d\\'Urbanisme en vigueur, le bien est classé en zone ${formData.zonage_plu || '[à compléter]'}...'. Décrire les règles d\\'urbanisme applicables (destination, COS, hauteur, prospect). NE PAS mentionner l\\'assainissement ni les servitudes ici — ces éléments figurent dans la description du terrain — 2 à 3 phrases style JALTA.",
  "situation_juridique": "Texte SITUATION JURIDIQUE — INTÉGRER OBLIGATOIREMENT la référence cadastrale '${formData.refs_cadastrales || '[référence à compléter]'}' dans le texte. Exemple d'ouverture : 'Le bien est cadastré sous la référence ${formData.refs_cadastrales || '[à compléter]'}...'. Mentionner le régime juridique (${formData.regime_juridique || '[à compléter]'}), la superficie du terrain (${formData.superficie_terrain || '[à compléter]'} m²), les mentions hypothécaires si connues — 3 à 5 phrases style JALTA.",
  "situation_locative_text": "Texte SITUATION LOCATIVE : si libre d'occupation ou occupé, conditions de l'occupation, incidence sur la valeur — 2 à 4 phrases. Si libre : le préciser clairement.",
  "description_terrain": "Texte section LE TERRAIN D\\'ASSIETTE — au moins 150 mots — style JALTA factuel. Inclure OBLIGATOIREMENT : surface (${formData.superficie_terrain || '[à compléter]'} m²), forme (${formData.forme_terrain || '[à compléter]'}), topographie, accès, clôtures, réseaux, zonage PLU (${formData.zonage_plu || '[à compléter]'}). Inclure OBLIGATOIREMENT la phrase sur l\\'assainissement : 'L\\'assainissement du bien est assuré par ${formData.assainissement || '[à compléter]'}.' Si servitude : mentionner. NE PAS mentionner le PPR. NE PAS écrire 'à l\\'examen visuel des photographies'.",
  "description_bati": "Texte section LA CONSTRUCTION — au moins 200 mots — style JALTA : 'Il s'agit d'un bâtiment en dur...', structure, toiture, façades, menuiseries, équipements (électricité, plomberie, chauffage), DPE. NE PAS mentionner les surfaces des pièces dans ce texte — elles figurent dans le tableau. NE PAS écrire 'à l'examen visuel des photographies'. Décrire uniquement la distribution fonctionnelle (nombre de niveaux, pièces principales) sans détailler les m².${(formData.type_bien === 'Appartement') ? ` Pour l'appartement : mentionner l'étage (${formData.appart_etage || '[à compléter]'}), le type (${formData.appart_type_pieces || '[à compléter]'}), et la copropriété (${formData.copro_type || '[à compléter]'}, ${formData.copro_nb_batiments || '[à compléter]'} bâtiment(s), tantièmes : ${formData.copro_tantiemes || '[à compléter]'}).` : (formData.type_bien === 'Immeuble') ? ` Pour l'immeuble : décrire la copropriété (${formData.copro_type || '[à compléter]'}), composition : ${formData.copro_composition || '[à compléter]'}, tantièmes : ${formData.copro_tantiemes || '[à compléter]'}.` : ''}",
  "desordres_texte": "Texte section ÉTAT DES LIEUX — liste tous les désordres constatés en style JALTA avec conditionnel. NE PAS écrire 'à l'examen visuel des photographies' — formuler directement les observations. Si aucun désordre : 'Au jour de notre visite, aucun désordre significatif n'a été constaté.'",
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

    const raw = response.content
      .filter(b => b.type === 'text')
      .map(b => b.text)
      .join('\n')
      .trim();

    // Le prompt retourne maintenant un JSON avec situation_geographique + marche_immobilier
    let situation_geographique = raw;
    let marche_immobilier = '';
    try {
      const parsed = JSON.parse(raw.replace(/^```json\n?/, '').replace(/\n?```$/, ''));
      situation_geographique = parsed.situation_geographique || raw;
      marche_immobilier = parsed.marche_immobilier || '';
    } catch (e) { /* ancien format texte brut — garder tel quel */ }

    res.json({ text: situation_geographique, marche_immobilier });
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
        max_tokens: 2000,
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
    const { report, sections, formData, refDossier, photos64, logo, photoResults } = req.body;
    let buffer;

    if (sections && formData) {
      buffer = await generateJaltaDocx(sections, formData, photos64 || {}, logo || null, photoResults || {});
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

**Référence dossier :** ${fd.ref_dossier || '[à rajouter par l\'expert]'}
**Date de visite :** ${formatDateFR(fd.date_visite)}
**Adresse :** ${fd.adresse_bien || '[à rajouter par l\'expert]'}
**Nature de la mission :** ${fd.type_mission || '[à rajouter par l\'expert]'}
**Donneur d'ordre :** ${fd.nom_donneur_ordre || ''} (${fd.donneur_ordre || ''})

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

function buildCoverPage(formData, logo, photos64 = {}) {
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

  // Bandeau titre principal
  items.push(
    new Table({
      width: { size: 100, type: WidthType.PERCENTAGE },
      borders: noBorders(),
      rows: [new TableRow({
        children: [shadedCell(C.NAVY, [
          new Paragraph({
            alignment: AlignmentType.CENTER,
            spacing: { before: 200, after: 80 },
            children: [new TextRun({ text: 'RAPPORT D\'EXPERTISE IMMOBILIÈRE', bold: true, color: C.WHITE, size: 36, font: 'Times New Roman' })]
          }),
          new Paragraph({
            alignment: AlignmentType.CENTER,
            spacing: { before: 40, after: 200 },
            children: [new TextRun({ text: '[à rajouter par l\'expert] — ex : ENSEMBLE IMMOBILIER', color: C.AMBER, size: 22, font: 'Times New Roman', italics: true })]
          })
        ], { borders: noBorders() })]
      })]
    }),
    spacer(200)
  );

  // Photo du bien (première photo extérieure si disponible)
  const coverPhotos = photos64.ext && photos64.ext.length ? [photos64.ext[0]] : [];
  if (coverPhotos.length) {
    const imgRun = buildImageRun(coverPhotos[0], { width: 460, height: 300 });
    if (imgRun) {
      items.push(new Paragraph({ children: [imgRun], alignment: AlignmentType.CENTER, spacing: { before: 60, after: 200 } }));
    }
  } else {
    items.push(imagePlaceholder('[à rajouter par l\'expert] — Photo du bien'));
    items.push(spacer(200));
  }

  // Tableau identité dossier — ordre : Référence dossier, Adresse du bien, Donneur d'ordre, Références cadastrales
  const donneurOrdre = `${fd.nom_donneur_ordre || ''} — ${fd.donneur_ordre || ''}`.replace(/^ — | — $/, '') || '[à rajouter par l\'expert]';
  items.push(
    new Table({
      width: { size: 100, type: WidthType.PERCENTAGE },
      borders: noBorders(),
      rows: [
        buildCoverRow('Référence dossier', fd.ref_dossier || '[à rajouter par l\'expert]'),
        buildCoverRow('Adresse du bien', fd.adresse_bien || '[à rajouter par l\'expert]'),
        buildCoverRow('Donneur d\'ordre', donneurOrdre),
        buildCoverRow('Références cadastrales', fd.refs_cadastrales || '[à rajouter par l\'expert]'),
      ]
    }),
    spacer(300)
  );

  // Pied de page couverture — adresse Cabinet JALTA
  items.push(
    new Table({
      width: { size: 100, type: WidthType.PERCENTAGE },
      borders: noBorders(),
      rows: [new TableRow({
        children: [new TableCell({
          borders: { top: { style: BorderStyle.SINGLE, size: 4, color: C.NAVY } },
          margins: { top: 120, bottom: 60, left: 0, right: 0 },
          children: [
            new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 60, after: 20 }, children: [new TextRun({ text: '09 Lotissement Bardinet Dillon – Route de Chateauboeuf – 97200 FORT DE FRANCE', size: 16, font: 'Times New Roman', color: C.DARK })] }),
            new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 20, after: 20 }, children: [new TextRun({ text: 'Tél. : 0596 75 08 90  -  Email : contact@cabinet-jalta.fr', size: 16, font: 'Times New Roman', color: C.DARK })] }),
            new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 20, after: 60 }, children: [new TextRun({ text: 'R.C.S. Fort-de-France 95B 488  -  N° SIRET : 402 038 285 000 19', size: 16, font: 'Times New Roman', color: C.DARK })] }),
          ]
        })]
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
              children: [new Paragraph({ children: [new TextRun({ text: '', size: 18, font: 'Times New Roman' })], alignment: AlignmentType.RIGHT, spacing: { before: 60, after: 60 } })]
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
    ['DATE DE L\'ÉVALUATION', formatDateFR(fd.date_visite)],
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

function parsePhotoResult(val) {
  if (!val || typeof val === 'object') return val || {};
  try {
    return JSON.parse(val.replace(/^```json\n?/, '').replace(/\n?```$/, ''));
  } catch { return {}; }
}

function buildMateriauxTable(fd, photoResults = {}) {
  const ext = parsePhotoResult(photoResults.ext);
  const int = parsePhotoResult(photoResults.int);

  // Fusion formulaire + observations IA photos (le formulaire prime si renseigné)
  const toiture_mat = fd.materiau_toiture || ext.toiture_materiau || '[à compléter]';
  const toiture_etat = fd.etat_toiture || ext.toiture_etat || '[à compléter]';
  const toiture_forme = fd.forme_toiture || '';
  const facades_mat = fd.materiau_facades || ext.facades_materiau || '[à compléter]';
  const facades_etat = fd.etat_facades || ext.facades_etat || '[à compléter]';
  const menus_mat = fd.menuiseries_ext || ext.menuiseries_materiau || '[à compléter]';
  const sols_mat = fd.sols_interieurs || int.sols_type || '[à rajouter par l\'expert]';
  const murs_mat = fd.revetements_murs || int.murs_revetement || '[à rajouter par l\'expert]';
  const elec_etat = fd.etat_electrique || int.electricite_obs || '[à compléter]';

  const corps = [
    ['Gros œuvre / Structure', fd.type_construction || '[à compléter]', '—'],
    ['Toiture / Couverture', `${toiture_mat}${toiture_forme ? ' — ' + toiture_forme : ''}`, toiture_etat],
    ['Façades / Enduits', facades_mat, facades_etat],
    ['Menuiseries extérieures', menus_mat, '—'],
    ['Revêtements de sols', sols_mat, '—'],
    ['Revêtements muraux', murs_mat, '—'],
    ['Plomberie / Sanitaires', '—', fd.etat_plomberie || '[à compléter]'],
    ['Installation électrique', '—', elec_etat],
    ['Chauffage / Climatisation', fd.chauffage || 'Néant (climat tropical)', '—'],
  ];
  return new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    borders: noBorders(),
    rows: [
      new TableRow({
        children: [
          shadedCell(C.NAVY, [new Paragraph({ children: [new TextRun({ text: 'Corps d\'état', bold: true, size: 19, color: C.WHITE, font: 'Times New Roman' })], spacing: { before: 40, after: 40 } })], { width: { size: 35, type: WidthType.PERCENTAGE } }),
          shadedCell(C.NAVY, [new Paragraph({ children: [new TextRun({ text: 'Matériaux / Description', bold: true, size: 19, color: C.WHITE, font: 'Times New Roman' })], spacing: { before: 40, after: 40 } })], { width: { size: 40, type: WidthType.PERCENTAGE } }),
          shadedCell(C.NAVY, [new Paragraph({ children: [new TextRun({ text: 'État', bold: true, size: 19, color: C.WHITE, font: 'Times New Roman' })], spacing: { before: 40, after: 40 } })], { width: { size: 25, type: WidthType.PERCENTAGE } }),
        ]
      }),
      ...corps.map(([label, desc, etat], i) => new TableRow({
        children: [
          shadedCell(i % 2 === 0 ? C.NAVY_L : C.GRAY, [new Paragraph({ children: [new TextRun({ text: label, bold: true, size: 19, color: C.NAVY, font: 'Times New Roman' })], spacing: { before: 40, after: 40 } })], { width: { size: 35, type: WidthType.PERCENTAGE } }),
          shadedCell(C.WHITE, [new Paragraph({ children: [new TextRun({ text: desc, size: 19, color: desc.includes('[à') ? C.AMBER : C.DARK, font: 'Times New Roman' })], spacing: { before: 40, after: 40 } })], { width: { size: 40, type: WidthType.PERCENTAGE }, borders: cellBorder(C.GRAY_MED) }),
          shadedCell(C.WHITE, [new Paragraph({ children: [new TextRun({ text: etat, size: 19, color: etat.includes('[à') ? C.AMBER : C.DARK, font: 'Times New Roman' })], spacing: { before: 40, after: 40 } })], { width: { size: 25, type: WidthType.PERCENTAGE }, borders: cellBorder(C.GRAY_MED) }),
        ]
      }))
    ]
  });
}

function buildDocumentsSection(documentsList) {
  const docs = (documentsList || '').split(',').map(d => d.trim()).filter(Boolean);
  if (!docs.length) {
    return [imagePlaceholder('[à rajouter par l\'expert] — Liste des pièces et documents consultés')];
  }
  const rows = docs.map(doc =>
    new TableRow({
      children: [
        new TableCell({
          borders: cellBorder(C.GRAY_MED),
          margins: { top: 60, bottom: 60, left: 120, right: 120 },
          children: [new Paragraph({ children: [new TextRun({ text: '✓', bold: true, color: C.NAVY, size: 19, font: 'Times New Roman' })], spacing: { before: 30, after: 30 } })]
        }),
        new TableCell({
          borders: cellBorder(C.GRAY_MED),
          margins: { top: 60, bottom: 60, left: 120, right: 120 },
          children: [new Paragraph({ children: [new TextRun({ text: doc, size: 19, font: 'Times New Roman' })], spacing: { before: 30, after: 30 } })]
        })
      ]
    })
  );
  return [
    new Table({ width: { size: 80, type: WidthType.PERCENTAGE }, borders: noBorders(), rows }),
    spacer(60),
    bodyPara('[à rajouter par l\'expert] — Compléter si nécessaire', { color: C.AMBER }),
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
    bodyPara(`La présente évaluation a été réalisée à la date du ${formatDateFR(fd.date_visite)}.`),
    spacer(120),
    subBanner('4   VISITE ET DOCUMENTS MIS À DISPOSITION'),
    spacer(80),
    ...buildDocumentsSection(fd.documents_fournis),
    spacer(120),
    subBanner('5   CLAUSE DE CONFIDENTIALITÉ'),
    spacer(80),
    bodyPara('Le présent rapport est établi à la demande et à l\'usage exclusif du donneur d\'ordre. Il ne peut être communiqué à des tiers sans l\'accord écrit de l\'expert signataire. Toute reproduction partielle ou totale est interdite sans autorisation préalable. Les valeurs vénales et conclusions définitives feront l\'objet d\'une validation complémentaire par l\'expert signataire conformément à la Charte de l\'Expertise Immobilière (5e édition) et au référentiel TEGOVA (6e édition).'),
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
    // Espace cartes IGN + Géoportail
    new Table({
      width: { size: 100, type: WidthType.PERCENTAGE },
      borders: noBorders(),
      rows: [new TableRow({
        children: [
          new TableCell({
            borders: { top: { style: BorderStyle.SINGLE, size: 2, color: C.GRAY_MED }, bottom: { style: BorderStyle.SINGLE, size: 2, color: C.GRAY_MED }, left: { style: BorderStyle.SINGLE, size: 2, color: C.GRAY_MED }, right: { style: BorderStyle.SINGLE, size: 2, color: C.GRAY_MED } },
            margins: { top: 300, bottom: 300, left: 120, right: 60 },
            width: { size: 49, type: WidthType.PERCENTAGE },
            shading: { fill: C.GRAY, type: ShadingType.CLEAR },
            children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: '[à rajouter par l\'expert] — Carte IGN / Plan de situation', color: C.AMBER, size: 18, font: 'Times New Roman', bold: true })] })]
          }),
          new TableCell({
            borders: { top: { style: BorderStyle.SINGLE, size: 2, color: C.GRAY_MED }, bottom: { style: BorderStyle.SINGLE, size: 2, color: C.GRAY_MED }, left: { style: BorderStyle.SINGLE, size: 2, color: C.GRAY_MED }, right: { style: BorderStyle.SINGLE, size: 2, color: C.GRAY_MED } },
            margins: { top: 300, bottom: 300, left: 60, right: 120 },
            width: { size: 49, type: WidthType.PERCENTAGE },
            shading: { fill: C.GRAY, type: ShadingType.CLEAR },
            children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: '[à rajouter par l\'expert] — Vue aérienne Géoportail®', color: C.AMBER, size: 18, font: 'Times New Roman', bold: true })] })]
          })
        ]
      })]
    }),
    spacer(80),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 20, after: 80 }, children: [new TextRun({ text: 'Localisation de l\'ensemble immobilier concerné', size: 17, italics: true, font: 'Times New Roman', color: C.DARK })] }),
    spacer(80),
    // Environnement économique
    subBanner('1 bis   ENVIRONNEMENT ÉCONOMIQUE'),
    spacer(80),
    ...splitParagraphs(s.marche_immobilier || '[à rajouter par l\'expert] — Analyse du marché immobilier local'),
    spacer(150),
    subBanner('2   SITUATION URBANISTIQUE'),
    spacer(80),
    ...splitParagraphs(s.situation_urbanistique || '[à rajouter par l\'expert]'),
    spacer(120),
    // Espace PPR
    subBanner('Plan de Prévention des Risques Naturels (P.P.R.)'),
    spacer(80),
    bodyPara('[à rajouter par l\'expert] — Selon le zonage du Plan de Prévention des Risques Naturels (P.P.R.), approuvé en Décembre 2013, le terrain d\'assiette est classé en zone... (aléa...), subissant l\'application de prescriptions particulières.'),
    spacer(80),
    imagePlaceholder('[à rajouter par l\'expert] — Carte PPR / Extrait du plan de zonage'),
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

function buildDescriptionSection(sections, formData, p64 = {}, photoResults = {}) {
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
      new TableRow({
        children: [
          shadedCell(C.NAVY, [new Paragraph({ children: [new TextRun({ text: 'TOTAL SURFACE UTILE PONDÉRÉE', bold: true, size: 19, font: 'Times New Roman', color: C.WHITE })], spacing: { before: 60, after: 60 } })], { columnSpan: 3 }),
          shadedCell(C.NAVY, [new Paragraph({ children: [new TextRun({ text: (totalHab + totalAnnexes) > 0 ? `${(totalHab + totalAnnexes).toFixed(2)} m²` : '[à rajouter par l\'expert]', bold: true, size: 19, font: 'Times New Roman', color: C.WHITE })], alignment: AlignmentType.RIGHT, spacing: { before: 60, after: 60 } })], { columnSpan: 2 }),
        ]
      }),
    ]
  });

  // Intro composition du bien
  const niveaux = fd.nb_niveaux ? `élevé à ${fd.nb_niveaux}` : '';
  const introLines = [
    'L\'ensemble immobilier, objet du présent rapport, est composé de :',
  ];

  return [
    pageBreak(),
    navyBanner('IV/ DESCRIPTION DU BIEN'),
    spacer(120),
    bodyPara(introLines[0]),
    new Paragraph({ children: [new TextRun({ text: `— d'un terrain d'une superficie de ${fd.superficie_terrain || '[à compléter]'} m²${fd.forme_terrain ? ', de forme ' + fd.forme_terrain : ''}, sur lequel sont édifiés :`, size: 20, font: 'Times New Roman' })], bullet: { level: 0 }, spacing: { before: 40, after: 20 } }),
    new Paragraph({ children: [new TextRun({ text: `— un (1) bâtiment à usage d'habitation ${niveaux}${fd.type_bien ? ' (' + fd.type_bien + ')' : ''},`, size: 20, font: 'Times New Roman' })], bullet: { level: 0 }, spacing: { before: 20, after: 20 } }),
    new Paragraph({ children: [new TextRun({ text: '— et des aménagements extérieurs (abords).', size: 20, font: 'Times New Roman' })], bullet: { level: 0 }, spacing: { before: 20, after: 80 } }),
    spacer(100),
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
    subBanner('DESCRIPTIF DES MATÉRIAUX PAR CORPS D\'ÉTAT'),
    spacer(80),
    buildMateriauxTable(fd, photoResults),
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

// Formate une date YYYY-MM-DD en DD/MM/YYYY
function formatDateFR(dateStr) {
  if (!dateStr) return '[à rajouter par l\'expert]';
  const m = dateStr.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (m) return `${m[3]}/${m[2]}/${m[1]}`;
  return dateStr;
}

// Découpe un texte multiligne en paragraphes docx
function splitParagraphs(text) {
  if (!text) return [bodyPara('[à rajouter par l\'expert]')];
  return text.split(/\n+/).filter(l => l.trim()).map(line => bodyPara(line));
}

async function generateJaltaDocx(sections, formData, photos64 = {}, logo = null, photoResults = {}) {
  const fd = formData || {};
  const p64 = photos64 || {};

  const children = [
    // Page de couverture (avec logo et photo de couverture si disponible)
    ...buildCoverPage(fd, logo, p64),
    // Sommaire
    ...buildSommaire(),
    // I — Résumé
    ...buildResumeSection(sections, fd),
    // II — Expertise détaillée
    ...buildExpertiseDetaillee(sections, fd),
    // III — Situation
    ...buildSituationSection(sections),
    // IV — Description (terrain + bâti + surfaces + désordres + photos intégrées)
    ...buildDescriptionSection(sections, fd, p64, photoResults),
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
                new TextRun({ text: 'Page n°', size: 16, color: C.DARK, font: 'Times New Roman' }),
                new TextRun({ children: [PageNumber.CURRENT], size: 16, color: C.DARK, font: 'Times New Roman' }),
                new TextRun({ text: ` — Dossier (${fd.nom_donneur_ordre || fd.donneur_ordre || 'Donneur d\'ordre'}) — `, size: 16, color: C.DARK, font: 'Times New Roman' }),
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
// FICHE DOSSIER — DEVIS DOCX
// ─────────────────────────────────────────────────────────────────────────────

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

async function generateDevisDocx(data) {
  const { num_dossier, suivi_par, ordonnateur, email_ord, objet,
          description_bien, lieu, honoraires_ttc, logo } = data;

  const signataire = suivi_par === 'CL' ? 'Claude LUCE' : 'Roméo VULCAIN';
  const montant = parseFloat(honoraires_ttc) || 0;
  const montantStr = montant.toFixed(2) + ' €';
  const montantLettres = amountToWordsFr(Math.round(montant));

  const MONTHS_FR = ['JANVIER','FÉVRIER','MARS','AVRIL','MAI','JUIN','JUILLET','AOÛT','SEPTEMBRE','OCTOBRE','NOVEMBRE','DÉCEMBRE'];
  const now = new Date();
  const dateStr = `Fort-de-France, le ${now.getDate()} ${MONTHS_FR[now.getMonth()]} ${now.getFullYear()}`;

  const genreMatch = (ordonnateur || '').match(/^(Madame|Mme|M\.|Monsieur)/i);
  const genre = genreMatch ? genreMatch[0] : 'Monsieur';

  const nb = { style: BorderStyle.NONE, size: 0, color: 'auto' };
  const noBorders = { top: nb, bottom: nb, left: nb, right: nb, insideH: nb, insideV: nb };

  // Taille de base : 17 = 8.5pt, 16 = 8pt, 15 = 7.5pt
  const S = 17;  // taille corps
  const SM = 15; // taille secondaire
  const p0 = { before: 0, after: 0 };
  const p1 = { before: 20, after: 20 };
  const p2 = { before: 30, after: 30 };

  const para = (children, opts = {}) => new Paragraph({
    children: Array.isArray(children) ? children : [new TextRun({ text: String(children), size: S, font: 'Calibri' })],
    spacing: p1,
    ...opts
  });

  const tr = (text, opts = {}) => new TextRun({ text, size: S, font: 'Calibri', ...opts });

  const sectionBar = (title) => new Paragraph({
    children: [new TextRun({ text: title, bold: true, color: C.WHITE, size: S, font: 'Calibri' })],
    shading: { fill: C.NAVY, type: ShadingType.SOLID },
    spacing: { before: 60, after: 20 },
    indent: { left: 80 }
  });

  const dashItem = (runs) => new Paragraph({
    children: Array.isArray(runs) ? runs : [tr('- ' + runs)],
    indent: { left: 360 },
    spacing: { before: 14, after: 14 }
  });

  // Logo cell
  const logoCell = [];
  if (logo?.data) {
    try {
      const buf = Buffer.from(logo.data, 'base64');
      const imgType = (logo.mimeType || '').includes('png') ? 'png' : 'jpg';
      logoCell.push(new Paragraph({
        children: [new ImageRun({ data: buf, transformation: { width: 70, height: 70 }, type: imgType })],
        spacing: p0
      }));
    } catch { logoCell.push(para('')); }
  } else {
    logoCell.push(
      new Paragraph({ children: [tr('CJ', { bold: true, size: 32, color: C.NAVY })], spacing: p0 }),
      new Paragraph({ children: [tr('CABINET JALTA', { size: S, color: C.NAVY })], spacing: p0 })
    );
  }

  const headerTable = new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    borders: noBorders,
    rows: [new TableRow({ children: [
      new TableCell({ width: { size: 20, type: WidthType.PERCENTAGE }, borders: noBorders, verticalAlign: 'center', children: logoCell }),
      new TableCell({ width: { size: 80, type: WidthType.PERCENTAGE }, borders: noBorders, children: [
        new Paragraph({
          children: [tr('EXPERTISES IMMOBILIERES', { bold: true, italics: true, size: 19, color: C.NAVY })],
          spacing: p0
        }),
        new Paragraph({
          children: [tr('Territoires : Martinique, Guadeloupe, Saint-Martin, Saint Barthélémy et Guyane Française', { italics: true, size: SM })],
          spacing: p0,
          border: { bottom: { style: BorderStyle.SINGLE, size: 3, color: C.GRAY_MED } }
        }),
        new Paragraph({ children: [tr(dateStr, { size: SM })], alignment: AlignmentType.RIGHT, spacing: { before: 24, after: 10 } }),
        new Paragraph({ children: [tr(ordonnateur || 'Monsieur', { bold: true })], alignment: AlignmentType.RIGHT, spacing: p0 }),
        ...(email_ord ? [new Paragraph({ children: [tr('Email. : ' + email_ord, { size: SM })], alignment: AlignmentType.RIGHT, spacing: p0 })] : []),
      ]})
    ]})]
  });

  const affaireBox = new Table({
    width: { size: 48, type: WidthType.PERCENTAGE },
    borders: {
      top: { style: BorderStyle.SINGLE, size: 6, color: C.BLACK },
      bottom: { style: BorderStyle.SINGLE, size: 6, color: C.BLACK },
      left: { style: BorderStyle.SINGLE, size: 6, color: C.BLACK },
      right: { style: BorderStyle.SINGLE, size: 6, color: C.BLACK },
      insideH: nb, insideV: nb
    },
    rows: [new TableRow({ children: [
      new TableCell({ borders: noBorders, margins: { top: 60, bottom: 60, left: 100, right: 100 }, children: [
        new Paragraph({ children: [tr('Affaire : ', { bold: true }), tr(num_dossier || '')], spacing: p0 }),
        new Paragraph({ children: [tr('Objet ', { bold: true, underline: {} }), tr(': proposition de services')], spacing: { before: 10, after: 0 } })
      ]})
    ]})]
  });

  const delaiTable = new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    borders: noBorders,
    rows: [
      new TableRow({ children: [
        new TableCell({ width: { size: 33, type: WidthType.PERCENTAGE }, borders: noBorders,
          children: [para('- Intervention sur site :', { spacing: p1 })] }),
        new TableCell({ width: { size: 67, type: WidthType.PERCENTAGE }, borders: noBorders,
          children: [para('A déterminer ultérieurement, après confirmation de la mission et versement de la provision', { spacing: p1 })] })
      ]}),
      new TableRow({ children: [
        new TableCell({ width: { size: 33, type: WidthType.PERCENTAGE }, borders: noBorders,
          children: [para('- Remise rapport :', { spacing: p1 })] }),
        new TableCell({ width: { size: 67, type: WidthType.PERCENTAGE }, borders: noBorders,
          children: [para('15 jours ouvrés environ, après la visite des lieux et versement de la totalité des honoraires', { spacing: p1 })] })
      ]})
    ]
  });

  // Footer intégré dans le contenu (ligne séparateur + adresse)
  const footerInline = [
    new Paragraph({
      children: [],
      border: { top: { style: BorderStyle.SINGLE, size: 4, color: C.GRAY_MED } },
      spacing: { before: 40, after: 10 }
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER, spacing: p0,
      children: [tr('Espace LAOUCHEZ – Boulevard N.MANDELA – 97200 FORT DE FRANCE  ·  Tél. : 0596 75 08 90  ·  contact@cabinet-jalta.fr', { size: SM })]
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER, spacing: p0,
      children: [tr('R.C.S. Fort-de-France 95B 488  –  N° SIRET : 402 038 285 000 19', { size: SM })]
    }),
  ];

  const children = [
    headerTable,
    new Paragraph({ children: [], spacing: { before: 60, after: 30 } }),
    affaireBox,
    new Paragraph({ children: [], spacing: { before: 60, after: 10 } }),
    para(genre + ','),
    para('Faisant suite à votre demande, nous avons l\'avantage de vous communiquer notre proposition pour la mission en objet.', { spacing: { before: 10, after: 30 } }),

    sectionBar('CONCERNE'),
    new Paragraph({ children: [tr(description_bien || `Un bien immobilier sis ${lieu || '[lieu à compléter]'}.`)], spacing: p1 }),

    sectionBar('MISSION'),
    new Paragraph({ children: [tr(`Détermination de la ${objet} du bien immobilier ci-dessus.`, { bold: true })], spacing: p1 }),

    sectionBar('DOCUMENTS A FOURNIR'),
    dashItem('Extraits documents cadastraux,'),
    dashItem([tr('- Extrait titre de propriété (partie désignation) ('), tr('s\'il est en votre possession', { italics: true }), tr('),')]),
    dashItem([tr('- Plans de distribution de la construction ('), tr('s\'ils existent', { italics: true }), tr('),')]),
    dashItem([tr('- Etat locatif actuel ('), tr('si loué', { italics: true }), tr(').')]),

    sectionBar('HONORAIRES*'),
    new Paragraph({
      alignment: AlignmentType.CENTER, spacing: p1,
      children: [
        tr(`${montantStr}  (${montantLettres} EUROS)  TTC*`, { bold: true, size: 19 }),
        tr('  ( TVA 8.5 % incluse)', { bold: true })
      ]
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER, spacing: p0,
      children: [tr('*Honoraires sous réserve de conformité des informations communiquées à l\'existant', { italics: true, size: SM })]
    }),

    sectionBar('CONDITIONS DE REGLEMENT'),
    new Paragraph({ children: [tr('- Provision : '), tr('50 %', { bold: true }), tr(', à la confirmation de la mission')], indent: { left: 360 }, spacing: p1 }),
    new Paragraph({ children: [tr('- Solde : à la remise du rapport')], indent: { left: 360 }, spacing: { before: 14, after: 20 } }),

    sectionBar('DELAI'),
    delaiTable,
    new Paragraph({ children: [], spacing: { before: 50, after: 10 } }),

    para('Restant à votre disposition,'),
    new Paragraph({ children: [tr(`Nous vous prions d'agréer, ${genre}, l'expression de nos sentiments dévoués.`)], spacing: { before: 10, after: 80 } }),
    new Paragraph({ children: [tr(signataire, { bold: true })], alignment: AlignmentType.RIGHT, spacing: p0 }),
    new Paragraph({ children: [tr('CABINET JALTA', { bold: true })], alignment: AlignmentType.RIGHT, spacing: { before: 0, after: 60 } }),
    new Paragraph({ children: [tr('BON POUR ACCORD,'), tr('          LE : _______________')], spacing: p0 }),

    ...footerInline
  ];

  const doc = new Document({
    sections: [{
      properties: {
        page: { margin: { top: convertInchesToTwip(0.45), right: convertInchesToTwip(0.75), bottom: convertInchesToTwip(0.35), left: convertInchesToTwip(0.75) } }
      },
      children
    }]
  });

  return Packer.toBuffer(doc);
}

app.post('/api/generate-devis', async (req, res) => {
  try {
    const buffer = await generateDevisDocx(req.body);
    const slug = (req.body.num_dossier || 'JALTA').replace(/[^a-zA-Z0-9\-_]/g, '_');
    const filename = `Devis_${slug}_${new Date().toISOString().slice(0, 10)}.docx`;
    res.set({
      'Content-Type': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
      'Content-Disposition': `attachment; filename="${filename}"`,
      'Content-Length': buffer.length
    });
    res.send(buffer);
  } catch (e) {
    console.error('Devis error:', e);
    res.status(500).json({ error: e.message });
  }
});

// ─────────────────────────────────────────────────────────────────────────────
// POST /api/generate-fiche — Fiche dossier (cover page JALTA + tableau récap)
// ─────────────────────────────────────────────────────────────────────────────
async function generateFicheDocx(data) {
  const { num_dossier, suivi_par, ordonnateur, email_ord, objet,
          lieu, description_bien, date_visite, date_butoir, logo } = data;

  const signataire = suivi_par === 'CL' ? 'Claude LUCE' : 'Roméo VULCAIN';
  const nb = { style: BorderStyle.NONE, size: 0, color: 'auto' };
  const noB = { top: nb, bottom: nb, left: nb, right: nb, insideH: nb, insideV: nb };

  // ── Helpers locaux ────────────────────────────────────────────────────────
  const p = (children, opts = {}) => new Paragraph({
    children: Array.isArray(children) ? children : [new TextRun({ text: String(children || ''), size: 20, font: 'Times New Roman' })],
    spacing: { before: 40, after: 40 },
    ...opts
  });
  const t = (text, opts = {}) => new TextRun({ text: String(text || ''), size: 20, font: 'Times New Roman', ...opts });

  const CELL_LABEL_BG = 'D6E8F7';  // bleu clair
  const CELL_VALUE_BG = 'F4F6F8';  // gris très clair
  const SEP = { style: BorderStyle.SINGLE, size: 2, color: 'C8D8E8' };
  const cellBorders = { top: nb, bottom: SEP, left: nb, right: nb, insideH: nb, insideV: nb };

  // ShadingType.CLEAR = fond coloré sans motif (compatible docx library)
  const shadeLabel = { fill: CELL_LABEL_BG, type: ShadingType.CLEAR, color: 'auto' };
  const shadeValue = { fill: CELL_VALUE_BG, type: ShadingType.CLEAR, color: 'auto' };

  const coverRow = (label, value) => new TableRow({
    children: [
      new TableCell({
        width: { size: 38, type: WidthType.PERCENTAGE },
        shading: shadeLabel,
        borders: cellBorders,
        margins: { top: 60, bottom: 60, left: 120, right: 120 },
        children: [p(label, { children: [t(label, { bold: true, color: '1A1A1A' })] })]
      }),
      new TableCell({
        width: { size: 62, type: WidthType.PERCENTAGE },
        shading: shadeValue,
        borders: cellBorders,
        margins: { top: 60, bottom: 60, left: 120, right: 120 },
        children: [p(value || '[à compléter]', { children: [t(value || '[à compléter]', { color: value ? '1A1A1A' : C.AMBER })] })]
      })
    ]
  });

  const infoRow = (label, value) => new TableRow({
    children: [
      new TableCell({
        width: { size: 38, type: WidthType.PERCENTAGE },
        shading: shadeLabel,
        borders: cellBorders,
        margins: { top: 70, bottom: 70, left: 120, right: 120 },
        children: [new Paragraph({ children: [new TextRun({ text: label, bold: true, size: 19, font: 'Times New Roman', color: '1A1A1A' })], spacing: { before: 40, after: 40 } })]
      }),
      new TableCell({
        width: { size: 62, type: WidthType.PERCENTAGE },
        shading: shadeValue,
        borders: cellBorders,
        margins: { top: 70, bottom: 70, left: 120, right: 120 },
        children: [new Paragraph({ children: [new TextRun({ text: value || '[à compléter]', size: 19, font: 'Times New Roman', color: value ? '1A1A1A' : C.AMBER })], spacing: { before: 40, after: 40 } })]
      })
    ]
  });

  // ── Page 1 : Cover ────────────────────────────────────────────────────────
  const coverItems = [];

  // Logo
  if (logo?.data) {
    try {
      const buf = Buffer.from(logo.data, 'base64');
      const imgType = (logo.mimeType || '').includes('png') ? 'png' : 'jpg';
      coverItems.push(new Paragraph({
        children: [new ImageRun({ data: buf, transformation: { width: 180, height: 70 }, type: imgType })],
        alignment: AlignmentType.LEFT,
        spacing: { before: 0, after: 200 }
      }));
    } catch { /* ignore */ }
  }

  // Bandeau navy "RAPPORT D'EXPERTISE IMMOBILIÈRE"
  coverItems.push(
    new Table({
      width: { size: 100, type: WidthType.PERCENTAGE },
      borders: noB,
      rows: [new TableRow({
        children: [new TableCell({
          shading: { fill: C.NAVY, type: ShadingType.SOLID },
          borders: noB,
          margins: { top: 160, bottom: 160, left: 200, right: 200 },
          children: [
            new Paragraph({
              alignment: AlignmentType.CENTER,
              spacing: { before: 0, after: 80 },
              children: [new TextRun({ text: 'RAPPORT D\'EXPERTISE IMMOBILIÈRE', bold: true, color: 'FFFFFF', size: 36, font: 'Times New Roman' })]
            }),
            new Paragraph({
              alignment: AlignmentType.CENTER,
              spacing: { before: 0, after: 0 },
              children: [new TextRun({ text: objet ? `(${objet})` : '[TYPE DE MISSION]', color: C.AMBER, size: 22, font: 'Times New Roman', italics: true })]
            })
          ]
        })]
      })]
    }),
    new Paragraph({ children: [], spacing: { before: 0, after: 200 } })
  );

  // Photo placeholder
  coverItems.push(
    new Table({
      width: { size: 100, type: WidthType.PERCENTAGE },
      borders: noB,
      rows: [new TableRow({
        children: [new TableCell({
          shading: { fill: 'EEF4FA', type: ShadingType.CLEAR, color: 'auto' },
          borders: { top: { style: BorderStyle.SINGLE, size: 2, color: 'C8D8E8' }, bottom: { style: BorderStyle.SINGLE, size: 2, color: 'C8D8E8' }, left: nb, right: nb },
          margins: { top: 200, bottom: 200, left: 0, right: 0 },
          children: [new Paragraph({
            alignment: AlignmentType.CENTER,
            children: [new TextRun({ text: '[PHOTO DU BIEN]', size: 22, font: 'Times New Roman', color: 'AABDCC', italics: true })]
          })]
        })]
      })]
    }),
    new Paragraph({ children: [], spacing: { before: 0, after: 240 } })
  );

  // Tableau d'identité dossier (4 lignes)
  coverItems.push(
    new Table({
      width: { size: 100, type: WidthType.PERCENTAGE },
      borders: noB,
      rows: [
        coverRow('Référence dossier', num_dossier),
        coverRow('Adresse du bien', lieu),
        coverRow('Donneur d\'ordre', ordonnateur),
        coverRow('Objet de la mission', objet),
      ]
    }),
    new Paragraph({ children: [], spacing: { before: 0, after: 300 } })
  );

  // Footer couverture — adresse JALTA
  coverItems.push(
    new Table({
      width: { size: 100, type: WidthType.PERCENTAGE },
      borders: noB,
      rows: [new TableRow({
        children: [new TableCell({
          borders: { top: { style: BorderStyle.SINGLE, size: 4, color: C.NAVY }, bottom: nb, left: nb, right: nb },
          margins: { top: 120, bottom: 60 },
          children: [
            new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 60, after: 20 }, children: [new TextRun({ text: '09 Lotissement Bardinet Dillon – Route de Chateauboeuf – 97200 FORT DE FRANCE', size: 16, font: 'Times New Roman', color: C.DARK })] }),
            new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 20, after: 20 }, children: [new TextRun({ text: 'Tél. : 0596 75 08 90  -  Email : contact@cabinet-jalta.fr', size: 16, font: 'Times New Roman', color: C.DARK })] }),
            new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 20, after: 60 }, children: [new TextRun({ text: 'R.C.S. Fort-de-France 95B 488  -  N° SIRET : 402 038 285 000 19', size: 16, font: 'Times New Roman', color: C.DARK })] }),
          ]
        })]
      })]
    })
  );

  // ── Page 2 : Tableau récapitulatif ────────────────────────────────────────
  const infoItems = [
    new Paragraph({ pageBreakBefore: true, children: [new TextRun('')] }),
    new Paragraph({
      children: [new TextRun({ text: 'RÉCAPITULATIF FICHE DOSSIER', bold: true, size: 26, font: 'Times New Roman', color: C.WHITE })],
      shading: { fill: C.NAVY, type: ShadingType.SOLID },
      alignment: AlignmentType.CENTER,
      spacing: { before: 0, after: 160 },
      indent: { left: 200, right: 200 }
    }),
    new Table({
      width: { size: 100, type: WidthType.PERCENTAGE },
      borders: noB,
      rows: [
        infoRow('N° dossier', num_dossier),
        infoRow('Affaire suivie par', signataire),
        infoRow('Ordonnateur', ordonnateur),
        infoRow('Email', email_ord),
        infoRow('Objet de la mission', objet),
        infoRow('Lieu du bien', lieu),
        infoRow('Date de visite', date_visite ? formatDateFR(date_visite) : ''),
        infoRow('Date butoir de remise', date_butoir ? formatDateFR(date_butoir) : ''),
        infoRow('Description du bien', description_bien),
      ]
    })
  ];

  const jaltaFooter = new Footer({
    children: [
      new Paragraph({
        alignment: AlignmentType.CENTER,
        border: { top: { style: BorderStyle.SINGLE, size: 4, color: C.GRAY_MED } },
        spacing: { before: 60, after: 20 },
        children: [new TextRun({ text: '09 Lotissement Bardinet Dillon – Route de Chateauboeuf – 97200 FORT DE FRANCE', size: 16, font: 'Times New Roman' })]
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER, spacing: { before: 0, after: 0 },
        children: [new TextRun({ text: 'Tél. : 0596 75 08 90  -  Email : contact@cabinet-jalta.fr', size: 16, font: 'Times New Roman' })]
      })
    ]
  });

  const doc = new Document({
    sections: [
      {
        properties: {
          page: { margin: { top: convertInchesToTwip(1.0), right: convertInchesToTwip(1.0), bottom: convertInchesToTwip(1.0), left: convertInchesToTwip(1.0) } }
        },
        footers: { default: jaltaFooter },
        children: [...coverItems, ...infoItems]
      }
    ]
  });

  return Packer.toBuffer(doc);
}

app.post('/api/generate-fiche', async (req, res) => {
  try {
    const buffer = await generateFicheDocx(req.body);
    const slug = (req.body.num_dossier || 'JALTA').replace(/[^a-zA-Z0-9\-_]/g, '_');
    const filename = `Fiche_${slug}_${new Date().toISOString().slice(0, 10)}.docx`;
    res.set({
      'Content-Type': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
      'Content-Disposition': `attachment; filename="${filename}"`,
      'Content-Length': buffer.length
    });
    res.send(buffer);
  } catch (e) {
    console.error('Fiche error:', e);
    res.status(500).json({ error: e.message });
  }
});

// ─────────────────────────────────────────────────────────────────────────────
const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`ExpertIA → http://localhost:${PORT}`));
