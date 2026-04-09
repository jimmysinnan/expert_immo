require('dotenv').config();
const express = require('express');
const multer  = require('multer');
const Anthropic = require('@anthropic-ai/sdk');
const mammoth  = require('mammoth');
const {
  Document, Packer, Paragraph, TextRun, HeadingLevel,
  Table, TableRow, TableCell, WidthType, AlignmentType,
  BorderStyle, Header, PageBreak, UnderlineType, ShadingType
} = require('docx');

const app    = express();
const upload = multer({ storage: multer.memoryStorage(), limits: { fileSize: 25 * 1024 * 1024 } });
const client = new Anthropic({ apiKey: process.env.ANTHROPIC_API_KEY });

app.use(express.json({ limit: '50mb' }));
app.use(express.static('public'));

// ─────────────────────────────────────────────────────────────────────────────
// PROMPTS
// ─────────────────────────────────────────────────────────────────────────────

const SYSTEM_PROMPT = `Tu es un expert immobilier certifié, spécialisé dans la rédaction de rapports d'expertise immobilière professionnels. Tu rédiges pour le compte d'un expert immobilier dont tu dois reproduire EXACTEMENT le style, le ton, le vocabulaire et la structure — tels qu'ils apparaissent dans les rapports de référence de sa base de connaissance.

Ta mission : générer les sections narratives d'un pré-rapport d'expertise à partir des données saisies dans le formulaire, des analyses de photos fournies, des données géographiques, et des rapports de référence.

RÈGLES ABSOLUES :
1. Reproduire EXACTEMENT le style des exemples fournis — jamais un style générique
2. Ne JAMAIS générer de valeurs vénales, prix, estimations ou calculs de valeur
3. Si une information manque → indiquer [À COMPLÉTER PAR L'EXPERT], jamais inventer
4. Toute observation issue des photos → conditionnel ou "à l'examen visuel"
5. Langue : français professionnel, vocabulaire technique immobilier
6. Longueur : adaptée au type de bien et à la richesse des données`;

function buildChapter1Prompt(adresse) {
  return `Tu es un expert immobilier certifié. Rédige le Chapitre 1 — Situation Géographique et Environnement d'un rapport d'expertise pour le bien situé à :

${adresse}

Recherche et inclus les données suivantes (INSEE, DVF, PLU, transports) et rédige en 5 sous-parties :

### 1.1 Présentation de la commune
[Commune, département, région, population INSEE, dynamisme économique, bassin d'emploi]

### 1.2 Situation dans la commune
[Quartier ou secteur, caractère résidentiel/commercial/mixte, standing, proximité centre-ville]

### 1.3 Accessibilité et transports
[Axes routiers, autoroutes, gare, transports en commun, temps trajets grandes villes]

### 1.4 Environnement immédiat
[Commerces de proximité, écoles, services, espaces verts, nuisances éventuelles]

### 1.5 Analyse du marché immobilier local
[Prix m² médian maison/appartement, évolution 12-24 mois, tension locative, données DVF récentes]

Règles : ton professionnel d'expert immobilier, 300 à 500 mots, indiquer la source et la date de chaque donnée chiffrée. Retourner uniquement le texte du chapitre.`;
}

function buildStylePrompt(docText) {
  return `Analyse ce rapport d'expertise immobilière de référence et extrais UNIQUEMENT les éléments stylistiques en JSON :
{
  "ton_general": "description du niveau de langue",
  "formules_introduction": ["liste des formules récurrentes d'introduction"],
  "formules_conclusion": ["formules de conclusion"],
  "style_terrain": "extrait de 2-3 phrases caractéristiques du chapitre terrain",
  "style_bati": "extrait de 2-3 phrases caractéristiques du chapitre bâti",
  "style_desordres": "comment les désordres sont présentés",
  "formules_conditionnelles": ["formules utilisées pour observations visuelles"],
  "vocabulaire_technique": ["termes techniques caractéristiques"]
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
  return `=== STYLE DE L'EXPERT (reproduire exactement) ===
${style || 'Style professionnel standard — français technique immobilier.'}

=== DONNÉES DU DOSSIER ===
Référence : ${formData.ref_dossier}
Date de visite : ${formData.date_visite}
Type de mission : ${formData.type_mission}
Donneur d'ordre : ${formData.nom_donneur_ordre} (${formData.donneur_ordre})
Adresse : ${formData.adresse_bien}
Références cadastrales : ${formData.refs_cadastrales || '[À COMPLÉTER PAR L\'EXPERT]'}
Régime juridique : ${formData.regime_juridique}
DPE : Classe ${formData.dpe_classe || 'NC'} — GES : Classe ${formData.ges_classe || 'NC'}
Type de bien : ${formData.type_bien}
Année de construction : ${formData.annee_construction || '[À COMPLÉTER PAR L\'EXPERT]'}
Niveaux : ${formData.nb_niveaux || '[À COMPLÉTER PAR L\'EXPERT]'}

=== TERRAIN ===
Superficie : ${formData.superficie_terrain} m²
Forme : ${formData.forme_terrain}
Topographie : ${formData.topographie}
Orientation : ${formData.orientation}
Accès : ${formData.acces_terrain}
Clôtures : ${formData.clotures}
Réseaux : ${formData.reseaux}
Zonage PLU : ${formData.zonage_plu || '[À COMPLÉTER PAR L\'EXPERT]'}
Contraintes : ${formData.contraintes}
Notes expert : ${formData.notes_terrain || ''}
Observations photos terrain : ${photos.terrain || 'Aucune photo fournie'}

=== BÂTI ===
Structure : ${formData.type_construction}
Toiture : ${formData.materiau_toiture} — ${formData.forme_toiture} — État : ${formData.etat_toiture}
Façades : ${formData.materiau_facades} — État : ${formData.etat_facades}
Menuiseries ext : ${formData.menuiseries_ext}
Sols intérieurs : ${formData.sols_interieurs || '[À COMPLÉTER PAR L\'EXPERT]'}
Chauffage : ${formData.chauffage}
Électricité : ${formData.etat_electrique}
Plomberie : ${formData.etat_plomberie}
Notes expert : ${formData.notes_bati || ''}
Observations photos extérieures : ${photos.ext || 'Aucune photo fournie'}
Observations photos intérieures : ${photos.int || 'Aucune photo fournie'}

=== DÉSORDRES CONSTATÉS ===
${desordres || 'Aucun désordre renseigné.'}
Observations photos désordres : ${photos.desordres || 'Aucune photo fournie'}

=== SURFACES ===
| Désignation | Niveau | Surface (m²) |
|-------------|--------|--------------|
${surfaces || '[À COMPLÉTER PAR L\'EXPERT]'}

=== CHAPITRE 1 DÉJÀ RÉDIGÉ (intégrer tel quel) ===
${chapter1}

---

GÉNÈRE LE PRÉ-RAPPORT COMPLET avec cette structure exacte :

# RAPPORT D'EXPERTISE IMMOBILIÈRE
## Pré-rapport soumis à validation

**Référence dossier :** ${formData.ref_dossier}
**Date de visite :** ${formData.date_visite}
**Adresse :** ${formData.adresse_bien}
**Nature de la mission :** ${formData.type_mission}
**Donneur d'ordre :** ${formData.nom_donneur_ordre} (${formData.donneur_ordre})

*Le présent document constitue un pré-rapport préparatoire. Les valeurs vénales et conclusions définitives feront l'objet d'une analyse complémentaire par l'expert signataire.*

---

## CHAPITRE 1 — SITUATION GÉOGRAPHIQUE ET ENVIRONNEMENT
[Intégrer le chapitre 1 déjà rédigé EXACTEMENT, sans modification]

---

## CHAPITRE 2 — DESCRIPTION DU TERRAIN
[Rédiger : situation/accès → caractéristiques physiques → superficie → réseaux → contraintes → PLU. 150-300 mots]

---

## CHAPITRE 3 — ÉTAT DU BIEN ET DESCRIPTION DES DÉSORDRES

### 3.1 Description générale et matériaux
[Rédiger : construction → toiture → façades/menuiseries → intérieurs → équipements → DPE. 250-400 mots]

### 3.2 Désordres constatés
[Format par désordre, du plus grave au moins grave :
**Désordre n°X — [Localisation]**
Nature : ...
Gravité : ...
Observation : ... (conditionnel / "à l'examen visuel")
Origine probable : ...
Si aucun désordre : "Aucun désordre significatif n'a été constaté lors de la visite."]

### 3.3 Tableau récapitulatif des surfaces
| Désignation | Niveau | Surface (m²) |
|-------------|--------|--------------|
[Reproduire les lignes + totaux]
| **Surface habitable totale** | | **X,XX m²** |
| **Surface annexes** | | **X,XX m²** |
| **Surface totale** | | **X,XX m²** |`;
}

// ─────────────────────────────────────────────────────────────────────────────
// ROUTES
// ─────────────────────────────────────────────────────────────────────────────

// POST /api/chapter1 — Recherche géographique (web_search)
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

// POST /api/extract-style — Extraction style depuis rapport de référence
app.post('/api/extract-style', upload.single('document'), async (req, res) => {
  try {
    if (!req.file) return res.json({ style: null });

    let docText = '';
    if (req.file.originalname.endsWith('.docx')) {
      const result = await mammoth.extractRawText({ buffer: req.file.buffer });
      docText = result.value.slice(0, 8000); // Limite pour le contexte
    } else if (req.file.originalname.endsWith('.pdf') || req.file.mimetype === 'application/pdf') {
      // Envoi direct en base64 pour les PDF
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
      return res.json({ style: response.content[0].text });
    }

    const response = await client.messages.create({
      model: process.env.CLAUDE_MODEL || 'claude-sonnet-4-6',
      max_tokens: 1500,
      temperature: 0,
      messages: [{ role: 'user', content: buildStylePrompt(docText) }]
    });

    res.json({ style: response.content[0].text });
  } catch (err) {
    console.error('[extract-style]', err.message);
    res.status(500).json({ error: err.message });
  }
});

// POST /api/analyze-photos — Analyse Vision IA (4 catégories)
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
        source: {
          type: 'base64',
          media_type: f.mimetype,
          data: f.buffer.toString('base64')
        }
      }));

      const response = await client.messages.create({
        model: process.env.CLAUDE_MODEL || 'claude-sonnet-4-6',
        max_tokens: 1000,
        temperature: 0,
        messages: [{
          role: 'user',
          content: [
            ...imageBlocks,
            { type: 'text', text: buildPhotoPrompt(cat) }
          ]
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

// POST /api/generate — Génération principale du pré-rapport
app.post('/api/generate', async (req, res) => {
  try {
    const response = await client.messages.create({
      model: process.env.CLAUDE_MODEL || 'claude-sonnet-4-6',
      max_tokens: 8000,
      temperature: 0.2,
      system: SYSTEM_PROMPT,
      messages: [{ role: 'user', content: buildMainPrompt(req.body) }]
    });

    const report = response.content[0].text;
    res.json({ report });
  } catch (err) {
    console.error('[generate]', err.message);
    res.status(500).json({ error: err.message });
  }
});

// POST /api/export-docx — Export Word
app.post('/api/export-docx', async (req, res) => {
  try {
    const { report, refDossier } = req.body;
    const buffer = await generateDocx(report);
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
// GÉNÉRATION DOCX
// ─────────────────────────────────────────────────────────────────────────────

async function generateDocx(markdown) {
  const children = parseMarkdown(markdown);
  const doc = new Document({
    styles: {
      paragraphStyles: [
        {
          id: 'expertTitle',
          name: 'Expert Title',
          basedOn: 'Normal',
          run: { font: 'Times New Roman', size: 36, bold: true, color: '1a2f4e' },
          paragraph: { spacing: { after: 200 }, alignment: AlignmentType.CENTER }
        }
      ]
    },
    sections: [{
      properties: {
        page: {
          margin: { top: 1440, right: 1134, bottom: 1440, left: 1134 }
        }
      },
      headers: {
        default: new Header({
          children: [
            new Paragraph({
              children: [
                new TextRun({ text: 'RAPPORT D\'EXPERTISE IMMOBILIÈRE — CONFIDENTIEL', size: 16, color: '6b6457' })
              ],
              border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: '9a7c38' } }
            })
          ]
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
              children: [new Paragraph({
                children: [new TextRun({
                  text: cell.replace(/\*\*/g, ''),
                  bold: ri === 0 || cell.startsWith('**'),
                  size: 18
                })]
              })],
              shading: ri === 0 ? { fill: 'E8E4DC', type: ShadingType.CLEAR } : undefined
            })
          )
        })
      );

      result.push(new Table({
        rows: tableRows,
        width: { size: 100, type: WidthType.PERCENTAGE }
      }));
    } catch {}
    tableBuffer = [];
    inTable = false;
  };

  for (const raw of lines) {
    const line = raw.trimEnd();

    if (line.startsWith('|')) {
      inTable = true;
      tableBuffer.push(line);
      continue;
    }
    if (inTable) { flushTable(); }

    if (line.startsWith('# ')) {
      result.push(new Paragraph({
        text: line.slice(2).trim(),
        heading: HeadingLevel.TITLE,
        alignment: AlignmentType.CENTER,
        run: { color: '1a2f4e' }
      }));
    } else if (line.startsWith('## ')) {
      result.push(new Paragraph({
        text: line.slice(3).trim(),
        heading: HeadingLevel.HEADING_1,
        run: { color: '1a2f4e' }
      }));
    } else if (line.startsWith('### ')) {
      result.push(new Paragraph({
        text: line.slice(4).trim(),
        heading: HeadingLevel.HEADING_2
      }));
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
  let lastIndex = 0;
  let match;

  while ((match = pattern.exec(text)) !== null) {
    if (match.index > lastIndex) {
      runs.push(new TextRun({ text: text.slice(lastIndex, match.index), size: 22 }));
    }
    const inner = match[0];
    if (inner.startsWith('**')) {
      runs.push(new TextRun({ text: inner.slice(2, -2), bold: true, size: 22 }));
    } else {
      runs.push(new TextRun({ text: inner.slice(1, -1), italics: true, size: 22, color: '6b6457' }));
    }
    lastIndex = match.index + match[0].length;
  }
  if (lastIndex < text.length) {
    runs.push(new TextRun({ text: text.slice(lastIndex), size: 22 }));
  }
  return runs.length ? runs : [new TextRun({ text, size: 22 })];
}

// ─────────────────────────────────────────────────────────────────────────────
const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`ExpertIA → http://localhost:${PORT}`));
