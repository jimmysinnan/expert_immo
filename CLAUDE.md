# CLAUDE.md — expert-immo-app (ExpertIA)

## Identité du projet
**ExpertIA** est une application web standalone de génération automatique de pré-rapports d'expertise immobilière pour le **Cabinet JALTA** en Martinique.
Elle utilise Claude AI (Vision + web_search + génération texte) pour produire des rapports conformes aux standards TEGOVA 5e/6e édition et à la Charte de l'Expertise Immobilière.

**GitHub :** https://github.com/jimmysinnan/expert_immo

---

## Stack technique
| Élément | Technologie |
|---|---|
| Backend | Node.js + Express |
| IA | Claude claude-sonnet-4-6 (Anthropic SDK) — Vision, web_search, génération |
| Export | docx (Word natif via librairie `docx`) |
| Frontend | Vanilla HTML/CSS/JS — zéro framework |
| Port | 3002 (défini dans .env) |

---

## Structure du projet
```
expert-immo-app/
├── server.js          → Backend Express + toute la logique IA (1273 lignes)
├── public/
│   ├── index.html     → Interface utilisateur (formulaire 3 étapes)
│   ├── app.js         → Logique frontend (650 lignes)
│   └── style.css      → Styles JALTA (navy + vert)
├── .env               → Clé API + config (⚠️ NE JAMAIS COMMITTER)
├── .env.example       → Template sans clé réelle
├── .gitignore         → node_modules/, .env, uploads/ exclus ✓
└── package.json       → express, @anthropic-ai/sdk, docx, mammoth, multer
```

---

## Workflow de l'application
1. **Formulaire** — 3 sections : Informations bien (42 champs) + Base de connaissance (.docx upload) + Photos par catégorie
2. **Pipeline de génération en 6 étapes :**
   - Chapitre 1 — Situation (via `web_search` : INSEE, DVF, marché local)
   - Extraction du style expert depuis base de connaissance (.docx → mammoth)
   - Analyse Vision IA des photos (terrain, extérieur, intérieur, désordres)
   - Génération pré-rapport complet (~8 000 tokens)
3. **Export** — Téléchargement direct `.docx` formaté aux couleurs JALTA

---

## Variables d'environnement (.env)
```env
ANTHROPIC_API_KEY=sk-ant-api-...   ← clé réelle (ne jamais committer)
CLAUDE_MODEL=claude-sonnet-4-6
PORT=3002
```

---

## Lancer le projet
```bash
cd C:\Users\jimmy\Projets\expert-immo-app
npm start          # production
npm run dev        # développement avec nodemon (hot reload)
```
Ouvrir : http://localhost:3002

---

## Sécurité
- ✅ `.env` exclu du git via `.gitignore`
- ✅ `.env.example` contient uniquement le placeholder `sk-ant-api-xxx`
- ✅ Validation de la clé API au démarrage du serveur (process.exit si invalide)
- ⚠️ Ne jamais mettre la vraie clé dans `.env.example` ni dans le code

---

## État actuel du développement

### Ce qui est fait ✅
- Application fonctionnelle de bout en bout
- Formulaire complet (42 champs + 5 désordres + 15 surfaces)
- Pipeline IA en 6 étapes opérationnel
- Export Word formaté aux couleurs JALTA (navy + vert)
- Structure rapport TEGOVA 5e/6e édition intégrée
- Base de connaissance : extraction de style depuis .docx de référence
- Analyse Vision des photos par catégorie
- Git initialisé + remote GitHub (expert_immo)

### Derniers commits
- `100da16` — feat: intégration structure et design JALTA (rapport TEGOVA 5e/6e)
- `d210400` — Initial commit — ExpertIA v1.0

### Prochaines étapes identifiées
- [ ] Améliorer la qualité des prompts de génération (affiner le style JALTA)
- [ ] Ajouter la gestion des erreurs côté frontend (feedback utilisateur)
- [ ] Optimiser le pipeline de génération (paralléliser certaines étapes ?)
- [ ] Tester sur de vrais dossiers JALTA et itérer sur les prompts
- [ ] Éventuellement : authentification simple (usage interne Cabinet)
- [ ] Déploiement possible sur VPS ou Railway si besoin d'accès distant

---

## Comment reprendre le travail

### Terminal
```bash
cd C:\Users\jimmy\Projets\expert-immo-app
claude --continue    # reprend la dernière conversation Claude Code
```

### VS Code
```bash
code C:\Users\jimmy\Projets\expert-immo-app
```
Puis `Ctrl+\`` pour ouvrir le terminal intégré, puis `claude --continue`.

---

## Notes stratégiques
Ce projet a une vraie valeur business immédiate : il automatise un travail de rédaction répétitif et technique pour un cabinet existant (JALTA). C'est un démonstrateur concret d'IA appliquée à un métier spécialisé. Priorité : stabiliser le pipeline, tester sur de vrais dossiers, mesurer le gain de temps réel.
