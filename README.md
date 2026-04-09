# ExpertIA — Rapport d'Expertise Immobilière Automatisé

Application web standalone pour la génération automatique de pré-rapports d'expertise immobilière via Claude AI.

## Stack

- **Backend** : Node.js + Express
- **IA** : Claude claude-sonnet-4-6 (Anthropic SDK) — Vision, web_search, génération texte
- **Export** : docx (Word natif)
- **Frontend** : Vanilla HTML/CSS/JS — zéro framework

## Installation

```bash
git clone https://github.com/jimmysinnan/expert-immo-app.git
cd expert-immo-app
npm install
cp .env.example .env
# Éditer .env et renseigner ANTHROPIC_API_KEY
npm start
```

Ouvrir : http://localhost:3000

## Workflow

1. **Section A–F** — Formulaire guidé (42 champs + 5 désordres + 15 surfaces)
2. **Base de connaissance** — Upload optionnel d'un rapport .docx de référence (extraction du style de l'expert)
3. **Photos** — Upload par catégorie (terrain, extérieures, intérieures, désordres) — analysées par Vision IA
4. **Génération** — Pipeline automatique en 6 étapes :
   - Chapitre 1 via web_search (INSEE, DVF, marché local)
   - Extraction style depuis la base de connaissance
   - Analyse Vision IA des photos
   - Génération du pré-rapport complet (8 000 tokens)
5. **Export** — Téléchargement direct en .docx

## Variables d'environnement

```env
ANTHROPIC_API_KEY=sk-ant-api-xxx
CLAUDE_MODEL=claude-sonnet-4-6
PORT=3000
```

## Coût estimé

~0,08 à 0,15 € par rapport (selon nombre de photos)
