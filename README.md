# SYREQ — Analyseur d'Exigences Naval / Sous-marin

Outil local d'extraction et d'analyse automatique d'exigences systèmes à partir de demandes clients.
Domaine : **communications sous-marines et navales** (VLF/ELF, SATCOM, liaisons acoustiques, réseaux embarqués).
100 % local — aucune donnée envoyée sur internet.

---

## Démarrage rapide

```powershell
python dev/analyse_exigences.py
```

Ouvrir : **http://localhost:5000**

> Ollama doit être installé et un modèle téléchargé. Voir [installation/INSTALLATION.md](installation/INSTALLATION.md).

---

## Documentation

### Installation

| Document | Contenu |
|----------|---------|
| [installation/INSTALLATION.md](installation/INSTALLATION.md) | Installation classique — Ollama, Python, dépendances, lancement |
| [installation/MODE_PORTABLE.md](installation/MODE_PORTABLE.md) | Mode clé USB — fonctionnement sans installation sur l'ordinateur cible |

### Développement

| Document | Contenu |
|----------|---------|
| [dev/ARCHITECTURE.md](dev/ARCHITECTURE.md) | Architecture technique, flux d'analyse, prompt système |
| [dev/API.md](dev/API.md) | Routes Flask, formats requête/réponse, intégration DOORS/Jazz |
| [dev/JSON_STRUCTURE.md](dev/JSON_STRUCTURE.md) | Schéma complet du JSON, règles métier, compatibilité |
| [dev/CONFIGURATION.md](dev/CONFIGURATION.md) | Paramètres serveur, Ollama, contexte PDF, historique |

---

## Fonctionnalités principales

- **Analyse en 3 niveaux ISO 29148** : Explicite `[E]` / Déduite `[D]` / Hypothèse `[H]`
- **Édition inline** de toutes les exigences avec glisser-déposer
- **Statuts de validation** : À confirmer / Confirmé / Rejeté / En révision
- **Contrôle qualité** : détection termes vagues, fausses classifications
- **Contexte PDF** : ancrage des valeurs sur un document de référence
- **Mode Affiner** : amélioration ISO 29148 par le modèle
- **Exports** : JSON · Word · Excel (avec listes déroulantes) · DOORS CSV · ReqIF

---

## Licence

Usage interne — Projet SYREQ.
100 % local — aucune donnée transmise à des serveurs externes.
