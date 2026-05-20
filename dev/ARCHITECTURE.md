# Architecture technique

## Vue d'ensemble

```
Navigateur (index.html)
    │  HTML/CSS/JS — interface utilisateur complète
    │  localStorage — historique des 20 dernières analyses
    │
    ↓  HTTP POST/GET  localhost:5000
    │
Flask (analyse_exigences.py)
    │  Serveur web local — Python 3.12
    │  Pas de base de données — tout en mémoire ou fichiers
    │
    ↓  bibliothèque Python ollama
    │
Ollama (processus système)
    │  Moteur d'inférence local
    │  Modèle : Mistral 7B (par défaut)
    │
    ↓  CUDA / CPU
    │
GPU RTX 3060 / CPU
```

## Fichiers du projet

| Fichier | Rôle |
|---------|------|
| `analyse_exigences.py` | Serveur Flask, appel Ollama, exports |
| `index.html` | Interface complète (HTML + CSS + JS inline) |
| `requirements.txt` | Dépendances Python |
| `exports/` | Fichiers JSON exportés côté serveur |

## Flux d'une analyse

```
1. Utilisateur saisit le texte → [Optionnel] charge un PDF de contexte
2. Clic ANALYSER → fetch POST /analyse
3. Flask construit le prompt (SYSTEM_PROMPT + contexte PDF + demande)
4. ollama.chat() → Mistral génère le JSON
5. parse_json_robust() extrait et valide le JSON
6. Flask retourne jsonify(result)
7. JS renderResult() construit le DOM éditable
8. Statistiques mises à jour (E/D/H, Must/Should, statuts)
9. Historique sauvegardé dans localStorage
```

## Prompt système (SYSTEM_PROMPT)

Le prompt est en français pour Mistral. Il définit :
- Le profil expert (communications sous-marines, STANAG, MIL-STD)
- L'approche en 3 niveaux : Explicite / Déduite / Hypothèse
- La structure JSON exacte attendue
- Les règles anti-hallucination (pas de valeur inventée)
- Les règles de mix (minimum 2E + 2D + 2H)

Fichier : `analyse_exigences.py` — variable `SYSTEM_PROMPT`

## Parsing JSON robuste

`parse_json_robust()` tente dans l'ordre :
1. `json.loads()` direct
2. Extraction du bloc `{ ... }` par recherche des accolades
3. `json_repair()` si disponible

## Paramètres Ollama

```python
temperature = 0.1   # Quasi-déterministe
top_p       = 0.9
num_predict = 4096  # Tokens générés max
num_ctx     = 16384 # Fenêtre de contexte (adapter selon VRAM)
```
