# SYREQ — Analyseur d'Exigences Naval / Sous-marin

Outil local d'extraction et d'analyse automatique d'exigences systèmes à partir de demandes clients.  
Domaine : communications sous-marines et navales (VLF/ELF, SATCOM, liaisons acoustiques, réseaux embarqués).  
**100 % local — aucune donnée envoyée sur internet.**

---

## Installation

### 1. Ollama

Télécharger et installer : https://ollama.com/download/windows  
Ollama démarre automatiquement au boot (icône systray).

### 2. Modèle

```powershell
ollama pull mistral       # recommandé — français natif, ~4,4 Go
ollama pull llama3        # alternative
```

### 3. Python

Télécharger Python 3.12 : https://www.python.org/downloads/windows/  
Cocher **"Add Python to PATH"**.

### 4. Dépendances

Depuis la racine du projet :

```powershell
python -m venv .venv
.venv\Scripts\Activate.ps1
pip install -r installation\requirements.txt
```

### 5. Lancement

```powershell
python dev\analyse_exigences.py
```

Ouvrir dans Chrome : **http://localhost:5000**

Ou double-clic sur `LANCER.bat` à la racine.

---

## Mode portable — Clé USB

Fonctionne sans rien installer sur l'ordinateur cible.

**Préparation (une seule fois, sur votre machine) :**

1. Extraire `python-3.12.x-embed-amd64.zip` dans `installation/python/`  
   → https://www.python.org/downloads/windows/ — *"Windows embeddable package (64-bit)"*

2. Double-clic sur `installation/SETUP_USB.bat`  
   Installe les packages Python, copie Ollama et les modèles automatiquement.

3. Copier le dossier `SYREQ/` sur la clé USB (~5–6 Go avec Mistral).

**Utilisation :** double-clic sur `LANCER.bat`.  
Prérequis sur l'ordinateur cible : Windows 10/11 64-bit.

---

## Fonctionnalités

### Analyse en 3 niveaux ISO 29148

| Badge | Niveau | Description |
|-------|--------|-------------|
| `[E]` vert | **Explicite** | Directement dans le texte — citation fournie |
| `[D]` amber | **Déduite** | Implication logique (ex : 400 m → 40 bar) |
| `[H]` orange | **Hypothèse** | Standard du domaine naval/militaire, à valider client |

Le modèle déduit les implications physiques et propose les normes pertinentes (STANAG, MIL-STD, SDIP-27…) en les signalant clairement.

### Édition interactive

- Texte, origine (E/D/H), priorité MoSCoW, méthode TAID, statut de validation
- Glisser-déposer pour réordonner les exigences
- Préfixe personnalisable : `EF-01` → `SONAR-01`, `SYS-COM-01`…

### Contrôle qualité

- Détection des termes non-testables (`adéquat`, `suffisant`, `robuste`…)
- Détection des fausses classifications `[E]` sans citation réelle → badge `[?]`
- Statistiques : E/D/H, Must/Should/Could, confirmés/à valider

### Document de contexte PDF

Charger un PDF de référence (norme, STB, cahier des charges) — le modèle y ancre ses valeurs numériques.

### Mode Affiner

Soumet les exigences au modèle pour amélioration ISO 29148.

### Historique local

20 dernières analyses conservées dans le navigateur (localStorage), supprimables individuellement.

### Exports

| Format | Description |
|--------|-------------|
| JSON | Rechargeable dans SYREQ |
| Word (.docx) | Fiche d'exigences mise en page |
| Excel (.xlsx) | 5 feuilles avec listes déroulantes |
| DOORS CSV | Import IBM DOORS Next |
| ReqIF (.reqif) | Import natif DOORS Next et Jazz ELM |

---

## Structure du projet

```
SYREQ/
├── LANCER.bat                  ← double-clic pour démarrer
├── README.md
├── dev/
│   ├── analyse_exigences.py    ← serveur Flask + logique
│   └── index.html              ← interface web
└── installation/
    ├── LANCER.bat              ← script de lancement complet
    ├── STOPPER.bat             ← arrêt
    ├── SETUP_USB.bat           ← préparation clé USB
    ├── requirements.txt
    ├── python/                 ← Python portable (USB)
    ├── ollama/                 ← Ollama portable (USB)
    └── ollama_models/          ← modèles IA (USB)
```

---

## API

| Méthode | Route | Description |
|---------|-------|-------------|
| GET | `/` | Interface web |
| POST | `/analyse` | Analyse via Ollama |
| POST | `/affiner` | Affinage ISO 29148 |
| POST | `/extract-context` | Extraction texte PDF/TXT |
| POST | `/export/word` | Téléchargement .docx |
| POST | `/export/excel` | Téléchargement .xlsx |
| POST | `/export/doors-csv` | CSV IBM DOORS Next |
| POST | `/export/reqif` | ReqIF 1.0.1 |
| GET | `/models` | Liste modèles Ollama |

---

## Configuration

Dans `dev/analyse_exigences.py` :

| Variable | Défaut | Description |
|----------|--------|-------------|
| `DEFAULT_MODEL` | `"mistral"` | Modèle par défaut |
| `PORT` | `5000` | Port du serveur |
| `num_ctx` | `16384` | Fenêtre de contexte (réduire si VRAM < 4 Go) |
| `temperature` | `0.1` | Déterminisme du modèle |

---

## Dépannage

**Terminal se ferme immédiatement**  
Lancer depuis PowerShell pour voir l'erreur :
```powershell
.venv\Scripts\python.exe dev\analyse_exigences.py
```

**"OLLAMA HORS LIGNE"**
```powershell
ollama serve
```

**Modèle lent**  
Fermer les onglets Chrome avant l'analyse pour libérer la VRAM.

**Port 5000 déjà utilisé**  
Modifier `PORT = 5001` dans `dev/analyse_exigences.py`.

---

## Licence

Usage interne — Projet SYREQ. 100 % local.
