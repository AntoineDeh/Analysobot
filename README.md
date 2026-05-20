# SYREQ — Analyseur d'Exigences Naval / Sous-marin

> Outil local d'extraction et d'analyse automatique d'exigences systèmes à partir de demandes clients.
> Domaine spécialisé : **communications sous-marines et navales** (VLF/ELF, SATCOM, liaisons acoustiques, réseaux embarqués).
> Modèle IA : **Ollama / Mistral** — **100 % local, aucune donnée envoyée sur internet**.

---

## Fonctionnalités

### Analyse en 3 niveaux (ISO 29148)
Chaque exigence est classée automatiquement :

| Badge | Niveau | Description |
|-------|--------|-------------|
| `[E]` vert | **Explicite** | Directement formulé dans le texte — citation fournie |
| `[D]` amber | **Déduite** | Logiquement nécessaire (ex : profondeur 400 m → 40 bar) |
| `[H]` orange | **Hypothèse** | Standard du domaine naval/militaire, à valider client |

Le modèle ne se limite pas à extraire le texte — il déduit les implications techniques et propose les standards pertinents (STANAG, MIL-STD, SDIP-27...) en le signalant explicitement.

### Édition interactive des résultats
Tous les champs sont modifiables en ligne :
- Texte de chaque exigence
- Origine (E/D/H), priorité MoSCoW, méthode de vérification TAID, statut de validation
- Périmètre (système, acteurs, milieux)
- Besoin fondamental, contraintes techniques, mots-clés
- Ajout / suppression / réordonnancement par glisser-déposer des exigences

### Contrôle qualité intégré
- **Détecteur de termes vagues** : surligne en jaune les mots non-testables (`adéquat`, `suffisant`, `robuste`…)
- **Détecteur de fausse classification** : si une exigence marquée `[E]` n'a pas de citation réelle → badge `[?]` avec avertissement
- **Statistiques en temps réel** : total, répartition E/D/H, Must/Should/Could, nb à confirmer/confirmés/rejetés

### Exports
| Format | Usage |
|--------|-------|
| JSON | Rechargeable dans SYREQ, archivage |
| Word (.docx) | Fiche d'exigences mise en page, prête à diffuser |
| Excel (.xlsx) | 5 feuilles : Synthèse, EF, ENF, Contraintes, Matrice Traçabilité — avec listes déroulantes |
| DOORS CSV | Import direct IBM DOORS Next via assistant de mappage de colonnes |
| ReqIF (.reqif) | Format standard OMG 1.0.1 — import natif IBM DOORS Next et Jazz ELM |

### Autres fonctionnalités
- **Document de contexte** : charger un PDF ou TXT comme référence technique (normes, STB, spécifications) — le modèle priorise ses valeurs
- **Mode Affiner** : soumet les exigences extraites au modèle pour amélioration ISO 29148
- **Préfixe personnalisable** : `EF-01` → `SONAR-01`, `SYS-VLF-01`…
- **Historique local** : 20 dernières analyses conservées dans `localStorage`, supprimables individuellement ou en masse
- **Chargement JSON** : reprise d'une session précédente

---

## Architecture

```
Navigateur (index.html)
  ↓ HTTP localhost:5000
Flask (analyse_exigences.py)   — serveur local, API REST
  ↓ ollama Python lib
Ollama                         — modèle Mistral 7B (100 % local)
  ↓
GPU / CPU local                — inférence sur RTX 3060 ou CPU
```

Toutes les données restent sur la machine. Aucun appel réseau externe.

---

## Installation

### 1. Ollama (Windows)

1. Télécharger l'installateur : https://ollama.com/download/windows
2. Exécuter `OllamaSetup.exe` — Ollama démarre automatiquement au boot (icône systray)
3. Vérifier dans PowerShell : `ollama --version`

### 2. Modèle recommandé

SYREQ affiche automatiquement **tous les modèles installés** dans Ollama via le sélecteur de l'interface. N'importe quel modèle disponible peut être utilisé.

| Modèle | Commande | VRAM | Qualité FR | Vitesse | Recommandé pour |
|--------|----------|------|------------|---------|-----------------|
| **Mistral 7B** | `ollama pull mistral` | ~4,4 Go | ★★★★★ | ★★★★ | Usage principal — français natif (Mistral AI, Paris) |
| **Llama 3 8B** | `ollama pull llama3` | ~4,7 Go | ★★★ | ★★★★ | Alternative si Mistral indisponible |
| **phi4-mini** | `ollama pull phi4-mini` | ~2,2 Go | ★★★ | ★★★★★ | VRAM insuffisante (< 3 Go libres) |

```powershell
# Installation recommandée
ollama pull mistral

# Vérifier les modèles installés
ollama list
```

> **Configuration matérielle testée** : RTX 3060 6 Go VRAM — Mistral tient entièrement en VRAM
> si les autres applications libèrent ~2 Go (fermer les onglets navigateur).

### 3. Dépendances Python

```powershell
# Depuis le dossier du projet
python -m venv .venv
.venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

**`requirements.txt`** :
```
ollama>=0.2.0         # client Python Ollama
flask>=3.0.0          # serveur web local
json-repair>=0.28.0   # robustesse parsing JSON
python-docx>=1.1.0    # export Word
openpyxl>=3.1.0       # export Excel
pypdf>=4.0.0          # extraction texte PDF
```

---

## Lancement

```powershell
# Ollama tourne déjà en tâche de fond sur Windows
# Lancer l'application :
python analyse_exigences.py
```

Ouvrir : **http://localhost:5000**

---

## Workflow typique

```
1. Saisir la demande client  ─── ou charger un exemple (VLF, acoustique, SATCOM, réseau interne)
2. [Optionnel] Charger un PDF de contexte (cahier des charges, norme, STB)
3. Cliquer ANALYSER (Ctrl+Entrée)
4. Vérifier les résultats :
   ├─ [E] Exigences explicites → vérifier les citations
   ├─ [D] Exigences déduites   → valider le raisonnement
   └─ [H] Hypothèses           → soumettre au client pour confirmation
5. Éditer, reclasser, ajuster les statuts (À confirmer / Confirmé / Rejeté)
6. [Optionnel] AFFINER → le modèle améliore la qualité rédactionnelle
7. Personnaliser le préfixe (ex: SONAR, SYS-COM)
8. Exporter (JSON / Word / Excel / DOORS CSV / ReqIF)
```

---

## Structure JSON

```json
{
  "perimetre": {
    "systeme_sous_analyse": "Récepteur VLF sous-marin",
    "acteurs_externes": ["Opérateur Radio", "Station VLF Rosnay"],
    "milieux_externes": ["Milieu marin immergé", "Eau de mer 3,5% NaCl"]
  },
  "besoin_fondamental": "Recevoir les ordres de commandement OTAN depuis une station VLF terrestre à une profondeur de 200 m.",
  "exigences_fonctionnelles": [
    {
      "texte": "Le système DOIT recevoir les messages OTAN en bande VLF (3–30 kHz) à une profondeur minimale de 200 m.",
      "origine": "Explicite",
      "justification": "«réception à une profondeur de 200 mètres»",
      "priority": "Must",
      "verification": "Test",
      "statut": "Confirmé"
    },
    {
      "texte": "Le système DOIT résister à une pression hydrostatique minimale de 20 bar.",
      "origine": "Déduite",
      "justification": "Profondeur 200 m → 1 bar ≈ 10 m → 20 bar minimum (loi de Pascal).",
      "priority": "Must",
      "verification": "Test",
      "statut": "À confirmer"
    },
    {
      "texte": "Le système DOIT être conforme au standard MIL-STD-461G pour la compatibilité électromagnétique.",
      "origine": "Hypothèse",
      "justification": "Hypothèse domaine : système militaire embarqué → conformité CEM militaire. A valider avec le client.",
      "priority": "Should",
      "verification": "Test",
      "statut": "À confirmer"
    }
  ],
  "exigences_non_fonctionnelles": [
    {
      "type": "Performance",
      "texte": "Le système DOIT atteindre une sensibilité minimale de -120 dBm.",
      "origine": "Explicite",
      "justification": "«sensibilité minimale du récepteur : -120 dBm»",
      "priority": "Must",
      "verification": "Test",
      "statut": "À confirmer"
    }
  ],
  "contraintes_techniques": {
    "acoustique": "Non spécifié",
    "frequences": "VLF 3–30 kHz, ELF 30–300 Hz",
    "etancheite_pression": "IP68, 20 bar (200 m)",
    "debit_latence": "Latence décodage < 5 s",
    "environnement_physique": "Eau de mer 3,5% NaCl, 0–25°C",
    "autres": ["STANAG 4204", "MIL-STD-188-110C", "SDIP-27 niveau B"]
  },
  "mots_cles_domaine": ["VLF", "sous-marin", "OTAN", "Rosnay", "STANAG 4204"],
  "_meta": {
    "model": "mistral",
    "timestamp": "2026-05-20T10:15:00",
    "demande_source": "Système de réception VLF pour sous-marin...",
    "demande_full": "Texte complet de la demande...",
    "context_file": "STB_VLF_v2.pdf"
  }
}
```

---

## API Flask

| Méthode | Route | Description |
|---------|-------|-------------|
| GET | `/` | Interface web |
| POST | `/analyse` | Analyse via Ollama → JSON résultat |
| POST | `/affiner` | Affinage ISO 29148 des exigences existantes |
| POST | `/extract-context` | Extraction texte PDF/TXT pour contexte |
| POST | `/export/word` | Génère et télécharge le .docx |
| POST | `/export/excel` | Génère et télécharge le .xlsx |
| POST | `/export/doors-csv` | Génère CSV IBM DOORS Next |
| POST | `/export/reqif` | Génère ReqIF 1.0.1 (DOORS/Jazz) |
| GET | `/models` | Liste les modèles Ollama disponibles |

---

## Intégration IBM DOORS Next / Jazz ELM

### Import DOORS CSV
1. Exporter depuis SYREQ → bouton **DOORS CSV**
2. Dans DOORS Next : `Artefacts > Importer > Importer depuis un fichier CSV`
3. Mapper les colonnes : `Primary Text` → Texte, `Priority` → Priorité, etc.

### Import ReqIF
1. Exporter depuis SYREQ → bouton **ReqIF**
2. Dans DOORS Next : `Artefacts > Importer > Importer depuis ReqIF`
3. Import direct sans configuration de mappage

Le fichier ReqIF contient les attributs : Texte, Type, Module, Priorité (enum), Vérification (enum).

---

## Dépannage

**"OLLAMA HORS LIGNE" au démarrage**
```powershell
ollama serve   # Démarrer manuellement si pas dans le systray
```

**JSON invalide retourné par le modèle**
- Relancer l'analyse (les LLMs ont une part non-déterministe)
- Si répété : vérifier que `json-repair` est installé (`pip install json-repair`)

**La sortie est vide ou les exigences manquent**
- Mistral et llama3 ne supportent pas le prefill `{` — le code n'utilise plus cette technique
- Vérifier la version Ollama : `ollama --version` (>=0.2.0 requis)

**Modèle lent / CPU au lieu de GPU**
```powershell
ollama run mistral "test"
# Vérifier dans Gestionnaire des tâches > GPU que la charge GPU monte
```
Si CPU uniquement : installer/réinstaller les drivers NVIDIA CUDA.

**VRAM insuffisante**
Fermer Chrome/Edge (libère 1–2 Go VRAM) avant de lancer une analyse.
Alternative : `ollama pull phi4-mini` (~2,2 Go VRAM).

**Port 5000 déjà utilisé**
Modifier `PORT = 5001` dans `analyse_exigences.py`.

---

## Configuration

Paramètres modifiables dans `analyse_exigences.py` :

| Variable | Valeur par défaut | Description |
|----------|-------------------|-------------|
| `DEFAULT_MODEL` | `"mistral"` | Modèle Ollama utilisé par défaut |
| `HOST` | `"127.0.0.1"` | Adresse du serveur Flask |
| `PORT` | `5000` | Port du serveur Flask |
| `OUTPUT_DIR` | `"exports"` | Dossier de sauvegarde JSON |
| `temperature` | `0.1` | Déterminisme du modèle (0 = max déterministe) |
| `num_ctx` | `16384` | Fenêtre de contexte tokens (adapter selon VRAM) |

---

## Licence

Usage interne — Projet SYREQ.
100 % local — aucune donnée transmise à des serveurs externes.
