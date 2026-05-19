# SYREQ — Analyseur d'Exigences Naval / Sous-marin

> Outil local d'extraction automatique d'exigences systèmes à partir de demandes clients brutes.
> Domaine : Télécommunications navales, systèmes sous-marins (sonar, VLF/ELF, acoustique, torpilles, AUV...).
> Modèle IA : **Ollama** (llama3 ou mistral) — **100% local, aucune donnée envoyée sur internet**.

---

## Architecture

```
Navigateur (index.html)         → Interface JS/CSS
        ↓ HTTP localhost:5000
Flask (analyse_exigences.py)    → Bridge HTTP ↔ Ollama
        ↓ ollama Python lib
Ollama (ollama serve)           → llama3 / mistral (local)
```

---

## 1. Installer Ollama

### Windows
1. Télécharger depuis : https://ollama.com/download/windows
2. Exécuter `OllamaSetup.exe` et suivre l'assistant.
3. Vérifier dans PowerShell : `ollama --version`

### Linux (Ubuntu / Debian)
```bash
curl -fsSL https://ollama.com/install.sh | sh
ollama --version
```

### macOS
```bash
brew install ollama
```

> llama3 (8B) requiert ~5 Go de RAM, mistral (7B) ~4.5 Go.

---

## 2. Télécharger un modèle

```bash
# Llama 3 — recommandé (meilleure compréhension du français)
ollama pull llama3

# OU Mistral 7B — plus rapide
ollama pull mistral

# Vérifier vos modèles
ollama list
```

---

## 3. Installer les dépendances Python

Prérequis : Python 3.9+ et pip.

```bash
# Aller dans le dossier du projet
cd /chemin/vers/syreq/

# (Recommandé) Créer un environnement virtuel
python -m venv .venv

# Activer l'environnement
# Linux/macOS :
source .venv/bin/activate
# Windows CMD :
.venv\Scripts\activate.bat
# Windows PowerShell :
.venv\Scripts\Activate.ps1

# Installer les dépendances
pip install -r requirements.txt
```

`requirements.txt` contient :
- `ollama>=0.2.0` — client Python officiel pour Ollama
- `flask>=3.0.0` — serveur web local

---

## 4. Lancer le script

### Étape 1 — Démarrer Ollama

```bash
# Dans un terminal séparé (si pas déjà actif) :
ollama serve
```

> Sur Windows, Ollama tourne en tâche de fond après installation.
> "address already in use" = déjà actif, c'est bon.

### Étape 2 — Lancer l'application

```bash
python analyse_exigences.py
```

Sortie attendue :
```
==============================================================
   Analyseur d'Exigences — Domaine Naval / Sous-marin
==============================================================
   Modèle par défaut  : llama3
   Dossier exports    : ./exports/
   Interface web      : http://127.0.0.1:5000
==============================================================
   [OK] Ollama est accessible.
```

### Étape 3 — Ouvrir l'interface

Navigateur : http://localhost:5000

---

## 5. Utilisation

1. Saisir la demande client dans la zone de texte (ou utiliser un exemple rapide).
2. Sélectionner le modèle en haut à droite (llama3 ou mistral).
3. Cliquer sur **ANALYSER**.
4. Résultats extraits automatiquement :
   - Périmètre (système, acteurs, milieux externes)
   - Besoin fondamental
   - Exigences fonctionnelles ("Le système DOIT...")
   - Contraintes techniques (acoustique, fréquences, pression, débit...)
   - Mots-clés domaine
5. **Exporter JSON** → sauvegardé dans `./exports/exigences_YYYYMMDD_HHMMSS.json`
6. **Copier JSON** → dans le presse-papiers

---

## 6. Structure du JSON exporté

```json
{
  "perimetre": {
    "systeme_sous_analyse": "Sonar passif basse fréquence",
    "acteurs_externes": ["Opérateur sonar", "PC commandement"],
    "milieux_externes": ["Eau de mer", "Fond marin"]
  },
  "besoin_fondamental": "Détecter et classifier des contacts sous-marins à longue distance...",
  "exigences_fonctionnelles": [
    "Le système DOIT détecter des contacts à une distance minimale de 25 km.",
    "Le système DOIT traiter les signaux acoustiques avec une latence < 200 ms."
  ],
  "contraintes_techniques": {
    "acoustique": "Fréquences 100–2000 Hz, rapport signal/bruit > 15 dB",
    "frequences": "Bande VLF (3–30 kHz) et basse fréquence acoustique",
    "etancheite_pression": "Pression 30 bar (300 m), IP68",
    "debit_latence": "Traitement temps réel, latence < 200 ms",
    "environnement_physique": "Eau 2–20°C, corrosion marine NaCl 3.5%",
    "autres": ["Compatibilité EM avec propulsion nucléaire"]
  },
  "mots_cles_domaine": ["OTAN", "sonar passif", "MIL-STD"],
  "_meta": {
    "model": "llama3",
    "timestamp": "2024-01-15T14:32:10.123456",
    "demande_source": "Nous avons besoin d'un sonar passif..."
  }
}
```

---

## 7. Routes API Flask

| Méthode | Route     | Description                          |
|---------|-----------|--------------------------------------|
| GET     | /         | Interface web (index.html)           |
| POST    | /analyse  | Lance l'analyse via Ollama           |
| POST    | /export   | Sauvegarde le résultat en JSON local |
| GET     | /models   | Liste les modèles Ollama disponibles |

---

## 8. Dépannage

**"Ollama non détecté"**
```bash
ollama serve   # Démarrer Ollama manuellement
```

**"JSON invalide retourné par le modèle"**
- Préférer llama3 (meilleur respect du format JSON)
- Réduire la longueur de la demande si > 2000 caractères
- Relancer l'analyse (les LLMs ont un part non-déterministe)

**Le modèle est lent**
- Vérifier la détection GPU : `ollama info`
- Utiliser mistral (plus léger)

**Port 5000 déjà utilisé**
Modifier `PORT = 5001` dans `analyse_exigences.py`.

---

## 9. Personnalisation du prompt

Le prompt système se trouve dans la variable `SYSTEM_PROMPT` dans `analyse_exigences.py`.
Adaptable à d'autres domaines : aéronautique (DO-178), ferroviaire (EN 50128), spatial...

---

## Licence

Usage interne — Projet SYREQ.
100% local — aucune donnée transmise à des serveurs externes.
