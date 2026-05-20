# Mode portable — Clé USB

SYREQ fonctionne entièrement depuis une clé USB, sans rien installer sur l'ordinateur cible.

---

## Structure de la clé

```
SYREQ/
├── LANCER.bat             ← Double-clic pour démarrer
├── STOPPER.bat            ← Arrêter proprement
├── SETUP_USB.bat          ← Préparation initiale (une seule fois)
├── analyse_exigences.py
├── index.html
├── exports/               ← Exports sauvegardés sur la clé
├── python/                ← Python portable (~30 Mo)
├── ollama/                ← Ollama portable (~50 Mo)
└── ollama_models/         ← Modèles IA (~4,5 Go pour Mistral)
```

---

## Préparation (une seule fois, sur votre machine)

### Étape 1 — Télécharger Python embeddable

Télécharger `python-3.12.x-embed-amd64.zip` depuis :
https://www.python.org/downloads/windows/

Section : **"Windows embeddable package (64-bit)"**

Extraire le contenu du ZIP dans le dossier `python/` du projet.

### Étape 2 — Exécuter SETUP_USB.bat

Double-clic sur `SETUP_USB.bat` (nécessite Ollama installé et connexion internet).

Le script fait automatiquement :
- Active les packages dans Python embeddable
- Installe Flask, Ollama, python-docx, openpyxl, pypdf via pip
- Copie `ollama.exe` depuis votre installation locale
- Copie les librairies CUDA (support GPU)
- Copie les modèles depuis `%USERPROFILE%\.ollama\models`

### Étape 3 — Copier sur la clé

Copier l'intégralité du dossier `SYREQ/` sur la clé USB.

**Taille totale** : ~5–6 Go avec Mistral.

> Utiliser une clé USB 3.0 minimum pour des performances acceptables.

---

## Utilisation sur un autre PC

Double-clic sur `LANCER.bat`.

- Lance Ollama avec les modèles de la clé
- Ouvre automatiquement http://localhost:5000
- Fermer la fenêtre noire arrête tout proprement

**Prérequis sur l'ordinateur cible** : Windows 10/11 64-bit uniquement.

> Sur CPU sans GPU NVIDIA, l'analyse prend 2 à 5 minutes au lieu de 20–40 secondes.

---

## Dépannage portable

**Ollama ne démarre pas**
Vérifier que `ollama/ollama.exe` est bien présent. Si absent, exécuter à nouveau `SETUP_USB.bat`.

**"Module not found"**
Les packages Python ne sont pas installés dans `python/Lib/site-packages/`.
Exécuter `SETUP_USB.bat` sur une machine avec connexion internet.

**Modèles absents**
Vérifier que `ollama_models/manifests/` et `ollama_models/blobs/` existent et ne sont pas vides.
