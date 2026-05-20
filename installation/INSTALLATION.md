# Installation classique

Installation sur une machine fixe. Tout s'installe une seule fois.

---

## 1. Ollama

Télécharger et exécuter l'installateur Windows :
https://ollama.com/download/windows

Ollama se lance automatiquement au démarrage de Windows (icône dans la barre des tâches). Pas besoin de le lancer manuellement.

---

## 2. Modèle de langage

Ouvrir PowerShell et exécuter :

```powershell
ollama pull mistral
```

Taille : ~4,4 Go. Téléchargement unique.

| Modèle | Commande | VRAM | Qualité FR | Recommandé pour |
|--------|----------|------|------------|-----------------|
| **Mistral 7B** | `ollama pull mistral` | ~4,4 Go | ★★★★★ | Usage principal |
| **Llama 3 8B** | `ollama pull llama3` | ~4,7 Go | ★★★ | Alternative |
| **phi4-mini** | `ollama pull phi4-mini` | ~2,2 Go | ★★★ | VRAM < 3 Go |

Le sélecteur de modèle dans l'interface affiche automatiquement tous les modèles installés.

> **Configuration testée** : RTX 3060 6 Go VRAM.
> Fermer les onglets navigateur avant l'analyse pour libérer la VRAM.

---

## 3. Python

Télécharger Python 3.12 (Windows 64-bit) :
https://www.python.org/downloads/windows/

Cocher **"Add Python to PATH"** lors de l'installation.

---

## 4. Dépendances du projet

Ouvrir PowerShell **à la racine du projet** (dossier contenant `dev/` et `installation/`) :

```powershell
python -m venv .venv
.venv\Scripts\Activate.ps1
pip install -r installation\requirements.txt
```

Packages installés :

| Package | Rôle |
|---------|------|
| `flask` | Serveur web local |
| `ollama` | Client Python pour Ollama |
| `python-docx` | Export Word |
| `openpyxl` | Export Excel |
| `pypdf` | Extraction texte PDF |
| `json-repair` | Robustesse parsing JSON |

---

## 5. Lancement

Depuis la racine du projet :

```powershell
python dev\analyse_exigences.py
```

Ouvrir : **http://localhost:5000**

---

## Dépannage

**"OLLAMA HORS LIGNE"**
```powershell
ollama serve
```

**Analyse vide ou JSON invalide**
Relancer l'analyse (comportement non-déterministe du LLM). Si répété, vérifier que `json-repair` est installé.

**Modèle lent**
Vérifier que le GPU est utilisé : Gestionnaire des tâches > GPU — la charge doit monter pendant l'analyse.
Si CPU seulement : réinstaller les drivers NVIDIA.

**Port 5000 déjà utilisé**
Modifier `PORT = 5001` dans `analyse_exigences.py`.
