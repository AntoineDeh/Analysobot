# API Flask

Toutes les routes écoutent sur `http://localhost:5000`.

---

## GET /

Retourne `index.html`.

---

## POST /analyse

Lance l'analyse d'une demande client.

**Corps (JSON) :**
```json
{
  "demande": "Texte de la demande client",
  "model": "mistral",
  "context": "Texte extrait du PDF de contexte (optionnel)",
  "context_filename": "nom_fichier.pdf (optionnel)"
}
```

**Retour (JSON) :** objet exigences complet (voir `dev/JSON_STRUCTURE.md`)

**Erreurs :**
```json
{ "error": "message d'erreur" }
```

---

## POST /affiner

Améliore la qualité rédactionnelle ISO 29148 des exigences existantes.

**Corps (JSON) :**
```json
{
  "result": { /* objet exigences complet */ },
  "model": "mistral"
}
```

**Retour :** objet exigences amélioré avec `_meta.refined_at` ajouté.

---

## POST /extract-context

Extrait le texte d'un fichier PDF, TXT ou MD pour l'utiliser comme contexte.

**Corps :** `multipart/form-data` avec champ `file`

**Types acceptés :** `.pdf`, `.txt`, `.md`

**Retour (JSON) :**
```json
{
  "text": "texte extrait...",
  "chars": 15420,
  "truncated": false
}
```

Limite : 80 000 caractères (tronqué si dépassé).

---

## POST /export/word

Génère et télécharge un fichier `.docx` structuré.

**Corps (JSON) :**
```json
{ "result": { /* objet exigences complet */ } }
```

**Retour :** fichier binaire `.docx`

Contenu : titre, périmètre, besoin fondamental, table EF (ID / Exigence / Priorité / Vérif. / Statut), table ENF, contraintes techniques, mots-clés.

---

## POST /export/excel

Génère et télécharge un fichier `.xlsx`.

**Corps (JSON) :**
```json
{ "result": { /* objet exigences complet */ } }
```

**Retour :** fichier binaire `.xlsx`

Feuilles :
1. **Synthèse** — informations générales + texte source
2. **Exigences Fonctionnelles** — avec listes déroulantes Priorité / Vérif. / Statut
3. **Contraintes Techniques**
4. **Exigences Non-Fonctionnelles** — si présentes
5. **Matrice Traçabilité** — toutes exigences + colonnes Référence Source / Responsable

---

## POST /export/doors-csv

Génère un CSV importable dans IBM DOORS Next.

**Corps (JSON) :**
```json
{ "result": { /* objet exigences complet */ } }
```

**Retour :** fichier `.csv` encodé UTF-8 avec BOM

Colonnes : `Identifier`, `Artifact Type`, `Module`, `Primary Text`, `Priority`, `Verification Method`, `Status`, `System`, `Fundamental Need`, `Created By`, `Created On`

**Import dans DOORS Next :** `Artefacts > Importer > Importer depuis un fichier CSV`

---

## POST /export/reqif

Génère un fichier ReqIF 1.0.1 importable dans IBM DOORS Next et Jazz ELM.

**Corps (JSON) :**
```json
{ "result": { /* objet exigences complet */ } }
```

**Retour :** fichier `.reqif` (XML)

Attributs définis : Texte, Type, Module, Priorité (enum), Vérification (enum).

**Import dans DOORS Next :** `Artefacts > Importer > Importer depuis ReqIF`

---

## GET /models

Liste les modèles Ollama disponibles localement.

**Retour (JSON) :**
```json
{
  "models": ["mistral:latest", "llama3:latest"]
}
```

En cas d'erreur Ollama, retourne `["llama3", "mistral"]` par défaut avec un champ `warning`.
