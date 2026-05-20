# Structure JSON des exigences

Format retourné par `/analyse` et attendu par tous les endpoints `/export/*`.

---

## Schéma complet

```json
{
  "perimetre": {
    "systeme_sous_analyse": "string — nom court du système",
    "acteurs_externes": ["string", "..."],
    "milieux_externes": ["string", "..."]
  },

  "besoin_fondamental": "string — mission principale en une phrase",

  "exigences_fonctionnelles": [
    {
      "id":           "string — identifiant stable (ex: EF-A3F2)",
      "texte":        "string — commence par 'Le système DOIT'",
      "origine":      "Explicite | Déduite | Hypothèse",
      "justification":"string — citation ou raisonnement",
      "priority":     "Must | Should | Could | Won't",
      "verification": "Test | Analyse | Inspection | Démonstration",
      "statut":       "À confirmer | Confirmé | Rejeté | En révision"
    }
  ],

  "exigences_non_fonctionnelles": [
    {
      "id":           "string — identifiant stable (ex: ENF-B1C4)",
      "texte":        "string — commence par 'Le système DOIT'",
      "type":         "Performance | Fiabilite | Interface | Physique | Securite",
      "origine":      "Explicite | Déduite | Hypothèse",
      "justification":"string",
      "priority":     "Must | Should | Could | Won't",
      "verification": "Test | Analyse | Inspection | Démonstration",
      "statut":       "À confirmer | Confirmé | Rejeté | En révision"
    }
  ],

  "contraintes_techniques": {
    "acoustique":             "string",
    "frequences":             "string",
    "etancheite_pression":    "string",
    "debit_latence":          "string",
    "environnement_physique": "string",
    "autres":                 ["string", "..."]
  },

  "mots_cles_domaine": ["string", "..."],

  "_meta": {
    "model":          "string — modèle utilisé",
    "timestamp":      "string — ISO 8601",
    "demande_source": "string — 200 premiers caractères de la demande",
    "demande_full":   "string — texte complet de la demande",
    "context_file":   "string — nom du fichier PDF de contexte",
    "refined_at":     "string — ISO 8601, présent si affiné"
  }
}
```

---

## Règles métier

### origine

| Valeur | Signification | Justification attendue |
|--------|---------------|----------------------|
| `Explicite` | Information présente mot pour mot dans le texte | Citation entre guillemets `«...»` |
| `Déduite` | Implication logique d'une contrainte explicite | Raisonnement technique (`400 m → 40 bar`) |
| `Hypothèse` | Standard du domaine non mentionné | `"Hypothèse domaine : ... A valider."` |

### Détection de fausse classification (frontend)

Une exigence `Explicite` est marquée `[?]` (suspect) si :
- La justification contient le texte `"citation exacte"` (placeholder non rempli)
- La justification ne contient pas de guillemets `«»`
- La justification est vide ou < 5 caractères

### id

Identifiant stable généré côté frontend au format `EF-XXXX` ou `ENF-XXXX` (XXXX = 4 hex aléatoires). Ne change pas lors des réordonnements ou suppressions d'autres items.

Le numéro d'affichage (`EF-01`, `EF-02`…) est calculé dynamiquement et ne fait pas partie du champ `id`.

---

## Compatibilité ascendante

Le parsing dans `renderResult()` gère les formats anciens :
- Exigence sous forme de chaîne simple → convertie en objet avec valeurs par défaut
- Champ `texte` (prompt) ou `text` (interne) → normalisé vers `text`
- `extrait_source` (ancien nom) → mappé vers `justification`
