# Configuration

## Paramètres serveur

Dans `analyse_exigences.py` :

| Variable | Défaut | Description |
|----------|--------|-------------|
| `DEFAULT_MODEL` | `"mistral"` | Modèle Ollama sélectionné par défaut dans l'interface |
| `HOST` | `"127.0.0.1"` | Adresse d'écoute du serveur Flask |
| `PORT` | `5000` | Port du serveur Flask |
| `OUTPUT_DIR` | `"exports"` | Dossier de sauvegarde des exports JSON serveur |

## Paramètres Ollama

Dans la fonction `analyse()` de `analyse_exigences.py` :

| Paramètre | Valeur | Description |
|-----------|--------|-------------|
| `temperature` | `0.1` | Déterminisme (0 = max répétable, 1 = créatif) |
| `top_p` | `0.9` | Diversité du sampling |
| `num_predict` | `4096` | Nombre maximum de tokens générés |
| `num_ctx` | `16384` | Fenêtre de contexte en tokens |

**Adapter `num_ctx` selon la VRAM disponible :**

| VRAM libre | num_ctx recommandé |
|------------|--------------------|
| > 3 Go | 16384 (défaut) |
| 2–3 Go | 8192 |
| < 2 Go | 4096 |

## Contexte PDF

Limite du texte extrait envoyé au modèle : **20 000 caractères**.

Modifiable dans `analyse_exigences.py` :
```python
parts.append(f"DOCUMENT DE CONTEXTE :\n{context[:20000]}")
```

## Historique frontend

Stocké dans `localStorage` du navigateur sous la clé `syreq_history`.
Capacité : 20 entrées.

Modifiable dans `index.html` :
```js
sessionHistory = [entry, ...sessionHistory].slice(0, 20);
```

## SYSTEM_PROMPT

Le prompt système est dans `analyse_exigences.py` — variable `SYSTEM_PROMPT`.

Points clés à adapter si changement de domaine :
- Profil expert (ligne 2–6)
- Standards applicables (STANAG, MIL-STD…)
- Exemples de déductions valides
- Exemples d'hypothèses du domaine
