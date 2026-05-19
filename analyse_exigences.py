#!/usr/bin/env python3
"""
analyse_exigences.py
--------------------
Serveur Flask local pour l'analyse de demandes clients en contexte naval/sous-marin.
Utilise Ollama (llama3 / mistral) pour extraire des exigences structurées.

Usage : python analyse_exigences.py
Puis ouvrir http://localhost:5000 dans un navigateur.
"""

import json
import re
import datetime
import ollama
from flask import Flask, request, jsonify, send_from_directory
from pathlib import Path

# Tentative d'import de json_repair (pip install json-repair)
# Si absent, on utilise le fallback manuel
try:
    from json_repair import repair_json
    JSON_REPAIR_AVAILABLE = True
except ImportError:
    JSON_REPAIR_AVAILABLE = False


def parse_json_robust(raw_text: str) -> dict:
    """
    Tente de parser le JSON retourné par le LLM avec plusieurs stratégies :
    1. json.loads direct
    2. Extraction du bloc {...} le plus large, puis json.loads
    3. json_repair si disponible
    4. Extraction manuelle des champs clés en dernier recours
    """
    # Nettoyage markdown
    text = re.sub(r"```json\s*", "", raw_text)
    text = re.sub(r"```\s*", "", text).strip()

    # Stratégie 1 — parse direct
    try:
        return json.loads(text)
    except json.JSONDecodeError:
        pass

    # Stratégie 2 — extraire le bloc JSON le plus long (du premier { au dernier })
    first = text.find("{")
    last = text.rfind("}")
    if first != -1 and last != -1 and last > first:
        candidate = text[first:last + 1]
        try:
            return json.loads(candidate)
        except json.JSONDecodeError:
            pass

        # Stratégie 3 — json_repair sur le candidat
        if JSON_REPAIR_AVAILABLE:
            try:
                repaired = repair_json(candidate)
                return json.loads(repaired)
            except Exception:
                pass

    # Stratégie 4 — json_repair sur le texte brut entier
    if JSON_REPAIR_AVAILABLE:
        try:
            repaired = repair_json(text)
            return json.loads(repaired)
        except Exception:
            pass

    # Aucune stratégie n'a fonctionné
    raise json.JSONDecodeError("Impossible de parser le JSON retourné par le modèle", text, 0)

# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------
DEFAULT_MODEL = "llama3"      # Modèle Ollama par défaut (llama3 ou mistral)
HOST = "127.0.0.1"
PORT = 5000
OUTPUT_DIR = Path("exports")  # Dossier de sortie pour les JSON exportés

OUTPUT_DIR.mkdir(exist_ok=True)

# ---------------------------------------------------------------------------
# Prompt système ultra-directif — domaine naval/sous-marin
# ---------------------------------------------------------------------------
SYSTEM_PROMPT = """Tu es un Ingénieur Système expert en télécommunications et systèmes sous-marins \
(sonar, VLF/ELF, acoustique sous-marine, étanchéité, systèmes embarqués navals).
Tu reçois une demande client brute. Tu dois l'analyser rigoureusement et extraire les éléments \
d'ingénierie système suivants.

RÉPONDS UNIQUEMENT EN JSON VALIDE, sans markdown, sans commentaires, sans texte avant ou après.

Le JSON doit respecter EXACTEMENT ce schéma :
{
  "perimetre": {
    "systeme_sous_analyse": "<nom court du système principal>",
    "acteurs_externes": ["<acteur1>", "<acteur2>"],
    "milieux_externes": ["<milieu1>", "<milieu2>"]
  },
  "besoin_fondamental": "<phrase unique décrivant la mission principale du système>",
  "exigences_fonctionnelles": [
    "Le système DOIT <verbe d'action> <objet> <condition optionnelle>."
  ],
  "contraintes_techniques": {
    "acoustique": "<contraintes acoustiques ou N/A>",
    "frequences": "<bandes de fréquences : VLF, ELF, SHF, etc. ou N/A>",
    "etancheite_pression": "<spécifications d'étanchéité/pression ou N/A>",
    "debit_latence": "<contraintes de débit et latence ou N/A>",
    "environnement_physique": "<température, corrosion, vibrations, etc. ou N/A>",
    "autres": ["<contrainte supplémentaire>"]
  },
  "mots_cles_domaine": ["<mot-clé1>", "<mot-clé2>"]
}

RÈGLES IMPÉRATIVES :
- Les exigences fonctionnelles commencent TOUJOURS par "Le système DOIT"
- Sois précis et technique. Utilise le vocabulaire du domaine (OTAN, MIL-STD, etc.) si pertinent.
- Si une information est absente dans la demande, mets "Non spécifié" (jamais null ni vide).
- Génère entre 4 et 10 exigences fonctionnelles selon la complexité.
- Milieux externes possibles : eau de mer, atmosphère, fond marin, navire de surface, satellite, etc.
- Acteurs externes possibles : opérateur, équipage, PC de commandement, systèmes tiers, etc.
"""

# ---------------------------------------------------------------------------
# Application Flask
# ---------------------------------------------------------------------------
app = Flask(__name__, static_folder=".")


@app.route("/")
def index():
    """Sert la page principale (index.html dans le même dossier)."""
    return send_from_directory(".", "index.html")


@app.route("/analyse", methods=["POST"])
def analyse():
    """
    Reçoit la demande client, l'envoie à Ollama, retourne le JSON structuré.

    Body JSON attendu :
        { "demande": "texte brut...", "model": "llama3" }

    Retourne :
        JSON structuré avec périmètre, besoin fondamental, exigences, contraintes.
    """
    data = request.get_json(force=True)
    demande = data.get("demande", "").strip()
    model = data.get("model", DEFAULT_MODEL).strip()

    if not demande:
        return jsonify({"error": "Demande vide."}), 400

    # Construction du message utilisateur injecté dans le prompt
    user_message = (
        "Voici la demande client à analyser :\n\n"
        "---\n"
        f"{demande}\n"
        "---\n\n"
        "Extrais toutes les informations selon le schéma JSON demandé."
    )

    try:
        # ------------------------------------------------------------------
        # Appel à Ollama via la bibliothèque officielle Python
        # ------------------------------------------------------------------
        response = ollama.chat(
            model=model,
            messages=[
                {"role": "system", "content": SYSTEM_PROMPT},
                {"role": "user",   "content": user_message},
            ],
            options={
                "temperature": 0.1,   # Faible température → réponses déterministes
                "top_p": 0.9,
                "num_predict": 4096,  # Augmenté pour éviter la troncature du JSON
                "num_ctx": 8192,      # Fenêtre de contexte élargie
            }
        )

        raw_text = response["message"]["content"]

        # Parsing robuste multi-stratégies (voir parse_json_robust)
        result = parse_json_robust(raw_text)

        # Ajout de métadonnées d'analyse
        result["_meta"] = {
            "model": model,
            "timestamp": datetime.datetime.now().isoformat(),
            "demande_source": demande[:200] + ("..." if len(demande) > 200 else "")
        }

        return jsonify(result)

    except ollama.ResponseError as e:
        # Erreur retournée par le serveur Ollama (modèle absent, timeout, etc.)
        return jsonify({"error": f"Erreur Ollama : {e.error}"}), 500

    except json.JSONDecodeError as e:
        return jsonify({
            "error": f"JSON invalide retourné par le modèle : {str(e)}",
            "raw": raw_text
        }), 500

    except Exception as e:
        return jsonify({"error": f"Erreur inattendue : {str(e)}"}), 500


@app.route("/export", methods=["POST"])
def export_json():
    """
    Exporte le résultat JSON dans un fichier horodaté dans exports/.

    Body JSON attendu :
        { "result": { ... } }

    Retourne :
        { "message": "...", "filename": "exigences_YYYYMMDD_HHMMSS.json" }
    """
    data = request.get_json(force=True)
    result = data.get("result")

    if not result:
        return jsonify({"error": "Rien à exporter."}), 400

    timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"exigences_{timestamp}.json"
    filepath = OUTPUT_DIR / filename

    with open(filepath, "w", encoding="utf-8") as f:
        json.dump(result, f, ensure_ascii=False, indent=2)

    return jsonify({
        "message": f"Fichier exporté avec succès.",
        "filename": filename,
        "path": str(filepath)
    })


@app.route("/models", methods=["GET"])
def list_models():
    """
    Retourne la liste des modèles disponibles dans l'instance Ollama locale.
    En cas d'échec, retourne une liste par défaut.
    """
    try:
        models_response = ollama.list()
        # ollama.list() retourne un dict avec la clé "models"
        models = models_response.get("models", [])
        # Chaque entrée peut avoir "name" ou "model" selon la version d'Ollama
        names = [m.get("name", m.get("model", "")) for m in models]
        names = [n for n in names if n]  # Filtre les entrées vides
        return jsonify({"models": names if names else ["llama3", "mistral"]})
    except Exception as e:
        # Ollama non démarré : on retourne des valeurs par défaut
        return jsonify({"models": ["llama3", "mistral"], "warning": str(e)})


# ---------------------------------------------------------------------------
# Point d'entrée principal
# ---------------------------------------------------------------------------
if __name__ == "__main__":
    print("=" * 62)
    print("   Analyseur d'Exigences — Domaine Naval / Sous-marin")
    print("=" * 62)
    print(f"   Modèle par défaut  : {DEFAULT_MODEL}")
    print(f"   Dossier exports    : ./{OUTPUT_DIR}/")
    print(f"   Interface web      : http://{HOST}:{PORT}")
    print("=" * 62)

    # Vérification qu'Ollama est accessible avant de démarrer Flask
    try:
        ollama.list()
        print("   [OK] Ollama est accessible.\n")
    except Exception:
        print("   [ATTENTION] Ollama ne semble pas démarré.")
        print("               Exécutez 'ollama serve' dans un terminal séparé.\n")

    app.run(host=HOST, port=PORT, debug=False)
