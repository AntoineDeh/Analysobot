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
SYSTEM_PROMPT = """You are a JSON API. You only output valid JSON. No explanations, no markdown, no text before or after the JSON object.

You are also an expert Systems Engineer specializing in naval/submarine telecommunications (sonar, VLF/ELF, underwater acoustics, watertight systems).

Analyze the client request and fill this exact JSON structure. Replace ALL placeholder values. Never output null. Use "Non specifie" if information is missing.

OUTPUT ONLY THIS JSON OBJECT:
{
  "perimetre": {
    "systeme_sous_analyse": "short system name here",
    "acteurs_externes": ["actor1", "actor2"],
    "milieux_externes": ["environment1", "environment2"]
  },
  "besoin_fondamental": "single sentence describing the main mission",
  "exigences_fonctionnelles": [
    "Le systeme DOIT perform action on object under condition.",
    "Le systeme DOIT perform second action.",
    "Le systeme DOIT perform third action.",
    "Le systeme DOIT perform fourth action."
  ],
  "contraintes_techniques": {
    "acoustique": "acoustic constraints or Non specifie",
    "frequences": "frequency bands VLF/ELF/etc or Non specifie",
    "etancheite_pression": "waterproofing and pressure specs or Non specifie",
    "debit_latence": "throughput and latency constraints or Non specifie",
    "environnement_physique": "temperature corrosion vibration or Non specifie",
    "autres": ["additional constraint if any"]
  },
  "mots_cles_domaine": ["keyword1", "keyword2", "keyword3"]
}

RULES:
- exigences_fonctionnelles items MUST start with "Le systeme DOIT"
- Generate 4 to 8 items in exigences_fonctionnelles
- All strings must be in French
- Output ONLY the JSON object, starting with { and ending with }
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

    # Construction du message utilisateur
    user_message = (
        f"Client request to analyze:\n\n{demande}\n\n"
        "Output the JSON object now:"
    )

    try:
        # ------------------------------------------------------------------
        # Appel à Ollama — prefill "{" force le modèle à démarrer en JSON pur
        # ------------------------------------------------------------------
        response = ollama.chat(
            model=model,
            messages=[
                {"role": "system",    "content": SYSTEM_PROMPT},
                {"role": "user",      "content": user_message},
                {"role": "assistant", "content": "{"},  # prefill : force JSON dès le 1er token
            ],
            options={
                "temperature": 0.1,
                "top_p": 0.9,
                "num_predict": 4096,
                "num_ctx": 8192,
            }
        )

        # Le modèle continue après "{" — on le réinjecte en tête
        raw_text = "{" + response["message"]["content"]

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
