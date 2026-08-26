from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.staticfiles import StaticFiles
from fastapi.responses import HTMLResponse, JSONResponse
from fastapi.middleware.cors import CORSMiddleware
import anthropic
import pandas as pd
import openpyxl
import pdfplumber
import io
import os
import re
import json

app = FastAPI()

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)

app.mount("/static", StaticFiles(directory="static"), name="static")

ANTHROPIC_API_KEY = os.environ.get("ANTHROPIC_API_KEY", "")


# ── Parseurs fichiers ──────────────────────────────────────────────────────────

def parse_excel(content: bytes) -> tuple[list[dict], str]:
    """Lit le fichier Excel commande KFC — retourne (produits, texte brut)"""
    wb = openpyxl.load_workbook(io.BytesIO(content), data_only=True)

    # Chercher la feuille la plus récente (dernière dans la liste)
    sheet_name = wb.sheetnames[-1]
    ws = wb[sheet_name]

    produits = []
    raw_lines = [f"Feuille: {sheet_name}"]

    for row in ws.iter_rows(max_row=300, values_only=True):
        if not row[0] or not isinstance(row[0], str):
            continue
        nom = row[0].strip()
        if nom in ["Produit", "Date de passation de commande", "Date de livraison", ""]:
            continue
        # Colonne I (index 8) = quantité commandée
        qte = row[8] if len(row) > 8 else None
        try:
            qte_val = float(qte) if qte is not None else 0.0
            produits.append({"nom": nom, "qte": qte_val})
            raw_lines.append(f"{nom} | {qte_val}")
        except (TypeError, ValueError):
            pass

    return produits, "\n".join(raw_lines)


def parse_csv_bon(content: bytes) -> tuple[list[dict], str]:
    """Lit le CSV bon de validation F4R — retourne (produits, texte brut)"""
    text = content.decode("utf-8", errors="replace")

    # Header réel à la ligne 2 (index 1), ligne 0 = metadata
    df = pd.read_csv(
        io.StringIO(text),
        sep=None,
        engine="python",
        header=1,
        skiprows=[0],
    )

    produits = []
    raw_lines = []

    for _, row in df.iterrows():
        try:
            nom = str(row.get("Designation article", "")).strip()
            qte = float(row.get("Quantite", 0))
            prix = str(row.get("prix HT unitaire", "")).strip()
            if nom and nom != "nan":
                produits.append({"nom": nom, "qte": qte, "prix_ht": prix})
                raw_lines.append(f"{nom} | {qte} | {prix} €")
        except (TypeError, ValueError):
            pass

    return produits, "\n".join(raw_lines)


def parse_pdf_bon(content: bytes) -> tuple[list[dict], str]:
    """Lit un PDF bon de commande — retourne ([], texte brut pour l'IA)"""
    with pdfplumber.open(io.BytesIO(content)) as pdf:
        text = "\n".join(
            page.extract_text() or "" for page in pdf.pages[:15]
        )
    return [], text


def parse_excel_bon(content: bytes) -> tuple[list[dict], str]:
    """Lit un Excel bon de validation"""
    wb = openpyxl.load_workbook(io.BytesIO(content), data_only=True)
    ws = wb.active
    rows = list(ws.iter_rows(max_row=300, values_only=True))

    produits = []
    raw_lines = []
    for row in rows:
        if not row[0]:
            continue
        try:
            nom = str(row[0]).strip()
            qte = float(row[1]) if len(row) > 1 and row[1] is not None else 0
            produits.append({"nom": nom, "qte": qte})
            raw_lines.append(f"{nom} | {qte}")
        except (TypeError, ValueError):
            pass

    return produits, "\n".join(raw_lines)


# ── Normalisation & lettrage ───────────────────────────────────────────────────

def normalize(s: str) -> str:
    s = s.upper()
    s = re.sub(r"\s+", " ", s).strip()
    return s


def letter_products(commande: list[dict], bon: list[dict]) -> dict:
    """Lettrage fuzzy entre commande et bon — retourne le rapport structuré"""

    bon_index = {normalize(p["nom"]): p for p in bon}
    matched_bon_keys = set()
    lettrage = []

    for item in commande:
        norm_cmd = normalize(item["nom"])
        qte_cmd = item["qte"]

        # 1. Match exact
        match = bon_index.get(norm_cmd)
        match_key = norm_cmd if match else None

        # 2. Match fuzzy (mots clés communs ≥ 50 %)
        if not match:
            words_cmd = set(norm_cmd.split())
            best_score, best_key, best_item = 0, None, None
            for bk, bp in bon_index.items():
                words_bon = set(bk.split())
                common = words_cmd & words_bon
                score = len(common) / max(len(words_cmd), len(words_bon), 1)
                if score > best_score and score >= 0.45:
                    best_score, best_key, best_item = score, bk, bp
            if best_item:
                match, match_key = best_item, best_key

        if match:
            matched_bon_keys.add(match_key)
            qte_rec = match["qte"]
            ecart = qte_rec - qte_cmd
            if abs(ecart) < 0.01:
                statut = "OK"
            else:
                statut = "ECART_QTE"
        else:
            qte_rec = None
            ecart = None
            statut = "NON_PASSE" if qte_cmd == 0 else "MANQUANT"

        lettrage.append({
            "produit_commande": item["nom"],
            "produit_bon": match["nom"] if match else None,
            "qte_commandee": int(qte_cmd) if qte_cmd is not None else None,
            "qte_recue": int(qte_rec) if qte_rec is not None else None,
            "ecart": int(ecart) if ecart is not None else None,
            "statut": statut,
        })

    # Produits sur bon non présents dans la commande
    for bk, bp in bon_index.items():
        if bk not in matched_bon_keys:
            lettrage.append({
                "produit_commande": None,
                "produit_bon": bp["nom"],
                "qte_commandee": None,
                "qte_recue": int(bp["qte"]),
                "ecart": None,
                "statut": "NON_COMMANDE",
            })

    stats = {
        "total_commande": sum(1 for l in lettrage if l["produit_commande"]),
        "total_bon": len(bon),
        "ok": sum(1 for l in lettrage if l["statut"] == "OK"),
        "ecart_qte": sum(1 for l in lettrage if l["statut"] == "ECART_QTE"),
        "manquant": sum(1 for l in lettrage if l["statut"] == "MANQUANT"),
        "non_commande": sum(1 for l in lettrage if l["statut"] == "NON_COMMANDE"),
        "non_passe": sum(1 for l in lettrage if l["statut"] == "NON_PASSE"),
    }

    return {"lettrage": lettrage, "stats": stats}


# ── Synthèse IA ───────────────────────────────────────────────────────────────

def build_synthese_ia(lettrage_result: dict, commande_raw: str, bon_raw: str) -> str:
    if not ANTHROPIC_API_KEY:
        return "⚠️ Clé API Anthropic non configurée sur le serveur."

    client = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)

    stats = lettrage_result["stats"]
    issues = [l for l in lettrage_result["lettrage"] if l["statut"] not in ("OK", "NON_PASSE")]

    issues_txt = "\n".join(
        f"- [{l['statut']}] {l['produit_commande'] or l['produit_bon']} "
        f"(commandé: {l['qte_commandee']}, reçu: {l['qte_recue']}, écart: {l['ecart']})"
        for l in issues
    )

    prompt = f"""Tu es un expert en gestion opérationnelle pour un réseau de restaurants KFC franchisés.

Voici le résultat du lettrage automatique entre la commande et le bon de validation :

STATS :
- {stats['total_commande']} produits commandés
- {stats['total_bon']} produits sur le bon de validation
- {stats['ok']} lettrés sans écart ✓
- {stats['ecart_qte']} écarts de quantité
- {stats['manquant']} produits manquants sur le bon
- {stats['non_commande']} produits sur le bon non commandés

INCOHÉRENCES DÉTECTÉES :
{issues_txt if issues_txt else "Aucune incohérence majeure."}

Rédige une synthèse opérationnelle concise (5-7 lignes max) à destination du manager de restaurant.
Structure ta réponse en 3 parties avec ces titres exacts :
### Bilan global
### Points d'attention prioritaires
### Actions recommandées

Sois direct, concret, et utilise le vocabulaire terrain d'un responsable KFC."""

    message = client.messages.create(
        model="claude-sonnet-4-6",
        max_tokens=600,
        messages=[{"role": "user", "content": prompt}],
    )

    return message.content[0].text


# ── Routes API ─────────────────────────────────────────────────────────────────

@app.get("/", response_class=HTMLResponse)
async def index():
    with open("templates/index.html", "r", encoding="utf-8") as f:
        return f.read()


@app.post("/analyze")
async def analyze(
    commande: UploadFile = File(...),
    bon: UploadFile = File(...),
):
    try:
        commande_bytes = await commande.read()
        bon_bytes = await bon.read()

        # Parse commande
        ext_cmd = commande.filename.lower().split(".")[-1]
        if ext_cmd in ("xlsx", "xls"):
            commande_produits, commande_raw = parse_excel(commande_bytes)
        elif ext_cmd == "csv":
            df = pd.read_csv(io.BytesIO(commande_bytes), sep=None, engine="python", header=0)
            commande_produits = [
                {"nom": str(row.iloc[0]), "qte": float(row.iloc[1])}
                for _, row in df.iterrows()
                if len(row) >= 2
            ]
            commande_raw = df.to_string()
        else:
            raise HTTPException(400, "Format commande non supporté (xlsx, xls, csv)")

        # Parse bon
        ext_bon = bon.filename.lower().split(".")[-1]
        if ext_bon == "pdf":
            bon_produits, bon_raw = parse_pdf_bon(bon_bytes)
        elif ext_bon == "csv":
            bon_produits, bon_raw = parse_csv_bon(bon_bytes)
        elif ext_bon in ("xlsx", "xls"):
            bon_produits, bon_raw = parse_excel_bon(bon_bytes)
        else:
            raise HTTPException(400, "Format bon non supporté (pdf, csv, xlsx)")

        # Lettrage
        result = letter_products(commande_produits, bon_produits if bon_produits else [])

        # Synthèse IA (si bon parsé en produits structurés)
        synthese = build_synthese_ia(result, commande_raw, bon_raw)
        result["synthese"] = synthese

        return JSONResponse(result)

    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(500, f"Erreur serveur : {str(e)}")


@app.get("/health")
async def health():
    return {"status": "ok", "api_key_set": bool(ANTHROPIC_API_KEY)}
