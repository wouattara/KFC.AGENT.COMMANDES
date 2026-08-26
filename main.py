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
    wb = openpyxl.load_workbook(io.BytesIO(content), data_only=True)
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
        qte = row[8] if len(row) > 8 else None
        try:
            qte_val = float(qte) if qte is not None else 0.0
            produits.append({"nom": nom, "qte": qte_val})
            raw_lines.append(f"{nom} | {qte_val}")
        except (TypeError, ValueError):
            pass
    return produits, "\n".join(raw_lines)


def parse_csv_bon(content: bytes) -> tuple[list[dict], str]:
    text = content.decode("utf-8", errors="replace")
    df = pd.read_csv(io.StringIO(text), sep=None, engine="python", header=1, skiprows=[0])
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
    with pdfplumber.open(io.BytesIO(content)) as pdf:
        text = "\n".join(page.extract_text() or "" for page in pdf.pages[:15])
    return [], text


def parse_excel_bon(content: bytes) -> tuple[list[dict], str]:
    wb = openpyxl.load_workbook(io.BytesIO(content), data_only=True)
    ws = wb.active
    produits = []
    raw_lines = []
    for row in ws.iter_rows(max_row=300, values_only=True):
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


# ── Normalisation ──────────────────────────────────────────────────────────────

# Mots parasites à ignorer pour le matching (conditionnements, unités, etc.)
STOP_WORDS = {
    'X', 'G', 'KG', 'ML', 'CL', 'L', 'PCS', 'PC', 'SAC', 'SACS',
    'BTE', 'BOITE', 'BOITES', 'COLIS', 'LOT', 'PACK', 'SACHET', 'SACHETS',
    'PM', 'GM', 'TENDER', 'PREMAR', 'FORMAT',
    '1', '2', '3', '4', '5', '6', '7', '8', '9', '10',
    '12', '18', '20', '24', '36', '48', '50', '55', '83', '100', '200',
    '500', '700', '760', '1000', '0',
}

def normalize(s: str) -> str:
    s = s.upper()
    # Supprimer les caractères spéciaux sauf espaces et tirets
    s = re.sub(r'[^\w\s\-]', ' ', s)
    s = re.sub(r'\s+', ' ', s).strip()
    return s

def keywords(s: str) -> list[str]:
    """Extrait les mots-clés significatifs (sans stop words ni nombres)"""
    words = normalize(s).split()
    result = []
    for w in words:
        # Ignorer ponctuations seules, tirets, chiffres purs
        if w in ('-', '–', '/') or re.match(r'^\d+$', w):
            continue
        # Ignorer stop words
        if w in STOP_WORDS:
            continue
        # Ignorer les tokens conditionnement : 6X18PCS, 3X24PCS, 8X500G, etc.
        if re.match(r'^\d+[X×]\d+.*$', w) or re.match(r'^\d+X$', w):
            continue
        # Ignorer dimensions seules : 30CM, 25CL, 50CL, 66OZ
        if re.match(r'^\d+[A-Z]+$', w):
            continue
        result.append(w)
    return result

def score_match(a: str, b: str) -> float:
    """
    Score de similarité entre deux noms de produits.
    Retourne un score entre 0 et 1.
    Privilégie les débuts de nom (les premiers mots sont les plus discriminants).
    """
    kw_a = keywords(a)
    kw_b = keywords(b)

    if not kw_a or not kw_b:
        return 0.0

    # Score 1 : mots en commun / union (Jaccard)
    set_a, set_b = set(kw_a), set(kw_b)
    intersection = set_a & set_b
    union = set_a | set_b
    jaccard = len(intersection) / len(union) if union else 0

    # Score 2 : préfixe — les N premiers mots du plus court sont-ils dans l'autre ?
    shorter = kw_a if len(kw_a) <= len(kw_b) else kw_b
    longer  = kw_b if len(kw_a) <= len(kw_b) else kw_a
    prefix_len = min(2, len(shorter))  # On regarde les 2 premiers mots clés
    prefix_matches = sum(1 for w in shorter[:prefix_len] if w in set(longer))
    prefix_score = prefix_matches / prefix_len if prefix_len > 0 else 0

    # Score 3 : containment — tous les mots du plus court sont dans le plus long ?
    containment = len(intersection) / len(set(shorter)) if shorter else 0

    # Score final = max des 3 approches pondérées
    return max(jaccard, prefix_score * 0.9, containment * 0.85)


# ── Lettrage ──────────────────────────────────────────────────────────────────

SEUIL_OK        = 0.72   # Match fiable → lettré automatiquement
SEUIL_APPROCHE  = 0.55   # Match possible → signalé "APPROCHE" pour révision

def letter_products(commande: list[dict], bon: list[dict]) -> dict:
    bon_items = {normalize(p["nom"]): p for p in bon}
    matched_bon_keys = set()
    lettrage = []

    for item in commande:
        norm_cmd  = normalize(item["nom"])
        qte_cmd   = item["qte"]

        # Chercher le meilleur match dans le bon
        best_score, best_key, best_item = 0.0, None, None
        for bk, bp in bon_items.items():
            s = score_match(norm_cmd, bk)
            if s > best_score:
                best_score, best_key, best_item = s, bk, bp

        if best_score >= SEUIL_OK:
            # Match fiable
            matched_bon_keys.add(best_key)
            qte_rec = best_item["qte"]
            ecart   = qte_rec - qte_cmd
            statut  = "OK" if abs(ecart) < 0.01 else "ECART_QTE"
            lettrage.append({
                "produit_commande": item["nom"],
                "produit_bon":      best_item["nom"],
                "qte_commandee":    int(qte_cmd),
                "qte_recue":        int(qte_rec),
                "ecart":            int(ecart),
                "statut":           statut,
                "score":            round(best_score, 2),
            })

        elif best_score >= SEUIL_APPROCHE:
            # Match probable — à vérifier
            matched_bon_keys.add(best_key)
            qte_rec = best_item["qte"]
            ecart   = qte_rec - qte_cmd
            lettrage.append({
                "produit_commande": item["nom"],
                "produit_bon":      best_item["nom"],
                "qte_commandee":    int(qte_cmd),
                "qte_recue":        int(qte_rec),
                "ecart":            int(ecart),
                "statut":           "APPROCHE",
                "score":            round(best_score, 2),
            })

        else:
            # Aucun match
            statut = "NON_PASSE" if qte_cmd == 0 else "MANQUANT"
            lettrage.append({
                "produit_commande": item["nom"],
                "produit_bon":      None,
                "qte_commandee":    int(qte_cmd) if qte_cmd else 0,
                "qte_recue":        None,
                "ecart":            None,
                "statut":           statut,
                "score":            0,
            })

    # Produits sur le bon non matchés
    for bk, bp in bon_items.items():
        if bk not in matched_bon_keys:
            lettrage.append({
                "produit_commande": None,
                "produit_bon":      bp["nom"],
                "qte_commandee":    None,
                "qte_recue":        int(bp["qte"]),
                "ecart":            None,
                "statut":           "NON_COMMANDE",
                "score":            0,
            })

    stats = {
        "total_commande":  sum(1 for l in lettrage if l["produit_commande"]),
        "total_bon":       len(bon),
        "ok":              sum(1 for l in lettrage if l["statut"] == "OK"),
        "ecart_qte":       sum(1 for l in lettrage if l["statut"] == "ECART_QTE"),
        "approche":        sum(1 for l in lettrage if l["statut"] == "APPROCHE"),
        "manquant":        sum(1 for l in lettrage if l["statut"] == "MANQUANT"),
        "non_commande":    sum(1 for l in lettrage if l["statut"] == "NON_COMMANDE"),
        "non_passe":       sum(1 for l in lettrage if l["statut"] == "NON_PASSE"),
    }

    return {"lettrage": lettrage, "stats": stats}


# ── Synthèse IA ───────────────────────────────────────────────────────────────

def build_synthese_ia(lettrage_result: dict) -> str:
    if not ANTHROPIC_API_KEY:
        return "⚠️ Clé API Anthropic non configurée sur le serveur."

    client = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)
    stats  = lettrage_result["stats"]
    issues = [l for l in lettrage_result["lettrage"] if l["statut"] not in ("OK", "NON_PASSE")]

    issues_txt = "\n".join(
        f"- [{l['statut']}] commande='{l['produit_commande'] or ''}' / bon='{l['produit_bon'] or ''}' "
        f"(commandé: {l['qte_commandee']}, reçu: {l['qte_recue']}, écart: {l['ecart']})"
        for l in issues
    )

    prompt = f"""Tu es un expert en gestion opérationnelle pour un réseau de restaurants KFC franchisés.

Résultat du lettrage automatique commande vs bon de validation :

STATS :
- {stats['total_commande']} produits commandés / {stats['total_bon']} sur le bon
- {stats['ok']} lettrés OK ✓
- {stats['approche']} rapprochements à vérifier (noms différents entre commande et bon)
- {stats['ecart_qte']} écarts de quantité
- {stats['manquant']} produits manquants sur le bon
- {stats['non_commande']} produits sur le bon non commandés

INCOHÉRENCES ET POINTS D'ATTENTION :
{issues_txt if issues_txt else "Aucune incohérence majeure."}

Rédige une synthèse opérationnelle concise à destination du manager de restaurant.
Structure en 3 parties avec ces titres exacts :
### Bilan global
### Points d'attention prioritaires
### Actions recommandées

Sois direct et concret. Maximum 8 lignes au total."""

    message = client.messages.create(
        model="claude-sonnet-4-6",
        max_tokens=600,
        messages=[{"role": "user", "content": prompt}],
    )
    return message.content[0].text


# ── Routes ─────────────────────────────────────────────────────────────────────

@app.get("/", response_class=HTMLResponse)
async def index():
    with open("templates/index.html", "r", encoding="utf-8") as f:
        return f.read()


@app.post("/analyze")
async def analyze(
    commande: UploadFile = File(...),
    bon:      UploadFile = File(...),
):
    try:
        commande_bytes = await commande.read()
        bon_bytes      = await bon.read()

        ext_cmd = commande.filename.lower().split(".")[-1]
        if ext_cmd in ("xlsx", "xls"):
            commande_produits, commande_raw = parse_excel(commande_bytes)
        elif ext_cmd == "csv":
            df = pd.read_csv(io.BytesIO(commande_bytes), sep=None, engine="python", header=0)
            commande_produits = [
                {"nom": str(row.iloc[0]), "qte": float(row.iloc[1])}
                for _, row in df.iterrows() if len(row) >= 2
            ]
            commande_raw = df.to_string()
        else:
            raise HTTPException(400, "Format commande non supporté (xlsx, xls, csv)")

        ext_bon = bon.filename.lower().split(".")[-1]
        if ext_bon == "pdf":
            bon_produits, bon_raw = parse_pdf_bon(bon_bytes)
        elif ext_bon == "csv":
            bon_produits, bon_raw = parse_csv_bon(bon_bytes)
        elif ext_bon in ("xlsx", "xls"):
            bon_produits, bon_raw = parse_excel_bon(bon_bytes)
        else:
            raise HTTPException(400, "Format bon non supporté (pdf, csv, xlsx)")

        result   = letter_products(commande_produits, bon_produits if bon_produits else [])
        synthese = build_synthese_ia(result)
        result["synthese"] = synthese

        return JSONResponse(result)

    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(500, f"Erreur serveur : {str(e)}")


@app.get("/health")
async def health():
    return {"status": "ok", "api_key_set": bool(ANTHROPIC_API_KEY)}
