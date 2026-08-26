From __future__ import annotations
From fastapi import FastAPI, UploadFile, File, HTTPException
From fastapi.staticfiles import StaticFiles
From fastapi.responses import HTMLResponse, JSONResponse
From fastapi.middleware.cors import CORSMiddleware
Import anthropic
Import pandas as pd
Import openpyxl
Import pdfplumber
Import io
Import os
Import re

App = FastAPI()
App.add_middleware(CORSMiddleware, allow_origins=[« * »], allow_methods=[« * »], allow_headers=[« * »])
App.mount(« /static », StaticFiles(directory= »static »), name= »static »)
ANTHROPIC_API_KEY = os.environ.get(« ANTHROPIC_API_KEY », « « )


# ── Parseurs ───────────────────────────────────────────────────────────────────

Def parse_excel(content : bytes) :
    Wb = openpyxl.load_workbook(io.BytesIO(content), data_only=True)
    Ws = wb[wb.sheetnames[-1]]
    Produits, lines = [], []
    For row in ws.iter_rows(max_row=300, values_only=True) :
        If not row[0] or not isinstance(row[0], str) : continue
        Nom = row[0].strip()
        If nom in [« Produit », »Date de passation de commande », »Date de livraison », » »] : continue
        Try :
            Qte = float(row[8]) if len(row) > 8 and row[8] is not None else 0.0
            Produits.append({« nom » : nom, « qte » : qte})
            Lines.append(f »{nom} | {qte} »)
        Except : pass
    Return produits, « \n ».join(lines)

Def parse_csv_bon(content : bytes) :
    Text = content.decode(« utf-8 », errors= »replace »)
    Df = pd.read_csv(io.StringIO(text), sep=None, engine= »python », header=1, skiprows=[0])
    Produits, lines = [], []
    For _, row in df.iterrows() :
        Try :
            Nom = str(row.get(« Designation article », » »)).strip()
            Qte = float(row.get(« Quantite », 0))
            If nom and nom != « nan » :
                Produits.append({« nom » : nom, « qte » : qte})
                Lines.append(f »{nom} | {qte} »)
        Except : pass
    Return produits, « \n ».join(lines)

Def parse_pdf_bon(content : bytes) :
    « « « 
    Parseur PDF pour bon de confirmation STEF (et formats similaires).
    Format attendu : Ligne | Article | Désignation | Temp. | Unité | Qté cdée | Qté Conf | Rupt
    « « « 
    # Pattern ligne produit : numéro ligne, éventuellement X, code article 6 chiffres,
    # désignation, température, COL, qté commandée, qté confirmée, rupture
    Pattern = re.compile(
        R’^(\d+)\s+( ?:X\s+) ?(\d{5,6})\s+(.+ ?)\s+(SEC|SURG|FRAIS)\s+COL\s+(\d+)\s+(\d+)\s+\d+\s*$’
    )

    With pdfplumber.open(io.BytesIO(content)) as pdf :
        Full_text = « \n ».join(page.extract_text() or « «  for page in pdf.pages)

    # Reconstituer les lignes fragmentées (désignation sur 2 lignes)
    Raw_lines = full_text.split(‘\n’)
    Merged = []
    I = 0
    While i < len(raw_lines) :
        Line = raw_lines[i].strip()
        If line and i + 1 < len(raw_lines) :
            Nxt = raw_lines[i + 1].strip()
            # Si la ligne suivante est une suite de désignation (pas un n° de ligne ni en-tête)
            If nxt and not re.match(r’^\d{2,3}\s’, nxt) and not re.match(
                R’^(Page|Ligne|Totaux|CONFIRMATION|Notre|Pour|STEF|KFC|Code|Commandé|Livré)’, nxt
            ) :
                Line = line + ‘ ‘ + nxt
                I += 1
        Merged.append(line)
        I += 1

    Produits = []
    Raw_lines_out = []
    Seen_articles = set()

    For line in merged :
        M = pattern.match(line)
        If m :
            Article = m.group(2)
            Designation = m.group(3).strip()
            Qte_conf = int(m.group(6))
            # Dédupliquer par code article (les lignes « X » sont des anciens tarifs)
            If article not in seen_articles :
                Seen_articles.add(article)
                Produits.append({« nom » : designation, « qte » : qte_conf, « article » : article})
                Raw_lines_out.append(f »{designation} | {qte_conf} »)

    # Fallback texte brut si aucun produit parsé (PDF image ou format inconnu)
    If not produits :
        Return [], full_text

    Return produits, « \n ».join(raw_lines_out)

Def parse_excel_bon(content : bytes) :
    Wb = openpyxl.load_workbook(io.BytesIO(content), data_only=True)
    Ws = wb.active
    Produits, lines = [], []
    For row in ws.iter_rows(max_row=300, values_only=True) :
        If not row[0] : continue
        Try :
            Nom = str(row[0]).strip()
            Qte = float(row[1]) if len(row) > 1 and row[1] is not None else 0
            Produits.append({« nom » : nom, « qte » : qte})
            Lines.append(f »{nom} | {qte} »)
        Except : pass
    Return produits, « \n ».join(lines)


# ── Matching ───────────────────────────────────────────────────────────────────

# Mots vides : conditionnements, unités, prépositions
STOP = {
    ‘X’,’G’,’KG’,’ML’,’CL’,’L’,’PCS’,’PC’,’SAC’,’SACS’,’BTE’,’BOITE’,’BOITES’,
    ‘COLIS’,’LOT’,’PACK’,’SACHET’,’SACHETS’,’PM’,’GM’,’TENDER’,’FORMAT’,
    ‘DE’,’DU’,’LE’,’LA’,’LES’,’ET’,’EN’,’AU’,’AUX’,
    ‘KRUNCHY’,’PREMAR’,’LIGHT’,’SANS’,’SUCRE’,’PET’,’CODE’,’UNIQUE’,
}

Def normalize(s : str) -> str :
    S = s.upper()
    S = re.sub(r’[^\w\s]’, ‘ ‘, s)
    Return re.sub(r’\s+’, ‘ ‘, s).strip()

Def extract_dimensions(s : str) -> set :
    « « « 
    Extrait les tokens discriminants de taille/contenance/quantité :
    33CL, 50CL, 33G, 500G, GRANDE, MOYENNE, PETITE, MINI, LARGE, etc.
    Ces tokens doivent correspondre entre commande et bon pour valider le match.
    « « « 
    S = s.upper()
    Dims = set()
    # Dimensions numériques : 33CL, 50CL, 700G, 500G, 2500G, 30CM…
    For m in re.finditer(r’\b(\d+( ?:[.,]\d+) ?)\s*(CL|ML|KG|G|CM|L|OZ)\b’, s) :
        Dims.add(m.group(1).replace(‘,’,’.’) + m.group(2))
    # Quantités conditionnement : x10pcs, x12pcs, x1000pcs, x24pcs
    For m in re.finditer(r’[Xx]\s*(\d+)\s*( ?:PCS|pcs) ?\b’, s) :
        Dims.add(‘QTE’ + m.group(1))
    # Mots de taille
    For mot in [‘GRANDE’,’MOYEN’,’MOYENNE’,’PETITE’,’PETIT’,’LARGE’,’MAXI’,’BIG’] :
        If re.search(r’\b’ + mot + r’\b’, s) :
            Dims.add(mot)
    Return dims

Def keywords(s : str) -> list :
    Words = normalize(s).split()
    Result = []
    For w in words :
        If w in STOP or w in (‘-‘,’–‘,’/’,’&’) : continue
        If re.match(r’^\d+$’, w) : continue
        # Tokens conditionnement : 6X18PCS, 3X24PCS, etc.
        If re.match(r’^\d+[X×]\d+.*$’, w) or re.match(r’^\d+X$’, w) : continue
        # Dimensions seules déjà capturées : 30CM, 25CL, 50CL
        If re.match(r’^\d+[A-Z]+$’, w) : continue
        Result.append(w)
    Return result

Def dimensions_compatibles(a : str, b : str) -> bool :
    « « « 
    Vérifie que les dimensions discriminantes sont compatibles.
    Si les deux ont des dimensions et qu’elles diffèrent → incompatible.
    Si l’un n’a pas de dimension → compatible (dimension absente = pas discriminant).
    « « « 
    Dims_a = extract_dimensions(a)
    Dims_b = extract_dimensions(b)
    If not dims_a or not dims_b :
        Return True
    # Au moins une dimension commune (ou aucune contradiction)
    Return bool(dims_a & dims_b) or not (dims_a – dims_b) or not (dims_b – dims_a)

Def score_match(a : str, b : str) :
    « « « 
    Retourne (score, dimensions_ok).
    Le score seul ne suffit pas : on vérifie aussi la compatibilité des dimensions.
    « « « 
    Kw_a = keywords(a)
    Kw_b = keywords(b)
    If not kw_a or not kw_b :
        Return 0.0, False

    Set_a, set_b = set(kw_a), set(kw_b)
    Inter = set_a & set_b
    Union = set_a | set_b

    # Jaccard
    Jaccard = len(inter) / len(union) if union else 0

    # Containment : tous les mots du plus court sont dans le plus long
    Shorter = set_a if len(set_a) <= len(set_b) else set_b
    Containment = len(inter) / len(shorter) if shorter else 0

    # Préfixe : 2 premiers mots-clés du plus court dans le plus long
    Kw_short = kw_a if len(kw_a) <= len(kw_b) else kw_b
    Kw_long  = kw_b if len(kw_a) <= len(kw_b) else kw_a
    Pfx = min(2, len(kw_short))
    Prefix_score = sum(1 for w in kw_short[ :pfx] if w in set(kw_long)) / pfx if pfx else 0

    Score = max(jaccard, containment * 0.88, prefix_score * 0.92)
    Dims_ok = dimensions_compatibles(a, b)

    Return score, dims_ok

SEUIL_OK       = 0.72
SEUIL_APPROCHE = 0.55

Def letter_products(commande : list, bon : list) -> dict :
    Bon_index = {normalize(p[« nom »]) : p for p in bon}
    Matched_bon = set()
    Lettrage = []

    For item in commande :
        Norm_cmd = normalize(item[« nom »])
        Qte_cmd  = item[« qte »]

        # Chercher le meilleur match avec contrainte de dimensions
        Best_score, best_key, best_item = 0.0, None, None
        For bk, bp in bon_index.items() :
            S, dims_ok = score_match(norm_cmd, bk)
            # Pénaliser fortement si dimensions incompatibles
            If not dims_ok :
                S *= 0.4
            If s > best_score :
                Best_score, best_key, best_item = s, bk, bp

        If best_score >= SEUIL_OK :
            Matched_bon.add(best_key)
            Qte_rec = best_item[« qte »]
            Ecart   = qte_rec – qte_cmd
            Statut  = « OK » if abs(ecart) < 0.01 else « ECART_QTE »
            Lettrage.append({
                « produit_commande » : item[« nom »],
                « produit_bon » :      best_item[« nom »],
                « qte_commandee » :    int(qte_cmd),
                « qte_recue » :        int(qte_rec),
                « ecart » :            int(ecart),
                « statut » :           statut,
                « score » :            round(best_score, 2),
            })
        Elif best_score >= SEUIL_APPROCHE :
            Matched_bon.add(best_key)
            Qte_rec = best_item[« qte »]
            Ecart   = qte_rec – qte_cmd
            Lettrage.append({
                « produit_commande » : item[« nom »],
                « produit_bon » :      best_item[« nom »],
                « qte_commandee » :    int(qte_cmd),
                « qte_recue » :        int(qte_rec),
                « ecart » :            int(ecart),
                « statut » :           « APPROCHE »,
                « score » :            round(best_score, 2),
            })
        Else :
            Statut = « NON_PASSE » if qte_cmd == 0 else « MANQUANT »
            Lettrage.append({
                « produit_commande » : item[« nom »],
                « produit_bon » :      None,
                « qte_commandee » :    int(qte_cmd) if qte_cmd else 0,
                « qte_recue » :        None,
                « ecart » :            None,
                « statut » :           statut,
                « score » :            0,
            })

    # Produits du bon non matchés
    For bk, bp in bon_index.items() :
        If bk not in matched_bon :
            Lettrage.append({
                « produit_commande » : None,
                « produit_bon » :      bp[« nom »],
                « qte_commandee » :    None,
                « qte_recue » :        int(bp[« qte »]),
                « ecart » :            None,
                « statut » :           « NON_COMMANDE »,
                « score » :            0,
            })

    Stats = {
        « total_commande » : sum(1 for l in lettrage if l[« produit_commande »]),
        « total_bon » :      len(bon),
        « ok » :             sum(1 for l in lettrage if l[« statut »] == « OK »),
        « ecart_qte » :      sum(1 for l in lettrage if l[« statut »] == « ECART_QTE »),
        « approche » :       sum(1 for l in lettrage if l[« statut »] == « APPROCHE »),
        « manquant » :       sum(1 for l in lettrage if l[« statut »] == « MANQUANT »),
        « non_commande » :   sum(1 for l in lettrage if l[« statut »] == « NON_COMMANDE »),
        « non_passe » :      sum(1 for l in lettrage if l[« statut »] == « NON_PASSE »),
    }
    Return {« lettrage » : lettrage, « stats » : stats}


# ── Synthèse IA ────────────────────────────────────────────────────────────────

Def build_synthese_ia(result : dict) -> str :
    If not ANTHROPIC_API_KEY :
        Return « ⚠️ Clé API Anthropic non configurée. »
    Client = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)
    S = result[« stats »]
    Issues = [l for l in result[« lettrage »] if l[« statut »] not in (« OK », »NON_PASSE »)]
    Issues_txt = « \n ».join(
        F »- [{l[‘statut’]}] cmd=’{l[‘produit_commande’] or ‘’}’ / bon=’{l[‘produit_bon’] or ‘’}’ « 
        F »(commandé :{l[‘qte_commandee’]}, reçu :{l[‘qte_recue’]}, écart :{l[‘ecart’]}) »
        For l in issues
    )
    Prompt = f » » »Tu es expert en gestion opérationnelle KFC franchisé.

Lettrage commande vs bon de validation :
- {s[‘total_commande’]} produits commandés / {s[‘total_bon’]} sur le bon
- {s[‘ok’]} lettrés OK ✓ | {s[‘approche’]} à vérifier | {s[‘ecart_qte’]} écarts qté | {s[‘manquant’]} manquants | {s[‘non_commande’]} non commandés

INCOHÉRENCES :
{issues_txt or ‘Aucune.’}

Synthèse opérationnelle concise pour le manager. 3 parties :
### Bilan global
### Points d’attention prioritaires  
### Actions recommandées
Maximum 8 lignes. » » »
    Msg = client.messages.create(model= »claude-sonnet-4-6 », max_tokens=600,
                                  Messages=[{« role » : »user », »content » :prompt}])
    Return msg.content[0].text


# ── Routes ─────────────────────────────────────────────────────────────────────

@app.get(« / », response_class=HTMLResponse)
Async def index() :
    With open(« templates/index.html », »r »,encoding= »utf-8 ») as f : return f.read()

@app.post(« /analyze »)
Async def analyze(commande : UploadFile = File(…), bon : UploadFile = File(…)) :
    Try :
        Cb = await commande.read()
        Bb = await bon.read()
        Ext_c = commande.filename.lower().split(« . »)[-1]
        If ext_c in (« xlsx », »xls ») : cp, cr = parse_excel(cb)
        Elif ext_c == « csv » :
            Df = pd.read_csv(io.BytesIO(cb), sep=None, engine= »python », header=0)
            Cp = [{« nom » :str(r.iloc[0]), »qte » :float(r.iloc[1])} for _,r in df.iterrows() if len®>=2]
            Cr = df.to_string()
        Else : raise HTTPException(400, »Format commande non supporté »)
        Ext_b = bon.filename.lower().split(« . »)[-1]
        If ext_b == « pdf » : bp, br = parse_pdf_bon(bb)
        Elif ext_b == « csv » : bp, br = parse_csv_bon(bb)
        Elif ext_b in (« xlsx », »xls ») : bp, br = parse_excel_bon(bb)
        Else : raise HTTPException(400, »Format bon non supporté »)
        Result = letter_products(cp, bp if bp else [])
        Result[« synthese »] = build_synthese_ia(result)
        Return JSONResponse(result)
    Except HTTPException : raise
    Except Exception as e : raise HTTPException(500, f »Erreur : {str€} »)

@app.get(« /health »)
Async def health() :
    Return {« status » : »ok », »api_key_set » :bool(ANTHROPIC_API_KEY)}
