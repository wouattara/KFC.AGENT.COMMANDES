# KFC Agent IA — Vérification Commandes

Application web pour analyser et lettrer automatiquement les fichiers de commande KFC
contre les bons de validation fournisseur.

## Déploiement sur Railway (5 minutes)

### 1. Pousser sur GitHub

```bash
git init
git add .
git commit -m "init kfc-agent"
git branch -M main
git remote add origin https://github.com/TON_USER/kfc-agent.git
git push -u origin main
```

### 2. Créer le projet sur Railway

1. Aller sur **railway.app** → "New Project"
2. Choisir **"Deploy from GitHub repo"**
3. Sélectionner ton repo `kfc-agent`
4. Railway détecte automatiquement le `Procfile` → déploiement lancé

### 3. Ajouter la clé API Anthropic

Dans Railway → ton projet → onglet **Variables** :

| Clé | Valeur |
|-----|--------|
| `ANTHROPIC_API_KEY` | `sk-ant-...` (ta clé Anthropic) |

### 4. Partager le lien

Une fois déployé, Railway génère une URL du type :
```
https://kfc-agent-production.up.railway.app
```

Partage cette URL à tes managers — accessible depuis n'importe quel appareil.

---

## Fichiers supportés

| Fichier | Format |
|---------|--------|
| Commande | `.xlsx`, `.xls`, `.csv` |
| Bon de validation | `.pdf`, `.csv`, `.xlsx` |

## Structure projet

```
kfc-agent/
├── main.py          # Backend FastAPI
├── templates/
│   └── index.html   # Frontend
├── static/          # Fichiers statiques (vide pour l'instant)
├── requirements.txt
├── Procfile
└── README.md
```

## Variables d'environnement

| Variable | Description |
|----------|-------------|
| `ANTHROPIC_API_KEY` | Clé API Anthropic (obligatoire) |
| `PORT` | Port serveur (géré automatiquement par Railway) |
