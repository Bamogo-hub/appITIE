# Système de Gestion des Bénéficiaires Effectifs
## Ministère de la Justice — République du Sénégal

---

## 📋 Description

Application web institutionnelle pour la collecte et la gestion des déclarations de bénéficiaires effectifs des entreprises, conformément aux exigences légales du Sénégal.

---

## 🚀 Installation

### Prérequis
- Python 3.9 ou supérieur
- pip (gestionnaire de paquets Python)

### Étapes d'installation

1. **Décompresser le fichier ZIP**
   ```bash
   unzip beneficiaires_effectifs.zip
   cd beneficiaires_effectifs
   ```

2. **Créer un environnement virtuel** (recommandé)
   ```bash
   python -m venv venv
   
   # Windows
   venv\Scripts\activate
   
   # Linux / macOS
   source venv/bin/activate
   ```

3. **Installer les dépendances**
   ```bash
   pip install -r requirements.txt
   ```

4. **Lancer l'application**
   ```bash
   python app.py
   ```

5. **Accéder à l'application**
   Ouvrir votre navigateur à l'adresse : http://localhost:5000

---

## 🔐 Compte administrateur par défaut

| Email | Mot de passe |
|-------|-------------|
| admin@justice.sn | Admin@2024 |

> ⚠️ Changez le mot de passe dès la première connexion.

---

## 📁 Structure des fichiers

```
beneficiaires_effectifs/
├── app.py                    # Application Flask principale
├── requirements.txt          # Dépendances Python
├── README.md                 # Cette documentation
├── templates/                # Pages HTML
│   ├── base.html            # Template de base (navbar, footer)
│   ├── login.html           # Page de connexion
│   ├── dashboard.html       # Tableau de bord
│   ├── declaration.html     # Formulaire de déclaration
│   ├── liste.html           # Liste des déclarations
│   ├── detail.html          # Détail d'une déclaration
│   └── admin.html           # Administration
├── static/
│   ├── css/style.css        # Styles CSS institutionnels
│   └── js/app.js            # JavaScript
├── uploads/                  # Fichiers téléchargés (créé automatiquement)
└── exports/                  # Exports Excel (créé automatiquement)
```

---

## 🔧 Fonctionnalités

### Authentification
- Connexion sécurisée email/mot de passe
- Mots de passe chiffrés (bcrypt)
- Gestion des rôles : Administrateur / Utilisateur

### Formulaire de déclaration (5 sections)
- **A** — Informations sur l'entité déclarante
- **B** — Informations sur le bénéficiaire effectif
- **C** — Participation et contrôle
- **D** — Documents justificatifs (PDF, ZIP, DOCX, XLSX, max 100 Mo)
- **E** — Certification complète

### Tableau de bord
- Statistiques en temps réel
- Graphiques des déclarations par mois
- Répartition des formes juridiques

### Gestion des données
- Recherche multi-critères (NINEA, RCCM, entreprise, bénéficiaire)
- Filtres (région, forme juridique, PPE, statut)
- Export Excel avec mise en forme
- Export CSV

### Administration
- Gestion des utilisateurs (création, activation/désactivation, suppression)
- Suivi des statuts des déclarations
- Exports complets

---

## 💾 Base de données

L'application utilise **SQLite** par défaut (fichier `declarations.db`).

Pour migrer vers **PostgreSQL**, modifier dans `app.py` :
```python
app.config['SQLALCHEMY_DATABASE_URI'] = 'postgresql://user:password@localhost/declarations_db'
```

L'export Excel (`declarations_beneficiaires.xlsx`) est mis à jour automatiquement à chaque soumission.

---

## 🌐 Déploiement en production

### Avec Gunicorn (Linux)
```bash
pip install gunicorn
gunicorn -w 4 -b 0.0.0.0:8000 app:app
```

### Variables d'environnement recommandées
```bash
export SECRET_KEY="votre-clé-secrète-très-longue"
export DATABASE_URL="postgresql://..."
```

---

## 📞 Support

Ministère de la Justice — Direction des Technologies
