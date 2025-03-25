# Audit Sémantique SEO

Outil d'analyse sémantique pour comparer le positionnement de différents domaines sur des mots-clés communs. Cette application permet d'identifier rapidement les opportunités sémantiques en analysant les positions de plusieurs sites concurrents.

![Capture d'écran de l'application](https://via.placeholder.com/800x450.png?text=Audit+Sémantique+SEO)

## Fonctionnalités

- Import de fichiers d'export Ahrefs (CSV, Excel)
- Analyse comparative de plusieurs domaines simultanément
- Organisation claire des données avec colonnes de positions et URLs séparées
- Filtres personnalisables pour les positions et nombres de domaines
- Code couleur pour les positions (vert, jaune, rose, rouge)
- Export Excel formaté prêt à être partagé

## Prérequis

- Python 3.7 ou supérieur
- Pip (gestionnaire de paquets Python)

## Installation

1. Cloner ce dépôt :
```bash
git clone https://github.com/votre-nom-utilisateur/audit-semantique-seo.git
cd audit-semantique-seo
```

2. Installer les dépendances requises :
```bash
pip install -r requirements.txt
```

## Utilisation

1. Lancer l'application :
```bash
streamlit run audit_semantique.py
```

2. Ouvrez votre navigateur à l'adresse indiquée (généralement http://localhost:8501)

3. Importez vos fichiers d'export Ahrefs :
   - Chaque fichier doit être nommé avec le nom du domaine (ex: example.com.xlsx)
   - Les fichiers doivent contenir les colonnes Keyword, Volume, Current position et Current URL

4. Configurez les filtres selon vos besoins :
   - Type de filtre (ex: "Au moins 2 sites positionnés, dont 1 top 10")
   - Nombre minimum de sites sur le mot-clé
   - Nombre minimum de sites dans les X premières positions

5. Générez l'audit et téléchargez le fichier Excel résultant

## Structure des fichiers

- `audit_semantique.py` : Script principal de l'application
- `requirements.txt` : Liste des dépendances Python
- `README.md` : Documentation du projet

## Format des données d'entrée

L'application accepte les formats de fichiers suivants :
- CSV (.csv)
- Excel (.xlsx, .xls)

Les fichiers doivent contenir au minimum les colonnes suivantes :
- Keyword : Le mot-clé recherché
- Volume : Le volume de recherche mensuel
- Current position : La position actuelle du site pour ce mot-clé
- Current URL : L'URL positionnée pour ce mot-clé

## Conseils d'utilisation

- Nommez vos fichiers avec le nom de domaine exact pour faciliter l'identification
- Pour de meilleurs résultats, utilisez les mêmes dates d'export pour tous les domaines analysés
- L'application fonctionne mieux avec un nombre raisonnable de domaines (jusqu'à 10)

## Licence

[Précisez votre licence ici, par exemple MIT, GPL, etc.]

## Contact

[Vos coordonnées ou informations de contact]

---

Créé par IlanL - [2025]