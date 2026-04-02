# LOI-BAIL Generator v2 - Design Specification

> Refonte complete du generateur de documents juridiques LOI (Lettre d'Intention) et BAIL (Bail Commercial) pour l'immobilier commercial francais.

## Contexte

Projet existant (oct-dec 2025, 95 commits) avec LOI fonctionnelle en production et BAIL partiellement implementee. L'ancien code souffre de duplication massive, bugs connus, code mort, et absence de tests automatises. Cette refonte repart de zero avec une architecture propre tout en preservant chaque comportement de formatage valide.

**Contraintes** :
- Application Streamlit deployee sur Streamlit Cloud (version gratuite)
- Pas d'IA a l'execution (zero token) — code deterministe uniquement
- Le client utilise l'ancien lien en parallele pendant la migration
- Un seul fichier Excel uploade genere les deux documents (LOI + BAIL)

## Architecture

Architecture en couches avec modules partages. 4 couches, chaque module a une seule responsabilite.

```
LOI-BAIL-superpower/
├── app.py                              # UI Streamlit uniquement
├── core/
│   ├── __init__.py
│   ├── models.py                       # Dataclasses typees
│   ├── excel_parser.py                 # Extraction Excel unifiee
│   ├── formula_engine.py               # Resolution formules (RECHERCHEV, ARRONDI)
│   ├── inpi_client.py                  # Client INPI (API + scraping fallback)
│   └── number_to_french.py             # Conversion nombres -> lettres francaises
├── generators/
│   ├── __init__.py
│   ├── shared.py                       # Calculs derives partages (adresse, paliers, etc.)
│   ├── loi_generator.py                # Logique specifique LOI
│   └── bail_generator.py               # Logique conditionnelle BAIL
├── renderers/
│   ├── __init__.py
│   ├── word_engine.py                  # Moteur Word partage (placeholders, formatage)
│   ├── loi_renderer.py                 # Specificites LOI (headers, sections bleues)
│   └── bail_renderer.py                # Specificites BAIL (articles, headings, tags HTML, TOC)
├── config/
│   ├── settings.py                     # Config, secrets, chemins
│   ├── redaction_loi.xlsx
│   └── redaction_bail.xlsx
├── templates/
│   ├── Template LOI avec placeholder.docx
│   └── Template BAIL avec placeholder.docx
├── output/
├── tests/
│   ├── conftest.py
│   ├── test_number_to_french.py
│   ├── test_formula_engine.py
│   ├── test_excel_parser.py
│   ├── test_shared.py
│   ├── test_bail_generator.py
│   ├── test_word_engine.py
│   ├── test_inpi_client.py
│   └── test_integration.py
├── .env.example
├── .gitignore
├── .streamlit/
│   └── config.toml
├── requirements.txt
└── README.md
```

## Section 1 : Modeles de donnees (`core/models.py`)

Structures de donnees typees au lieu de dicts bruts.

```python
@dataclass
class DossierData:
    """Toutes les variables extraites d'un dossier (Excel + INPI + derivees)"""
    variables: dict[str, str]          # Variables brutes extraites de l'Excel
    variables_derivees: dict[str, str] # Calculees (adresse, paliers, surfaces, type bail, dates)
    inpi_data: InpiData | None         # Donnees INPI enrichies
    source_file: Path                  # Chemin vers l'Excel source (pour resolution formules BAIL)

@dataclass
class InpiData:
    nom_societe: str
    type_societe: str          # Forme juridique
    capital_social: str        # Formate "XXX XXX EUR"
    localite_rcs: str          # Ville RCS (sans arrondissement pour Paris/Lyon/Marseille)
    adresse_domiciliation: str
    president: str             # "Prenom Nom"
    fonction: str              # "President", "Gerant", etc.
    status: str                # "success" | "failed"
    error_message: str | None

@dataclass
class SocieteInfo:
    nom: str
    header_text: str           # Texte en-tete du document
    footer_text: str           # Texte pied de page (avec \n pour multi-lignes)

@dataclass
class ArticleResult:
    designation: str           # Cle dans le template (ex: "Article 1", "Comparution Bailleur")
    contenu: str               # Texte genere avec tags <b>/<i>/<u> et markers **
    placeholders_manquants: list[str]  # Placeholders non resolus
```

**Principes** :
- `DossierData` est construit une seule fois et passe aux deux generateurs (LOI + BAIL)
- Les variables derivees sont calculees une seule fois dans `generators/shared.py`
- `InpiData` est un objet type au lieu d'un dict

## Section 2 : Extraction Excel (`core/excel_parser.py` + `core/formula_engine.py`)

### `excel_parser.py`

Un seul parser pour LOI et BAIL :

- Charge le fichier Excel uploade (`data_only=True` pour les valeurs calculees)
- Charge le fichier de config (`Redaction LOI.xlsx` ou `Redaction BAIL.xlsx`) en double mode : `data_only=True` + `data_only=False` (pour les formules)
- Lit la config mapping : colonne A = nom variable, colonne B = source (cellule ou formule)
- Extrait toutes les variables dans un `dict[str, str]`
- Ajoute `"Date d'aujourd'hui"` = date du jour en DD/MM/YYYY
- Extrait les infos Societe Bailleur (nom, header, footer) depuis l'onglet dedie
- La cellule SIRET est lue depuis la config comme toute autre variable (plus de hardcode `Validation!B25`)
- Le fichier source est conserve dans `DossierData.source_file` pour la resolution formules BAIL

### `formula_engine.py`

Moteur de resolution de formules isole :

- `resolve_formula(formula, workbook) -> valeur`
- References simples : `='Sheet'!B5`
- `RECHERCHEV` (VLOOKUP) : lookup_value, table_range, col_index, match_type. Argument splitting respectant les quotes simples. Essai comparaison numerique d'abord, fallback string.
- `ARRONDI` (ROUND) : inner_formula recursive, num_decimals. Splitting respectant le nesting de parentheses.
- Plages de cellules (`E38:E41`) : retourne une liste de valeurs non-vides
- Formatage : dates en DD/MM/YYYY, nombres en string, texte strippe

## Section 3 : Client INPI (`core/inpi_client.py`)

Trois sources de donnees avec fallback en cascade :

### Source 1 — API REST INPI (prioritaire)
- Auth par login/password vers `registre-national-entreprises.inpi.fr/api/sso/login`
- Token JWT cache 1h
- Recherche par SIREN : GET `/companies?siren[]={siren}`
- Extraction dirigeants depuis `composition.pouvoirs` (filtre `actif=True`, roles President/Gerant)

### Source 2 — Scraping BeautifulSoup (fallback)
- Parsing HTML statique de `data.inpi.fr/entreprises/{siren}`
- Via `cloudscraper` pour bypass anti-bot
- Utilise si l'API ne retourne rien

### Source 3 — Scraping Playwright (fallback ultime)
- Rendu JS complet, headless Chromium
- Import conditionnel (pas toujours dispo sur Streamlit Cloud)

### Corrections par rapport a l'ancien code
- Format dirigeant unifie `"Prenom Nom"` partout (l'API retournait `"Nom Prenom"`)
- Cache via `st.cache_data` avec TTL 1h (au lieu de `lru_cache` process-level)
- Rate limiting conserve : 5 req/min

### Contrat de sortie
`get_company_info(siren, credentials) -> InpiData` : toujours un objet `InpiData`, avec `status="success"` ou `status="failed"` + `error_message`. Plus de `None`.

## Section 4 : Calculs partages (`generators/shared.py`)

Module unique eliminant toute duplication entre LOI et BAIL.

### Normalisation des noms de variables

Mapping centralise unique applique une seule fois apres extraction Excel :
- `"Duree du Bail"` -> `"Duree Bail"`
- `"Duree du DG"` -> `"Duree DG"`
- `"Montant Palier X"` / `"Montant du Palier X"` -> `"Montant du palier X"`
- `"Date prise d'effet"` / `"Date de prise d'effet du bail"` / `"Date debut bail"` -> `"Date de prise d'effet"`
- `"Statut Locaux loues"` -> `"Statut Locaux Loues"`
- `"Duree ferme bail"` -> `"Duree ferme Bail"`
- `"Dure GAPD"` -> `"Duree GAPD"` (correction typo)

### Calculs derives

`calculer_variables_derivees(variables, inpi_data) -> dict[str, str]` :

| Variable | Logique | Correction |
|----------|---------|-----------|
| `Adresse Locaux Loues` | `"{rue}, {ville}"` | BAIL utilisait `"ville, rue"` |
| `Montant du palier {1-6}` | `loyer_base - loyer_annee_i`, formate `"XXX XXX"` | BAIL stockait des floats, LOI des strings |
| `Surface R-1` | `totale - RDC`, en entier | Unifie |
| `Type Bail` | 9->`"3/6/9"`, 10->`"6/9/10"`, autre->`"{n} ans"` | Identique |
| `Date de signature` | `Date d'aujourd'hui + 21 jours` | BAIL utilisait `datetime.now()` au lieu de la variable |
| `Date offre valable` | `Date d'aujourd'hui + 7 jours` | Identique |
| `Date de prise d'effet + 9 ans` | `relativedelta(years=9)` | Ancien utilisait `timedelta(days=365*9)` |
| `Montant du DG` | `(loyer / 12) * duree_dg` | Identique |
| `Periode DG` | `{3: "quart", 4: "tiers", 6: "moitie"}` | Typo `"moitier"` corrigee |

### Detection societe

`est_societe(type_preneur) -> bool` :
- Match exact sur un set : `{"SAS", "SARL", "EURL", "SA", "SCI", "SNC", "SASU"}` + variantes avec accents
- Plus de substring matching (l'ancien matchait `"SA"` dans `"PASSAGE"`)

### Formatage nombres

- `formater_nombre(valeur) -> str` : `160000` -> `"160 000"`, `1234.56` -> `"1 234,56"` (supprime `,00`)
- Utilise `number_to_french.py` pour les placeholders `"en lettres"`

## Section 5 : Moteur Word partage (`renderers/word_engine.py`)

Coeur critique — toute la logique de remplacement de placeholders unifiee en un seul moteur.

### Algorithme de remplacement `[Variable]`

3 etapes :

1. **Detection** : regex `\[([^\]]+)\]` sur `paragraph.text` (texte reconstitue de tous les runs)
2. **Evaluation du chemin** :
   - Placeholder en entier dans un seul run -> **remplacement in-place** (`run.text.replace(...)`) — preserve 100% du formatage original
   - Placeholder fragmente entre runs -> **reconstruction char-map**
3. **Reconstruction char-map** (quand necessaire) :
   - Construction `char_to_run_map` : chaque caractere -> run source
   - Segmentation : texte normal / placeholder remplace / placeholder manquant
   - Creation de nouveaux runs avec copie de format depuis le run source
   - Couleur : manquant -> rouge `(255,0,0)`, trouve -> pas de override (garde couleur originale). Exception : si le paragraphe contenait un mix de donnees presentes et manquantes ET qu'il etait optionnel (bleu) et garde, tous les runs sont passes en noir (pour enlever le bleu)

### Copie de format

`copy_run_format(source, target, override_color=None)` :
- Copie : `name`, `size`, `bold`, `italic`, `underline`, `strike`, `sub/superscript`, `all_caps`, `small_caps`
- Couleur : si `override_color` fourni, l'applique. Sinon copie RGB explicite uniquement (`color.type == 1`). Couleur theme -> pas de copie.
- Correction BAIL : plus de forcage Calibri 11pt — police du run source preservee

### Sections optionnelles (bleues)

`is_paragraph_optional(paragraph) -> bool` :
- Un run avec `blue > red AND blue > green` en RGB -> paragraphe optionnel
- Un seul run bleu suffit pour marquer tout le paragraphe

### API

```python
class WordEngine:
    def replace_placeholders(self, paragraph, variables, clear_list=None) -> str | None:
        """Retourne 'delete' si optionnel sans donnees, None sinon"""

    def process_document_body(self, document, variables, clear_list=None) -> None:
        """Traite tous les paragraphes + cellules de tableaux"""

    def is_paragraph_optional(self, paragraph) -> bool:

    def copy_run_format(self, source, target, override_color=None) -> None:
```

## Section 6 : Renderer LOI (`renderers/loi_renderer.py`)

S'appuie sur `WordEngine` et ajoute la logique specifique LOI.

### Sections speciales

| Section | Condition de suppression | Comportement si gardee |
|---------|--------------------------|----------------------|
| Paragraphes bleus avec placeholders | Toutes les variables vides | Supprime. Si donnees presentes -> runs passes en noir |
| "Honoraires de commercialisation" | Pas de `Honoraires Preneurs` | Supprime (detection par contenu texte, pas couleur) |
| "Remises (sur loyer annuel indexe)" | Aucune donnee palier | Supprime. Si paliers presents -> passe en noir |
| "Condition(s) suspensive(s)" | Aucune condition remplie | Supprime. Si conditions presentes -> passe en noir, puis traitement `[.]` |

### Headers / Footers

`update_headers_footers(document, societe_info)` :
- Header : un paragraphe centre, bold 22pt, `space_after=12pt`
- Footer : lignes separees par `\n`, centrees, 9pt, `space_before=12pt` sur la premiere
- Marges : `top_margin=0.5"`, `header_distance=0.3"`, idem bottom/footer
- Applique a toutes les sections du document

### Liste de placeholders a effacer

- Variable d'instance (plus de liste mutable de classe)
- Si preneur non-societe (via `est_societe()`) : `"PRESIDENT DE LA SOCIETE"` et `"FONCTION INPI"` ajoutes a la clear list -> remplaces par du vide

### Flux de generation

1. Charger le template DOCX
2. Pre-scan : detecter presence paliers, conditions suspensives, honoraires
3. Traiter les sections speciales (suppression ou passage en noir)
4. `WordEngine.process_document_body()` pour le remplacement standard
5. Supprimer les paragraphes marques "delete"
6. Traiter les tableaux
7. Mettre a jour headers/footers
8. Sauvegarder

## Section 7 : Generateur BAIL + Renderer BAIL

### `generators/bail_generator.py` — Logique conditionnelle

**Chargement des regles** depuis `Redaction BAIL.xlsx` :
- Chaque ligne : Article, Designation, Condition, Source, Option 1, Condition Option 2, Option 2
- Lignes sans nom d'article = continuations de l'article precedent

**Evaluateur de conditions** `evaluer_condition(condition, variables) -> bool` :
- Condition vide -> `True`
- `"Si [X] = 'Oui'"` -> comparaison string (strip, case-insensitive)
- `"Si [X] > 9"` -> comparaison numerique (float)
- `"Si [X] non vide"` -> `bool(value)` et `value != "0"` et `value != 0`
- `"plusieurs conditions suspensives"` -> compte conditions 1-4 non vides
- Operateurs : `=`, `!=`, `>`, `<`, `>=`, `<=`, `superieur a`
- Normalisation guillemets typographiques avant parsing
- Condition non reconnue -> `False` + log warning

**Resolution de sources** (corrige — recoit le fichier Excel source) :
- Formules `='Sheet'!E47` -> resolues via `formula_engine`
- Plages `E38:E41` -> liste de valeurs
- Pattern `"Conditions suspensives 1, 2, 3, 4."` -> noms individuels

**Conditions suspensives** :
- 1 condition -> texte Option 1 direct
- Plusieurs -> lettrees `a.`, `b.`, `c.`, `d.` avec textes juridiques par type :
  - Financement, Autorisations administratives, Extraction, Liberation des locaux, Autre -> `[.]`

**Articles generes** dans l'ordre :
Comparution Bailleur/Preneur -> Article preliminaire -> Articles 1, 2, 3, 5.3, 7.1, 7.2, 7.3, 7.6, 8, 19, 22.2, 26, 26.1, 26.2

Retourne une `list[ArticleResult]`.

### `renderers/bail_renderer.py` — Rendu Word

**Remplacement `{{ARTICLE_xxx}}`** :
- Mapping designation -> placeholder template (le double espace `"Article  7.3"` gere dans le mapping)
- Split `\n\n` -> paragraphes separes (paragraphe vide entre)
- Split `\n` + markers `**`/`***`/`****` -> headings
- Premier paragraphe reutilise l'element XML existant, suivants inseres via `addnext()`

**Tags de formatage HTML** `parse_formatting_tags(text)` :
- Stack-based : `<b>`, `<i>`, `<u>` et fermetures
- Supporte nesting (`<b><i>texte</i></b>`)
- Retourne segments `(texte, {bold, italic, underline})`

**Heading markers** :
- `****` -> Heading 4, `***` -> Heading 3, `**` -> Heading 2
- Detection du plus specifique d'abord
- Reset indentation, alignment left

**Post-traitement** :
1. `WordEngine.process_document_body()` pour les `[Variable]` restants
2. Nettoyage `{{...}}` non remplaces (paragraphes composes uniquement de placeholders -> supprimes)
3. Fix indentation headings
4. Mark TOC dirty (`updateFields=true`)
5. Sauvegarder

## Section 8 : Application Streamlit (`app.py`)

Interface minimaliste — uniquement de l'UI, zero logique metier.

### Flux utilisateur
1. Upload unique `st.file_uploader` pour l'Excel "Fiche de decision"
2. Verification fichiers requis au demarrage (templates + configs)
3. Parsing : appel unique `excel_parser.parse()` -> `DossierData` (cache `@st.cache_data` avec cle `"{filename}_{date}"` pour invalidation quotidienne)
4. Deux onglets : LOI et BAIL
5. Bouton Generer dans chaque onglet -> generateur + renderer -> `st.download_button`

### Noms fichiers sortie
- LOI : `"YYYY MM DD - LOI NomPreneur.docx"`
- BAIL : `"BAIL - NomPreneur - DateLOI.docx"`

## Section 9 : Config (`config/settings.py`)

```python
TEMPLATE_LOI = "templates/Template LOI avec placeholder.docx"
TEMPLATE_BAIL = "templates/Template BAIL avec placeholder.docx"
CONFIG_LOI = "config/redaction_loi.xlsx"
CONFIG_BAIL = "config/redaction_bail.xlsx"
OUTPUT_DIR = "output/"

# INPI
INPI_USERNAME = secrets("INPI_USERNAME")
INPI_PASSWORD = secrets("INPI_PASSWORD")
INPI_MAX_CALLS = 5
INPI_PERIOD = 60
```

## Section 10 : Tests

Framework : pytest + fixtures partagees dans `conftest.py`.

| Fichier test | Couverture |
|-------------|-----------|
| `test_number_to_french.py` | Cas limites : 0, 70, 80, 100, 1000, composes |
| `test_formula_engine.py` | RECHERCHEV, ARRONDI, references simples, erreurs |
| `test_excel_parser.py` | Extraction variables, gestion dates, valeurs None |
| `test_shared.py` | Calculs derives : adresse, paliers, surfaces, type bail, dates, detection societe |
| `test_bail_generator.py` | Evaluation conditions (tous operateurs), selection articles, conditions suspensives |
| `test_word_engine.py` | Remplacement simple, placeholders fragmentes, sections bleues, missing -> rouge |
| `test_inpi_client.py` | Parsing reponse API, parsing HTML, fallback chain (avec mocks) |
| `test_integration.py` | Flux complet : Excel -> DossierData -> LOI DOCX + BAIL DOCX |

## Bugs corriges dans cette refonte

| Bug | Ancien code | Correction |
|-----|-------------|-----------|
| Adresse inversee BAIL | `"ville, rue"` dans `bail_generator.py:206` | Unifie `"rue, ville"` partout |
| Typo "moitier" | `bail_generator.py:296` | Corrige en `"moitie"` |
| Calcul 9 ans inexact | `timedelta(days=365*9)` | `relativedelta(years=9)` |
| Variable mutable partagee | `PLACEHOLDERS_TO_CLEAR` class-level | Variable d'instance |
| Police forcee Calibri 11pt BAIL | `_apply_default_font` ecrasait la police source | Police du template preservee |
| Detection societe par substring | `"SA"` matchait `"PASSAGE"` | Match exact sur set de formes juridiques |
| Fichier source non passe au BAIL | `BailGenerator` sans `source_file` | `DossierData.source_file` accessible |
| Variables derivees calculees 2 fois | `calculer_variables_derivees` appele dans `generer_bail` + explicitement | Calcul unique dans `shared.py` |
| Date signature BAIL vs LOI | BAIL utilisait `datetime.now()`, LOI la variable | Unifie sur la variable `"Date d'aujourd'hui"` |
| Format dirigeant inconsistant | API `"Nom Prenom"`, scraping `"Prenom Nom"` | Unifie `"Prenom Nom"` |

## Dependances (`requirements.txt`)

```
streamlit
python-docx
openpyxl
pandas
python-dotenv
cloudscraper
beautifulsoup4
python-dateutil
ratelimit
```

Playwright reste en import conditionnel (non liste car rarement disponible sur Streamlit Cloud).
