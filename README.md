# Primerjava epitopov

*Deterministic comparison of predicted linear and conformational epitopes between two proteins, with formatted Excel reports.*

Skripti analizirata računalniško napovedane epitope z namenom vrednotenja potencialne navzkrižne reaktivnosti med proteinoma. Nastali sta kot orodje za magistrsko nalogo s področja farmacije, kjer AI orodja niso vračala ponovljivih rezultatov.

## Kaj počneta

**Linearni epitopi** (`linearni_epitopi.py`)
Posamezen epitop se razčleni na prekrivajoče se dvočrkovne kombinacije (dimere). Skripta nato preveri ujemanje in pojavnost dimerov med epitopi prvega in drugega proteina.

**Konformacijski epitopi** (`konformacijski_epitopi.py`)
Primerjava poteka na ravni posamezne aminokisline. Kot pozitivno ujemanje se upošteva le par z identično aminokislino na enakem položaju v proteinu (npr. `K108` ↔ `K108`).

Obe skripti ustvarita Excel datoteko z enim delovnim listom na vsak epitop referenčnega proteina. Ujemanja so obarvana zeleno.

## Namestitev

```bash
pip install -r requirements.txt
```

Potreben je Python 3.9 ali novejši.

## Uporaba

```bash
python linearni_epitopi.py primeri/linearni_protein_a.csv primeri/linearni_protein_b.csv rezultat_linearni.xlsx
python konformacijski_epitopi.py primeri/konformacijski_protein_a.csv primeri/konformacijski_protein_b.csv rezultat_konformacijski.xlsx
```

Prvi argument je referenčni protein, drugi primerjalni, tretji izhodna datoteka. Končnici `.csv` in `.xlsx` lahko izpustiš.

## Format vhodnih datotek

**Linearni epitopi:** stolpci `Protein`, `ID Epitopa`, `Sekvenca`

```csv
Protein,ID Epitopa,Sekvenca
Protein_A,EP_1,KALVEGTWQK
Protein_A,EP_2,GSDLYTNKAL
```

**Konformacijski epitopi:** stolpci `Protein`, `ID Epitopa`, `Residues` (ostanki ločeni z vejico, oznaka verige `A:` je neobvezna)

```csv
Protein,ID Epitopa,Residues
Protein_A,EP_1,"A:K108, A:A109, A:L110"
```

CSV, pri katerem je izvoz iz Excela celotno vrstico ovil v narekovaje, konformacijska skripta prebere samodejno.

## Primer izhoda

Konformacijski epitopi, delovni list `Protein_A_EP_1`:

| Protein_A - EP_1 | Protein_B |
|---|---|
| K108 | EP_1 = 1 |
| A109 | EP_2 = 1 |
| L110 | EP_1 = 1 |
| E114 | / |
| W118 | EP_2 = 1 |

Linearni epitopi, delovni list `Protein_A_EP_1` (prve vrstice):

| Par | Protein_A - EP_1 | Protein_B - EP_1 | Protein_B - EP_2 | Protein_B - EP_3 | Protein_B - EP_4 |
|---|---|---|---|---|---|
| KA | 1 | 1 | 1 | 0 | 0 |
| AL | 1 | 1 | 1 | 0 | 0 |
| LV | 1 | 0 | 0 | 0 | 0 |

Podatki v mapi `primeri/` so izmišljeni.
