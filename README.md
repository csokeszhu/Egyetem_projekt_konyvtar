# Kölcsönzési Export Feldolgozó (KEF)

IT modellezés és alkalmazásfejlesztés – gyakorlati feladat.

Egy városi könyvtárhálózat könyvtári rendszeréből exportált, kb. 90 000 soros, 20 oszlopos
kölcsönzési lista **betöltése és validálása → formázása (6 számított oszlop) → szétbontása
fiókkönyvtáranként** (vagy kategóriánként / könyvenként) külön, formázott Excel-fájlokba.
A téma egy valós munkahelyi feladatból jön (SAP-export → Power Automate + VBA + Office Script →
cégenkénti fájlok); a KEF ezt egyetlen asztali alkalmazásba fogja össze.

## Telepítés

Python 3.12 kell (Windows 10/11), majd a projekt gyökeréből:

```powershell
python -m pip install -r requirements.txt
```

## Futtatás

```powershell
python tools/generate_sample.py      # 90 000 soros mintaadat -> data/sample/
python -m kef.main                   # az alkalmazás
python -m pytest -q tests            # 26 automatikus teszt (benne a dokumentált T01–T10)
```

### Bemutató felhasználók

Első indításkor a program létrehozza a helyi adatbázist (`data/kef.db`) három teszt-felhasználóval
(a jelszavak a `kef/db.py` `SEED_USERS` listájában vannak):

| Azonosító    | Szerepkör      | Mit tehet                                              |
|--------------|----------------|--------------------------------------------------------|
| `admin`      | Adminisztrátor | mindent, szabályok és felhasználók kezelése            |
| `feldolgozo` | Feldolgozó     | betöltés, formázás, szétbontás, előzmények             |
| `megtekinto` | Megtekintő     | csak előzmények és saját jelszó                        |

## Dokumentáció

A kód működésének rövid leírása: [OSSZEFOGLALO.md](OSSZEFOGLALO.md).

Képernyőképek és ábrák újragenerálása:

```powershell
python tools/screenshots.py          # végigkattintja a GUI-t és képernyőképeket ment
python docs/diagrams.py              # funkciófa, bejárhatósági gráf, ER-diagram, folyamatábra
```

## Felépítés

```
kef/            alkalmazás (config, db, auth, service, pipeline/, gui/)
tools/          mintaadat-generátor, képernyőkép-készítő
tests/          pytest tesztek
docs/           dokumentáció-generálás, ábrák, képernyőképek
data/sample/    generált mintaadat
```
