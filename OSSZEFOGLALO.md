# Összefoglaló – mit csinál a program és hogyan

## Mi ez?

Egy könyvtárhálózatnak van 12 fiókkönyvtára. A központi rendszerből havonta kijön egy nagy Excel
(kb. 90 000 sor, 20 oszlop) az összes kölcsönzésről. Ezt eddig valakinek kézzel kellett szétszednie,
hogy minden fiók megkapja a saját részét.

A mi programunk (KEF – Kölcsönzési Export Feldolgozó) ezt csinálja meg:

1. **beolvassa** a nagy fájlt és kiszűri a hibás sorokat,
2. **hozzáad pár hasznos oszlopot** (pl. hány napot késett a könyv, mennyi a késedelmi díj),
3. **szétbontja fiókonként** külön Excel-fájlokba, szépen formázva.

Az ötlet egy valós munkahelyi feladatból jött: ott SAP-ból kellett ugyanígy céget szétbontani, és
Power Automate + VBA makró + Office Script csinálta. Mi ezt egy programba raktuk össze.

## Hogyan működik a kód (lépésről lépésre)

```
nagy Excel  →  loader.py  →  transform.py  →  splitter.py  →  fiókonkénti Excelek
               (beolvasás,     (új oszlopok,     (csoportosítás,
                ellenőrzés)     díjszámítás)      fájlok írása)
```

### 0. Mintaadat – `tools/generate_sample.py`
Mivel valódi könyvtári adatunk nincs, ez a szkript kitalál 90 000 kölcsönzést (olvasók, könyvek,
dátumok véletlenszerűen). Direkt beletesz kb. 270 hibás sort is (üres mező, rossz dátum, nem létező
fiók stb.), hogy legyen mit kiszűrni.

### 1. Beolvasás és ellenőrzés – `kef/pipeline/loader.py`
- Beolvassa az .xlsx vagy .csv fájlt (pandas-szal).
- Megnézi, hogy megvan-e mind a 20 oszlop. Ha nincs, hibát dob, pl. „Hiányzó oszlop(ok): cim”.
- Soronként ellenőriz: kötelező mezők nem üresek, a dátum értelmes, nincs kétszer ugyanaz az
  azonosító, a fiókkód létezik (F01–F12) stb.
- A végén két táblázatot ad vissza: a **jó sorokat** és a **hibás sorok listáját** (melyik sor, mi a baj).

### 2. „Formázás” / új oszlopok – `kef/pipeline/transform.py`
Ez a régi VBA-makró helyett van. Hozzáad 6 oszlopot:
- **kölcsönzési napok** – meddig volt kint a könyv
- **késés (nap)** – hány napot késett a lejárathoz képest
- **késedelmi díj** – késés × 50 Ft, de max. 3000 Ft, diáknak és nyugdíjasnak fele annyi
- **státusz** – visszahozva / kint / késésben
- **év-hónap** és **korcsoport** – kimutatásokhoz

Ezeket a számokat (50 Ft, 3000 Ft, 50%) az admin át tudja írni a programban.

### 3. Szétbontás – `kef/pipeline/splitter.py`
Ez a régi Office Script helyett van.
- Csoportosít fiók szerint (vagy kategória / könyv szerint, ha azt választod).
- Minden csoportnak ír egy Excel-fájlt (xlsxwriterrel), pl. `KEF_202411-202611_F02_Belvarosi_Fiok.xlsx`.
- Formázás: magyar fejlécek, szűrő, rögzített első sor, a késésben lévő sorok pirosak, plusz egy
  „Összesítő” lap.
- Készít egy összesítő fájlt is az összes fiókról, a hibás sorokkal együtt.
- Ha közben megszakítod, kitörli a félkész fájlokat.

### 4. Összefogás – `kef/service.py`
Ez hívja meg sorban a fenti lépéseket, előtte megnézi, hogy a felhasználónak van-e joga hozzá, és
minden futtatást elment az adatbázisba (ki, mikor, mennyi sor, hány fájl lett).

### 5. Adatbázis és belépés – `kef/db.py`, `kef/auth.py`
- Egy kis **SQLite** adatbázis (`data/kef.db`), 4 táblával: felhasználók, beállítások, futtatások,
  létrejött fájlok.
- 3 szerepkör van: **admin** (mindent), **feldolgozó** (feldolgozhat), **megtekintő** (csak nézelődik).
- A jelszavakat nem simán tároljuk, hanem hash-elve (PBKDF2).

### 6. A felület – `kef/gui/`
**Tkinter**-rel csináltuk, 10 ablak van, mindegyik egy külön fájl:

| # | Ablak | Fájl |
|---|-------|------|
| 1 | Bejelentkezés | `login.py` |
| 2 | Főmenü | `menu.py` |
| 3 | Adatbetöltés | `load.py` |
| 4 | Formázási szabályok | `rules.py` |
| 5 | Szétbontás beállításai | `splitscreen.py` |
| 6 | Feldolgozás folyamatban | `progress.py` |
| 7 | Eredmény | `result.py` |
| 8 | Előzmények | `history.py` |
| 9 | Felhasználók és jelszó | `users.py` |
| 10 | Megerősítő ablak („Biztos?”) | `dialog.py` |

Az `app.py` váltogat az ablakok között. A feldolgozás külön szálon fut, hogy a program ne fagyjon le,
amíg dolgozik (90 000 sornál kb. 40 mp).

## Használt technológiák

| Mi | Mire használjuk |
|----|-----------------|
| **Python 3.12** | ebben van írva az egész |
| **pandas**, **numpy** | táblázatkezelés: beolvasás, szűrés, számolás, csoportosítás |
| **python-calamine** | gyors Excel-olvasás (sokkal gyorsabb, mint az openpyxl) |
| **xlsxwriter** | a formázott Excel-fájlok kiírása |
| **Tkinter** | a grafikus felület (a Pythonban alapból benne van) |
| **SQLite** | kis adatbázis egy fájlban, nem kell hozzá szerver |
| **pytest** | automatikus tesztek |
| **matplotlib** | ábrák legenerálása |

## Tesztek – `tests/`
26 teszt van. A tervben leírt 10 teszteset (T01–T10) mind meg van írva tesztként is a
`tests/test_tesztesetek.py`-ban (pl. rossz jelszó, hiányzó oszlop, késedelmi díj számítása,
szétbontás 3 fiókra). Futtatás:

```powershell
python -m pytest -q tests
```

## A terv
A terv tartalma (ahogy a feladatban kérték):
1. Feladatspecifikáció
2. Funkcionális terv (funkciófa + a funkciók leírása)
3. Képernyőtervek (10 ablak + bejárhatósági gráf)
4. Adatmodell (ER-diagram + mintaadatok)
5. Tesztelési terv (10 teszteset)

## Indítás

```powershell
python -m pip install -r requirements.txt
python tools/generate_sample.py
python -m kef.main
```

Belépni az `admin` / `Admin123!` párossal lehet (a többi felhasználó a `kef/db.py`-ban van).
