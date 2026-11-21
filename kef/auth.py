"""Bejelentkezés, jelszókezelés és szerepkör alapú jogosultságok."""
from __future__ import annotations

import hashlib
import hmac
import re
import secrets
from dataclasses import dataclass

from kef.db import Database, now

LOGIN_ERROR = "Helytelen felhasználói azonosító és/vagy jelszó!"
PASSWORD_RULE = "A jelszó 8–30 karakter, betűt és számjegyet is tartalmaz, szóköz nélkül."
LOGIN_PATTERN = re.compile(r"^[A-Za-z0-9._-]{3,30}$")
DEFAULT_RESET_PASSWORD = "Kef12345"
ITERATIONS = 120_000

ROLE_NAMES = {"admin": "Adminisztrátor", "feldolgozo": "Feldolgozó", "megtekinto": "Megtekintő"}

# jogosultság -> szerepkörök, akik gyakorolhatják
PERMISSIONS: dict[str, set[str]] = {
    "feldolgozas": {"admin", "feldolgozo"},
    "beallitas_modositas": {"admin"},
    "felhasznalo_kezeles": {"admin"},
    "elozmenyek": {"admin", "feldolgozo", "megtekinto"},
}


class AuthError(Exception):
    pass


class PermissionDenied(Exception):
    pass


@dataclass(frozen=True)
class User:
    id: int
    azonosito: str
    nev: str
    szerepkor: str

    def can(self, permission: str) -> bool:
        return self.szerepkor in PERMISSIONS.get(permission, set())

    @property
    def role_name(self) -> str:
        return ROLE_NAMES.get(self.szerepkor, self.szerepkor)


def hash_password(password: str, salt: bytes | None = None) -> str:
    salt = salt or secrets.token_bytes(16)
    digest = hashlib.pbkdf2_hmac("sha256", password.encode("utf-8"), salt, ITERATIONS)
    return f"pbkdf2_sha256${ITERATIONS}${salt.hex()}${digest.hex()}"


def verify_password(password: str, stored: str) -> bool:
    try:
        _, iters, salt_hex, hash_hex = stored.split("$")
        digest = hashlib.pbkdf2_hmac("sha256", password.encode("utf-8"), bytes.fromhex(salt_hex), int(iters))
        return hmac.compare_digest(digest.hex(), hash_hex)
    except (ValueError, TypeError):
        return False


def validate_new_password(password: str) -> None:
    if (not 8 <= len(password) <= 30 or " " in password
            or not re.search(r"[A-Za-zÁÉÍÓÖŐÚÜŰáéíóöőúüű]", password) or not re.search(r"\d", password)):
        raise AuthError(PASSWORD_RULE)


def require(user: User | None, permission: str) -> None:
    if user is None or not user.can(permission):
        raise PermissionDenied("Ehhez a művelethez nincs jogosultságod.")


def _to_user(row) -> User:
    return User(id=row["id"], azonosito=row["azonosito"], nev=row["nev"], szerepkor=row["szerepkor"])


def login(db: Database, azonosito: str, jelszo: str) -> User:
    azonosito = (azonosito or "").strip()
    if not azonosito or not jelszo or not LOGIN_PATTERN.match(azonosito):
        raise AuthError(LOGIN_ERROR)
    row = db.get_user_by_login(azonosito)
    if row is None or not verify_password(jelszo, row["jelszo_hash"]):
        raise AuthError(LOGIN_ERROR)
    if not row["aktiv"]:
        raise AuthError("A felhasználói fiók le van tiltva. Fordulj az adminisztrátorhoz!")
    db.execute("UPDATE felhasznalo SET utolso_belepes = ? WHERE id = ?", (now(), row["id"]))
    return _to_user(row)


def change_password(db: Database, user: User, old: str, new: str, new_again: str) -> None:
    row = db.get_user(user.id)
    if row is None or not verify_password(old or "", row["jelszo_hash"]):
        raise AuthError("A régi jelszó hibás.")
    if new != new_again:
        raise AuthError("Az új jelszó és a megerősítése nem egyezik.")
    if new == old:
        raise AuthError("Az új jelszó nem egyezhet meg a régivel.")
    validate_new_password(new)
    db.execute("UPDATE felhasznalo SET jelszo_hash = ? WHERE id = ?", (hash_password(new), user.id))


def create_user(db: Database, admin: User, azonosito: str, nev: str, szerepkor: str, jelszo: str) -> int:
    require(admin, "felhasznalo_kezeles")
    azonosito = azonosito.strip()
    if not LOGIN_PATTERN.match(azonosito):
        raise AuthError("Az azonosító 3–30 karakter: betű, szám, pont, kötőjel, aláhúzás.")
    if not nev.strip():
        raise AuthError("A név megadása kötelező.")
    if szerepkor not in ROLE_NAMES:
        raise AuthError("Ismeretlen szerepkör.")
    if db.get_user_by_login(azonosito):
        raise AuthError("Ilyen azonosítójú felhasználó már létezik.")
    validate_new_password(jelszo)
    cur = db.execute(
        "INSERT INTO felhasznalo (azonosito, nev, jelszo_hash, szerepkor, letrehozva) VALUES (?,?,?,?,?)",
        (azonosito, nev.strip(), hash_password(jelszo), szerepkor, now()),
    )
    return int(cur.lastrowid)


def update_user(db: Database, admin: User, user_id: int, szerepkor: str, aktiv: bool) -> None:
    require(admin, "felhasznalo_kezeles")
    if user_id == admin.id and (szerepkor != "admin" or not aktiv):
        raise AuthError("Saját magadat nem tilthatod le és nem fokozhatod le.")
    if szerepkor not in ROLE_NAMES:
        raise AuthError("Ismeretlen szerepkör.")
    db.execute("UPDATE felhasznalo SET szerepkor = ?, aktiv = ? WHERE id = ?", (szerepkor, int(aktiv), user_id))


def reset_password(db: Database, admin: User, user_id: int) -> str:
    require(admin, "felhasznalo_kezeles")
    db.execute("UPDATE felhasznalo SET jelszo_hash = ? WHERE id = ?",
               (hash_password(DEFAULT_RESET_PASSWORD), user_id))
    return DEFAULT_RESET_PASSWORD
