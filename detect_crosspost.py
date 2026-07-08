#!/usr/bin/env python3
"""
Detecte si des publications Instagram sont des "crossposts owned"
(postees par le compte Instagram officiel de la marque).

Usage:
    python3 detect_crosspost.py chemin/vers/fichier.xlsx

Le fichier Excel doit contenir au moins les colonnes :
    - brand            : nom de la marque
    - lien             : URL du post/reel Instagram (ex: https://www.instagram.com/p/XXXX/)
    - crosspost        : colonne resultat, remplie avec True/False par le script
    - handle_instagram : (optionnel mais recommande) le(s) handle(s) Instagram
                         officiel(s) de la marque, ex: "nike" ou "nike, nike.france"
                         si plusieurs comptes. A remplir a la main une fois par marque.

Le script recupere le handle Instagram (@compte) de l'auteur de chaque post en
parsant la page publique du post, puis le compare :
    - aux handles listes dans "handle_instagram" si la colonne est renseignee
      pour cette ligne (comparaison fiable, insensible aux differences entre
      nom de marque et handle reel) ;
    - sinon, en repli, au nom de la marque (comparaison approximative, a
      verifier manuellement en cas de doute).
Le fichier Excel est mis a jour en place (memes onglet et colonnes).
"""

import re
import sys
import time
import unicodedata
from pathlib import Path

import requests
from bs4 import BeautifulSoup
import pandas as pd

REQUEST_HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
    ),
    "Accept-Language": "fr-FR,fr;q=0.9,en;q=0.8",
}

REQUEST_DELAY_SECONDS = 2  # pause entre deux requetes pour eviter le blocage Instagram
REQUEST_TIMEOUT_SECONDS = 15

# Formats observes dans le <title> / og:title des pages Instagram publiques :
#   "N Likes, N Comments - username (@username) on Instagram: ..."
#   "username on Instagram: ..."
#   "username (@username) on Instagram"
TITLE_PATTERNS = [
    re.compile(r"\(@([\w.]+)\)\s+on Instagram", re.IGNORECASE),
    re.compile(r"^([\w.]+)\s+on Instagram", re.IGNORECASE),
]


def normalize(text: str) -> str:
    """Minuscule, sans accents, sans caracteres non alphanumeriques."""
    if not text:
        return ""
    text = unicodedata.normalize("NFKD", str(text))
    text = text.encode("ascii", "ignore").decode("ascii")
    return re.sub(r"[^a-z0-9]", "", text.lower())


def extract_username(url: str) -> str | None:
    """Recupere le @handle Instagram de l'auteur d'un post via la page publique."""
    try:
        response = requests.get(url, headers=REQUEST_HEADERS, timeout=REQUEST_TIMEOUT_SECONDS)
    except requests.RequestException as exc:
        print(f"  [erreur reseau] {url} -> {exc}")
        return None

    if response.status_code != 200:
        print(f"  [http {response.status_code}] {url}")
        return None

    soup = BeautifulSoup(response.text, "html.parser")

    candidates = []
    meta_title = soup.find("meta", property="og:title")
    if meta_title and meta_title.get("content"):
        candidates.append(meta_title["content"])
    if soup.title and soup.title.string:
        candidates.append(soup.title.string)

    for candidate in candidates:
        for pattern in TITLE_PATTERNS:
            match = pattern.search(candidate)
            if match:
                return match.group(1)

    return None


def parse_known_handles(handle_instagram: str | None) -> list[str]:
    """Decoupe la colonne handle_instagram (valeurs separees par des virgules)."""
    if not handle_instagram or pd.isna(handle_instagram):
        return []
    return [h.strip() for h in str(handle_instagram).split(",") if h.strip()]


def is_owned_crosspost(brand: str, username: str | None, known_handles: list[str]) -> bool:
    """Compare le handle Instagram trouve aux handles connus, ou a defaut au nom de marque."""
    if not username:
        return False
    norm_user = normalize(username)
    if not norm_user:
        return False

    if known_handles:
        return any(normalize(h) == norm_user for h in known_handles)

    # Repli si aucun handle de reference n'est fourni pour cette marque.
    norm_brand = normalize(brand)
    if not norm_brand:
        return False
    return norm_brand in norm_user or norm_user in norm_brand


def process_file(xlsx_path: Path) -> None:
    df = pd.read_excel(xlsx_path, dtype=str)

    required_columns = {"brand", "lien"}
    missing = required_columns - set(df.columns)
    if missing:
        raise SystemExit(f"Colonnes manquantes dans l'Excel : {', '.join(sorted(missing))}")

    if "crosspost" not in df.columns:
        df["crosspost"] = None

    total = len(df)
    for index, row in df.iterrows():
        brand = row.get("brand")
        lien = row.get("lien")

        if not lien or pd.isna(lien):
            print(f"[{index + 1}/{total}] lien manquant, ligne ignoree")
            continue

        known_handles = parse_known_handles(row.get("handle_instagram"))

        print(f"[{index + 1}/{total}] {brand} -> {lien}")
        username = extract_username(str(lien))
        result = is_owned_crosspost(str(brand), username, known_handles)
        mode = "handle_instagram" if known_handles else "nom de marque (repli)"
        print(f"  compte detecte: {username!r} -> crosspost owned = {result}  [reference: {mode}]")

        df.at[index, "crosspost"] = result

        time.sleep(REQUEST_DELAY_SECONDS)

    df.to_excel(xlsx_path, index=False)
    print(f"\nFichier mis a jour : {xlsx_path}")


def main() -> None:
    if len(sys.argv) != 2:
        raise SystemExit("Usage: python3 detect_crosspost.py chemin/vers/fichier.xlsx")

    xlsx_path = Path(sys.argv[1])
    if not xlsx_path.exists():
        raise SystemExit(f"Fichier introuvable : {xlsx_path}")

    process_file(xlsx_path)


if __name__ == "__main__":
    main()
