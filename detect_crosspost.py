#!/usr/bin/env python3
"""
Detecte si des publications Instagram sont des "crossposts owned"
(postees par le compte Instagram officiel de la marque).

Usage:
    python3 detect_crosspost.py chemin/vers/fichier.xlsx

Le fichier Excel doit contenir au moins les colonnes :
    - Brand            : nom de la marque
    - Permalink         : URL du post/reel Instagram (ex: https://www.instagram.com/p/XXXX/)
    - Is Crosspost      : colonne resultat, remplie avec True/False par le script
    - handle_instagram  : (optionnel mais recommande) le(s) handle(s) Instagram
                          officiel(s) de la marque, ex: "nike" ou "nike, nike.france"
                          si plusieurs comptes. A remplir sur UNE SEULE ligne par
                          marque : le script propage automatiquement la valeur a
                          toutes les autres lignes de la meme marque.

Le script se connecte a Instagram avec un compte (identifiants dans un fichier
.env local, jamais commite) pour recuperer de facon fiable le compte auteur de
chaque post, puis le compare :
    - aux handles connus pour cette marque (renseignes sur n'importe quelle
      ligne de cette marque dans "handle_instagram") si disponibles
      (comparaison fiable, insensible aux differences entre nom de marque et
      handle reel) ;
    - sinon, en repli, au nom de la marque (comparaison approximative, a
      verifier manuellement en cas de doute).
Le fichier Excel est mis a jour en place (memes onglet et colonnes).

Configuration requise (une seule fois) :
    1. Copier .env.example en .env
    2. Remplir IG_USERNAME et IG_PASSWORD avec un compte Instagram
       (idealement un compte secondaire/test, pas ton compte principal :
       l'automatisation de connexions viole les CGU d'Instagram et peut
       entrainer une limitation temporaire du compte utilise).
"""

import os
import re
import sys
import time
import unicodedata
from pathlib import Path

import instaloader
import pandas as pd
from dotenv import load_dotenv

REQUEST_DELAY_SECONDS = 2  # pause entre deux requetes pour eviter le blocage Instagram
SESSION_FILE = Path(__file__).parent / ".ig_session"

SHORTCODE_PATTERN = re.compile(r"instagram\.com/(?:p|reel|tv)/([A-Za-z0-9_-]+)")


def normalize(text: str) -> str:
    """Minuscule, sans accents, sans caracteres non alphanumeriques."""
    if not text:
        return ""
    text = unicodedata.normalize("NFKD", str(text))
    text = text.encode("ascii", "ignore").decode("ascii")
    return re.sub(r"[^a-z0-9]", "", text.lower())


def get_instagram_loader() -> instaloader.Instaloader:
    """Connecte un Instaloader au compte configure dans .env, avec cache de session."""
    load_dotenv()
    username = os.environ.get("IG_USERNAME")
    password = os.environ.get("IG_PASSWORD")
    if not username or not password:
        raise SystemExit(
            "IG_USERNAME / IG_PASSWORD manquants. "
            "Copie .env.example en .env et renseigne un compte Instagram."
        )

    loader = instaloader.Instaloader(
        download_pictures=False,
        download_videos=False,
        download_video_thumbnails=False,
        download_geotags=False,
        download_comments=False,
        save_metadata=False,
        compress_json=False,
    )

    if SESSION_FILE.exists():
        try:
            loader.load_session_from_file(username, str(SESSION_FILE))
            return loader
        except Exception:
            pass  # session invalide/expiree, on se reconnecte ci-dessous

    loader.login(username, password)
    loader.save_session_to_file(str(SESSION_FILE))
    return loader


def extract_shortcode(url: str) -> str | None:
    match = SHORTCODE_PATTERN.search(url)
    return match.group(1) if match else None


def extract_username(loader: instaloader.Instaloader, url: str) -> str | None:
    """Recupere le handle Instagram de l'auteur d'un post via l'API Instaloader."""
    shortcode = extract_shortcode(url)
    if not shortcode:
        print(f"  [lien invalide] impossible d'extraire le shortcode de {url}")
        return None

    try:
        post = instaloader.Post.from_shortcode(loader.context, shortcode)
        return post.owner_username
    except Exception as exc:
        print(f"  [erreur instagram] {url} -> {exc}")
        return None


def parse_known_handles(handle_instagram: str | None) -> list[str]:
    """Decoupe la colonne handle_instagram (valeurs separees par des virgules)."""
    if not handle_instagram or pd.isna(handle_instagram):
        return []
    return [h.strip() for h in str(handle_instagram).split(",") if h.strip()]


def build_brand_handles_map(df: pd.DataFrame) -> dict[str, list[str]]:
    """Construit brand normalise -> handles, a partir de n'importe quelle ligne
    de l'Excel ou handle_instagram est renseigne (inutile de le repeter sur
    chaque ligne d'une meme marque)."""
    brand_handles: dict[str, list[str]] = {}
    if "handle_instagram" not in df.columns:
        return brand_handles

    for _, row in df.iterrows():
        handles = parse_known_handles(row.get("handle_instagram"))
        if not handles:
            continue
        norm_brand = normalize(row.get("Brand"))
        if not norm_brand:
            continue
        brand_handles.setdefault(norm_brand, [])
        for handle in handles:
            if handle not in brand_handles[norm_brand]:
                brand_handles[norm_brand].append(handle)

    return brand_handles


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

    required_columns = {"Brand", "Permalink"}
    missing = required_columns - set(df.columns)
    if missing:
        raise SystemExit(f"Colonnes manquantes dans l'Excel : {', '.join(sorted(missing))}")

    if "Is Crosspost" not in df.columns:
        df["Is Crosspost"] = None

    brand_handles_map = build_brand_handles_map(df)
    loader = get_instagram_loader()

    total = len(df)
    for index, row in df.iterrows():
        brand = row.get("Brand")
        permalink = row.get("Permalink")

        if not permalink or pd.isna(permalink):
            print(f"[{index + 1}/{total}] lien manquant, ligne ignoree")
            continue

        known_handles = parse_known_handles(row.get("handle_instagram"))
        if not known_handles:
            known_handles = brand_handles_map.get(normalize(brand), [])

        print(f"[{index + 1}/{total}] {brand} -> {permalink}")
        username = extract_username(loader, str(permalink))
        result = is_owned_crosspost(str(brand), username, known_handles)
        mode = "handle_instagram" if known_handles else "nom de marque (repli)"
        print(f"  compte detecte: {username!r} -> Is Crosspost = {result}  [reference: {mode}]")

        df.at[index, "Is Crosspost"] = result

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
