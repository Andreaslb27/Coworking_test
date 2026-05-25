"""
SCRAPER COWORKINGS PARISIENS (leportagesalarial.com) -> XLSX
===========================================================

Objectif (à partir de https://www.leportagesalarial.com/coworking/) :
1) Parcourir tous les liens "parisiens" (Paris + 75xxx dans le texte d'ancre)
2) Pour chaque page coworking Paris :
   - Titre (h1)
   - Image principale (1ère image pertinente)
   - Description (texte avant la section "Contacter ...")
   - Adresse / Téléphone / Accès / Site / Twitter / Facebook / LinkedIn
   - Meta title (<title>)
   - Meta description (<meta name="description">)
   - meta_title_lt_150 : True si len(meta_title) < 150
   - Date de publication (ou autre date dispo dans le HTML : metas/time)
3) Exporter le tout dans un fichier .xlsx (Excel)

Dépendances:
pip install requests pyquery pandas openpyxl
"""

import re
import time
from urllib.parse import urljoin

import pandas as pd
import requests
from pyquery import PyQuery as pq


# --------------------------------------------------------------------
# Constantes
# --------------------------------------------------------------------
START_URL = "https://www.leportagesalarial.com/coworking/"
BASE = "https://www.leportagesalarial.com"

# Headers pour éviter d'être pris pour un bot trop basique
HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/120.0 Safari/537.36"
    )
}


# --------------------------------------------------------------------
# Fonctions utilitaires
# --------------------------------------------------------------------
def clean_text(s: str) -> str:
    """
    Nettoie un texte:
    - remplace les sauts de ligne/tabulations par un espace
    - supprime les espaces multiples
    - supprime espaces début/fin
    """
    return re.sub(r"\s+", " ", (s or "")).strip()


def get_doc(url: str, timeout: int = 30) -> pq:
    """
    Télécharge la page (requests) puis la parse en HTML avec PyQuery.
    On renvoie un objet PyQuery (doc) qu'on pourra interroger avec des sélecteurs CSS.
    """
    r = requests.get(url, headers=HEADERS, timeout=timeout)
    r.raise_for_status()  # lève une erreur si HTTP != 200
    return pq(r.text)


def is_paris_link(anchor_text: str) -> bool:
    """
    Détermine si un lien (dans l'annuaire) correspond à un coworking parisien.
    Heuristique: le texte du lien contient "Paris" + référence au département 75.
    Ex: "Coworking X à Paris (75012)".
    """
    t = (anchor_text or "").lower()
    return ("paris" in t) and ("(75" in t or " 75" in t)


# --------------------------------------------------------------------
# 1) Extraction des liens parisiens depuis la page annuaire
# --------------------------------------------------------------------
def extract_paris_links(index_doc: pq) -> list[str]:
    """
    Sur la page /coworking/, les liens "IDF" sont listés sous un bloc
    "Coworking Paris – Île de France". On fait :

    - repérer le <h3> contenant "Coworking Paris" (ou fallback "Île de France")
    - parcourir les noeuds suivants (siblings) jusqu'au prochain <h3>
    - récupérer tous les <a> (nom + href)
    - garder uniquement ceux dont le texte indique Paris (75...)
    - convertir href en URL absolue
    - dédoublonner
    """
    # On essaye de localiser le titre de la section
    h3 = index_doc('h3:contains("Coworking Paris")')
    if not h3:
        h3 = index_doc('h3:contains("Île de France")')

    if not h3:
        raise RuntimeError("Section 'Coworking Paris – Île de France' introuvable.")

    links: list[str] = []

    # node = élément juste après le <h3>
    node = h3.next()

    # On avance jusqu'à rencontrer un autre <h3> (fin de section)
    while node and node.length:
        if node.is_("h3"):
            break

        # Cherche tous les liens présents dans ce bloc
        for a in node.find("a").items():
            anchor_text = clean_text(a.text())
            href = clean_text(a.attr("href") or "")
            if not anchor_text or not href:
                continue

            # Filtre Paris
            if not is_paris_link(anchor_text):
                continue

            # Construire l'URL absolue (si href est relatif)
            full_url = urljoin(START_URL, href)

            # Petit filtre supplémentaire: on ne garde que les pages /coworking/...
            if "/coworking/" in full_url:
                links.append(full_url)

        node = node.next()

    # Dédoublonnage (ordre non garanti) -> on peut aussi garder l'ordre avec un dict
    unique_links = list(dict.fromkeys(links))
    return unique_links


# --------------------------------------------------------------------
# 2) Extraction des données sur une page coworking
# --------------------------------------------------------------------
def extract_meta(doc: pq) -> tuple[str, str]:
    """
    Récupère:
    - meta title (balise <title>)
    - meta description (<meta name="description">)
    """
    meta_title = clean_text(doc("title").text())
    meta_desc = clean_text(doc('meta[name="description"]').attr("content") or "")
    return meta_title, meta_desc


def extract_title(doc: pq) -> str:
    """
    Récupère le titre principal visible de la page: le premier <h1>
    """
    return clean_text(doc("h1").eq(0).text())


def extract_main_image(doc: pq) -> str:
    """
    Récupère l'image principale.
    Comme la structure peut varier, on utilise une heuristique:

    - Essayer la 1ère image dans des zones "contenu" courantes
    - Fallback: la 1ère image de la page

    On renvoie un URL absolu (src peut être relatif).
    """
    candidates = [
        "article img",
        ".entry-content img",
        ".post-content img",
        ".wp-block-image img",
        ".content img",
    ]
    for sel in candidates:
        src = doc(sel).eq(0).attr("src")
        if src:
            return urljoin(BASE, src)

    # fallback ultime
    src = doc("img").eq(0).attr("src")
    return urljoin(BASE, src) if src else ""


def extract_description(doc: pq) -> str:
    """
    Récupère une description textuelle (intro) de la page.

    Heuristique:
    - On prend le "contenu principal" (souvent article .entry-content)
    - On concatène les <p> jusqu'à tomber sur un <h2> contenant "Contacter"
      (car après, ce sont les coordonnées / réseaux sociaux etc.)

    Avantage: ça marche même si la page n'a pas une balise "description" dédiée.
    """
    # On choisit un conteneur principal probable
    container = doc("article .entry-content")
    if not container:
        container = doc(".entry-content")
    if not container:
        container = doc("article")
    if not container:
        container = doc("body")

    parts: list[str] = []

    # Parcourt les éléments enfants de ce conteneur
    for child in container.children().items():
        # Stop si on arrive à la section contact
        if child.is_("h2") and "contacter" in clean_text(child.text()).lower():
            break

        # Ajouter les paragraphes directs
        if child.is_("p"):
            t = clean_text(child.text())
            if t:
                parts.append(t)

        # Ajouter les paragraphes éventuellement imbriqués
        for p_ in child.find("p").items():
            t = clean_text(p_.text())
            if t:
                parts.append(t)

        # Stop "sécurité" : éviter de récupérer un texte énorme si la page est bizarre
        if len(" ".join(parts)) > 2500:
            break

    return clean_text(" ".join(parts))


def extract_contact_fields(doc: pq) -> dict:
    """
    Extrait:
    Adresse, Téléphone, Accès, Site, Twitter, Facebook, LinkedIn

    Le site affiche souvent ces infos sous forme d'une liste <ul><li>...</li></ul>
    juste après un <h2> contenant "Contacter ...".

    On cherche:
    - <h2> 'Contacter'
    - le <ul> qui suit
    - pour chaque <li>, on split sur ":" -> label + value
    - si le <li> contient un <a>, on prend href pour les réseaux / site
    """
    out = {
        "Adresse": "",
        "Téléphone": "",
        "Accès": "",
        "Site": "",
        "Twitter": "",
        "Facebook": "",
        "LinkedIn": "",
    }

    # Localiser le h2 de contact
    h2 = doc('h2:contains("Contacter")').eq(0)
    if not h2:
        return out

    # Le <ul> juste après est souvent celui des coordonnées
    ul = h2.nextAll("ul").eq(0)
    if not ul:
        return out

    for li in ul.find("li").items():
        txt = clean_text(li.text())
        if ":" not in txt:
            continue

        label, value = [clean_text(x) for x in txt.split(":", 1)]
        lab = label.lower()

        # si un lien est présent, on préfère href
        href = clean_text(li.find("a").eq(0).attr("href") or "")

        if "adresse" in lab:
            out["Adresse"] = value
        elif "téléphone" in lab or "telephone" in lab:
            out["Téléphone"] = value
        elif "accès" in lab or "acces" in lab:
            out["Accès"] = value
        elif "site" in lab:
            out["Site"] = href or value
        elif "twitter" in lab:
            out["Twitter"] = href or value
        elif "facebook" in lab:
            out["Facebook"] = href or value
        elif "linkedin" in lab:
            out["LinkedIn"] = href or value

    return out


def extract_publication_date(doc: pq) -> str:
    """
    Récupère la date de publication depuis le HTML.
    Suivant le thème WordPress, ça peut être stocké différemment.

    On essaie dans cet ordre:
    1) meta property article:published_time
    2) meta property article:modified_time
    3) meta property og:updated_time
    4) <time datetime="...">
    5) sinon -> ""

    (Si tu veux absolument une date même si la page ne l'affiche pas,
    on peut ajouter un fallback sur "was last modified: ..." dans le texte.)
    """
    for sel in [
        'meta[property="article:published_time"]',
        'meta[property="article:modified_time"]',
        'meta[property="og:updated_time"]',
    ]:
        v = doc(sel).attr("content")
        if v:
            return clean_text(v)

    dt = doc("time").eq(0).attr("datetime")
    if dt:
        return clean_text(dt)

    return ""


# --------------------------------------------------------------------
# 3) Pipeline complet + export XLSX
# --------------------------------------------------------------------
def run(output_xlsx: str = "coworkings_paris.xlsx", sleep_s: float = 0.5):
    """
    - Télécharge l'annuaire
    - Récupère les liens Paris
    - Parcourt chaque page et extrait les infos
    - Exporte en XLSX
    """
    # Télécharger la page annuaire
    index_doc = get_doc(START_URL)

    # 1) Obtenir la liste des URLs parisiens
    paris_links = extract_paris_links(index_doc)

    rows = []

    for url in paris_links:
        # Structure standard d'une ligne (sera une ligne Excel)
        row = {
            "URL": url,
            "Titre": "",
            "Image_principale": "",
            "Description": "",
            "Adresse": "",
            "Téléphone": "",
            "Accès": "",
            "Site": "",
            "Twitter": "",
            "Facebook": "",
            "LinkedIn": "",
            "Meta_title": "",
            "Meta_description": "",
            "meta_title_lt_150": False,
            "Date_publication": "",
            "Erreur": "",
        }

        try:
            # Télécharger la page détail
            doc = get_doc(url)

            # 2) Titre
            row["Titre"] = extract_title(doc)

            # 3) Image principale
            row["Image_principale"] = extract_main_image(doc)

            # 4) Description
            row["Description"] = extract_description(doc)

            # 5) Infos contact / réseaux
            row.update(extract_contact_fields(doc))

            # 6-7) Meta title / meta description
            meta_title, meta_desc = extract_meta(doc)
            row["Meta_title"] = meta_title
            row["Meta_description"] = meta_desc

            # 8) Bool sur longueur meta title
            row["meta_title_lt_150"] = (len(meta_title) < 150) if meta_title else False

            # 9) Date publication
            row["Date_publication"] = extract_publication_date(doc)

        except Exception as e:
            # si une page plante, on garde l'erreur dans la colonne
            row["Erreur"] = str(e)

        rows.append(row)

        # petite pause pour ne pas spam le site (bonne pratique)
        time.sleep(sleep_s)

    # 10) Export xlsx via pandas (engine openpyxl)
    df = pd.DataFrame(rows)
    df.to_excel(output_xlsx, index=False, engine="openpyxl")

    print(f"OK -> {output_xlsx} | Pages Paris: {len(df)}")


if __name__ == "__main__":
    # Lance le scraping + export Excel
    run()