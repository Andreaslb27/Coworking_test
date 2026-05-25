# Imports principaux
import pandas as pd
import streamlit as st
import folium

# Géocodage des adresses
from geopy.geocoders import Nominatim
from geopy.extra.rate_limiter import RateLimiter

# Permet d'afficher Folium dans Streamlit
from streamlit_folium import st_folium


# --------------------------------------------------
# Configuration de la page Streamlit
# --------------------------------------------------

st.set_page_config(
    page_title="Coworkings Paris",
    page_icon="🧑‍💻",
    layout="wide"
)

st.title("Carte des coworkings parisiens")
st.write("Visualisation des coworkings récupérés via le scraping.")


# --------------------------------------------------
# Chargement du fichier Excel
# --------------------------------------------------

@st.cache_data
def load_data():
    """
    Lecture du fichier Excel généré par le scraper.
    fillna("") permet d'éviter les NaN.
    """
    df = pd.read_excel("coworkings_paris.xlsx")
    df = df.fillna("")
    return df


# --------------------------------------------------
# Géocodage des adresses
# --------------------------------------------------

@st.cache_data
def geocode_addresses(df):
    """
    Convertit les adresses en coordonnées GPS.
    On utilise Nominatim (OpenStreetMap).
    """

    geolocator = Nominatim(
        user_agent="coworking_paris_app"
    )

    # Limite les requêtes pour éviter d'être bloqué
    geocode = RateLimiter(
        geolocator.geocode,
        min_delay_seconds=1
    )

    latitudes = []
    longitudes = []

    # Parcours des adresses du DataFrame
    for address in df["Adresse"]:

        # Si adresse vide -> coordonnées nulles
        if address == "":
            latitudes.append(None)
            longitudes.append(None)
            continue

        # Ajout de "Paris, France" pour améliorer le résultat
        full_address = f"{address}, Paris, France"

        # Géocodage
        location = geocode(full_address)

        # Si coordonnées trouvées
        if location:
            latitudes.append(location.latitude)
            longitudes.append(location.longitude)

        else:
            latitudes.append(None)
            longitudes.append(None)

    # Création des colonnes GPS
    df["Latitude"] = latitudes
    df["Longitude"] = longitudes

    return df


# --------------------------------------------------
# Chargement et préparation des données
# --------------------------------------------------

df = load_data()

# On garde uniquement les lignes avec une adresse
df = df[df["Adresse"] != ""]

# Géocodage des adresses
df = geocode_addresses(df)

# Suppression des lignes sans coordonnées GPS
df_map = df.dropna(
    subset=["Latitude", "Longitude"]
)


# --------------------------------------------------
# Sidebar : recherche
# --------------------------------------------------

st.sidebar.header("Filtres")

search = st.sidebar.text_input(
    "Rechercher un coworking"
)

# Filtre sur le titre
if search:
    df_map = df_map[
        df_map["Titre"].str.contains(
            search,
            case=False,
            na=False
        )
    ]


# --------------------------------------------------
# Affichage du tableau
# --------------------------------------------------

st.subheader("Tableau des coworkings")

st.dataframe(df_map)


# --------------------------------------------------
# Création de la carte Folium
# --------------------------------------------------

st.subheader("Carte des coworkings")

# Carte centrée sur Paris
map_paris = folium.Map(
    location=[48.8566, 2.3522],
    zoom_start=12
)


# --------------------------------------------------
# Ajout des marqueurs
# --------------------------------------------------

for _, row in df_map.iterrows():

    # Contenu affiché dans le popup
    popup_html = f"""
    <b>{row['Titre']}</b><br>
    {row['Adresse']}<br><br>

    <a href="{row['URL']}" target="_blank">
        Voir la page
    </a>
    """

    # Création du marqueur
    folium.Marker(
        location=[
            row["Latitude"],
            row["Longitude"]
        ],

        popup=folium.Popup(
            popup_html,
            max_width=300
        ),

        tooltip=row["Titre"]

    ).add_to(map_paris)


# --------------------------------------------------
# Affichage de la carte dans Streamlit
# --------------------------------------------------

st_folium(
    map_paris,
    width=1200,
    height=600
)