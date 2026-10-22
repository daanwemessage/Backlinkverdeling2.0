"""
Backlinkverdeling: anchors en target-URL's

Rekenregels
1. Redirects: regels die bij dezelfde link horen worden één link,
   toegewezen aan de pagina waar hij uiteindelijk uitkomt.
2. Target-URL's: omgezet naar https, zonder www, met slash aan het eind.
3. Anchors: hoofdletters tellen niet mee; getoond wordt de schrijfwijze
   die het vaakst voorkomt. Lege anchors worden "(geen anchor)".
   Anchors die de merknaam bevatten worden "Branded". De merknaam komt
   automatisch uit het domein waar de links naartoe gaan. Uitzondering:
   een kale URL met een pad erachter (alcedo.nl/brandoverslag/) blijft
   een eigen anchor. Alleen de kale homepage-URL telt als Branded.
4. Verwijzend domein: het adres van de verwijzende pagina zonder www.
   Subdomeinen tellen als aparte domeinen.
5. Per verwijzend domein telt elke anchor en elke target-URL één keer.
6. Percentage = aantal per anchor of target-URL / totaal. Samen 100%.
"""

import io
import re
from urllib.parse import urlparse

import pandas as pd

GEEN_ANCHOR = "(geen anchor)"
BRANDED = "Branded"


# ---------- Inlezen ----------

def lees_export(data: bytes) -> pd.DataFrame:
    """Ahrefs levert UTF-8 (komma's) of UTF-16 (tabs)."""
    if data[:2] in (b"\xff\xfe", b"\xfe\xff"):
        return pd.read_csv(io.BytesIO(data), encoding="utf-16", sep="\t")
    return pd.read_csv(io.BytesIO(data), encoding="utf-8-sig", sep=None, engine="python")


def controleer_kolommen(df: pd.DataFrame) -> list:
    nodig = ["Referring page URL", "Target URL", "Anchor", "Redirect Chain URLs", "First seen"]
    return [k for k in nodig if k not in df.columns]


# ---------- Rekenregels ----------

def zonder_www(host: str) -> str:
    host = host.lower().split(":")[0]
    return host[4:] if host.startswith("www.") else host


def verwijzend_domein(url: str) -> str:
    return zonder_www(urlparse(str(url)).netloc)


def basisvorm_url(url: str) -> str:
    """https, zonder www, met slash aan het eind."""
    p = urlparse(str(url).strip())
    pad = p.path or "/"
    if not pad.endswith("/"):
        pad += "/"
    basis = f"https://{zonder_www(p.netloc)}{pad}"
    return f"{basis}?{p.query}" if p.query else basis


def voeg_redirects_samen(df: pd.DataFrame) -> pd.DataFrame:
    """Eén regel per echte link: de meest recente versie is de eindbestemming."""
    df = df.copy()
    keten = df["Redirect Chain URLs"].fillna("").astype(str)
    eerste = keten.str.split(",").str[0].str.strip()
    df["_oorspronkelijk"] = eerste.where(eerste != "", df["Target URL"])
    df["_anchor_ruw"] = df["Anchor"].fillna("").astype(str).str.strip()
    df["_first_seen"] = pd.to_datetime(df["First seen"], errors="coerce")
    df = df.sort_values("_first_seen")
    sleutel = ["Referring page URL", "_anchor_ruw", "_oorspronkelijk"]
    return df.drop_duplicates(subset=sleutel, keep="last")


def merknaam(targets: pd.Series) -> str:
    """Het meest voorkomende doeldomein, zonder extensie: alcedo.nl wordt alcedo."""
    hosts = targets.map(lambda u: zonder_www(urlparse(str(u)).netloc))
    host = hosts.mode().iloc[0] if len(hosts) else ""
    delen = host.split(".")
    return delen[-2] if len(delen) >= 2 else host


def alleen_letters(tekst: str) -> str:
    """Kleine letters en cijfers, zonder spaties, streepjes of punten."""
    return re.sub(r"[^a-z0-9]", "", tekst.lower())


URL_PATROON = re.compile(r"^(https?://)?(www\.)?[a-z0-9.-]+\.[a-z]{2,}(/.*)?$")


def is_url_met_pad(anchor: str) -> bool:
    """True voor anchors als alcedo.nl/brandoverslag/, False voor alcedo.nl/."""
    tekst = anchor.strip().lower()
    if not URL_PATROON.match(tekst):
        return False
    tekst = re.sub(r"^(https?://)?(www\.)?", "", tekst)
    pad = tekst.split("/", 1)[1] if "/" in tekst else ""
    return pad.strip("/") != ""


def bereid_voor(df: pd.DataFrame) -> pd.DataFrame:
    links = voeg_redirects_samen(df)
    links["Domein"] = links["Referring page URL"].map(verwijzend_domein)
    links["Target-URL"] = links["Target URL"].map(basisvorm_url)

    anchor = links["_anchor_ruw"].where(links["_anchor_ruw"] != "", GEEN_ANCHOR)
    links["_anchor_sleutel"] = anchor.str.lower()
    # Meest voorkomende schrijfwijze per anchor
    weergave = (
        pd.DataFrame({"sleutel": links["_anchor_sleutel"], "schrijfwijze": anchor})
        .groupby("sleutel")["schrijfwijze"]
        .agg(lambda s: s.value_counts().index[0])
    )
    links["Anchor"] = links["_anchor_sleutel"].map(weergave)

    merk = alleen_letters(merknaam(links["Target URL"]))
    if merk:
        is_branded = links["_anchor_ruw"].map(
            lambda a: merk in alleen_letters(a) and not is_url_met_pad(a)
        )
        links.loc[is_branded, "Anchor"] = BRANDED
    links.attrs["merknaam"] = merk
    return links


def verdeling(links: pd.DataFrame, kolom: str) -> pd.DataFrame:
    uniek = links[["Domein", kolom]].drop_duplicates()
    tabel = uniek[kolom].value_counts().rename_axis(kolom).reset_index(name="Verwijzende domeinen")
    tabel["Percentage"] = tabel["Verwijzende domeinen"] / len(uniek)
    return tabel.sort_values(["Percentage", kolom], ascending=[False, True]).reset_index(drop=True)


def maak_excel(anchors: pd.DataFrame, targets: pd.DataFrame) -> bytes:
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        for naam, tabel in (("Anchors", anchors), ("Target-URLs", targets)):
            tabel.to_excel(writer, sheet_name=naam, index=False)
            blad = writer.sheets[naam]
            blad.column_dimensions["A"].width = 60
            blad.column_dimensions["B"].width = 22
            blad.column_dimensions["C"].width = 14
            for cel in blad["C"][1:]:
                cel.number_format = "0.0%"
            blad.freeze_panes = "A2"
    return buffer.getvalue()


# ---------- Scherm ----------

def main():
    import streamlit as st

    st.set_page_config(page_title="Backlinkverdeling", page_icon="🔗")
    st.title("Backlinkverdeling")
    st.write(
        "Upload een Ahrefs Backlinks-export. Je krijgt de verdeling van anchors en "
        "target-URL's terug. Per verwijzend domein telt elke anchor en elke target-URL "
        "maar één keer, zodat footerlinks en redirects de cijfers niet vertekenen."
    )

    with st.expander("Hoe kom ik aan die export?"):
        st.markdown(
            "1. Ahrefs, Site Explorer, vul het domein in.\n"
            "2. Ga naar Backlink profile, **Backlinks**.\n"
            "3. Zet de filters **Dofollow** en **Best links: No spam** aan.\n"
            "4. Zet Show history op **Don't show history**.\n"
            "5. Klik op **Export**, kies alle regels en CSV."
        )

    upload = st.file_uploader("Backlinks-export", type=["csv"])
    if not upload:
        return

    try:
        df = lees_export(upload.getvalue())
    except Exception:
        st.error("Dit bestand kan niet gelezen worden. Upload een CSV-export uit Ahrefs.")
        return

    ontbreekt = controleer_kolommen(df)
    if ontbreekt:
        st.error("Dit lijkt geen Backlinks-export. Ontbrekende kolommen: " + ", ".join(ontbreekt))
        return

    links = bereid_voor(df)
    anchors = verdeling(links, "Anchor")
    targets = verdeling(links, "Target-URL")

    if links.attrs.get("merknaam"):
        st.caption(f"Herkende merknaam: **{links.attrs['merknaam']}**. "
                   "Anchors die deze naam bevatten vallen onder Branded.")

    k1, k2, k3, k4 = st.columns(4)
    k1.metric("Backlinks", f"{len(links):,}".replace(",", "."))
    k2.metric("Verwijzende domeinen", links["Domein"].nunique())
    k3.metric("Anchors", len(anchors))
    k4.metric("Target-URL's", len(targets))

    domein = links["Target-URL"].map(lambda u: urlparse(u).netloc).mode()
    naam = domein.iloc[0] if len(domein) else "backlinks"
    st.download_button(
        "Download Excel",
        data=maak_excel(anchors, targets),
        file_name=f"verdeling_{naam}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    procent = st.column_config.NumberColumn("Percentage", format="%.1f%%")
    tab1, tab2 = st.tabs(["Anchors", "Target-URLs"])
    for tab, tabel in ((tab1, anchors), (tab2, targets)):
        with tab:
            weergave = tabel.assign(Percentage=tabel["Percentage"] * 100)
            st.dataframe(weergave, hide_index=True, use_container_width=True,
                         column_config={"Percentage": procent})


if __name__ == "__main__":
    main()
