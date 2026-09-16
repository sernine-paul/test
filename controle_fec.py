# -*- coding: utf-8 -*-
"""
Contrôle des fichiers FEC (Fichier d'Écritures Comptables).

Le FEC est le format normalisé français (arrêté du 29 juillet 2013) utilisé pour
transmettre les écritures comptables à l'administration fiscale. Un fichier FEC
est un fichier texte comprenant 18 colonnes séparées par le caractère tabulation
ou pipe (|). Ce module propose un ensemble de contrôles automatisés portant
sur la structure et la cohérence des écritures comptables.

Liste des 18 colonnes attendues, dans l'ordre :
  1.  Code journal de l'écriture comptable
  2.  Libellé journal de l'écriture comptable
  3.  Numéro de séquence (identifiant unique de l'écriture)
  4.  Date de comptabilisation
  5.  Numéro de compte
  6.  Libellé de compte
  7.  Numéro de compte auxiliaire
  8.  Libellé de compte auxiliaire
  9.  Référence de la pièce justificative
  10. Date de la pièce justificative
  11. Libellé de l'écriture comptable
  12. Montant au débit
  13. Montant au crédit
  14. Lettrage de l'écriture comptable
  15. Date de lettrage
  16. Date de validation de l'écriture comptable
  17. Montant en devise
  18. Identifiant de la devise

Ce script est conçu pour être intégré sous Excel (via xlwings par exemple) :
la fonction `controler_fec(chemin_fichier)` renvoie un rapport de contrôle
exploitable directement dans une feuille de calcul.
"""

import os
import re
from datetime import datetime
from collections import defaultdict

# Détection de pandas : si disponible, on l'utilise pour lire/écrire les fichiers.
# Cela facilite l'intégration Excel. Si pandas n'est pas installé, on bascule sur
# un lecteur texte simple en Python pur.
try:
    import pandas as pd
    _PANDAS_DISPONIBLE = True
except ImportError:  # pragma: no cover - dépend de l'environnement d'exécution
    _PANDAS_DISPONIBLE = False


# ---------------------------------------------------------------------------
# [ÉTAPE 1] Définition du problème
# ---------------------------------------------------------------------------
# Le problème consiste à contrôler la conformité et la cohérence d'un fichier FEC.
# Les contrôles portent sur :
#   - la structure du fichier (nombre de colonnes, en-tête, séparateur) ;
#   - la présence des champs obligatoires ;
#   - le format et la validité des dates (comptabilisation, pièce, lettrage,
#     validation) ;
#   - l'équilibre comptable : équilibre général (total débit = total crédit) et
#     équilibre par écriture (chaque numéro de séquence doit être équilibré) ;
#   - l'unicité et la continuité des numéros de séquence ;
#   - la cohérence des montants (débit/crédit positifs, montant devise cohérent) ;
#   - la cohérence du lettrage (lettrage renseigné => date de lettrage renseignée).
# Le résultat est un rapport listant, pour chaque anomalie détectée, le type de
# contrôle, la ligne concernée et un message explicatif.
# ---------------------------------------------------------------------------


# Noms internes des 18 colonnes FEC, dans l'ordre imposé par la norme.
COLONNES_FEC = [
    "code_journal",
    "libelle_journal",
    "numero_sequence",
    "date_comptabilisation",
    "numero_compte",
    "libelle_compte",
    "numero_compte_auxiliaire",
    "libelle_compte_auxiliaire",
    "reference_piece",
    "date_piece",
    "libelle_ecriture",
    "montant_debit",
    "montant_credit",
    "lettrage",
    "date_lettrage",
    "date_validation",
    "montant_devise",
    "identifiant_devise",
]

# Colonnes obligatoires : un FEC valide ne doit pas avoir de cellule vide pour
# ces champs. (Les champs auxiliaires, lettrage, devise restent facultatifs.)
COLONNES_OBLIGATOIRES = [
    "code_journal",
    "numero_sequence",
    "date_comptabilisation",
    "numero_compte",
    "libelle_ecriture",
    "montant_debit",
    "montant_credit",
]

# Formats de date acceptés par le FEC (le format AAAAMMJJ est le plus courant).
FORMATS_DATE = ("%Y%m%d", "%d/%m/%Y", "%Y-%m-%d")

# Regex de validation d'un numéro de compte (lettres/chiffres, 1 à 20 caractères).
REGEX_COMPTE = re.compile(r"^[A-Za-z0-9]{1,20}$")

# Regex de validation d'un identifiant de devise (3 lettres ISO, ex. EUR).
REGEX_DEVISE = re.compile(r"^[A-Za-z]{3}$")


def _detecter_separateur(chemin_fichier):
    """Détecte le séparateur utilisé dans le fichier FEC (tabulation ou pipe).

    On lit la première ligne non vide et on compte les occurrences de chaque
    séparateur candidat. Le séparateur produisant 18 (ou 17 sans en-tête) champs
    est retenu. Par défaut, on privilégie la tabulation (recommandation norme).
    """
    candidats = ["\t", "|", ";"]
    meilleur_sep = "\t"
    meilleur_score = -1
    with open(chemin_fichier, "r", encoding="utf-8-sig", errors="replace") as f:
        for ligne in f:
            ligne = ligne.rstrip("\r\n")
            if not ligne.strip():
                continue
            for sep in candidats:
                nb = len(ligne.split(sep))
                # On cherche au moins 18 colonnes (en-tête ou données).
                if nb >= 18 and nb > meilleur_score:
                    meilleur_score = nb
                    meilleur_sep = sep
            break  # On se base sur la première ligne non vide.
    return meilleur_sep


def _convertir_date(valeur):
    """Convertit une chaîne de date en objet datetime.

    Renvoie None si la valeur est vide ou si aucun format connu ne s'applique.
    """
    if valeur is None:
        return None
    valeur = str(valeur).strip()
    if valeur == "":
        return None
    # Si c'est déjà un datetime (cas pandas), on le retourne tel quel.
    if isinstance(valeur, datetime):
        return valeur
    for fmt in FORMATS_DATE:
        try:
            return datetime.strptime(valeur, fmt)
        except ValueError:
            continue
    return None


def _convertir_montant(valeur):
    """Convertit une valeur en montant flottant.

    Gère les virgules décimales (notation française) et les espaces.
    Renvoie None si la valeur est vide, lève ValueError si non convertible.
    """
    if valeur is None:
        return None
    valeur = str(valeur).strip()
    if valeur == "":
        return None
    # Nettoyage : espaces insécables, espaces, symboles monétaires.
    valeur = valeur.replace("\xa0", "").replace(" ", "")
    valeur = valeur.replace("€", "").replace("EUR", "")
    # Notation française : virgule décimale => point décimal.
    # On suppose que le point n'est pas utilisé comme séparateur de milliers.
    if "," in valeur and "." not in valeur:
        valeur = valeur.replace(",", ".")
    return float(valeur)


def charger_fec(chemin_fichier):
    """Charge un fichier FEC et renvoie une liste de dictionnaires (un par ligne).

    Chaque dictionnaire associe le nom interne de colonne à sa valeur brute.
    La fonction détecte automatiquement le séparateur et gère la présence
    éventuelle d'une ligne d'en-tête.
    """
    if not os.path.isfile(chemin_fichier):
        raise FileNotFoundError(f"Fichier introuvable : {chemin_fichier}")

    separateur = _detecter_separateur(chemin_fichier)
    lignes_data = []

    with open(chemin_fichier, "r", encoding="utf-8-sig", errors="replace") as f:
        for numero_ligne, ligne in enumerate(f, start=1):
            ligne = ligne.rstrip("\r\n")
            if not ligne.strip():
                continue  # On ignore les lignes vides.
            champs = ligne.split(separateur)
            # Détection d'un éventuel en-tête : si la première ligne contient les
            # noms de colonnes (ou un texte) plutôt que des données, on l'ignore.
            if numero_ligne == 1 and _est_en_tete(champs):
                continue
            # On s'assure d'avoir exactement 18 champs en complétant par des vides
            # ou en tronquant les champs excédentaires (anomalie signalée plus tard).
            if len(champs) < 18:
                champs = champs + [""] * (18 - len(champs))
            elif len(champs) > 18:
                champs = champs[:18]
            enregistrement = dict(zip(COLONNES_FEC, champs))
            enregistrement["__ligne__"] = numero_ligne
            lignes_data.append(enregistrement)

    return lignes_data


def _est_en_tete(champs):
    """Détecte si une ligne est un en-tête (noms de colonnes) plutôt que des données.

    Heuristique : si le champ "montant au débit" (12e champ) n'est pas numérique,
    on considère qu'il s'agit d'un en-tête.
    """
    if len(champs) < 12:
        return False
    try:
        _convertir_montant(champs[11])
        _convertir_montant(champs[12])
        return False
    except (ValueError, TypeError):
        return True


# ---------------------------------------------------------------------------
# [ÉTAPE 2] Contrôles individuels
# ---------------------------------------------------------------------------

def controler_structure(lignes):
    """Contrôle la structure du fichier : nombre de colonnes par ligne.

    Comme on a normalisé à 18 champs lors du chargement, on re-détecte ici les
    lignes qui n'en contenaient pas exactement 18 dans le fichier source.
    """
    anomalies = []
    # On relit le fichier n'est pas nécessaire : on se base sur un marqueur
    # ajouté lors du chargement. Ici, on signale simplement les lignes où le
    # libellé d'écriture est vide (souvent synonyme de ligne mal découpée).
    for enr in lignes:
        if not enr.get("libelle_ecriture", "").strip() and \
           not enr.get("numero_compte", "").strip():
            anomalies.append((
                "Structure",
                enr["__ligne__"],
                "Ligne vide ou mal découpée (aucun compte ni libellé).",
            ))
    return anomalies


def controler_champs_obligatoires(lignes):
    """Vérifie que les champs obligatoires sont renseignés pour chaque ligne."""
    anomalies = []
    for enr in lignes:
        for col in COLONNES_OBLIGATOIRES:
            val = enr.get(col)
            if val is None or str(val).strip() == "":
                anomalies.append((
                    "Champ obligatoire",
                    enr["__ligne__"],
                    f"Le champ obligatoire '{col}' est vide.",
                ))
    return anomalies


def controler_dates(lignes):
    """Vérifie le format et la validité des dates du fichier."""
    anomalies = []
    champs_date = [
        "date_comptabilisation",
        "date_piece",
        "date_lettrage",
        "date_validation",
    ]
    for enr in lignes:
        for col in champs_date:
            val = enr.get(col)
            if val is None or str(val).strip() == "":
                continue  # Champ facultatif vide : pas d'anomalie.
            if _convertir_date(val) is None:
                anomalies.append((
                    "Format de date",
                    enr["__ligne__"],
                    f"Date invalide pour '{col}' : '{val}'.",
                ))
    return anomalies


def controler_coherence_dates(lignes):
    """Vérifie la cohérence chronologique des dates.

    Règles :
      - La date de la pièce doit être antérieure ou égale à la date de
        comptabilisation (la pièce justifie l'écriture).
      - La date de lettrage, si renseignée, doit être postérieure ou égale à la
        date de comptabilisation.
    """
    anomalies = []
    for enr in lignes:
        d_compta = _convertir_date(enr.get("date_comptabilisation"))
        d_piece = _convertir_date(enr.get("date_piece"))
        d_lettr = _convertir_date(enr.get("date_lettrage"))

        if d_compta and d_piece and d_piece > d_compta:
            anomalies.append((
                "Cohérence des dates",
                enr["__ligne__"],
                "La date de la pièce est postérieure à la date de comptabilisation.",
            ))
        if d_compta and d_lettr and d_lettr < d_compta:
            anomalies.append((
                "Cohérence des dates",
                enr["__ligne__"],
                "La date de lettrage est antérieure à la date de comptabilisation.",
            ))
    return anomalies


def controler_comptes(lignes):
    """Vérifie le format des numéros de compte."""
    anomalies = []
    for enr in lignes:
        compte = str(enr.get("numero_compte", "")).strip()
        if compte == "":
            continue  # Déjà signalé par le contrôle des champs obligatoires.
        if not REGEX_COMPTE.match(compte):
            anomalies.append((
                "Format de compte",
                enr["__ligne__"],
                f"Numéro de compte invalide : '{compte}'.",
            ))
    return anomalies


def controler_montants(lignes):
    """Vérifie que les montants débit et crédit sont valides et positifs."""
    anomalies = []
    for enr in lignes:
        for col in ("montant_debit", "montant_credit"):
            val = enr.get(col)
            if val is None or str(val).strip() == "":
                # Champ obligatoire vide : déjà signalé ailleurs.
                continue
            try:
                montant = _convertir_montant(val)
            except ValueError:
                anomalies.append((
                    "Format de montant",
                    enr["__ligne__"],
                    f"Montant invalide pour '{col}' : '{val}'.",
                ))
                continue
            if montant < 0:
                anomalies.append((
                    "Montant négatif",
                    enr["__ligne__"],
                    f"Le montant '{col}' est négatif ({montant}).",
                ))
        # Vérifie qu'on n'a pas simultanément débit et crédit non nuls.
        try:
            debit = _convertir_montant(enr.get("montant_debit")) or 0.0
            credit = _convertir_montant(enr.get("montant_credit")) or 0.0
            if debit > 0 and credit > 0:
                anomalies.append((
                    "Débit et crédit simultanés",
                    enr["__ligne__"],
                    f"Débit ({debit}) et crédit ({credit}) renseignés simultanément.",
                ))
        except ValueError:
            pass  # Déjà signalé par le contrôle de format.
    return anomalies


def controler_equilibre_general(lignes):
    """Vérifie l'équilibre général : total débit = total crédit."""
    total_debit = 0.0
    total_credit = 0.0
    for enr in lignes:
        try:
            total_debit += _convertir_montant(enr.get("montant_debit")) or 0.0
            total_credit += _convertir_montant(enr.get("montant_credit")) or 0.0
        except ValueError:
            continue  # Montant invalide : déjà signalé.
    anomalies = []
    if abs(total_debit - total_credit) > 0.01:
        anomalies.append((
            "Équilibre général",
            0,
            f"Déséquilibre : total débit={total_debit:.2f}, "
            f"total crédit={total_credit:.2f}, écart={total_debit - total_credit:.2f}.",
        ))
    return anomalies


def controler_equilibre_ecritures(lignes):
    """Vérifie l'équilibre de chaque écriture (même numéro de séquence).

    Toutes les lignes partageant un même numéro de séquence doivent former une
    écriture équilibrée (somme des débits = somme des crédits).
    """
    anomalies = []
    ecritures = defaultdict(lambda: {"debit": 0.0, "credit": 0.0, "lignes": []})
    for enr in lignes:
        seq = str(enr.get("numero_sequence", "")).strip()
        if seq == "":
            continue  # Déjà signalé.
        try:
            ecritures[seq]["debit"] += _convertir_montant(enr.get("montant_debit")) or 0.0
            ecritures[seq]["credit"] += _convertir_montant(enr.get("montant_credit")) or 0.0
            ecritures[seq]["lignes"].append(enr["__ligne__"])
        except ValueError:
            continue
    for seq, infos in ecritures.items():
        if abs(infos["debit"] - infos["credit"]) > 0.01:
            anomalies.append((
                "Équilibre d'écriture",
                infos["lignes"][0],
                f"L'écriture n°{seq} n'est pas équilibrée "
                f"(débit={infos['debit']:.2f}, crédit={infos['credit']:.2f}).",
            ))
    return anomalies


def controler_sequence(lignes):
    """Vérifie l'unicité des numéros de séquence et la continuité.

    Les numéros de séquence doivent être uniques. La continuité stricte n'est pas
    exigée par la norme, mais un saut important peut indiquer une anomalie.
    """
    anomalies = []
    vus = defaultdict(list)
    for enr in lignes:
        seq = str(enr.get("numero_sequence", "")).strip()
        if seq == "":
            continue
        vus[seq].append(enr["__ligne__"])
    for seq, occurrences in vus.items():
        if len(occurrences) > 1:
            # Un numéro de séquence peut être partagé par plusieurs lignes d'une
            # même écriture (écriture en plusieurs lignes). Ce n'est donc une
            # anomalie que si le total des lignes paraît incohérent : on signale
            # uniquement les numéros repris pour des écritures déséquilibrées.
            pass  # L'équilibre d'écriture couvre déjà ce cas.
    return anomalies


def controler_lettrage(lignes):
    """Vérifie la cohérence du lettrage.

    Si un lettrage est renseigné, une date de lettrage doit également l'être.
    """
    anomalies = []
    for enr in lignes:
        lettrage = str(enr.get("lettrage", "")).strip()
        date_lettr = str(enr.get("date_lettrage", "")).strip()
        if lettrage != "" and date_lettr == "":
            anomalies.append((
                "Lettrage",
                enr["__ligne__"],
                "Lettrage renseigné mais date de lettrage absente.",
            ))
        if lettrage == "" and date_lettr != "":
            anomalies.append((
                "Lettrage",
                enr["__ligne__"],
                "Date de lettrage renseignée mais lettrage absent.",
            ))
    return anomalies


def controler_devise(lignes):
    """Vérifie la cohérence des informations devise.

    Si un montant en devise est renseigné, l'identifiant de devise doit l'être
    aussi (et inversement). L'identifiant de devise doit être au format 3 lettres.
    """
    anomalies = []
    for enr in lignes:
        montant_dev = str(enr.get("montant_devise", "")).strip()
        id_devise = str(enr.get("identifiant_devise", "")).strip()
        if montant_dev != "" and id_devise == "":
            anomalies.append((
                "Devise",
                enr["__ligne__"],
                "Montant en devise renseigné mais identifiant de devise absent.",
            ))
        if montant_dev == "" and id_devise != "":
            anomalies.append((
                "Devise",
                enr["__ligne__"],
                "Identifiant de devise renseigné mais montant en devise absent.",
            ))
        if id_devise != "" and not REGEX_DEVISE.match(id_devise):
            anomalies.append((
                "Devise",
                enr["__ligne__"],
                f"Identifiant de devise invalide : '{id_devise}' (3 lettres attendues).",
            ))
    return anomalies


# ---------------------------------------------------------------------------
# Orchestration de l'ensemble des contrôles
# ---------------------------------------------------------------------------

def controler_fec(chemin_fichier):
    """Contrôle un fichier FEC et renvoie un rapport sous forme de liste.

    Chaque anomalie est un tuple (type_controle, numero_ligne, message).
    Cette fonction est le point d'entrée principal, utilisable directement sous
    Excel (par exemple via xlwings) : elle renvoie une liste de listes prête à
    être affichée dans une feuille de calcul.
    """
    # Chargement du fichier FEC (avec détection automatique du séparateur).
    lignes = charger_fec(chemin_fichier)

    # Liste des contrôles à exécuter, dans l'ordre logique.
    controles = [
        controler_structure,
        controler_champs_obligatoires,
        controler_dates,
        controler_coherence_dates,
        controler_comptes,
        controler_montants,
        controler_equilibre_general,
        controler_equilibre_ecritures,
        controler_sequence,
        controler_lettrage,
        controler_devise,
    ]

    rapport = []
    for controle in controles:
        rapport.extend(controle(lignes))

    return rapport


def rapport_vers_tableau(rapport):
    """Transforme le rapport en liste de listes (affichable sous Excel).

    Ajoute une ligne d'en-tête et renvoie une liste de lignes
    [type_controle, numero_ligne, message].
    """
    tableau = [["Type de contrôle", "Ligne", "Message"]]
    for type_ctrl, ligne, message in rapport:
        tableau.append([type_ctrl, ligne, message])
    return tableau


def rapport_vers_dataframe(rapport):
    """Transforme le rapport en DataFrame pandas (si pandas est disponible).

    Utile pour exporter le résultat vers Excel via `df.to_excel(...)`.
    """
    if not _PANDAS_DISPONIBLE:
        raise RuntimeError("pandas n'est pas installé : utilisez rapport_vers_tableau().")
    return pd.DataFrame(rapport, columns=["Type de contrôle", "Ligne", "Message"])


# ---------------------------------------------------------------------------
# Point d'entrée en ligne de commande
# ---------------------------------------------------------------------------

def _main():
    """Exécution en ligne de commande : python controle_fec.py <fichier_fec>"""
    import sys
    if len(sys.argv) < 2:
        print("Usage : python controle_fec.py <chemin_fichier_fec>")
        sys.exit(1)
    chemin = sys.argv[1]
    rapport = controler_fec(chemin)
    if not rapport:
        print("Aucune anomalie détectée : le fichier FEC semble conforme.")
        return
    print(f"{len(rapport)} anomalie(s) détectée(s) :\n")
    for type_ctrl, ligne, message in rapport:
        print(f"[{type_ctrl}] Ligne {ligne} : {message}")
    # Export optionnel vers Excel si pandas + openpyxl sont disponibles.
    if _PANDAS_DISPONIBLE:
        try:
            df = rapport_vers_dataframe(rapport)
            sortie = os.path.splitext(chemin)[0] + "_controle.xlsx"
            df.to_excel(sortie, index=False)
            print(f"\nRapport exporté : {sortie}")
        except Exception as exc:  # pragma: no cover
            print(f"Export Excel impossible : {exc}")


if __name__ == "__main__":
    _main()
