# =============================================================================
# helper.py – Génération PDF professionnelle (Leadway Auto Renewal)
# =============================================================================
# -*- coding: utf-8 -*-
import os
import pandas as pd
from datetime import datetime

from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_LEFT
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.units import cm
from reportlab.lib.utils import ImageReader
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import (
    BaseDocTemplate,
    Frame,
    PageTemplate,
    Paragraph,
    Spacer,
    Table,
    TableStyle,
)


class Helper:
    def __init__(self):
        """Initialise le chemin du logo."""
        base_dir = os.path.dirname(os.path.abspath(__file__))
        self.HEADER_IMAGE_PATH = os.path.join(base_dir, "assets", "logo_leadway.jpg")

    # -------------------------------------------------------------------------
    # Génération du PDF compact + bandeau orange
    # -------------------------------------------------------------------------
    def report_generator(self, data, file_path, mois_label: str = "JANVIER 2026", commission_total: float | None = None):
        """
        Rapport PDF pro et compact :
        - Bandeau orange + logo, titres centrés
        - Tableau police 6.2 pt
        - PRODUIT élargie, LIEN très large, EMAIL resserrée, NOM large
        - En-têtes de colonnes centrés
        - 'Liste détaillée des contrats' centré
        - Colonne durée (mois) convertie en entier
        """
        try:
            # ---------- Préparations ----------
            os.makedirs(os.path.dirname(file_path), exist_ok=True)

            n_cols = len(data.columns)
            page_size = landscape(A4) if n_cols > 7 else A4
            page_width, page_height = page_size

            left_margin, right_margin, top_margin, bottom_margin = (
                0.8 * cm, 0.8 * cm, 1.2 * cm, 0.8 * cm
            )

            # ---------- Document ----------
            doc = BaseDocTemplate(
                file_path,
                pagesize=page_size,
                leftMargin=left_margin,
                rightMargin=right_margin,
                topMargin=top_margin,
                bottomMargin=bottom_margin,
            )
            frame = Frame(doc.leftMargin, doc.bottomMargin, doc.width, doc.height, id="normal")

            # ---------- Header / Footer ----------
            def header_footer(canvas, _doc):
                width, height = page_size

                # Bande orange
                canvas.setFillColor(colors.HexColor("#FF8C00"))
                canvas.rect(0, height - 40, width, 40, fill=True, stroke=False)

                # Titre blanc centré
                canvas.setFillColor(colors.white)
                canvas.setFont("Helvetica-Bold", 10)
                canvas.drawCentredString(
                    width / 2, height - 25, "LEADWAY ASSURANCE IARD – AUTO RENOUVELLEMENT"
                )

                # Logo
                if os.path.exists(self.HEADER_IMAGE_PATH):
                    try:
                        img = ImageReader(self.HEADER_IMAGE_PATH)
                        canvas.drawImage(
                            img,
                            width - 115,
                            height - 38,
                            width=95,
                            height=26,
                            preserveAspectRatio=True,
                            mask="auto",
                        )
                    except Exception:
                        pass

                # Pied de page
                canvas.setFont("Helvetica", 7)
                canvas.setFillColor(colors.black)
                canvas.drawString(25, 20, f"Page {_doc.page}")
                canvas.drawRightString(width - 25, 20, f"Édité le {datetime.now():%d/%m/%Y}")

            doc.addPageTemplates([PageTemplate(id="page", frames=frame, onPage=header_footer)])

            # ---------- Styles ----------
            styles = getSampleStyleSheet()
            title_style = ParagraphStyle(
                "Title", parent=styles["Normal"], alignment=TA_CENTER,
                fontSize=11, leading=13, spaceAfter=4, textColor=colors.black,
                fontName="Helvetica-Bold",
            )
            total_style = ParagraphStyle(
                "Total", parent=styles["Normal"], alignment=TA_CENTER,
                fontSize=8, leading=9.5, textColor=colors.black,
            )
            small_center = ParagraphStyle(  # Liste détaillée centrée
                "SmallCenter", parent=styles["Normal"], alignment=TA_CENTER,
                fontSize=6.2, leading=7.8, textColor=colors.black,
            )
            small_left = ParagraphStyle(   # corps du tableau (même taille)
                "SmallLeft", parent=styles["Normal"], alignment=TA_LEFT,
                fontSize=6.2, leading=7.8, textColor=colors.black,
            )

            # ---------- En-tête de contenu ----------
            elements = []
            elements.append(Spacer(1, 50))
            elements.append(Paragraph(f"RENOUVELLEMENT {mois_label.upper()}", title_style))
            elements.append(Paragraph(
                f"Nombre total de contrats à renouveler : <b>{data.shape[0]}</b>",
                total_style,
            ))
            # Commission potentielle (Prime nette * taux selon POOL TPV)
            commission_text = "Commission potentielle si tout est renouvelé : <b>0</b>"
            if commission_total is not None:
                commission_text = (
                    "Commission potentielle si tout est renouvelé : "
                    f"<b>{int(round(commission_total)):,}</b> FCFA"
                ).replace(",", " ")
            elements.append(Paragraph(commission_text, total_style))
            elements.append(Spacer(1, 4))
            elements.append(Paragraph("Liste détaillée des contrats :", small_center))
            elements.append(Spacer(1, 4))

            # ---------- Pré-traitement des données ----------
            df = data.copy()
            # Masquer les colonnes internes du tableau PDF
            cols_to_hide = {"PRIME NETTE", "POOL TPV", "ROLE", "CODE TEAM"}
            df = df[[c for c in df.columns if c not in cols_to_hide]].copy()

            # Durée (mois) → entier si colonne présente
            duree_col = next((c for c in df.columns if "dur" in c.lower() and "mois" in c.lower()), None)
            if duree_col is not None:
                # conversion numérique + arrondi à l'entier (pas de décimales)
                df[duree_col] = (
                    df[duree_col]
                    .apply(lambda x: str(x).replace(",", "."))
                    .pipe(lambda s: s.replace(["", "nan", "None"], None))
                )
                df[duree_col] = (
                    df[duree_col]
                    .astype(float)
                    .round(0)
                    .astype("Int64")  # garde le NA propre
                )

            # Mise au propre (texte)
            df = df.fillna("").astype(str)

            cols = list(df.columns)
            table_data = [cols] + df.values.tolist()

            # ---------- Largeurs de colonnes ----------
            total_width = page_width - (left_margin + right_margin)
            pref_w = {
                "NOM COMPLET": 2.2,
                "PRODUIT": 1.0,
                "PRIME TTC": 0.9,
                "IMMATRICULATION": 1.3,
                "TELEPHONE": 1.0,
                "DUREE ( MOIS)": 0.9,
                "DATE ECHEANCE": 1.0,
                "LIEN DE RENOUVELLEMENT": 2.4,
                "EMAIL AGENT": 1.2,
            }
            weights = [pref_w.get(c, 1.0) for c in cols]
            total_weight = max(1.0, sum(weights))
            col_widths = []
            for c, w in zip(cols, weights):
                width = (w / total_weight) * total_width
                col_widths.append(max(32, min(170, width)))

            # ---------- Tableau ----------
            table = Table(table_data, colWidths=col_widths, repeatRows=1)
            table.setStyle(TableStyle([
                # En-tête (orange) + centrage des libellés
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#FF8C00")),
                ("TEXTCOLOR",  (0, 0), (-1, 0), colors.white),
                ("FONTNAME",   (0, 0), (-1, 0), "Helvetica-Bold"),
                ("FONTSIZE",   (0, 0), (-1, 0), 6.0),
                ("ALIGN",      (0, 0), (-1, 0), "CENTER"),   # <-- en-têtes centrés
                ("VALIGN",     (0, 0), (-1, 0), "MIDDLE"),

                # Corps
                ("FONTNAME",   (0, 1), (-1, -1), "Helvetica"),
                ("FONTSIZE",   (0, 1), (-1, -1), 6.2),
                ("TEXTCOLOR",  (0, 1), (-1, -1), colors.black),
                ("ALIGN",      (0, 1), (-1, -1), "LEFT"),
                ("VALIGN",     (0, 1), (-1, -1), "MIDDLE"),

                # Grille / alternance / retours à la ligne
                ("GRID",       (0, 0), (-1, -1), 0.25, colors.grey),
                ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.whitesmoke, colors.HexColor("#F9F9F9")]),
                ("WORDWRAP",   (0, 0), (-1, -1), "CJK"),

                # Paddings fins
                ("LEFTPADDING",  (0, 0), (-1, -1), 0.6),
                ("RIGHTPADDING", (0, 0), (-1, -1), 0.6),
                ("TOPPADDING",   (0, 0), (-1, -1), 0.3),
                ("BOTTOMPADDING",(0, 0), (-1, -1), 0.3),
            ]))

            elements.append(table)

            # ---------- Build ----------
            doc.build(elements)
            print(f"✅ PDF généré (centrage en-têtes, durée en entier) : {file_path}")

        except Exception as e:
            print(f"❌ Erreur lors de la génération du PDF pour {file_path} : {e}")
