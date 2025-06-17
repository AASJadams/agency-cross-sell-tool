# FULL Agency Cross-Sell Analysis Tool
# With Streamlit UI + Branded PDF Report

import os
import pandas as pd
import matplotlib.pyplot as plt
from io import BytesIO

import streamlit as st
from reportlab.lib.pagesizes import letter
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle,
    Image, PageBreak
)
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER
from reportlab.lib import colors

# Streamlit Theme
st.set_page_config(page_title="Agency Cross-Sell Analysis", layout="centered")

# Custom CSS Styling
st.markdown(
    """
    <style>
        html, body, [data-testid="stAppViewContainer"] {
            background-color: #07385e;
            color: #fec52d;
        }
        header[data-testid="stHeader"] {
            background-color: #87212e;
        }
        h1, h2, h3, h4, h5, h6, p, label, span, div {
            color: #fec52d !important;
        }
        input, textarea, .stTextInput, .stTextArea, .stFileUploader label {
            color: #07385e !important;
            font-weight: bold;
        }
        section[data-testid="stFileUploader"] label {
            color: #07385e !important;
            font-weight: bold;
        }
        button[kind="primary"] {
            background-color: #07385e !important;
            color: #fec52d !important;
            border: 1px solid #fec52d !important;
        }
        button[kind="primary"]:hover {
            background-color: #052742 !important;
            color: #ffffff !important;
        }
    </style>
    """,
    unsafe_allow_html=True
)

file = st.file_uploader("Upload Book of Business Report", type=["xlsx"])

if file:
    df = pd.read_excel(file)
    df.columns = df.columns.str.strip()

    column_map = {
        "Division": ["Division", "Agency", "Agency Name", "Policy Division"],
        "Customer Name": ["Customer Name", "Client Name"],
        "Line of Business/Non-Premium": ["Line of Business/Non-Premium", "Line of Business"],
        "Carrier": ["Carrier", "Parent Company"],
        "Policy Number": ["Policy Number"]
    }

    for std_col, variants in column_map.items():
        for col in variants:
            if col in df.columns:
                df.rename(columns={col: std_col}, inplace=True)
                break

    required_cols = ["Division", "Customer Name", "Line of Business/Non-Premium", "Carrier", "Policy Number"]
    if not all(col in df.columns for col in required_cols):
        st.error("Missing one or more required columns.")
    else:
        df = df[df["Line of Business/Non-Premium"].isin([
            "Homeowners", "Private Passenger Auto", "Mobile Homeowners"
        ])]

        def classify(group):
            types = set(group["Line of Business/Non-Premium"])
            return pd.Series(["Cross-Sold" if len(types) > 1 else "Monoline"] * len(group), index=group.index)

        def classify_carrier(group):
            lines_by_carrier = group.groupby("Carrier")["Line of Business/Non-Premium"].apply(set)
            return pd.Series([
                "Cross-Sold" if any(len(lines) > 1 for lines in lines_by_carrier.values) else "Monoline"
            ] * len(group), index=group.index)

        df["Account Type"] = df.groupby("Customer Name").apply(classify).reset_index(level=0, drop=True)
        df["Carrier Account Type"] = df.groupby("Customer Name").apply(classify_carrier).reset_index(level=0, drop=True)

        agency_summary = df.groupby("Division").agg(
            Total_Clients=("Customer Name", "nunique"),
            Monoline=("Account Type", lambda x: (x == "Monoline").sum()),
            Cross_Sold=("Account Type", lambda x: (x == "Cross-Sold").sum())
        ).reset_index()
        agency_summary["% Monoline"] = (agency_summary["Monoline"] / agency_summary["Total_Clients"] * 100).round(2)
        agency_summary["% Cross-Sold"] = 100 - agency_summary["% Monoline"]

        selected_agency = st.selectbox("Select an Agency:", agency_summary["Division"].unique())

        if selected_agency:

            agency_df = df[df["Division"] == selected_agency].copy()
            if agency_df.empty:
                st.info("No cross-sell or monoline data available for this agency.")
            else:
                agency_row = agency_summary[agency_summary["Division"] == selected_agency].iloc[0]

                total_clients = agency_df["Customer Name"].nunique()
                account_counts = agency_df.groupby("Customer Name")["Account Type"].first().value_counts()
                total_monoline = account_counts.get("Monoline", 0)
                total_cross_sold = account_counts.get("Cross-Sold", 0)

                carrier_counts = agency_df.groupby("Customer Name")["Carrier Account Type"].first().value_counts()
                total_carrier_mono = carrier_counts.get("Monoline", 0)
                total_carrier_cross = carrier_counts.get("Cross-Sold", 0)

                # On-Screen Chart Totals
                st.subheader("Account Mix Overview")
                col1, col2 = st.columns(2, gap="large")

                with col1:
                    st.markdown("**Agency Account Type**")
                    agency_img = BytesIO()
                    fig, ax = plt.subplots(figsize=(6, 6))
                    ax.pie([total_monoline, total_cross_sold],
                           labels=["Monoline", "Cross-Sold"],
                           autopct='%1.1f%%',
                           colors=["#fec52d", "#87212e"])
                    ax.set_title("Agency Account Type", fontsize=12, fontweight='bold')
                    plt.savefig(agency_img, format='png', bbox_inches='tight')
                    plt.close()
                    st.image(agency_img.getvalue(), width=300)

                with col2:
                    st.markdown("**Carrier Account Type**")
                    carrier_img = BytesIO()
                    fig, ax = plt.subplots(figsize=(6, 6))
                    ax.pie([total_carrier_mono, total_carrier_cross],
                           labels=["Monoline", "Cross-Sold"],
                           autopct='%1.1f%%',
                           colors=["#fec52d", "#87212e"])
                    ax.set_title("Carrier Account Type", fontsize=12, fontweight='bold')
                    plt.savefig(carrier_img, format='png', bbox_inches='tight')
                    plt.close()
                    st.image(carrier_img.getvalue(), width=300)

                # On-Screen Customer Breakdown
                st.subheader("Customer Breakdown")
                styled_table = agency_df[[
                    "Customer Name", "Carrier", "Line of Business/Non-Premium", "Policy Number",
                    "Account Type", "Carrier Account Type"
                ]].copy()

                def highlight_cross(row):
                    if row["Account Type"] == "Cross-Sold" and row["Carrier Account Type"] == "Cross-Sold":
                        return ["background-color: lightgreen"] * len(row)
                    elif row["Account Type"] == "Cross-Sold" or row["Carrier Account Type"] == "Cross-Sold":
                        return ["background-color: khaki"] * len(row)
                    else:
                        return [""] * len(row)

                st.dataframe(styled_table.style.apply(highlight_cross, axis=1))

                # PDF generation continues here...


            
            if st.button("Generate PDF Report"):
                buffer = BytesIO()
                doc = SimpleDocTemplate(buffer, pagesize=letter)
                elements = []
                styles = getSampleStyleSheet()
                white_heading_left = ParagraphStyle(
                    name="WhiteHeadingLeft",
                    parent=styles["Heading4"],
                    alignment=1,  # Centered
                    textColor=colors.white
                )
                centered = ParagraphStyle(name="centered", parent=styles["Title"], alignment=TA_CENTER)
                centered_normal = ParagraphStyle(name="centered_normal", parent=styles["Normal"], alignment=TA_CENTER)

                # Title Page
                logo_path = "AAS Logo.png"
                if os.path.exists(logo_path):
                    elements.append(Spacer(1, 100))
                    elements.append(Image(logo_path, width=200, height=100))
                elements.append(Spacer(1, 24))
                elements.append(Paragraph("Monoline Report", centered))
                elements.append(Spacer(1, 12))
                elements.append(Paragraph(f"<b>{selected_agency}</b>", centered))
                elements.append(Spacer(1, 6))
                elements.append(Paragraph("Date: " + pd.Timestamp.today().strftime("%B %d, %Y"), centered_normal))
                elements.append(Spacer(1, 24))
                percent_cross = round(agency_row['% Cross-Sold'], 1)
                summary_text = f"This report provides an overview of {selected_agency}'s current book of business, highlighting how many clients are monoline versus cross-sold at both the agency and carrier levels."
                summary_text += f" Out of {agency_row.Total_Clients} total clients, {agency_row.Cross_Sold} ({percent_cross}%) are cross-sold."
                if agency_row['% Cross-Sold'] < 20:
                    summary_text += " This represents a significant opportunity to round out accounts and increase revenue through strategic cross-selling."
                elif agency_row['% Cross-Sold'] < 50:
                    summary_text += " While some cross-selling is occurring, there’s still room to improve and capture additional premium."
                else:
                    summary_text += " Great job! Your agency is effectively cross-selling accounts. Continue nurturing multi-policy relationships."
                elements.append(Spacer(1, 200))
                elements.append(Paragraph(summary_text, styles["Normal"]))
                percent_cross = round(agency_row['% Cross-Sold'], 1)
                summary_text = f"This report provides an overview of {selected_agency}'s current book of business, highlighting how many clients are monoline versus cross-sold at both the agency and carrier levels."
                summary_text += f" Out of {agency_row.Total_Clients} total clients, {agency_row.Cross_Sold} ({percent_cross}%) are cross-sold."
                if agency_row['% Cross-Sold'] < 20:
                    summary_text += " This represents a significant opportunity to round out accounts and increase revenue through strategic cross-selling."
                elif agency_row['% Cross-Sold'] < 50:
                    summary_text += " While some cross-selling is occurring, there’s still room to improve and capture additional premium."
                else:
                    summary_text += " Great job! Your agency is effectively cross-selling accounts. Continue nurturing multi-policy relationships."
                

                # Charts
                elements.append(PageBreak())
                elements.append(Paragraph("Agency & Carrier Account Mix", centered))
                fig, ax = plt.subplots(figsize=(6.5, 6.5))
                ax.pie([total_monoline, total_cross_sold],
                       labels=["Agency Monoline", "Agency Cross-Sold"],
                       autopct='%1.1f%%',
                       colors=["#fec52d", "#87212e"])
                ax.set_title("Agency Account Type", fontsize=12, fontweight='bold')
                chart_path = "agency_chart.png"
                plt.savefig(chart_path)
                plt.close()

                fig, ax = plt.subplots(figsize=(6.5, 6.5))
                ax.pie([total_carrier_mono, total_carrier_cross],
                       labels=["Carrier Monoline", "Carrier Cross-Sold"],
                       autopct='%1.1f%%',
                       colors=["#fec52d", "#87212e"])
                ax.set_title("Carrier Account Type", fontsize=12, fontweight='bold')
                chart_path2 = "carrier_chart.png"
                plt.savefig(chart_path2)
                plt.close()

                elements.append(Spacer(1, 12))
                elements.append(Image(chart_path, width=275, height=275))
                elements.append(Spacer(1, 12))
                elements.append(Image(chart_path2, width=275, height=275))
                elements.append(PageBreak())

                
                # Account Mix Summary
                elements.append(Spacer(1, 24))
                elements.append(Paragraph("Account Mix Comparison", centered))
                elements.append(Spacer(1, 24))
                elements.append(Spacer(1, 18))

                agency_data = [["Total Clients", agency_row.Total_Clients],
                               ["Monoline", agency_row.Monoline],
                               ["Cross-Sold", agency_row.Cross_Sold],
                               ["% Monoline", f"{agency_row['% Monoline']:.1f}%"],
                               ["% Cross-Sold", f"{agency_row['% Cross-Sold']:.1f}%"]]

                carrier_data = [["Total Clients", total_clients],
                                ["Monoline", total_carrier_mono],
                                ["Cross-Sold", total_carrier_cross],
                                ["% Monoline", f"{total_carrier_mono / total_clients * 100:.1f}%" if total_clients else "0.0%"],
                                ["% Cross-Sold", f"{total_carrier_cross / total_clients * 100:.1f}%" if total_clients else "0.0%"]]

                mix_table = Table(
                    [
                        [
                            Paragraph('<font color="white"><b>Agency</b></font>', styles["Heading4"]),
                            "",
                            Paragraph('<font color="white"><b>Carrier</b></font>', styles["Heading4"]),
                            ""
                        ]
                    ] +
                    [[a[0], a[1], b[0], b[1]] for a, b in zip(agency_data, carrier_data)],
                    colWidths=[100, 100, 100, 100]
                )

                mix_table.setStyle(TableStyle([
                    ("BACKGROUND", (0, 0), (1, 0), colors.HexColor("#07385e")),  # Agency header background
                    ("BACKGROUND", (2, 0), (3, 0), colors.HexColor("#87212e")),  # Carrier header background
                    ("TEXTCOLOR", (0, 0), (3, 0), colors.white),  # White text for all header cells
                    ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                    ("FONTNAME", (0, 1), (-1, -1), "Helvetica-Bold"),
                    ("TEXTCOLOR", (0, 1), (-1, -1), colors.HexColor("#07385e")),
                    ("FONTSIZE", (0, 0), (-1, -1), 14),
                    ("ALIGN", (0, 0), (-1, -1), "CENTER"),
                    ("BOTTOMPADDING", (0, 0), (-1, -1), 12),
                    ("GRID", (0, 0), (-1, -1), 0.25, colors.black)
                ]))





                elements.append(mix_table)
                elements.append(PageBreak())

                # Add legend before customer breakdown
                legend_data = [["Highlight", "Meaning"],
                               ["", "Cross-Sold with Carrier AND Agency"],
                               ["", "Cross-Sold with Agency ONLY"]]
                legend_table = Table(legend_data, colWidths=[80, 300])
                legend_style = TableStyle([
                    ("BACKGROUND", (0, 0), (-1, 0), colors.lightgrey),
                    ("TEXTCOLOR", (0, 0), (-1, 0), colors.black),
                    ("ALIGN", (0, 0), (-1, -1), "LEFT"),
                    ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                    ("FONTSIZE", (0, 0), (-1, -1), 8),
                    ("GRID", (0, 0), (-1, -1), 0.25, colors.black),
                    ("BACKGROUND", (0, 1), (-1, 1), colors.green),
                    ("BACKGROUND", (0, 2), (-1, 2), colors.yellow)
                ])
                legend_table.setStyle(legend_style)
                elements.append(legend_table)
                elements.append(Spacer(1, 12))

                # Customer Breakdown Table
                elements.append(Paragraph("Customer Breakdown", centered))
                customer_table_data = [["Customer Name", "Carrier", "LOB", "Policy #", "Agency Type", "Carrier Type"]]
                highlight_rows = []
                for _, row in agency_df.iterrows():
                    row_data = [row["Customer Name"], row["Carrier"], row["Line of Business/Non-Premium"],
                                row["Policy Number"], row["Account Type"], row["Carrier Account Type"]]
                    customer_table_data.append(row_data)
                    if row["Account Type"] == "Cross-Sold" and row["Carrier Account Type"] == "Cross-Sold":
                        highlight_rows.append("green")
                    elif row["Account Type"] == "Cross-Sold" or row["Carrier Account Type"] == "Cross-Sold":
                        highlight_rows.append("yellow")
                    else:
                        highlight_rows.append(None)

                table = Table(customer_table_data, repeatRows=1)
                style = TableStyle([
                    ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#f39c12")),
                    ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                    ("ALIGN", (0, 0), (-1, -1), "CENTER"),
                    ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                    ("FONTSIZE", (0, 0), (-1, -1), 7),
                    ("GRID", (0, 0), (-1, -1), 0.25, colors.black)
                ])
                for i, color in enumerate(highlight_rows, start=1):
                    if color:
                        style.add("BACKGROUND", (0, i), (-1, i), colors.green if color == "green" else colors.yellow)
                table.setStyle(style)
                elements.append(Spacer(1, 12))
                elements.append(table)
                elements.append(PageBreak())

                # Cross-Selling Strategy Page
                elements.append(Paragraph("Cross-Selling Strategy", centered))
                elements.append(Spacer(1, 18))
                elements.append(Paragraph("<b>Why Cross-Selling Matters:</b>", styles["Heading3"]))
                elements.append(Spacer(1, 6))
                elements.append(Paragraph("• Increased revenue per client – more policies mean more commission.", styles["Normal"]))
                elements.append(Paragraph("• Higher retention – multi-policy clients are less likely to leave.", styles["Normal"]))
                elements.append(Paragraph("• Enhanced client satisfaction – bundled protection simplifies their lives.", styles["Normal"]))
                elements.append(Paragraph("• Stronger carrier relationships – better loss ratios and production.", styles["Normal"]))
                elements.append(Spacer(1, 12))
                elements.append(Paragraph("<b>How to Cross-Sell Effectively:</b>", styles["Heading3"]))
                elements.append(Spacer(1, 6))
                elements.append(Paragraph("• Ask: 'Do you have your auto/home/business insured elsewhere?'", styles["Normal"]))
                elements.append(Paragraph("• Use life events: new home, car, baby, or business growth.", styles["Normal"]))
                elements.append(Paragraph("• Run checkups: use your monoline data to target opportunities.", styles["Normal"]))
                elements.append(Paragraph("• Educate: explain how bundling avoids coverage gaps and saves money.", styles["Normal"]))
                elements.append(Spacer(1, 18))
                elements.append(Paragraph("Use this report as a strategic tool to identify rounding opportunities and strengthen client relationships.", centered_normal))

                doc.build(elements)

                st.success("PDF Report Generated!")
                st.download_button("Download PDF", data=buffer.getvalue(), file_name="monoline_report.pdf", mime="application/pdf")
