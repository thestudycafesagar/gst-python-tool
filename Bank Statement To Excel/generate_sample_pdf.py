"""
Sample PDF Generator
Creates a demo bank statement PDF for testing the application.
Run: python generate_sample_pdf.py
Requires: reportlab (pip install reportlab)
"""
try:
    from reportlab.lib.pagesizes import A4
    from reportlab.lib import colors
    from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.units import cm
except ImportError:
    print("reportlab not installed. Run: pip install reportlab")
    exit(1)

import random
import os
from datetime import datetime, timedelta

OUTPUT_FILE = "sample_hdfc_statement.pdf"

TRANSACTIONS = [
    ("UPI-SWIGGY-FOOD", "D", 350.00),
    ("SALARY CREDIT-ACME CORP", "C", 75000.00),
    ("ATM WDL-SBI ATM MG ROAD", "D", 5000.00),
    ("NEFT-RENT PAYMENT", "D", 15000.00),
    ("UPI-AMAZON-SHOPPING", "D", 1299.00),
    ("INTEREST CREDIT", "C", 234.50),
    ("BILL PAY-ELECTRICITY", "D", 2450.00),
    ("UPI-PHONEPE-RECHARGE", "D", 399.00),
    ("RTGS-MUTUAL FUND", "D", 10000.00),
    ("UPI-ZEPTO-GROCERIES", "D", 876.00),
    ("EMI-HDFC LOAN", "D", 8500.00),
    ("UPI-ZOMATO-FOOD", "D", 420.00),
    ("DIVIDEND-RELIANCE", "C", 1500.00),
    ("NEFT-FREELANCE PAYMENT", "C", 25000.00),
    ("ATM WDL-HDFC ATM", "D", 2000.00),
    ("UPI-NETFLIX-SUBSCRIPTION", "D", 649.00),
    ("BILL PAY-BROADBAND", "D", 799.00),
    ("UPI-FUEL STATION", "D", 3200.00),
    ("INSURANCE PREMIUM-LIC", "D", 12000.00),
    ("UPI-MEESHO-SHOPPING", "D", 599.00),
]


def generate():
    doc = SimpleDocTemplate(OUTPUT_FILE, pagesize=A4, topMargin=1*cm, bottomMargin=1*cm)
    styles = getSampleStyleSheet()
    elements = []

    # Bank header
    header_style = ParagraphStyle(
        "header",
        fontSize=16,
        fontName="Helvetica-Bold",
        textColor=colors.HexColor("#003366"),
        spaceAfter=4,
    )
    sub_style = ParagraphStyle(
        "sub",
        fontSize=10,
        fontName="Helvetica",
        textColor=colors.HexColor("#555555"),
        spaceAfter=2,
    )

    elements.append(Paragraph("HDFC BANK LIMITED", header_style))
    elements.append(Paragraph("Account Statement", sub_style))
    elements.append(Paragraph("Customer Care: 1800-202-6161 | www.hdfcbank.com", sub_style))
    elements.append(Spacer(1, 0.3*cm))

    # Account info table
    acct_data = [
        ["Account Holder:", "ROHIT SHARMA", "Account No:", "XXXX XXXX 1234"],
        ["Branch:", "MG ROAD, BANGALORE", "IFSC:", "HDFC0001234"],
        ["Statement Period:", "01/01/2025 to 31/03/2025", "Account Type:", "Savings"],
        ["Opening Balance:", "₹ 48,750.00", "Closing Balance:", "₹ 89,341.50"],
    ]
    acct_table = Table(acct_data, colWidths=[4*cm, 6*cm, 4*cm, 5.5*cm])
    acct_table.setStyle(TableStyle([
        ("FONTNAME", (0, 0), (-1, -1), "Helvetica"),
        ("FONTSIZE", (0, 0), (-1, -1), 9),
        ("FONTNAME", (0, 0), (0, -1), "Helvetica-Bold"),
        ("FONTNAME", (2, 0), (2, -1), "Helvetica-Bold"),
        ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#F0F7FF")),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#CCE0FF")),
        ("PADDING", (0, 0), (-1, -1), 5),
        ("ROWBACKGROUNDS", (0, 0), (-1, -1), [colors.HexColor("#F0F7FF"), colors.white]),
    ]))
    elements.append(acct_table)
    elements.append(Spacer(1, 0.5*cm))

    # Transaction header
    elements.append(Paragraph("Transaction Details", ParagraphStyle(
        "txn_hdr", fontSize=12, fontName="Helvetica-Bold",
        textColor=colors.HexColor("#003366"), spaceAfter=6
    )))

    # Build transactions
    start_date = datetime(2025, 1, 1)
    balance = 48750.00
    txn_rows = [["Date", "Narration", "Value Dt", "Withdrawal Amt", "Deposit Amt", "Closing Balance"]]

    for i, (desc, txn_type, amount) in enumerate(TRANSACTIONS):
        txn_date = start_date + timedelta(days=i * 4 + random.randint(0, 3))
        date_str = txn_date.strftime("%d/%m/%y")
        if txn_type == "D":
            balance -= amount
            debit_str = f"{amount:,.2f}"
            credit_str = ""
        else:
            balance += amount
            debit_str = ""
            credit_str = f"{amount:,.2f}"

        txn_rows.append([
            date_str,
            desc,
            date_str,
            debit_str,
            credit_str,
            f"{balance:,.2f}",
        ])

    txn_table = Table(txn_rows, colWidths=[2*cm, 7.5*cm, 2*cm, 3*cm, 3*cm, 3.5*cm])
    txn_style = [
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, -1), 8),
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#003366")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("ALIGN", (3, 0), (5, -1), "RIGHT"),
        ("GRID", (0, 0), (-1, -1), 0.3, colors.HexColor("#CCCCCC")),
        ("PADDING", (0, 0), (-1, -1), 4),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#F5F9FF")]),
    ]
    # Color debit cells red, credit cells green
    for row_idx in range(1, len(txn_rows)):
        if txn_rows[row_idx][3]:  # has withdrawal
            txn_style.append(("TEXTCOLOR", (3, row_idx), (3, row_idx), colors.HexColor("#CC0000")))
        if txn_rows[row_idx][4]:  # has deposit
            txn_style.append(("TEXTCOLOR", (4, row_idx), (4, row_idx), colors.HexColor("#006600")))

    txn_table.setStyle(TableStyle(txn_style))
    elements.append(txn_table)
    elements.append(Spacer(1, 0.5*cm))
    elements.append(Paragraph(
        "This is a computer-generated statement. No signature required.",
        ParagraphStyle("footer_note", fontSize=8, textColor=colors.gray)
    ))

    doc.build(elements)
    print(f"✅ Sample PDF created: {os.path.abspath(OUTPUT_FILE)}")
    print(f"   Upload this file to test the Bank Statement → Excel Converter.")


if __name__ == "__main__":
    generate()
