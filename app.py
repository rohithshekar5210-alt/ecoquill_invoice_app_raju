import os, re, json, uuid, shutil, html, base64, io, zipfile, urllib.request, urllib.error
from pathlib import Path
from datetime import datetime, date
from decimal import Decimal, ROUND_HALF_UP
import pandas as pd
import streamlit as st
from openpyxl import Workbook, load_workbook
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_RIGHT
from reportlab.lib.units import mm
from reportlab.platypus import (
    SimpleDocTemplate,
    Table,
    TableStyle,
    Paragraph,
    Spacer,
    Image,
)
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

st.set_page_config(
    page_title="EcoQuill Invoice Generator",
    page_icon="🌿",
    layout="wide",
    initial_sidebar_state="expanded",
)
BASE = Path(__file__).resolve().parent
DATA, ASSETS, INVOICES, BACKUPS = [
    BASE / x for x in ("data", "assets", "invoices", "backups")
]
for p in (DATA, ASSETS, INVOICES, BACKUPS):
    p.mkdir(parents=True, exist_ok=True)
BOOK = DATA / "ecoquill_invoice_data.xlsx"
SETTINGS_FILE = BASE / "business_settings.json"
LOGO = ASSETS / "business_logo.png"
SIGN = ASSETS / "digital_signature.png"
PDF_FONTS_READY = True
for n, f in [
    ("EQRegular", "DejaVuSans.ttf"),
    ("EQBold", "DejaVuSans-Bold.ttf"),
    ("EQItalic", "DejaVuSans-Oblique.ttf"),
    ("EQBoldItalic", "DejaVuSans-BoldOblique.ttf"),
]:
    try:
        pdfmetrics.registerFont(TTFont(n, str(ASSETS / f)))
    except Exception:
        PDF_FONTS_READY = False

GREEN_DARK = "#0D3B2E"
INK = "#27342F"
MUTED = "#65736D"
PRODUCTS = [f"EcoQuill BioCarry {i} KG Compostable Bag" for i in range(1, 11)]

STYLE_DEFAULTS = {
    "invoice_title": {
        "size": 32,
        "color": "#FFFFFF",
        "bold": True,
        "italic": False,
        "underline": False,
    },
    "invoice_number": {
        "size": 12,
        "color": "#FFFFFF",
        "bold": True,
        "italic": False,
        "underline": False,
    },
    "bill_heading": {
        "size": 13,
        "color": GREEN_DARK,
        "bold": True,
        "italic": False,
        "underline": False,
    },
    "bill_body": {
        "size": 9,
        "color": INK,
        "bold": False,
        "italic": False,
        "underline": False,
    },
    "from_heading": {
        "size": 13,
        "color": GREEN_DARK,
        "bold": True,
        "italic": False,
        "underline": False,
    },
    "from_body": {
        "size": 9,
        "color": INK,
        "bold": False,
        "italic": False,
        "underline": False,
    },
    "date": {
        "size": 10,
        "color": MUTED,
        "bold": False,
        "italic": False,
        "underline": False,
    },
    "table_header": {
        "size": 9,
        "color": "#FFFFFF",
        "bold": True,
        "italic": False,
        "underline": False,
    },
    "table_body": {
        "size": 8,
        "color": INK,
        "bold": False,
        "italic": False,
        "underline": False,
    },
    "totals": {
        "size": 10,
        "color": INK,
        "bold": False,
        "italic": False,
        "underline": False,
    },
    "grand_total": {
        "size": 11,
        "color": "#FFFFFF",
        "bold": True,
        "italic": False,
        "underline": False,
    },
    "notes": {
        "size": 8,
        "color": MUTED,
        "bold": False,
        "italic": False,
        "underline": False,
    },
    "payment": {
        "size": 8,
        "color": INK,
        "bold": False,
        "italic": False,
        "underline": False,
    },
    "thank_you": {
        "size": 25,
        "color": GREEN_DARK,
        "bold": False,
        "italic": False,
        "underline": False,
    },
}
DEFAULT = {
    "invoice_title": "INVOICE",
    "company_name": "EcoQuill",
    "company_tagline": "Sustainability Starts With Every Bag",
    "company_address": "660, 9th Cross, Weavers Colony, Gottigere Post, Bannerghatta Road, Bengaluru - 560083",
    "company_phone": "+91 7899334559",
    "company_email": "ecoquill.biobags@gmail.com",
    "company_gstin": "AB290726105417Z",
    "company_website": "www.ecoquill.in",
    "authorized_signatory": "Your Name Here",
    "invoice_body": {
        "dark": "#20242E",
        "accent": "#A8C832",
        "grid": "#E1E4E8",
        "rows": 6,
        "row_mm": 13.0,
        "font": 7.2,
        "notes": 7.0,
        "show_hsn": False,
        "show_gst": False,
        "show_paid": False,
        "sl": "SL#",
        "description": "PRODUCT DESCRIPTION",
        "unit": "UNIT PRICE",
        "qty": "QTY.",
        "total": "TOTAL",
        "payment": "PAYMENT DETAILS:",
        "terms": "TERMS AND CONDITIONS:",
        "subtotal": "Subtotal:",
        "tax": "Taxation:",
        "grand": "G. Total",
        "thanks": "Thank you for your purchase here!",
        "sign_w": 31.0,
        "sign_h": 13.0,
    },
    "invoice_prefix": "EQ/INV",
    "default_place_of_supply": "Karnataka",
    "bank_name": "Bank Name Placeholder",
    "account_name": "EcoQuill",
    "account_no": "Account Number Placeholder",
    "ifsc_code": "IFSC Placeholder",
    "upi_id": "UPI Placeholder",
    "terms_conditions": "Goods once sold will not be taken back.\nPayment should be made as per agreed terms.\nAny dispute is subject to Bengaluru jurisdiction only.\nPlease verify quantity and product details at delivery.",
    "font_styles": STYLE_DEFAULTS,
}
SCHEMA = {
    "Sequence": ["Financial_Year", "Last_Number", "Updated_On"],
    "Products": [
        "Product_ID",
        "Product_Name",
        "Size_KG",
        "HSN_SAC",
        "Sale_Rate",
        "GST_Rate",
        "Created_On",
        "Is_Deleted",
        "Deleted_On",
        "Delete_Reason",
    ],
    "Customers": [
        "Customer_ID",
        "Customer_Name",
        "Address",
        "Phone",
        "WhatsApp",
        "Email",
        "GSTIN",
        "Created_On",
        "Is_Deleted",
        "Deleted_On",
        "Delete_Reason",
    ],
    "Invoices": [
        "Invoice_ID",
        "Invoice_Number",
        "Invoice_Date",
        "Customer_ID",
        "Customer_Name",
        "Taxable_Value",
        "CGST",
        "SGST",
        "IGST",
        "Total_GST",
        "Packing_Charges",
        "Grand_Total",
        "Paid_Amount",
        "Outstanding_Amount",
        "Payment_Status",
        "PDF_Path",
        "Created_On",
        "Is_Deleted",
        "Deleted_On",
        "Delete_Reason",
    ],
    "Invoice_Items": [
        "Item_ID",
        "Invoice_ID",
        "Invoice_Number",
        "Product_ID",
        "Product_Name",
        "Size_KG",
        "HSN_SAC",
        "Quantity_KG",
        "Rate_Per_KG",
        "Taxable_Value",
        "GST_Rate",
        "GST_Amount",
        "Line_Total",
        "Created_On",
        "Is_Deleted",
        "Deleted_On",
        "Delete_Reason",
    ],
}

# ============================================================
# OPTIONAL PERSISTENT CLOUD SNAPSHOT
# Configure Streamlit secrets: SUPABASE_URL, SUPABASE_SERVICE_KEY,
# and optionally SUPABASE_BUCKET (default: ecoquill-persistence).
# Without this, Streamlit Community Cloud can reset local files
# after the app sleeps / redeploys, wiping invoice history.
# ============================================================
REMOTE_STATE_NAME = "ecoquill_invoice_state.zip"
_REMOTE_SYNC_BUSY = False
_REMOTE_RESTORED = False


def _secret(name, default=""):
    try:
        return str(st.secrets.get(name, default) or default)
    except Exception:
        return str(os.getenv(name, default) or default)


def remote_configured():
    return bool(_secret("SUPABASE_URL") and _secret("SUPABASE_SERVICE_KEY"))


def _remote_url():
    root = _secret("SUPABASE_URL").rstrip("/")
    bucket = _secret("SUPABASE_BUCKET", "ecoquill-persistence")
    return f"{root}/storage/v1/object/{bucket}/{REMOTE_STATE_NAME}"


def _remote_request(method, data=None):
    key = _secret("SUPABASE_SERVICE_KEY")
    headers = {"Authorization": f"Bearer {key}", "apikey": key}
    if data is not None:
        headers.update({"Content-Type": "application/zip", "x-upsert": "true"})
    req = urllib.request.Request(
        _remote_url(), data=data, headers=headers, method=method
    )
    with urllib.request.urlopen(req, timeout=25) as response:
        return response.read()


def state_zip_bytes():
    """Snapshot the invoice workbook, settings, logo, signature and generated PDFs."""
    buff = io.BytesIO()
    with zipfile.ZipFile(buff, "w", zipfile.ZIP_DEFLATED) as z:
        for file in [BOOK, SETTINGS_FILE, LOGO, SIGN]:
            if file.exists() and file.is_file():
                z.write(file, file.relative_to(BASE).as_posix())
        if INVOICES.exists():
            for file in INVOICES.rglob("*"):
                if file.is_file():
                    z.write(file, file.relative_to(BASE).as_posix())
    return buff.getvalue()


def restore_remote_state_once():
    global _REMOTE_RESTORED
    if _REMOTE_RESTORED or not remote_configured():
        return
    _REMOTE_RESTORED = True
    try:
        payload = _remote_request("GET")
        with zipfile.ZipFile(io.BytesIO(payload)) as z:
            exact_allowed = {
                "data/ecoquill_invoice_data.xlsx",
                "business_settings.json",
                "assets/business_logo.png",
                "assets/digital_signature.png",
            }
            allowed_prefixes = ("invoices/",)
            for member in z.infolist():
                name = member.filename.replace("\\", "/")
                if name in exact_allowed or name.startswith(allowed_prefixes):
                    target = (BASE / name).resolve()
                    if (
                        BASE.resolve() not in target.parents
                        and target != BASE.resolve()
                    ):
                        continue
                    target.parent.mkdir(parents=True, exist_ok=True)
                    target.write_bytes(z.read(member))
    except urllib.error.HTTPError as exc:
        if exc.code != 404:
            st.warning(
                f"Persistent storage restore failed: HTTP {exc.code}. Local data will still work."
            )
    except Exception as exc:
        st.warning(
            f"Persistent storage restore failed: {exc}. Local data will still work."
        )


def sync_remote_state(show_error=True):
    global _REMOTE_SYNC_BUSY
    if _REMOTE_SYNC_BUSY or not remote_configured():
        return True
    _REMOTE_SYNC_BUSY = True
    try:
        _remote_request("POST", state_zip_bytes())
        return True
    except Exception as exc:
        if show_error:
            st.error(f"Local save succeeded, but persistent cloud backup failed: {exc}")
        return False
    finally:
        _REMOTE_SYNC_BUSY = False


def persistence_status():
    return (
        "Persistent cloud storage configured"
        if remote_configured()
        else "Local storage only. Configure Supabase secrets before Streamlit Cloud deployment."
    )


# ============================================================
# CORE HELPERS
# ============================================================
def now():
    return datetime.now().replace(microsecond=0)


def uid(p):
    return f"{p}-{datetime.now():%Y%m%d}-{uuid.uuid4().hex[:8].upper()}"


def qty(v):
    try:
        return float(
            Decimal(str(v or 0)).quantize(Decimal("0.001"), rounding=ROUND_HALF_UP)
        )
    except Exception:
        return 0.0


def money(v):
    try:
        return float(
            Decimal(str(v or 0)).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
        )
    except Exception:
        return 0.0


def settings():
    x = json.loads(json.dumps(DEFAULT))
    if SETTINGS_FILE.exists():
        try:
            saved = json.loads(SETTINGS_FILE.read_text(encoding="utf-8"))
            x.update(saved)
            x.setdefault("invoice_body", {})
            for k, v in DEFAULT.get("invoice_body", {}).items():
                x["invoice_body"].setdefault(k, v)
            for k, v in STYLE_DEFAULTS.items():
                x.setdefault("font_styles", {}).setdefault(k, v)
        except Exception:
            pass
    return x


def save_settings(x):
    SETTINGS_FILE.write_text(
        json.dumps(x, indent=2, ensure_ascii=False), encoding="utf-8"
    )
    sync_remote_state()


def init():
    """Initialize workbook, automatically recovering from an invalid XLSX file."""
    if BOOK.exists():
        try:
            wb = load_workbook(BOOK)
        except Exception:
            damaged = (
                BACKUPS / f"corrupted_workbook_{datetime.now():%Y%m%d_%H%M%S}.xlsx"
            )
            try:
                shutil.copy2(BOOK, damaged)
            except Exception:
                pass
            BOOK.unlink(missing_ok=True)
            Workbook().save(BOOK)
            wb = load_workbook(BOOK)
    else:
        Workbook().save(BOOK)
        wb = load_workbook(BOOK)
    if "Sheet" in wb.sheetnames and len(wb.sheetnames) == 1:
        wb.remove(wb["Sheet"])
    for n, cols in SCHEMA.items():
        if n not in wb.sheetnames:
            ws = wb.create_sheet(n)
            ws.append(cols)
        else:
            ws = wb[n]
            old = [x.value for x in ws[1]]
            for c in cols:
                if c not in old:
                    ws.cell(1, ws.max_column + 1, c)
    wb.save(BOOK)


def read(n):
    try:
        d = pd.read_excel(BOOK, sheet_name=n, engine="openpyxl", dtype=object)
    except Exception:
        return pd.DataFrame(columns=SCHEMA[n])
    for c in SCHEMA[n]:
        if c not in d:
            d[c] = ""
    return d[SCHEMA[n]].astype(object).where(pd.notna(d[SCHEMA[n]]), "")


def backup():
    if BOOK.exists():
        shutil.copy2(BOOK, BACKUPS / f"ecoquill_{datetime.now():%Y%m%d_%H%M%S_%f}.xlsx")


def write(n, d, b=True):
    if b:
        backup()
    t = BOOK.with_suffix(".tmp.xlsx")
    try:
        shutil.copy2(BOOK, t)
        with pd.ExcelWriter(
            t, engine="openpyxl", mode="a", if_sheet_exists="replace"
        ) as w:
            d.reindex(columns=SCHEMA[n]).to_excel(w, sheet_name=n, index=False)
        os.replace(t, BOOK)
        sync_remote_state()
        return True
    except Exception as e:
        st.error(f"Save failed; existing workbook preserved. {e}")
    finally:
        t.unlink(missing_ok=True)
    return False


def add(n, r):
    return write(
        n,
        pd.concat(
            [read(n), pd.DataFrame([{c: r.get(c, "") for c in SCHEMA[n]}])],
            ignore_index=True,
        ),
    )


def active(d):
    return (
        d
        if d.empty or "Is_Deleted" not in d
        else d[~d.Is_Deleted.astype(str).str.lower().isin(["true", "yes", "1"])]
    )


def ensure_products():
    d = active(read("Products"))
    have = set(d.Product_Name.astype(str)) if not d.empty else set()
    for i, n in enumerate(PRODUCTS, 1):
        if n not in have:
            add(
                "Products",
                {
                    "Product_ID": f"EQB-{i:02d}KG",
                    "Product_Name": n,
                    "Size_KG": i,
                    "HSN_SAC": "39232990",
                    "Sale_Rate": 0,
                    "GST_Rate": 5,
                    "Created_On": now(),
                    "Is_Deleted": False,
                },
            )


def fy(d):
    y = d.year if d.month >= 4 else d.year - 1
    return f"{y}-{str(y + 1)[-2:]}"


def invoice_number(d, consume=False):
    f = fy(d)
    seq = read("Sequence")
    ix = seq.index[seq.Financial_Year.astype(str) == f]
    last = int(float(seq.loc[ix[0], "Last_Number"])) if len(ix) else 0
    n = last + 1
    if consume:
        if len(ix):
            seq.loc[ix[0], ["Last_Number", "Updated_On"]] = [n, now()]
        else:
            seq = pd.concat(
                [
                    seq,
                    pd.DataFrame(
                        [{"Financial_Year": f, "Last_Number": n, "Updated_On": now()}]
                    ),
                ],
                ignore_index=True,
            )
        write("Sequence", seq)
    return f"{settings()['invoice_prefix']}/{f}/{n:06d}"


def upsert_customer(name, data):
    d = active(read("Customers"))
    z = d[d.Customer_Name.astype(str).str.lower() == name.lower()]
    if not z.empty:
        return str(z.iloc[0].Customer_ID)
    rid = uid("CUS")
    add(
        "Customers",
        {
            "Customer_ID": rid,
            "Customer_Name": name,
            **data,
            "Created_On": now(),
            "Is_Deleted": False,
        },
    )
    return rid


def soft_delete(sheet, idc, rid, reason):
    d = read(sheet)
    ix = d.index[d[idc].astype(str) == rid]
    if not len(ix):
        raise ValueError("Record not found")
    d.loc[ix[0], ["Is_Deleted", "Deleted_On", "Delete_Reason"]] = [True, now(), reason]
    write(sheet, d)


def delete_invoice(rid, reason):
    invoices = active(read("Invoices"))
    r = invoices[invoices.Invoice_ID.astype(str) == rid]
    if r.empty:
        raise ValueError("Invoice not found")
    items = read("Invoice_Items")
    mask = items.Invoice_ID.astype(str) == rid
    items.loc[mask, ["Is_Deleted", "Deleted_On", "Delete_Reason"]] = [
        True,
        now(),
        reason,
    ]
    write("Invoice_Items", items)
    soft_delete("Invoices", "Invoice_ID", rid, reason)


# ============================================================
# PDF RENDERING (unchanged EcoQuill signature invoice layout)
# ============================================================
def font_name(c):
    if not PDF_FONTS_READY:
        if c.get("bold") and c.get("italic"):
            return "Helvetica-BoldOblique"
        if c.get("bold"):
            return "Helvetica-Bold"
        if c.get("italic"):
            return "Helvetica-Oblique"
        return "Helvetica"
    return (
        "EQBoldItalic"
        if c.get("bold") and c.get("italic")
        else "EQBold"
        if c.get("bold")
        else "EQItalic"
        if c.get("italic")
        else "EQRegular"
    )


def rupee(v):
    return f"Rs. {float(v):,.2f}"


def page_brand(canvas, doc, invoice, s):
    canvas.saveState()
    w, h = A4
    dark = colors.HexColor("#20242E")
    accent = colors.HexColor("#A8C832")
    muted = colors.HexColor("#68707B")
    canvas.setFillColor(colors.white)
    canvas.rect(0, 0, w, h, fill=1, stroke=0)
    canvas.setFillColor(accent)
    canvas.rect(0, h - 5 * mm, w, 5 * mm, fill=1, stroke=0)
    canvas.rect(0, 0, w, 4 * mm, fill=1, stroke=0)
    p = canvas.beginPath()
    p.moveTo(76 * mm, h - 5 * mm)
    p.lineTo(w, h - 5 * mm)
    p.lineTo(w, h - 38 * mm)
    p.lineTo(91 * mm, h - 38 * mm)
    p.close()
    canvas.setFillColor(dark)
    canvas.drawPath(p, fill=1, stroke=0)
    a = canvas.beginPath()
    a.moveTo(0, h - 5 * mm)
    a.lineTo(160 * mm, h - 5 * mm)
    a.lineTo(149 * mm, h - 13 * mm)
    a.lineTo(0, h - 13 * mm)
    a.close()
    canvas.setFillColor(accent)
    canvas.drawPath(a, fill=1, stroke=0)
    canvas.setFillColor(dark)
    canvas.setFont("Helvetica-Bold", 25)
    invoice_title = (
        str(s.get("invoice_title", invoice.get("invoice_title", "INVOICE"))).strip()
        or "INVOICE"
    )
    canvas.drawString(16 * mm, h - 27 * mm, invoice_title[:40])
    canvas.setFont("Helvetica", 7.3)
    meta = [
        ("Invoice No.", invoice.get("invoice_no", "")),
        ("Invoice Date", invoice.get("invoice_date", "")),
        ("Due Date", invoice.get("due_date", "")),
    ]
    for i, (label, value) in enumerate(meta):
        y = h - (32 + i * 4.2) * mm
        canvas.setFillColor(muted)
        canvas.drawString(16 * mm, y, label + ":")
        canvas.setFillColor(dark)
        canvas.drawString(39 * mm, y, str(value))
    lx, ly, lw, lh = w - 87 * mm, h - 35 * mm, 38 * mm, 28 * mm
    canvas.setFillColor(colors.white)
    canvas.roundRect(lx, ly, lw, lh, 5 * mm, fill=1, stroke=0)
    if LOGO.exists():
        canvas.saveState()
        clip = canvas.beginPath()
        clip.roundRect(lx, ly, lw, lh, 5 * mm)
        canvas.clipPath(clip, stroke=0, fill=0)
        canvas.drawImage(
            str(LOGO),
            lx,
            ly,
            width=lw,
            height=lh,
            preserveAspectRatio=True,
            mask="auto",
            anchor="c",
        )
        canvas.restoreState()
    else:
        canvas.setFillColor(accent)
        canvas.setFont("Helvetica-Bold", 15)
        canvas.drawCentredString(lx + lw / 2, ly + lh / 2 - 3, "EQ")
    tx = lx + lw + 4 * mm
    maxw = w - 5 * mm - tx
    canvas.setFillColor(colors.white)
    canvas.setFont("Helvetica-Bold", 12)
    canvas.drawString(tx, h - 20 * mm, str(s.get("company_name", "EcoQuill")))
    tagline = str(s.get("company_tagline", ""))
    canvas.setFont("Helvetica", 7)
    lines = []
    cur = ""
    for wd in tagline.split():
        test = (cur + " " + wd).strip()
        if canvas.stringWidth(test, "Helvetica", 7) <= maxw:
            cur = test
        else:
            if cur:
                lines.append(cur)
            cur = wd
    if cur:
        lines.append(cur)
    for i, t in enumerate(lines[:3]):
        canvas.drawString(tx, h - (24.5 + i * 3.2) * mm, t)
    f = canvas.beginPath()
    f.moveTo(0, 4 * mm)
    f.lineTo(137 * mm, 4 * mm)
    f.lineTo(149 * mm, 18 * mm)
    f.lineTo(0, 18 * mm)
    f.close()
    canvas.setFillColor(dark)
    canvas.drawPath(f, fill=1, stroke=0)
    canvas.setFillColor(colors.white)
    canvas.setFont("Helvetica", 6.5)
    canvas.drawString(15 * mm, 10 * mm, str(s.get("company_phone", "")))
    canvas.drawString(65 * mm, 10 * mm, str(s.get("company_email", ""))[:38])
    canvas.drawString(
        116 * mm, 10 * mm, str(s.get("company_website", "www.ecoquill.in"))[:26]
    )
    canvas.restoreState()


def make_pdf(company, customer, invoice, items, totals, bank, terms, path=None):
    s = settings()
    c = s.get("invoice_body", DEFAULT["invoice_body"])
    dark, accent, grid = map(colors.HexColor, [c["dark"], c["accent"], c["grid"]])
    out = (
        Path(path)
        if path
        else INVOICES
        / f"EcoQuill_Invoice_{re.sub('[^A-Za-z0-9_-]', '_', invoice['invoice_no'])}.pdf"
    )
    doc = SimpleDocTemplate(
        str(out),
        pagesize=A4,
        leftMargin=15 * mm,
        rightMargin=15 * mm,
        topMargin=56 * mm,
        bottomMargin=25 * mm,
    )
    body = ParagraphStyle(
        "b",
        fontName="Helvetica",
        fontSize=float(c["font"]),
        leading=float(c["font"]) * 1.3,
        textColor=colors.HexColor("#4B5563"),
    )
    right = ParagraphStyle("r", parent=body, alignment=TA_RIGHT)
    center = ParagraphStyle("ct", parent=body, alignment=TA_CENTER)
    bold = ParagraphStyle("bd", parent=body, fontName="Helvetica-Bold", textColor=dark)
    head = ParagraphStyle(
        "hd",
        fontName="Helvetica-Bold",
        fontSize=7.2,
        leading=9,
        textColor=colors.white,
        alignment=TA_CENTER,
    )
    note = ParagraphStyle(
        "nt",
        fontName="Helvetica",
        fontSize=float(c["notes"]),
        leading=float(c["notes"]) * 1.35,
        textColor=colors.HexColor("#555B66"),
    )
    thank = ParagraphStyle(
        "th", fontName="Helvetica-BoldOblique", fontSize=10, leading=12, textColor=dark
    )
    story = []
    left = f"<b>Invoice To:</b><br/><b>{html.escape(str(customer.get('name', '')))}</b><br/>{html.escape(str(customer.get('address', '')))}<br/>{html.escape(str(customer.get('phone', '')))}<br/>{html.escape(str(customer.get('email', '')))}<br/>GSTIN: {html.escape(str(customer.get('gstin') or 'Not provided'))}"
    righttxt = f"""
<b>{html.escape(str(company.get("name", "EcoQuill")))}</b><br/>
{html.escape(str(company.get("address", "")))}<br/>
Phone: {html.escape(str(company.get("phone", "")))}<br/>
Email: {html.escape(str(company.get("email", "")))}<br/>
GSTIN: {html.escape(str(company.get("gstin", "")))}
"""
    t = Table(
        [[Paragraph(left, body), Paragraph(righttxt, right)]],
        colWidths=[112 * mm, 68 * mm],
    )
    t.setStyle(
        TableStyle(
            [
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("LEFTPADDING", (0, 0), (-1, -1), 0),
                ("RIGHTPADDING", (0, 0), (-1, -1), 0),
            ]
        )
    )
    story += [t, Spacer(1, 7 * mm)]
    headers = [c["sl"], c["description"]]
    fields = ["sl", "desc"]
    widths = [12, 76]
    if c["show_hsn"]:
        headers += ["HSN/SAC"]
        fields += ["hsn"]
        widths += [22]
    headers += [c["unit"], c["qty"]]
    fields += ["rate", "qty"]
    widths += [28, 24]
    if c["show_gst"]:
        headers += ["GST"]
        fields += ["gst"]
        widths += [18]
    headers += [c["total"]]
    fields += ["total"]
    widths += [40 if len(fields) == 4 else 28]
    f = 180 / sum(widths)
    widths = [x * f * mm for x in widths]
    rows = [[Paragraph(x, head) for x in headers]]
    for i in range(max(int(c["rows"]), len(items))):
        it = items[i] if i < len(items) else None
        rr = []
        for fld in fields:
            if fld == "sl":
                val = f"{i + 1:02d}" if it is not None else ""
            elif it is None:
                val = ""
            elif fld == "desc":
                val = html.escape(str(it["product_name"]))
            elif fld == "hsn":
                val = html.escape(str(it.get("hsn", "")))
            elif fld == "rate":
                val = rupee(it["rate"])
            elif fld == "qty":
                val = f"{it['quantity']:,.3f}"
            elif fld == "gst":
                val = f"{it['gst_percent']:,.2f}%"
            else:
                val = rupee(it["line_amount"] + it["gst_amount"])
            rr.append(
                Paragraph(
                    val, body if fld == "desc" else center if fld == "sl" else right
                )
            )
        rows.append(rr)
    table = Table(rows, colWidths=widths, repeatRows=1)
    pad = max(6, (float(c["row_mm"]) * mm - 9) / 2)
    table.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (1, 0), dark),
                ("BACKGROUND", (2, 0), (-1, 0), accent),
                ("GRID", (0, 0), (-1, -1), 0.35, grid),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("TOPPADDING", (0, 1), (-1, -1), pad),
                ("BOTTOMPADDING", (0, 1), (-1, -1), pad),
                ("TOPPADDING", (0, 0), (-1, 0), 8),
                ("BOTTOMPADDING", (0, 0), (-1, 0), 8),
                ("LEFTPADDING", (0, 0), (-1, -1), 6),
                ("RIGHTPADDING", (0, 0), (-1, -1), 6),
            ]
        )
    )
    story.append(table)
    pay = (
        f"<b>{html.escape(c['payment'])}</b><br/>Bank: {html.escape(bank['bank_name'])}<br/>Account Name: {html.escape(bank['account_name'])}<br/>Account No: {html.escape(bank['account_no'])}<br/>IFSC: {html.escape(bank['ifsc'])}<br/>UPI: {html.escape(bank['upi'])}<br/><br/><b>{html.escape(c['terms'])}</b><br/>"
        + "<br/>".join(html.escape(x) for x in terms)
    )
    tr = [
        [
            Paragraph(c["subtotal"], body),
            Paragraph(rupee(totals["taxable_value"]), right),
        ],
        [Paragraph(c["tax"], body), Paragraph(rupee(totals["gst_amount"]), right)],
    ]
    gi = len(tr)
    tr.append(
        [
            Paragraph(c["grand"], bold),
            Paragraph(
                rupee(totals["grand_total"]),
                ParagraphStyle(
                    "g", parent=right, fontName="Helvetica-Bold", textColor=colors.white
                ),
            ),
        ]
    )
    if c["show_paid"]:
        tr += [
            [
                Paragraph("Paid:", body),
                Paragraph(rupee(totals.get("paid_amount", 0)), right),
            ],
            [
                Paragraph("Balance Due:", bold),
                Paragraph(
                    rupee(totals.get("balance_due", totals["grand_total"])), right
                ),
            ],
        ]
    totals_table = Table(tr, colWidths=[37 * mm, 33 * mm])
    totals_table.setStyle(
        TableStyle(
            [
                ("GRID", (0, 0), (-1, -1), 0.35, grid),
                ("BACKGROUND", (0, gi), (-1, gi), accent),
                ("TOPPADDING", (0, 0), (-1, -1), 8),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 8),
                ("LEFTPADDING", (0, 0), (-1, -1), 7),
                ("RIGHTPADDING", (0, 0), (-1, -1), 7),
            ]
        )
    )
    low = Table([[Paragraph(pay, note), totals_table]], colWidths=[110 * mm, 70 * mm])
    low.setStyle(
        TableStyle(
            [
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("LEFTPADDING", (0, 0), (-1, -1), 0),
                ("RIGHTPADDING", (0, 0), (-1, -1), 0),
            ]
        )
    )
    story += [low, Spacer(1, 5 * mm)]
    sig = []
    if SIGN.exists():
        sig.append(
            Image(
                str(SIGN), width=float(c["sign_w"]) * mm, height=float(c["sign_h"]) * mm
            )
        )
    sig += [
        Paragraph(
            html.escape(str(s.get("authorized_signatory", "Your Name Here"))), bold
        ),
        Paragraph(html.escape(str(company.get("name", ""))), note),
    ]
    close = Table(
        [[Paragraph(html.escape(c["thanks"]), thank), sig]],
        colWidths=[116 * mm, 64 * mm],
    )
    close.setStyle(
        TableStyle(
            [
                ("VALIGN", (0, 0), (-1, -1), "BOTTOM"),
                ("ALIGN", (1, 0), (1, 0), "RIGHT"),
                ("LEFTPADDING", (0, 0), (-1, -1), 0),
                ("RIGHTPADDING", (0, 0), (-1, -1), 0),
            ]
        )
    )
    story.append(close)
    doc.build(
        story,
        onFirstPage=lambda x, y: page_brand(x, y, invoice, s),
        onLaterPages=lambda x, y: page_brand(x, y, invoice, s),
    )
    return out


# ============================================================
# UI CHROME
# ============================================================
def css():
    return """<style>
    :root{--bg:#d8d9dd;--panel:#f7f8fa;--white:#fff;--text:#17191f;--muted:#747983;--line:#e1e3e8;--black:#17181c;--blue:#315bec;--green:#1a8f70}
    .stApp,[data-testid=stAppViewContainer]{background:linear-gradient(135deg,#d7d8dc,#eef0f3)!important;color:var(--text)!important}
    [data-testid=stHeader]{background:transparent!important}
    .block-container{max-width:1500px;padding:1.25rem 1.8rem 3rem}
    [data-testid=stSidebar]{background:rgba(249,250,251,.98)!important;border-right:1px solid #d9dce2!important;box-shadow:10px 0 35px rgba(20,24,34,.08)!important}
    [data-testid=stSidebar] *{color:#343842!important}
    [data-testid=stSidebar] h3{color:#111318!important;font-size:21px!important}
    [data-testid=stSidebar] [role=radiogroup] label{padding:.66rem .78rem!important;margin:.1rem 0!important;border-radius:12px!important;font-weight:620!important}
    [data-testid=stSidebar] [role=radiogroup] label:hover{background:#eceef2!important}
    [data-testid=stSidebar] [role=radiogroup] label:has(input:checked){background:var(--black)!important;box-shadow:0 7px 15px rgba(0,0,0,.14)!important}
    [data-testid=stSidebar] [role=radiogroup] label:has(input:checked) *{color:white!important}
    [data-testid=stSidebar] img{border-radius:20px!important;object-fit:cover!important;border:4px solid white!important;box-shadow:0 9px 20px rgba(30,35,47,.16)!important}
    .hero{background:rgba(249,250,251,.98)!important;border:1px solid white!important;border-radius:24px!important;box-shadow:0 14px 34px rgba(30,35,47,.09)!important;color:var(--text)!important;padding:25px 28px;margin-bottom:18px}
    .hero h1{color:#111318!important;opacity:1!important;margin:0;font-size:30px}.hero p{color:#6f7480!important;opacity:1!important;margin:7px 0 0}
    .metrics{display:grid;grid-template-columns:repeat(auto-fit,minmax(180px,1fr));gap:13px;margin-bottom:18px}
    .metric,[data-testid=stForm],[data-testid=stExpander],[data-testid=stDataFrame]{background:rgba(250,251,252,.98)!important;border:1px solid white!important;border-radius:18px!important;box-shadow:0 10px 26px rgba(30,35,47,.07)!important}
    .metric{padding:16px}.metric label{font-size:12px;color:var(--muted)!important;font-weight:700}.metric strong{display:block;font-size:22px;color:var(--text)!important;margin:5px 0}.metric small{color:var(--muted)!important}
    .stApp h1,.stApp h2,.stApp h3,.stApp h4,.stApp p,.stApp label,.stApp span{color:#20232a!important}.stApp label{font-weight:650!important}
    .stTabs [data-baseweb=tab-list]{background:#f8f9fb!important;border:1px solid white!important;border-radius:14px!important;padding:5px!important;gap:5px!important;box-shadow:0 8px 20px rgba(30,35,47,.06)!important}
    .stTabs [data-baseweb=tab]{background:#eceef2!important;border-radius:10px!important;font-weight:700!important}.stTabs [data-baseweb=tab] *{color:#343943!important}
    .stTabs [aria-selected=true]{background:var(--black)!important}.stTabs [aria-selected=true] *{color:#fff!important}.stTabs [data-baseweb=tab-highlight]{display:none!important}
    .stTextInput input,.stNumberInput input,.stDateInput input,.stTextArea textarea,[data-baseweb=select]>div{background:#fff!important;color:#15171c!important;border:1px solid #ccd1da!important;border-radius:12px!important;opacity:1!important}
    input:disabled,[aria-disabled=true]{background:#eceff3!important;color:#505663!important;-webkit-text-fill-color:#505663!important;opacity:1!important}
    .stButton>button,.stDownloadButton>button,.stFormSubmitButton>button{background:var(--black)!important;color:#fff!important;border:0!important;border-radius:11px!important;font-weight:750!important;box-shadow:0 7px 15px rgba(0,0,0,.14)!important}.stButton>button *,.stDownloadButton>button *{color:#fff!important}.stButton>button:hover{background:#315bec!important}
    [data-testid=stAlert]{border-radius:14px!important}
    @media(max-width:800px){.block-container{padding:.8rem!important}.hero h1{font-size:25px!important}.metrics{grid-template-columns:1fr 1fr!important}}
    </style>"""


def header(t, s):
    st.markdown(
        f"<div class='hero'><h1>{html.escape(t)}</h1><p>{html.escape(s)}</p></div>",
        unsafe_allow_html=True,
    )


def metrics(xs):
    st.markdown(
        "<div class='metrics'>"
        + "".join(
            f"<div class='metric'><label>{html.escape(a)}</label><strong>{html.escape(b)}</strong><small>{html.escape(c)}</small></div>"
            for a, b, c in xs
        )
        + "</div>",
        unsafe_allow_html=True,
    )


def style_editor(s):
    labels = {
        "invoice_title": "Invoice Title",
        "invoice_number": "Invoice Number",
        "bill_heading": "Bill To Heading",
        "bill_body": "Bill To Body",
        "from_heading": "From Heading",
        "from_body": "From Body",
        "date": "Date / Supply / Payment",
        "table_header": "Product Table Header",
        "table_body": "Product Table Body",
        "totals": "Totals",
        "grand_total": "Grand Total",
        "notes": "Notes / Amount in Words",
        "payment": "Payment Information",
        "thank_you": "Thank You Message",
    }
    edited = {}
    for key, label in labels.items():
        c = s["font_styles"][key]
        st.markdown(f"**{label}**")
        a, b = st.columns(2)
        size = a.number_input("Size", 6.0, 40.0, float(c["size"]), 1.0, key=f"sz{key}")
        color = b.color_picker("Color", c["color"], key=f"cl{key}")
        x, y, z = st.columns(3)
        bold = x.checkbox("Bold", c["bold"], key=f"bo{key}")
        italic = y.checkbox("Italic", c["italic"], key=f"it{key}")
        underline = z.checkbox("Underline", c["underline"], key=f"un{key}")
        edited[key] = {
            "size": size,
            "color": color,
            "bold": bold,
            "italic": italic,
            "underline": underline,
        }
        st.divider()
    return edited


def delete_panel(sheet, idc, delete_func=None):
    d = active(read(sheet))
    if d.empty:
        st.info("No active records.")
        return
    labels = {
        str(r[idc]): " | ".join(
            str(r.get(c, ""))
            for c in [idc, "Invoice_Number", "Product_Name", "Customer_Name"]
            if str(r.get(c, ""))
        )
        for _, r in d.iterrows()
    }
    rid = st.selectbox(
        "Record", list(labels), format_func=lambda x: labels[x], key=f"del{sheet}"
    )
    reason = st.text_area("Deletion reason", key=f"reason{sheet}")
    confirm = st.checkbox("I confirm this controlled deletion.", key=f"confirm{sheet}")
    if st.button("Delete Selected Record", key=f"button{sheet}"):
        if not reason.strip():
            st.error("Deletion reason is required.")
        elif not confirm:
            st.error("Select confirmation.")
        else:
            try:
                (
                    delete_func(rid, reason)
                    if delete_func
                    else soft_delete(sheet, idc, rid, reason)
                )
                st.success("Record deleted and audit history retained.")
                st.rerun()
            except ValueError as e:
                st.error(str(e))


# ============================================================
# PAGE: INVOICE GENERATOR
# ============================================================
def invoice_page():
    s = settings()
    header(
        "Invoice Generator",
        "EcoQuill signature invoice with reliable ₹ typography. Generate on demand, no purchase entry needed.",
    )
    with st.sidebar.expander("Invoice Design Studio"):
        tabs = st.tabs(["Business", "Typography", "Invoice Body", "Assets"])
        x = s.copy()
        with tabs[0]:
            title_value = s.get("invoice_title", "INVOICE")
            title_presets = ["INVOICE", "TAX INVOICE", "Custom"]
            title_mode = st.selectbox(
                "Invoice Title",
                title_presets,
                index=0
                if title_value == "INVOICE"
                else 1
                if title_value == "TAX INVOICE"
                else 2,
                help="Choose a standard label or enter any custom invoice heading.",
            )
            custom_title = st.text_input(
                "Custom Invoice Title",
                value=title_value if title_mode == "Custom" else "",
                disabled=title_mode != "Custom",
            )
            x["invoice_title"] = (
                custom_title.strip() if title_mode == "Custom" else title_mode
            ) or "INVOICE"
            x["company_name"] = st.text_input("Business Name", s["company_name"])
            x["company_tagline"] = st.text_input("Tagline", s["company_tagline"])
            x["company_address"] = st.text_area("Address", s["company_address"])
            x["company_phone"] = st.text_input("Phone", s["company_phone"])
            x["company_email"] = st.text_input("Email", s["company_email"])
            x["company_gstin"] = st.text_input("GSTIN", s["company_gstin"])
            x["company_website"] = st.text_input(
                "Website", s.get("company_website", "www.ecoquill.in")
            )
            x["bank_name"] = st.text_input("Bank", s["bank_name"])
            x["account_name"] = st.text_input("Account Name", s["account_name"])
            x["account_no"] = st.text_input("Account Number", s["account_no"])
            x["ifsc_code"] = st.text_input("IFSC", s["ifsc_code"])
            x["upi_id"] = st.text_input("UPI", s["upi_id"])
            x["terms_conditions"] = st.text_area(
                "Terms", s["terms_conditions"], height=150
            )
        with tabs[1]:
            x["font_styles"] = style_editor(s)
        with tabs[2]:
            b = dict(s.get("invoice_body", DEFAULT["invoice_body"]))
            c1, c2, c3 = st.columns(3)
            b["dark"] = c1.color_picker("Dark header", b["dark"])
            b["accent"] = c2.color_picker("Accent", b["accent"])
            b["grid"] = c3.color_picker("Grid", b["grid"])
            c1, c2, c3 = st.columns(3)
            b["rows"] = c1.number_input("Visible rows", 1, 12, int(b["rows"]))
            b["row_mm"] = c2.number_input(
                "Row height mm", 8.0, 18.0, float(b["row_mm"]), 0.5
            )
            b["font"] = c3.number_input("Body font", 6.0, 11.0, float(b["font"]), 0.2)
            c1, c2, c3 = st.columns(3)
            b["show_hsn"] = c1.checkbox("Show HSN", b["show_hsn"])
            b["show_gst"] = c2.checkbox("Show GST", b["show_gst"])
            b["show_paid"] = c3.checkbox("Show Paid/Balance", b["show_paid"])
            for k, label in [
                ("sl", "SL heading"),
                ("description", "Description heading"),
                ("unit", "Unit price heading"),
                ("qty", "Quantity heading"),
                ("total", "Total heading"),
                ("payment", "Payment heading"),
                ("terms", "Terms heading"),
                ("subtotal", "Subtotal label"),
                ("tax", "Taxation label"),
                ("grand", "Grand total label"),
                ("thanks", "Thank-you text"),
            ]:
                b[k] = st.text_input(label, b[k], key="b_" + k)
            c1, c2 = st.columns(2)
            b["sign_w"] = c1.number_input(
                "Signature width mm", 20.0, 50.0, float(b["sign_w"])
            )
            b["sign_h"] = c2.number_input(
                "Signature height mm", 8.0, 25.0, float(b["sign_h"])
            )
            x["invoice_body"] = b
        with tabs[3]:
            logo = st.file_uploader(
                "Logo for top-right corner", type=["png", "jpg", "jpeg"]
            )
            sign = st.file_uploader(
                "Digital Signature for bottom-right sign section",
                type=["png", "jpg", "jpeg"],
            )
            st.caption("Transparent PNG logos produce the cleanest result.")
        if st.button("Save Invoice Design", width="stretch"):
            if logo:
                LOGO.write_bytes(logo.getbuffer())
            if sign:
                SIGN.write_bytes(sign.getbuffer())
            save_settings(x)
            st.success("Invoice design saved.")
            st.rerun()

    left, right = st.columns([1.1, 0.9])
    with left:
        name = st.text_input("Customer Name")
        address = st.text_area("Customer Address")
        phone = st.text_input("Customer Phone")
        wa = st.text_input("Customer WhatsApp")
        email = st.text_input("Customer Email Optional")
        gstin = st.text_input("Customer GSTIN Optional")
    with right:
        dt = st.date_input("Invoice Date")
        due = st.date_input("Due Date", value=dt)
        st.text_input("Automatic Invoice Number", invoice_number(dt), disabled=True)
        place = st.text_input("Place of Supply", s["default_place_of_supply"])
        paystatus = st.selectbox(
            "Payment Status", ["Pending", "Paid", "Partially Paid"]
        )
        paid = st.number_input("Paid Amount", 0.0)
        packing = st.number_input("Packing / Delivery Charges", 0.0, step=10.0)

    products = active(read("Products"))
    names = products.Product_Name.astype(str).tolist()
    st.subheader("Product Details")
    if not names:
        st.info(
            "No products in Product Master yet — add one on the Product Master page, or type a custom line below."
        )
    count = st.number_input("Product Rows", 1, 20, 1)
    rows = []
    for i in range(int(count)):
        a, b, c, d, e = st.columns([2, 1.2, 1, 1, 1])
        if names:
            pn = a.selectbox(
                f"Product {i + 1}", names + ["Custom / Other"], key=f"p{i}"
            )
        else:
            pn = "Custom / Other"
        if pn == "Custom / Other":
            pn = b.text_input(f"Custom Product Name {i + 1}", key=f"cp{i}")
            r_hsn = ""
            r_rate = 0.0
            r_gst = 5.0
        else:
            r = products[products.Product_Name.astype(str) == pn].iloc[0]
            b.text_input(
                f"HSN/SAC {i + 1}", value=str(r.HSN_SAC), disabled=True, key=f"h{i}"
            )
            r_hsn = r.HSN_SAC
            r_rate = float(r.Sale_Rate or 0)
            r_gst = float(r.GST_Rate or 5)
        kg = c.number_input(f"Quantity {i + 1}", 0.0, value=1.0, step=1.0, key=f"q{i}")
        rate = d.number_input(f"Rate {i + 1}", 0.0, value=r_rate, step=1.0, key=f"r{i}")
        gst = e.number_input(f"GST % {i + 1}", 0.0, 100.0, r_gst, 0.5, key=f"g{i}")
        line = money(kg * rate)
        rows.append(
            {
                "product_name": pn or "Item",
                "hsn": r_hsn,
                "quantity": kg,
                "rate": rate,
                "gst_percent": gst,
                "line_amount": line,
                "gst_amount": money(line * gst / 100),
            }
        )

    taxable = money(sum(x["line_amount"] for x in rows))
    gstamt = money(sum(x["gst_amount"] for x in rows))
    grand = money(taxable + gstamt + packing)
    metrics(
        [
            ("Taxable Value", rupee(taxable), "Item total"),
            ("GST", rupee(gstamt), "Calculated"),
            ("Packing", rupee(packing), "Additional"),
            ("Grand Total", rupee(grand), "Invoice total"),
        ]
    )

    def payload(ino, ps):
        return (
            {
                "name": s["company_name"],
                "tagline": s["company_tagline"],
                "address": s["company_address"],
                "phone": s["company_phone"],
                "email": s["company_email"],
                "gstin": s["company_gstin"],
            },
            {
                "name": name,
                "address": address,
                "phone": phone,
                "whatsapp": wa,
                "email": email,
                "gstin": gstin,
            },
            {
                "invoice_title": s.get("invoice_title", "INVOICE"),
                "invoice_no": ino,
                "invoice_date": dt.strftime("%d/%m/%Y"),
                "due_date": due.strftime("%d/%m/%Y"),
                "place_of_supply": place,
                "payment_status": ps,
            },
            {
                "taxable_value": taxable,
                "gst_amount": gstamt,
                "packing_charges": packing,
                "grand_total": grand,
                "paid_amount": paid,
                "balance_due": money(grand - paid),
            },
            {
                "bank_name": s["bank_name"],
                "account_name": s["account_name"],
                "account_no": s["account_no"],
                "ifsc": s["ifsc_code"],
                "upi": s["upi_id"],
            },
            [x.strip() for x in s["terms_conditions"].splitlines() if x.strip()],
        )

    pcol, gcol = st.columns(2)
    if pcol.button("Preview Invoice", width="stretch"):
        if not name or not address or not rows or taxable <= 0:
            st.error("Customer name, address and positive values are required.")
        else:
            company, customer, invoice, totals, bank, terms = payload(
                invoice_number(dt), paystatus
            )
            st.session_state.preview = str(
                make_pdf(
                    company,
                    customer,
                    invoice,
                    rows,
                    totals,
                    bank,
                    terms,
                    INVOICES / "_preview.pdf",
                )
            )
    if gcol.button("Generate Invoice", width="stretch"):
        if not name or not address or taxable <= 0:
            st.error("Customer name, address and positive values are required.")
        else:
            ino = invoice_number(dt, True)
            iid = uid("INV")
            cid = upsert_customer(
                name,
                {
                    "Address": address,
                    "Phone": phone,
                    "WhatsApp": wa,
                    "Email": email,
                    "GSTIN": gstin,
                },
            )
            intra = (
                place.lower().strip() == s["default_place_of_supply"].lower().strip()
            )
            cg = money(gstamt / 2) if intra else 0
            sg = cg if intra else 0
            ig = 0 if intra else gstamt
            out = money(grand - paid)
            ps = "Paid" if out == 0 else "Partially Paid" if paid else paystatus
            company, customer, invoice, totals, bank, terms = payload(ino, ps)
            path = make_pdf(company, customer, invoice, rows, totals, bank, terms)
            add(
                "Invoices",
                {
                    "Invoice_ID": iid,
                    "Invoice_Number": ino,
                    "Invoice_Date": dt,
                    "Customer_ID": cid,
                    "Customer_Name": name,
                    "Taxable_Value": taxable,
                    "CGST": cg,
                    "SGST": sg,
                    "IGST": ig,
                    "Total_GST": gstamt,
                    "Packing_Charges": packing,
                    "Grand_Total": grand,
                    "Paid_Amount": paid,
                    "Outstanding_Amount": out,
                    "Payment_Status": ps,
                    "PDF_Path": str(path),
                    "Created_On": now(),
                    "Is_Deleted": False,
                },
            )
            for x in rows:
                linepack = money(packing * x["line_amount"] / taxable) if taxable else 0
                linetotal = money(x["line_amount"] + x["gst_amount"] + linepack)
                add(
                    "Invoice_Items",
                    {
                        "Item_ID": uid("ITEM"),
                        "Invoice_ID": iid,
                        "Invoice_Number": ino,
                        "Product_ID": "",
                        "Product_Name": x["product_name"],
                        "Size_KG": "",
                        "HSN_SAC": x["hsn"],
                        "Quantity_KG": x["quantity"],
                        "Rate_Per_KG": x["rate"],
                        "Taxable_Value": x["line_amount"],
                        "GST_Rate": x["gst_percent"],
                        "GST_Amount": x["gst_amount"],
                        "Line_Total": linetotal,
                        "Created_On": now(),
                        "Is_Deleted": False,
                    },
                )
            st.session_state.generated = str(path)
            st.success(f"Invoice {ino} generated.")

    if st.session_state.get("preview") and Path(st.session_state.preview).exists():
        data = base64.b64encode(Path(st.session_state.preview).read_bytes()).decode()
        st.components.v1.html(
            f'<iframe src="data:application/pdf;base64,{data}" width="100%" height="850"></iframe>',
            height=870,
        )
    if st.session_state.get("generated") and Path(st.session_state.generated).exists():
        p = Path(st.session_state.generated)
        st.download_button(
            "Download Invoice", p.read_bytes(), p.name, "application/pdf"
        )


# ============================================================
# PAGE: INVOICE HISTORY
# ============================================================
def invoice_history():
    header(
        "Invoice History",
        "Every invoice generated, with PDF re-download and controlled deletion.",
    )
    a, b = st.tabs(["Invoices", "Delete Invoice"])
    with a:
        invoice_data = active(read("Invoices")).sort_values(
            "Created_On", ascending=False
        )
        st.dataframe(invoice_data, width="stretch", hide_index=True)
        if not invoice_data.empty:
            available = []
            for _, row in invoice_data.iterrows():
                pdf_path = Path(str(row.get("PDF_Path", "")))
                if not pdf_path.is_absolute():
                    pdf_path = BASE / pdf_path
                if pdf_path.exists() and pdf_path.is_file():
                    available.append(
                        (str(row.Invoice_ID), str(row.Invoice_Number), pdf_path)
                    )
            if available:
                labels = {rid: number for rid, number, _ in available}
                selected_pdf = st.selectbox(
                    "Download Historical Invoice",
                    list(labels),
                    format_func=lambda x: labels[x],
                )
                selected_path = next(
                    path for rid, _, path in available if rid == selected_pdf
                )
                st.download_button(
                    "Download Selected Invoice PDF",
                    selected_path.read_bytes(),
                    selected_path.name,
                    "application/pdf",
                    width="stretch",
                )
            else:
                st.info(
                    "Invoice records exist, but no historical PDF file is currently available in storage."
                )
    with b:
        delete_panel("Invoices", "Invoice_ID", delete_invoice)


# ============================================================
# PAGE: PRODUCT MASTER
# ============================================================
def masters_page():
    header(
        "Product Master",
        "Manage the product list, default rate and GST used by the invoice dropdown.",
    )
    st.subheader("Manage Products")
    with st.form("add_product_master", clear_on_submit=True):
        a, b, c = st.columns(3)
        product_name = a.text_input("Product Name")
        size_kg = b.number_input(
            "Size KG (optional)", min_value=0.0, value=0.0, step=0.001, format="%.3f"
        )
        hsn = c.text_input("HSN / SAC")
        a, b = st.columns(2)
        sale_rate = a.number_input("Default Rate", min_value=0.0, value=0.0, step=1.0)
        gst_rate = b.number_input(
            "Default GST %", min_value=0.0, max_value=100.0, value=5.0, step=0.5
        )
        add_product_clicked = st.form_submit_button("Add Product", width="stretch")
    if add_product_clicked:
        existing = active(read("Products"))
        duplicate = not existing.empty and product_name.strip().lower() in set(
            existing.Product_Name.astype(str).str.strip().str.lower()
        )
        if not product_name.strip():
            st.error("Product name is required.")
        elif duplicate:
            st.error("An active product with this name already exists.")
        else:
            rid = uid("PRD")
            if add(
                "Products",
                {
                    "Product_ID": rid,
                    "Product_Name": product_name.strip(),
                    "Size_KG": size_kg,
                    "HSN_SAC": hsn.strip(),
                    "Sale_Rate": sale_rate,
                    "GST_Rate": gst_rate,
                    "Created_On": now(),
                    "Is_Deleted": False,
                },
            ):
                st.success(
                    "Product added and is now available in the invoice dropdown."
                )
                st.rerun()
    st.dataframe(active(read("Products")), width="stretch", hide_index=True)
    st.subheader("Delete Product")
    st.caption(
        "Controlled deletion removes the product from future dropdowns but keeps historical invoices intact."
    )
    delete_panel("Products", "Product_ID")
    st.subheader("Customers on Record")
    st.dataframe(active(read("Customers")), width="stretch", hide_index=True)


def main():
    restore_remote_state_once()
    init()
    ensure_products()
    st.markdown(css(), unsafe_allow_html=True)
    s = settings()
    st.sidebar.markdown(f"### 🌿 {s['company_name']}\n{s['company_tagline']}")
    pages = {
        "▣  Invoice Generator": invoice_page,
        "▤  Invoice History": invoice_history,
        "◇  Product Master": masters_page,
    }
    page = st.sidebar.radio("Workspace", list(pages), label_visibility="collapsed")
    st.sidebar.caption("Signature invoice generator")
    st.sidebar.caption(persistence_status())
    pages[page]()


if __name__ == "__main__":
    main()
