import os
import sys
import re
import psycopg2
import win32com.client
import pythoncom
from dotenv import load_dotenv

def safe_float(val, default=0.0):
    """Safely parse various values into float, handling None, formatting, and other structures."""
    if val is None:
        return default
    if isinstance(val, (int, float)):
        return float(val)
    val_str = str(val).strip()
    if not val_str:
        return default
    cleaned = re.sub(r'[^\d.]', '', val_str.replace(',', ''))
    try:
        return float(cleaned) if cleaned else default
    except ValueError:
        return default

def generate_html_template(html_path, active_slots, items_data, terms_data):
    """Generate high-fidelity HTML matching standard landscape comparison layout."""
    import datetime
    
    n_companies = len(active_slots)
    total_cols = 4 + 3 * n_companies + 2
    
    html = []
    html.append("<html>")
    html.append("<head>")
    html.append('<meta charset="utf-8">')
    html.append("<style>")
    html.append("body { font-family: Calibri, Arial, sans-serif; font-size: 8.5pt; margin: 0; padding: 0; }")
    html.append("table { border-collapse: collapse; width: 100%; border: 1px solid #A0A0A0; }")
    html.append("th, td { border: 1px solid #A0A0A0; padding: 5px; vertical-align: middle; }")
    html.append(".logo-box { background-color: #FAFAFA; font-size: 14pt; font-weight: bold; font-style: italic; text-align: center; }")
    html.append(".title-company { font-size: 11pt; font-weight: bold; text-align: center; }")
    html.append(".title-address { font-size: 8pt; text-align: center; }")
    html.append(".bold-text { font-weight: bold; }")
    html.append(".center-text { text-align: center; }")
    html.append(".left-text { text-align: left; }")
    html.append(".right-text { text-align: right; }")
    html.append(".header-bg { background-color: #F2F2F2; font-weight: bold; text-align: center; }")
    html.append(".supplier-bg { background-color: #EAF2F8; font-weight: bold; text-align: center; }")
    html.append(".total-bg { background-color: #FCF3CF; font-weight: bold; }")
    html.append("</style>")
    html.append("</head>")
    html.append("<body>")
    
    html.append("<table>")
    
    # Title Row 1
    html.append("  <tr>")
    html.append(f'    <td class="logo-box" rowspan="3" colspan="4">Standard Engineering<br><span style="font-size: 9pt; font-weight: normal; font-style: normal;">Where excellence is standard</span></td>')
    html.append(f'    <td class="title-company" colspan="{total_cols - 7}">Standard Engineering Technology Ltd</td>')
    html.append(f'    <td class="bold-text center-text" colspan="2">DATE:</td>')
    html.append(f'    <td class="center-text">{datetime.date.today().strftime("%d.%m.%Y")}</td>')
    html.append("  </tr>")
    
    # Title Row 2
    html.append("  <tr>")
    html.append(f'    <td class="title-address" colspan="{total_cols - 7}">ALINAGAR, CHETLAPOTHARAM VILLAGE, SURVEY NO 424, GADDAPOTHARAM GRAMPANCHAYAT</td>')
    html.append(f'    <td class="bold-text center-text" colspan="2">REF:</td>')
    html.append(f'    <td class="center-text">SGL/26-27/07/415</td>')
    html.append("  </tr>")
    
    # Title Row 3
    html.append("  <tr>")
    html.append(f'    <td class="title-address" colspan="{total_cols - 7}">JINNARAM MANDAL, SANGAREDDY DIST, Sangareddy, Telangana, 502313</td>')
    html.append(f'    <td class="center-text" colspan="3">&nbsp;</td>')
    html.append("  </tr>")
    
    # CS Plate header
    html.append("  <tr>")
    html.append(f'    <td class="header-bg" colspan="{total_cols}" style="font-size: 10pt; height: 25px;">CS Plate/Comparison Statement</td>')
    html.append("  </tr>")
    
    # Header Row 5
    html.append("  <tr>")
    html.append('    <th class="header-bg" rowspan="2">Sl.No.</th>')
    html.append('    <th class="header-bg" rowspan="2" style="width: 30%;">Item Description</th>')
    html.append('    <th class="header-bg" rowspan="2">QTY</th>')
    html.append('    <th class="header-bg" rowspan="2">UOM</th>')
    for slot_num, supplier_name in active_slots:
        html.append(f'    <th class="supplier-bg" colspan="3">{supplier_name}</th>')
    html.append('    <th class="header-bg" rowspan="2">LPP</th>')
    html.append('    <th class="header-bg" rowspan="2">PO NO/ Date</th>')
    html.append("  </tr>")
    
    # Header Row 6
    html.append("  <tr>")
    for _ in range(n_companies):
        html.append('    <th class="header-bg">Quoted Price</th>')
        html.append('    <th class="header-bg">After Discount</th>')
        html.append('    <th class="header-bg">Total Value</th>')
    html.append("  </tr>")
    
    # Populate data rows
    company_totals = {slot_num: 0.0 for slot_num, _ in active_slots}
    
    for idx, row in enumerate(items_data):
        db_id, db_desc, db_qty, db_uom = row[0], row[1], row[2], row[3]
        
        html.append("  <tr>")
        html.append(f'    <td class="center-text">{idx + 1}</td>')
        html.append(f'    <td class="left-text">{db_desc}</td>')
        html.append(f'    <td class="center-text">{safe_float(db_qty):.1f}</td>')
        html.append(f'    <td class="center-text">{db_uom}</td>')
        
        # Per-supplier values
        for slot_num, _ in active_slots:
            base_offset = 4 + (slot_num - 1) * 4
            rate = row[base_offset + 1]
            disc = row[base_offset + 2]
            amt = row[base_offset + 3]
            
            rate_val = safe_float(rate)
            disc_val = safe_float(disc)
            amt_val = safe_float(amt)
            
            if rate is not None and rate_val > 0:
                price_str = f"₹{rate_val:,.2f}"
                if disc_val > 0:
                    discounted_rate = rate_val * (1 - disc_val / 100.0)
                    disc_str = f"₹{discounted_rate:,.2f} ({disc_val:.0f}%)"
                else:
                    disc_str = "-"
                amt_str = f"₹{amt_val:,.2f}"
                company_totals[slot_num] += amt_val
            else:
                price_str = "-"
                disc_str = "-"
                amt_str = "-"
                
            html.append(f'    <td class="right-text">{price_str}</td>')
            html.append(f'    <td class="center-text">{disc_str}</td>')
            html.append(f'    <td class="right-text">{amt_str}</td>')
            
        html.append('    <td>&nbsp;</td>') # LPP
        html.append('    <td>&nbsp;</td>') # PO NO/ Date
        html.append("  </tr>")
        
    # TOTAL BASE VALUE Row
    html.append("  <tr>")
    html.append('    <td class="header-bg" colspan="4">TOTAL BASE VALUE</td>')
    for slot_num, _ in active_slots:
        total_val = company_totals[slot_num]
        html.append(f'    <td class="right-text bold-text" colspan="3">₹{total_val:,.2f}</td>')
    html.append('    <td>&nbsp;</td>')
    html.append('    <td>&nbsp;</td>')
    html.append("  </tr>")
    
    # TERMS & CONDITIONS section
    html.append("  <tr>")
    html.append(f'    <td class="header-bg left-text" colspan="{total_cols}">TERMS & CONDITIONS:</td>')
    html.append("  </tr>")
    
    # Terms rows
    terms_order = [
        "DISCOUNT",
        "P&F CHARGES",
        "FREIGHT CHARGES",
        "GST",
        "TOTAL VALUE",
        "PAYMENT TERMS",
        "DELIVERY SCHEDULE",
        "TEST CERTIFICATES",
        "QUOTATION REF/ DATE",
        "CONTACT DETAILS"
    ]
    
    import json
    
    def format_term_value(val):
        if not val:
            return "&nbsp;"
        val_str = str(val).strip()
        # Parse JSON structures if present (e.g. contact details)
        if (val_str.startswith("{") and val_str.endswith("}")) or (val_str.startswith("[") and val_str.endswith("]")):
            try:
                data = json.loads(val_str)
                if isinstance(data, dict):
                    parts = []
                    for k, v in data.items():
                        if v:
                            parts.append(f"{k.capitalize()}: {v}")
                    if parts:
                        return ", ".join(parts)
                elif isinstance(data, list):
                    return ", ".join(str(x) for x in data if x)
            except Exception:
                pass
        return val_str

    terms_map = {}
    for row in terms_data:
        term_name, val1, val2, val3 = row
        terms_map[term_name.upper()] = [val1, val2, val3]
        
    for label in terms_order:
        html.append("  <tr>")
        html.append(f'    <td class="bold-text left-text" colspan="4">{label}</td>')
        for slot_num, _ in active_slots:
            if label == "TOTAL VALUE":
                total_val = company_totals[slot_num]
                html.append(f'    <td class="right-text total-bg" colspan="3">₹{total_val:,.2f}</td>')
            else:
                vals = terms_map.get(label, [None, None, None])
                term_val = vals[slot_num - 1]
                term_val_clean = format_term_value(term_val)
                html.append(f'    <td class="center-text" colspan="3">{term_val_clean}</td>')
        html.append('    <td>&nbsp;</td>')
        html.append('    <td>&nbsp;</td>')
        html.append("  </tr>")
        
    html.append("</table>")
    
    # Signatures
    html.append('<br><br>')
    html.append('<table border="0" style="border: none; width: 100%; font-weight: bold; font-family: Calibri, Arial;">')
    html.append('  <tr style="border: none;">')
    html.append('    <td style="border: none; text-align: left; width: 33%;">PREPARED BY</td>')
    html.append('    <td style="border: none; text-align: center; width: 33%;">CHECKED BY</td>')
    html.append('    <td style="border: none; text-align: right; width: 33%;">APPROVED BY</td>')
    html.append('  </tr>')
    html.append('</table>')
    
    html.append("</body>")
    html.append("</html>")
    
    os.makedirs(os.path.dirname(html_path), exist_ok=True)
    with open(html_path, "w", encoding="utf-8") as f:
        f.write("\n".join(html))
    print(f"[HTML] Generated comparison sheet: {html_path}")

def export_html_to_pdf(html_path, pdf_path):
    """Load the generated HTML using MS Word COM and export to high-fidelity PDF."""
    print(f"[PDF] Compiling HTML into PDF report...")
    word = None
    doc = None
    try:
        pythoncom.CoInitialize()
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False
        word.DisplayAlerts = False
        
        abs_html = os.path.abspath(html_path)
        abs_pdf = os.path.abspath(pdf_path)
        
        doc = word.Documents.Open(abs_html)
        
        # Format page setup: Landscape, A4, 20pt Margins
        doc.PageSetup.Orientation = 1  # Landscape
        doc.PageSetup.PaperSize = 7     # A4
        doc.PageSetup.TopMargin = 20
        doc.PageSetup.BottomMargin = 20
        doc.PageSetup.LeftMargin = 20
        doc.PageSetup.RightMargin = 20
        
        # Save as PDF (Format type 17)
        doc.SaveAs(abs_pdf, FileFormat=17)
        print(f"[PDF] Generated PDF successfully: {pdf_path}")
    except Exception as e:
        print(f"[-] Error during Word COM PDF generation: {e}")
        raise e
    finally:
        try:
            if doc:
                doc.Close(False)
        except Exception:
            pass
        try:
            if word:
                word.Quit()
        except Exception:
            pass
        
        # Release COM references from python memory
        doc = None
        word = None
        pythoncom.CoUninitialize()
        
        # Force terminate winword process in case it is still active/hanging
        import subprocess
        try:
            subprocess.run("taskkill /f /im winword.exe", shell=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        except Exception:
            pass


# Absolute path to the project directory for RPA compatibility
try:
    PROJECT_DIR = os.path.dirname(os.path.abspath(__file__))
except NameError:
    PROJECT_DIR = r"C:\UiPathProjects\BlankProcess\SE"
ENV_PATH = os.path.join(PROJECT_DIR, ".env")

def run_pdf_generation_logic(conn, cursor):
    # Check active suppliers count
    cursor.execute("SELECT slot_number, supplier_name FROM supplier_slots WHERE supplier_name IS NOT NULL ORDER BY slot_number;")
    active_slots = cursor.fetchall()
    supplier_count = len(active_slots)
    
    print(f"[*] Found {supplier_count} active supplier(s) in the database.")
    
    if supplier_count < 3:
        print(f"[-] Cannot compile Comparison Statement. Need at least 3 suppliers, currently have {supplier_count}.")
        sys.exit(0)
        
    print("[*] Fetching side-by-side data from database...")
    
    # Get supplier items
    cursor.execute("""
        SELECT db_item_id, item_description, qty, uom,
               supplier_1_name, supplier_1_rate, supplier_1_discount, supplier_1_amount,
               supplier_2_name, supplier_2_rate, supplier_2_discount, supplier_2_amount,
               supplier_3_name, supplier_3_rate, supplier_3_discount, supplier_3_amount
        FROM supplier_items
        ORDER BY db_item_id;
    """)
    items_data = cursor.fetchall()
    
    # Get supplier terms
    cursor.execute("""
        SELECT term_name, supplier_1_value, supplier_2_value, supplier_3_value
        FROM supplier_terms;
    """)
    terms_data = cursor.fetchall()
    
    # Generate Comparison Statement HTML and PDF
    html_path = os.path.join(PROJECT_DIR, "Output", "temp_compare.html")
    pdf_path = os.path.join(PROJECT_DIR, "Output", "Output.pdf")
    
    os.makedirs(os.path.join(PROJECT_DIR, "Output"), exist_ok=True)
    
    generate_html_template(html_path, active_slots, items_data, terms_data)
    export_html_to_pdf(html_path, pdf_path)
    
    # Clean up temporary HTML
    if os.path.exists(html_path):
        os.remove(html_path)
        print(f"[HTML] Cleaned up temporary comparison file.")
        
    # Copy PDF to the Extraction Quotation folder
    target_extraction_pdf = r"C:\Extraction Quotation\Output\Output.pdf"
    try:
        os.makedirs(os.path.dirname(target_extraction_pdf), exist_ok=True)
        import shutil
        shutil.copy2(pdf_path, target_extraction_pdf)
        print(f"[+] Synced copy of PDF report to: {target_extraction_pdf}")
    except Exception as e:
        print(f"[-] Could not copy to target location: {e}")
        
    print("\n[+] SUCCESS: Standalone comparison statement generated successfully from the database.")
    
    # Clean up database data as requested
    print("[*] Clearing database data (transient tables and cache) and resetting slots/terms...")
    try:
        cursor.execute("TRUNCATE comparison_data, supplier_items, processed_files, raw_quotations CASCADE;")
        cursor.execute("UPDATE supplier_slots SET supplier_name = NULL;")
        cursor.execute("UPDATE supplier_terms SET supplier_1_value = NULL, supplier_2_value = NULL, supplier_3_value = NULL;")
        conn.commit()
        print("[+] Database tables successfully cleared and reset.")
    except Exception as e:
        conn.rollback()
        print(f"[-] Error while clearing database: {e}")
        
def main():
    load_dotenv(ENV_PATH)
    
    # Get Database Connection
    db_host = os.getenv("DB_HOST")
    db_port = os.getenv("DB_PORT")
    db_user = os.getenv("DB_USER")
    db_password = os.getenv("DB_PASSWORD")
    db_name = os.getenv("DB_NAME")
    
    print("[*] Connecting to database...")
    try:
        conn_params = {
            "host": db_host,
            "port": db_port,
            "user": db_user,
            "password": db_password,
            "database": db_name
        }
        if db_host and "neon.tech" in db_host:
            conn_params["sslmode"] = "require"
        conn = psycopg2.connect(**conn_params)
        cursor = conn.cursor()
    except Exception as e:
        print(f"[-] Database connection failed: {e}")
        sys.exit(1)
        
    try:
        run_pdf_generation_logic(conn, cursor)
    finally:
        if cursor:
            cursor.close()
        if conn:
            conn.close()

if __name__ == "__main__":
    main()
