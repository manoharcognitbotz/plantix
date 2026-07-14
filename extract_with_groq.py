import os
import sys
import re
import json
import base64
import csv
import difflib
import email
from email import policy
from bs4 import BeautifulSoup
import pdfplumber
import psycopg2
import win32com.client
import pythoncom
from dotenv import load_dotenv
from groq import Groq

# ------------------------------------------------------------------------------
# DATABASE SETUP
# ------------------------------------------------------------------------------

def init_db(cursor):
    """Create tables for side-by-side incremental comparison statement storage."""
    # Check if we have the old vertical tables: if supplier_terms exists and has column supplier_name, drop old schema
    cursor.execute("""
        SELECT column_name 
        FROM information_schema.columns 
        WHERE table_name='supplier_terms' AND column_name='supplier_name';
    """)
    if cursor.fetchone():
        print("[*] Migrating schema: dropping old vertical database tables...")
        cursor.execute("DROP TABLE IF EXISTS supplier_items, supplier_terms, supplier_slots CASCADE;")
        
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS processed_files (
            id SERIAL PRIMARY KEY,
            filename VARCHAR(255) UNIQUE,
            processed_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        );
    """)
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS raw_quotations (
            filename VARCHAR(255) PRIMARY KEY,
            company_name VARCHAR(255),
            parsed_json TEXT,
            processed_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        );
    """)
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS supplier_slots (
            slot_number INTEGER PRIMARY KEY,
            supplier_name VARCHAR(255) UNIQUE
        );
    """)
    cursor.execute("""
        INSERT INTO supplier_slots (slot_number, supplier_name) 
        VALUES (1, NULL), (2, NULL), (3, NULL) 
        ON CONFLICT DO NOTHING;
    """)
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS supplier_items (
            db_item_id INTEGER PRIMARY KEY REFERENCES comparison_data(id),
            item_description TEXT,
            qty NUMERIC,
            uom VARCHAR(50),
            
            supplier_1_name VARCHAR(255),
            supplier_1_rate NUMERIC,
            supplier_1_discount NUMERIC,
            supplier_1_amount NUMERIC,
            
            supplier_2_name VARCHAR(255),
            supplier_2_rate NUMERIC,
            supplier_2_discount NUMERIC,
            supplier_2_amount NUMERIC,
            
            supplier_3_name VARCHAR(255),
            supplier_3_rate NUMERIC,
            supplier_3_discount NUMERIC,
            supplier_3_amount NUMERIC
        );
    """)
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS supplier_terms (
            term_name VARCHAR(100) PRIMARY KEY,
            supplier_1_value TEXT,
            supplier_2_value TEXT,
            supplier_3_value TEXT
        );
    """)
    cursor.execute("""
        INSERT INTO supplier_terms (term_name) VALUES
        ('DISCOUNT'),
        ('P&F CHARGES'),
        ('FREIGHT CHARGES'),
        ('GST'),
        ('PAYMENT TERMS'),
        ('DELIVERY SCHEDULE'),
        ('TEST CERTIFICATES'),
        ('QUOTATION REF/ DATE'),
        ('CONTACT DETAILS')
        ON CONFLICT DO NOTHING;
    """)


# ------------------------------------------------------------------------------
# TEXT & IMAGE EXTRACTION HELPERS
# ------------------------------------------------------------------------------

def extract_pdf_text(pdf_path):
    """Extract text from all pages of a PDF file using pdfplumber."""
    print(f"[PDF] Extracting text from: {os.path.basename(pdf_path)}")
    text = ""
    try:
        with pdfplumber.open(pdf_path) as pdf:
            for idx, page in enumerate(pdf.pages):
                page_text = page.extract_text()
                if page_text:
                    text += f"--- Page {idx+1} ---\n{page_text}\n"
    except Exception as e:
        print(f"[-] Error reading PDF {pdf_path}: {e}")
    return text

def extract_doc_text(doc_path):
    """Extract text from Word (.doc/.docx) documents using MS Word COM Automation."""
    print(f"[Word] Extracting text from: {os.path.basename(doc_path)}")
    text = ""
    word = None
    doc = None
    try:
        pythoncom.CoInitialize()
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False
        word.DisplayAlerts = False
        
        abs_path = os.path.abspath(doc_path)
        doc = word.Documents.Open(abs_path)
        text = doc.Content.Text
    except Exception as e:
        print(f"[-] Error reading Word document {doc_path}: {e}")
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
        pythoncom.CoUninitialize()
    return text

def extract_eml_text(eml_path):
    """Extract headers, plain text, and HTML body text from an EML file."""
    print(f"[EML] Extracting text from: {os.path.basename(eml_path)}")
    text = ""
    try:
        with open(eml_path, 'rb') as f:
            msg = email.message_from_binary_file(f, policy=policy.default)
            
        # Collect headers and protect email addresses from being treated as HTML tags
        from_header = str(msg.get('from', '')).replace('<', '[').replace('>', ']')
        to_header = str(msg.get('to', '')).replace('<', '[').replace('>', ']')
        headers = f"Subject: {msg.get('subject', '')}\nFrom: {from_header}\nTo: {to_header}\nDate: {msg.get('date', '')}\n\n"
        
        html_body = ''
        plain_text = ''
        for part in msg.walk():
            ct = part.get_content_type()
            if ct == 'text/html' and not html_body:
                html_body = part.get_payload(decode=True).decode(part.get_content_charset() or 'utf-8', errors='ignore')
            elif ct == 'text/plain' and not plain_text:
                plain_text = part.get_payload(decode=True).decode(part.get_content_charset() or 'utf-8', errors='ignore')
                
        # Fallbacks
        if not html_body and msg.get_content_type() == 'text/html':
            html_body = msg.get_payload(decode=True).decode(msg.get_content_charset() or 'utf-8', errors='ignore')
        if not plain_text and msg.get_content_type() == 'text/plain':
            plain_text = msg.get_payload(decode=True).decode(msg.get_content_charset() or 'utf-8', errors='ignore')
            
        body_text = ""
        if html_body:
            soup = BeautifulSoup(html_body, 'lxml')
            body_text = soup.get_text(separator='\n')
        else:
            body_text = plain_text
            
        text = headers + body_text
    except Exception as e:
        print(f"[-] Error reading EML file {eml_path}: {e}")
    return text

def extract_csv_text(csv_path):
    """Extract plain text from CSV file."""
    print(f"[CSV] Reading text from: {os.path.basename(csv_path)}")
    try:
        with open(csv_path, 'r', encoding='utf-8', errors='ignore') as f:
            return f.read()
    except Exception as e:
        print(f"[-] Error reading CSV file {csv_path}: {e}")
        return ""

def encode_image(image_path):
    """Encode an image file into base64 string."""
    with open(image_path, "rb") as image_file:
        return base64.b64encode(image_file.read()).decode('utf-8')

# ------------------------------------------------------------------------------
# GROQ PARSERS (TEXT & VISION)
# ------------------------------------------------------------------------------

def get_system_prompt():
    return (
        "You are an expert procurement and database assistant.\n"
        "Analyze the provided document and extract the quotation details.\n"
        "Ensure you detect if this document is actually a commercial quotation, estimate, or price sheet.\n"
        "If it is not a quotation/estimate or contains no products/prices, set 'is_quotation' to false.\n"
        "Otherwise, set 'is_quotation' to true and extract:\n"
        "- company_name: Name of the supplier/issuer. For emails, look for a clear supplier name in the email body (signatures, introductions, footer) first. If not found in the body, extract the company name from the sender's email address or domain name (found in the 'From:' header, e.g., 'cognitbotz' from '@cognitbotz.com'). Do NOT guess or concatenate item makes/brands (like 'Dowels' or 'MBP') as the company name to avoid printing inaccurate supplier data in the PDF.\n"
        "- items: a list of objects, each containing:\n"
        "    - description: description of the product or service\n"
        "    - uom: unit of measure (e.g. Nos, Set, Pcs, Mtr, or null if unknown)\n"
        "    - quantity: number of units (float, e.g., 30.0)\n"
        "    - rate: price per unit (float, e.g., 1031.0)\n"
        "    - discount_percent: discount percentage applied to this item (float, e.g. 42.0 for 42%, or 0.0 if none. Must be a raw float, do not write expressions)\n"
        "    - amount: total amount for this item (float). Compute this mathematically (e.g., rate * quantity * (1 - discount/100)). This MUST be a single raw numeric float value (e.g., 17939.4). Do NOT output formulas, expressions, or operators like *, /, -, or +.\n"
        "- terms_and_conditions: an object containing summary strings for:\n"
        "    - discount (general discount terms, if any)\n"
        "    - p_f_charges (packing and forwarding details)\n"
        "    - freight_charges (freight/shipping details)\n"
        "    - gst (tax/GST details, e.g. '18%' or 'inclusive')\n"
        "    - payment_terms (payment conditions)\n"
        "    - delivery_schedule (delivery timeline)\n"
        "    - test_certificates (test certificate details, if mentioned)\n"
        "    - quotation_ref_date (quotation ref number and date)\n"
        "    - contact_details (phone/email/address)\n\n"
        "You MUST return a valid JSON object matching the requested schema. Do not write any explanations."
    )

def parse_text_with_groq(text, api_key):
    """Call Groq API using JSON mode to extract structured quotation details from text."""
    import time
    client = Groq(api_key=api_key)
    models = ["llama-3.3-70b-versatile"]
    for model in models:
        max_retries = 3
        for attempt in range(max_retries):
            try:
                print(f"[*] Attempting parse with Groq model: {model} (Attempt {attempt+1}/{max_retries})...")
                response = client.chat.completions.create(
                    messages=[
                        {"role": "system", "content": get_system_prompt()},
                        {"role": "user", "content": f"Document Text:\n\n{text}"}
                    ],
                    model=model,
                    response_format={"type": "json_object"}
                )
                result = json.loads(response.choices[0].message.content)
                return result
            except Exception as e:
                err_str = str(e)
                if ("429" in err_str or "rate limit" in err_str.lower()) and attempt < max_retries - 1:
                    wait_time = (attempt + 1) * 3
                    print(f"[*] Groq rate limit hit. Retrying in {wait_time} seconds...")
                    time.sleep(wait_time)
                else:
                    print(f"[-] Error calling Groq Text API with model {model}: {e}")
                    break
    return None

def parse_image_with_groq(image_path, api_key):
    """Call Groq Vision API to parse an image of a quotation."""
    print(f"[Vision] Calling Groq LLM to parse image: {os.path.basename(image_path)}...")
    import time
    client = Groq(api_key=api_key)
    max_retries = 3
    for attempt in range(max_retries):
        try:
            base64_image = encode_image(image_path)
            response = client.chat.completions.create(
                messages=[
                    {
                        "role": "user",
                        "content": [
                            {"type": "text", "text": get_system_prompt()},
                            {
                                "type": "image_url",
                                "image_url": {
                                    "url": f"data:image/jpeg;base64,{base64_image}"
                                }
                            }
                        ]
                    }
                ],
                model="llama-3.2-90b-vision-preview",
                response_format={"type": "json_object"}
            )
            result = json.loads(response.choices[0].message.content)
            return result
        except Exception as e:
            err_str = str(e)
            if ("429" in err_str or "rate limit" in err_str.lower()) and attempt < max_retries - 1:
                wait_time = (attempt + 1) * 3
                print(f"[*] Groq rate limit hit. Retrying in {wait_time} seconds...")
                time.sleep(wait_time)
            else:
                print(f"[-] Error calling Groq Vision API: {e}")
                return None

# ------------------------------------------------------------------------------
# FUZZY TEXT NORMALIZATION & MATCHING
# ------------------------------------------------------------------------------

def normalize_text(text):
    """Normalize description text to improve similarity score matching."""
    if not text:
        return ""
    text = text.lower()
    text = re.sub(r'make\s*:\s*\w+', '', text)   # remove "Make: Dowels"
    text = re.sub(r'\bmake\b', '', text)           # remove "make"
    text = re.sub(r'[^a-z0-9\s]', ' ', text)      # keep alphanumeric only
    text = re.sub(r'\s+', ' ', text)               # collapse extra spaces
    return text.strip()

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
# HTML template generation and PDF export are now fully delegated to generate_pdf_from_db.py.
# ------------------------------------------------------------------------------
# MAIN PIPELINE
# ------------------------------------------------------------------------------

# Absolute path to the project directory for RPA compatibility
try:
    PROJECT_DIR = os.path.dirname(os.path.abspath(__file__))
except NameError:
    PROJECT_DIR = r"C:\UiPathProjects\BlankProcess\SE"
ENV_PATH = os.path.join(PROJECT_DIR, ".env")

def run_pipeline_logic(conn, cursor, groq_api_key):
    # Initialize database tables
    init_db(cursor)
    conn.commit()
    
    # Check for CLI flags
    compile_only = False
    if len(sys.argv) > 1:
        if sys.argv[1] == '--reset':
            print("[*] Resetting database tables as requested...")
            cursor.execute("TRUNCATE processed_files, supplier_slots, supplier_items, supplier_terms CASCADE;")
            cursor.execute("INSERT INTO supplier_slots (slot_number, supplier_name) VALUES (1, NULL), (2, NULL), (3, NULL) ON CONFLICT DO NOTHING;")
            cursor.execute("""
                INSERT INTO supplier_terms (term_name) VALUES
                ('DISCOUNT'), ('P&F CHARGES'), ('FREIGHT CHARGES'), ('GST'), ('PAYMENT TERMS'),
                ('DELIVERY SCHEDULE'), ('TEST CERTIFICATES'), ('QUOTATION REF/ DATE'), ('CONTACT DETAILS')
                ON CONFLICT DO NOTHING;
            """)
            conn.commit()
            print("[+] Tables reset completed.")
        elif sys.argv[1] == '--clear-cache':
            print("[*] Clearing raw quotations cache as requested...")
            cursor.execute("TRUNCATE raw_quotations CASCADE;")
            conn.commit()
            print("[+] Cache cleared.")
        elif sys.argv[1] == '--compile-only':
            compile_only = True
            print("[*] Running in compile-only mode. Generating PDF directly from the database (0 Groq tokens used).")
        
    # 3. Retrieve baseline items from database
    cursor.execute("SELECT id, pr_item_code, uom, indent_quantity FROM comparison_data ORDER BY id;")
    db_items = cursor.fetchall()
    
    if not db_items:
        print("[-] Error: No baseline items found in 'comparison_data' table.")
        sys.exit(1)
        
    supported_files = []
    new_files_processed = 0
    
    if compile_only:
        print("[*] Skipping scan of Downloads folder and Groq LLM calls.")
    else:
        # Persist existing data and processed files so we can add quotations one by one.
        # We do not clear comparison tables at startup anymore.
        print("[*] Persisting existing data. Scanning for new files...")
        
        # 4. Scan Downloads folder for files
        downloads_dir = os.path.join(PROJECT_DIR, "Download")
        if not os.path.exists(downloads_dir):
            print(f"[-] Downloads directory not found at {downloads_dir}")
            sys.exit(1)
            
        for f in os.listdir(downloads_dir):
            if f.startswith('~$'):
                continue
            ext = os.path.splitext(f)[1].lower()
            if ext in ['.pdf', '.doc', '.docx', '.eml', '.csv', '.png', '.jpg', '.jpeg']:
                supported_files.append(os.path.join(downloads_dir, f))
                
        # Sort files chronologically (oldest / first arrived file processed first)
        supported_files.sort(key=os.path.getmtime)
                
        print(f"[*] Found {len(supported_files)} file(s) in Download directory.")
    
    # 5. Extract text and parse with Groq or load from cache
    for f_path in supported_files:
        filename = os.path.basename(f_path)
        ext = os.path.splitext(filename)[1].lower()
        
        # Check if the file has already been successfully stored in this batch
        cursor.execute("SELECT 1 FROM processed_files WHERE filename = %s;", (filename,))
        if cursor.fetchone():
            print(f"[*] File '{filename}' has already been processed and stored. Skipping.")
            continue
            
        # Check if we have cached raw quotation JSON in database
        cursor.execute("SELECT parsed_json FROM raw_quotations WHERE filename = %s;", (filename,))
        cache_row = cursor.fetchone()
        
        parsed_data = None
        if cache_row:
            print(f"[*] Found cached raw quotation for '{filename}' in database. Skipping Groq API call.")
            try:
                parsed_data = json.loads(cache_row[0])
            except Exception as e:
                print(f"[-] Failed to parse cached JSON for {filename}: {e}")
                
        if not parsed_data:
            print(f"[*] Processing new file: {filename}")
            
            # Parse based on file type
            if ext in ['.png', '.jpg', '.jpeg']:
                # Call vision LLM
                parsed_data = parse_image_with_groq(f_path, groq_api_key)
            else:
                # Call text extraction + text LLM
                text = ""
                if ext == '.pdf':
                    text = extract_pdf_text(f_path)
                elif ext in ['.doc', '.docx']:
                    text = extract_doc_text(f_path)
                elif ext == '.eml':
                    text = extract_eml_text(f_path)
                elif ext == '.csv':
                    text = extract_csv_text(f_path)
                    
                if not text.strip():
                    print(f"[-] Warning: No text extracted from {filename}. Skipping.")
                    continue
                    
                print(f"[*] Calling Groq LLM to parse text: {filename}...")
                parsed_data = parse_text_with_groq(text, groq_api_key)
                
            if not parsed_data:
                print(f"[-] Failed to parse data from {filename}. Skipping.")
                continue
                

            
        co_name = (parsed_data.get("company_name") or "").strip() or f"Supplier ({filename})"
        supplier_items = parsed_data.get("items") or []
        terms = parsed_data.get("terms_and_conditions") or {}
        
        print(f"[+] Groq parsed: '{co_name}' with {len(supplier_items)} items.")
        
        # 1. Perform Fuzzy Match with Database Baseline in memory FIRST
        matched_any = False
        temp_matches = [] # list of tuples: (db_id, rate, disc, amt)
        
        for db_item in db_items:
            db_id, db_desc, db_uom, db_qty = db_item
            
            best_score = 0.0
            best_idx = -1
            
            norm_db = normalize_text(db_desc)
            for idx, item in enumerate(supplier_items):
                norm_supp = normalize_text(item.get('description', ''))
                score = difflib.SequenceMatcher(None, norm_db, norm_supp).ratio()
                
                # Boost match score if UOM matches
                if item.get('uom') and db_uom and item.get('uom').lower() == db_uom.lower():
                    score += 0.05
                    
                if score > best_score:
                    best_score = score
                    best_idx = idx
            
            if best_score >= 0.40 and best_idx != -1:
                matched_item = supplier_items[best_idx]
                matched_any = True
                
                rate = safe_float(matched_item.get('rate'))
                disc = safe_float(matched_item.get('discount_percent'))
                amt = rate * (1 - disc / 100.0) * safe_float(db_qty)
                
                temp_matches.append((db_id, rate, disc, amt))
                
        # 2. If no items matched existing database data, move the file to an Unmatched folder and do NOT store anything in the database!
        if not matched_any:
            print(f"[-] Quotation '{filename}' has 0 matching items against reference database items.")
            print("[*] Extracted items from this file were:")
            for item in supplier_items:
                print(f"    - Description: '{item.get('description')}', UOM: '{item.get('uom')}', Qty: '{item.get('quantity')}', Rate: '{item.get('rate')}'")
            
            # Move to Download/Unmatched folder instead of deleting
            unmatched_dir = os.path.join(downloads_dir, "Unmatched")
            try:
                os.makedirs(unmatched_dir, exist_ok=True)
                dest_path = os.path.join(unmatched_dir, filename)
                # If target file already exists in Unmatched, remove it first to avoid collision
                if os.path.exists(dest_path):
                    os.remove(dest_path)
                import shutil
                shutil.move(f_path, dest_path)
                print(f"[Move] Moved unmatched quotation file to: Download/Unmatched/{filename}")
            except Exception as e:
                print(f"[-] Failed to move unmatched file {filename} to Unmatched folder: {e}")
                
            cursor.execute("INSERT INTO processed_files (filename) VALUES (%s) ON CONFLICT DO NOTHING;", (filename,))
            conn.commit()
            continue
            
        # 3. IF matched successfully, now we save to raw_quotations cache
        co_name_temp = co_name
        cursor.execute("""
            INSERT INTO raw_quotations (filename, company_name, parsed_json)
            VALUES (%s, %s, %s)
            ON CONFLICT (filename) DO UPDATE SET
                company_name = EXCLUDED.company_name,
                parsed_json = EXCLUDED.parsed_json;
        """, (filename, co_name_temp, json.dumps(parsed_data)))
        conn.commit()
        
        # 4. Determine slot for this supplier
        cursor.execute("SELECT slot_number FROM supplier_slots WHERE supplier_name = %s;", (co_name,))
        slot_row = cursor.fetchone()
        if slot_row:
            slot = slot_row[0]
        else:
            cursor.execute("SELECT slot_number FROM supplier_slots WHERE supplier_name IS NULL ORDER BY slot_number LIMIT 1;")
            empty_row = cursor.fetchone()
            if empty_row:
                slot = empty_row[0]
                cursor.execute("UPDATE supplier_slots SET supplier_name = %s WHERE slot_number = %s;", (co_name, slot))
            else:
                print(f"[-] Warning: No empty slot left for {co_name}. Skipping slot assignment.")
                continue
                
        # 5. Keep supplier_items in sync and populate empty rows
        cursor.execute("DELETE FROM supplier_items WHERE db_item_id NOT IN (SELECT id FROM comparison_data);")
        cursor.execute("""
            INSERT INTO supplier_items (db_item_id, item_description, qty, uom)
            SELECT id, pr_item_code, indent_quantity, uom FROM comparison_data
            ON CONFLICT (db_item_id) DO UPDATE SET
                item_description = EXCLUDED.item_description,
                qty = EXCLUDED.qty,
                uom = EXCLUDED.uom;
        """)
        
        # 6. Store matching item values in DB
        for db_id, rate, disc, amt in temp_matches:
            update_query = f"""
                UPDATE supplier_items SET
                    supplier_{slot}_name = %s,
                    supplier_{slot}_rate = %s,
                    supplier_{slot}_discount = %s,
                    supplier_{slot}_amount = %s
                WHERE db_item_id = %s;
            """
            cursor.execute(update_query, (co_name, rate, disc, amt, db_id))
            
        # 7. Update terms in DB
        terms_mapping = {
            'DISCOUNT': terms.get('discount'),
            'P&F CHARGES': terms.get('p_f_charges'),
            'FREIGHT CHARGES': terms.get('freight_charges'),
            'GST': terms.get('gst'),
            'PAYMENT TERMS': terms.get('payment_terms'),
            'DELIVERY SCHEDULE': terms.get('delivery_schedule'),
            'TEST CERTIFICATES': terms.get('test_certificates'),
            'QUOTATION REF/ DATE': terms.get('quotation_ref_date'),
            'CONTACT DETAILS': terms.get('contact_details')
        }
        
        for term_name, term_val in terms_mapping.items():
            if term_val is not None:
                if isinstance(term_val, (dict, list)):
                    term_val = json.dumps(term_val)
                else:
                    term_val = str(term_val)
            
            update_term_query = f"""
                UPDATE supplier_terms SET
                    supplier_{slot}_value = %s
                WHERE term_name = %s;
            """
            cursor.execute(update_term_query, (term_val, term_name))
            
        # 8. Log file as processed
        cursor.execute("INSERT INTO processed_files (filename) VALUES (%s) ON CONFLICT DO NOTHING;", (filename,))
        conn.commit()
        new_files_processed += 1
        print(f"[+] File '{filename}' successfully imported to database in slot {slot}.")
        
    print(f"[*] Import pass finished. {new_files_processed} new file(s) imported.")
    
    # 6. Check unique suppliers count
    cursor.execute("SELECT slot_number, supplier_name FROM supplier_slots WHERE supplier_name IS NOT NULL ORDER BY slot_number;")
    active_slots = cursor.fetchall()
    supplier_count = len(active_slots)
    
    print(f"[*] Current unique suppliers in database: {supplier_count} ({', '.join([r[1] for r in active_slots]) if active_slots else 'None'})")
    
    if supplier_count >= 3:
        print("[*] Reached target: >= 3 companies in database. PDF compilation is deferred to the next stage (generate_pdf_from_db.py).")
    else:
        print(f"[*] PDF Compilation deferred. Need >= 3 suppliers, currently have {supplier_count}.")
        
def main():
    load_dotenv(ENV_PATH)
    
    # 1. Check Groq API Key
    groq_api_key = os.getenv("GROQ_API_KEY")
    if not groq_api_key:
        print("[-] Error: GROQ_API_KEY is not defined in the environment variables or .env file.")
        print("[*] Please open 'C:\\UiPathProjects\\BlankProcess\\SE\\.env' and add: GROQ_API_KEY=your_groq_api_key_here")
        sys.exit(1)
        
    # 2. Get Database Connection
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
        run_pipeline_logic(conn, cursor, groq_api_key)
    finally:
        if cursor:
            cursor.close()
        if conn:
            conn.close()

if __name__ == "__main__":
    main()
