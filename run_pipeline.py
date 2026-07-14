import sys
import subprocess
import os

# Absolute path to the project directory — ensures RPA tools can invoke this
# script from any working directory without path resolution issues.
try:
    PROJECT_DIR = os.path.dirname(os.path.abspath(__file__))
except NameError:
    PROJECT_DIR = r"C:\UiPathProjects\BlankProcess\SE"
PYTHON_EXE = os.path.join(PROJECT_DIR, ".venv313", "Scripts", "python.exe")

def run_script(script_name, args=[]):
    script_path = os.path.join(PROJECT_DIR, script_name)
    print(f"\n=======================================================")
    print(f"[*] RUNNING STAGE: {script_name} {' '.join(args)}")
    print(f"=======================================================\n")
    try:
        # Run using the venv Python with the project directory as cwd
        subprocess.run(
            [PYTHON_EXE, script_path] + args,
            check=True,
            cwd=PROJECT_DIR
        )
        print(f"\n[+] STAGE COMPLETED: {script_name} successfully.")
        return True
    except subprocess.CalledProcessError as e:
        print(f"\n[-] STAGE FAILED: {script_name} failed with exit code {e.returncode}.")
        return False
    except FileNotFoundError:
        print(f"\n[-] STAGE FAILED: Python executable not found at {PYTHON_EXE}")
        print(f"    Please verify the virtual environment path.")
        return False

def clean_leftover_word_processes():
    print("[*] Cleaning up any leftover Word background processes...")
    try:
        # Runs the Windows taskkill command to silently close all MS Word tasks
        subprocess.run("taskkill /f /im winword.exe", shell=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    except Exception:
        pass

def main():
    # Pass any arguments (like --reset) to the sub-scripts
    extra_args = sys.argv[1:]
    
    try:
        # Clean up background processes first
        clean_leftover_word_processes()
        
        # 1. Run create_tables.py to ensure database and all tables exist
        if not run_script("create_tables.py"):
            print("[-] Pipeline aborted at Stage 1 (create_tables.py).")
            sys.exit(1)
        
        # 2. Run store_comparison_data.py to import/update reference baseline items
        if not run_script("store_comparison_data.py"):
            print("[-] Pipeline aborted at Stage 2 (store_comparison_data.py).")
            sys.exit(1)
            
        # 3. Run extract_with_groq.py to scan, extract, and match quotation files
        if not run_script("extract_with_groq.py", extra_args):
            print("[-] Pipeline aborted at Stage 3 (extract_with_groq.py).")
            sys.exit(1)
            
        # 4. Run generate_pdf_from_db.py to compile the comparison PDF report from the database
        if not run_script("generate_pdf_from_db.py"):
            print("[-] Pipeline aborted at Stage 4 (generate_pdf_from_db.py).")
            sys.exit(1)
            
        print("\n=======================================================")
        print("[+] SUCCESS: Complete Comparison Pipeline executed successfully.")
        print("=======================================================\n")
        
    finally:
        # Guarantee cleanup at termination so UiPath/RPA tools do not hang on background Word locks
        clean_leftover_word_processes()

main()
