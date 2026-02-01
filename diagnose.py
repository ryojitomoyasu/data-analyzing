import sys
import os
import importlib

print(f"Python Version: {sys.version}")
print(f"Current Directory: {os.getcwd()}")

required_packages = [
    "streamlit",
    "pandas",
    "plotly",
    "google.genai",
    "dotenv",
    "openpyxl"
]

print("\n--- Checking Dependencies ---")
for package in required_packages:
    try:
        if package == "dotenv":
            importlib.import_module("dotenv")
        elif package == "google.genai":
            import google.genai
        else:
            importlib.import_module(package)
        print(f"[OK] {package} is installed.")
    except ImportError as e:
        print(f"[ERROR] {package} is NOT installed or failed to import: {e}")
    except Exception as e:
        print(f"[ERROR] {package} caused an error: {e}")

print("\n--- Checking File Structure ---")
files_to_check = ["app.py", "data_processor.py", "ai_engine.py", "pages/1_📊_Sales_Dashboard.py"]
for f in files_to_check:
    if os.path.exists(f):
        print(f"[OK] Found {f}")
    else:
        print(f"[ERROR] Missing {f}")
