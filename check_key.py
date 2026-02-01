
import os
from dotenv import load_dotenv

load_dotenv()
key = os.getenv("GOOGLE_API_KEY")

if key:
    print(f"Key found: {key[:4]}...{key[-4:]} (Length: {len(key)})")
else:
    print("No GOOGLE_API_KEY found in environment.")
