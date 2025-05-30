import os
from dotenv import load_dotenv

load_dotenv()

def load_config():
    return {
        "GEMINI_API_KEY": os.getenv("GEMINI_API_KEY", ""),
        "TIMEOUT": int(os.getenv("TIMEOUT", 30))
    }

