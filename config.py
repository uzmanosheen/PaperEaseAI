import os
from dotenv import load_dotenv

# the key is read from .env — copy .env.example to .env and paste yours in
load_dotenv()

GEMINI_API_KEY = os.getenv("GEMINI_API_KEY")
MODEL_NAME = "gemini-3.6-flash"  # flash tier: fast and cheap for one-shot extraction