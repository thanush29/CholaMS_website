# ai_utils.py
import google.generativeai as genai

# --- Direct API key import (⚠️ Not secure for production) ---
GEMINI_API_KEY = "AIzaSyD2Hf_IvqlC-e_Zxltm6C9YC0Dy5RxXCXo"

USE_GEMINI = bool(GEMINI_API_KEY)

if USE_GEMINI:
    genai.configure(api_key=GEMINI_API_KEY)

def generate_text(prompt, max_output_tokens=512):
    """Generate text with Gemini API key directly"""
    if not USE_GEMINI:
        return None
    try:
        model = genai.GenerativeModel("gemini-1.5-pro")
        resp = model.generate_content(
            prompt,
            generation_config={"max_output_tokens": max_output_tokens}
        )
        return resp.text if resp else None
    except Exception as e:
        print("⚠️ Gemini call failed:", e)
        return None
