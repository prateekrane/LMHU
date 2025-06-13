import requests
import logging
import time
import json
from config.settings import load_config

config = load_config()
API_KEY = config["GEMINI_API_KEY"]
TIMEOUT = config["TIMEOUT"]
ENDPOINT = "https://generativelanguage.googleapis.com/v1/models/gemini-2.0-flash:generateContent"

HEADERS = {
    "Content-Type": "application/json"
}


def call_gemini(prompt, retries=3):
    data = {
        "contents": [{"parts": [{"text": prompt}]}],
        "generationConfig": {
            "temperature": 0.7,
            "topK": 40,
            "topP": 0.95,
            "maxOutputTokens": 8192,
        }
    }
    params = {"key": API_KEY}

    for attempt in range(retries):
        try:
            response = requests.post(ENDPOINT, headers=HEADERS, json=data, params=params, timeout=TIMEOUT)
            response.raise_for_status()
            
            response_text = response.text

            # Clean the response text if it's wrapped in markdown code blocks
            if response_text.strip().startswith('```json') and response_text.strip().endswith('```'):
                response_text = response_text.strip()[7:-3].strip()
            elif response_text.strip().startswith('```') and response_text.strip().endswith('```'):
                 # Handle potential non-json code blocks as well
                 response_text = response_text.strip()[3:-3].strip()

            # --- START DEBUG PRINT --- 
            # Print the text content right before parsing to see exactly what was received
            print(f"DEBUG JSON Parse Attempt (Attempt {attempt+1}):\n---\n{response_text}\n---")
            # --- END DEBUG PRINT ---
            
            # Check if the cleaned text is empty
            if not response_text.strip():
                logging.error(f"AI service returned empty response body after stripping markdown (attempt {attempt+1}).")
                raise RuntimeError("AI service returned an empty response.")

            # Try to parse JSON, log response text if it fails
            try:
                return json.loads(response_text)
            except json.JSONDecodeError as e:
                logging.error(f"AI service returned non-JSON response (attempt {attempt+1}) during parsing: {response_text}. Error: {e}")
                if attempt < retries - 1:
                    time.sleep(2 ** attempt)
                else:
                    # Raise a RuntimeError with the problematic text for the UI to catch
                    raise RuntimeError(f"Failed to parse JSON from AI service after retries. Last response text causing error: {response_text}")

        except requests.RequestException as e:
            logging.error(f"AI service Request Error (attempt {attempt+1}): {e}")
            if attempt < retries - 1:
                time.sleep(2 ** attempt)
    # This part might be reached if all retries fail due to request exceptions
    raise RuntimeError("Failed to fetch response from AI service after retries due to request errors.")
