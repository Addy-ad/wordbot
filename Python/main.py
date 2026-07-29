# --- CONFIG AT THE BEGINNING ---
LLM_URL = "http://localhost:8080/v1" # llamacpp endpoint
MODEL_NAME = "Gemma4-E2B-Uncensored-IVA"
API_KEY = "XXXX"

from flask import Flask
from waitress import serve
import logging
from openai import OpenAI
import os

# Import route blueprints
from routes import process_callback, research_callback

# Check the operating system and clear the terminal screen accordingly
if os.name == 'nt':
    # Windows system command to clear screen
    os.system('cls')
else:
    # macOS and Linux system command to clear screen
    os.system('clear')

# Configure basic logging to show server events
logging.basicConfig(level=logging.INFO)
logging.getLogger("httpx").setLevel(logging.WARNING)
logging.getLogger("httpcore").setLevel(logging.WARNING)
logging.getLogger('waitress')

# Initialize Flask app
app = Flask(__name__)

# Initialize OpenAI client (make available to routes)
app.config['OPENAI_CLIENT'] = OpenAI(base_url=LLM_URL, api_key=API_KEY, timeout=120.0)
app.config['MODEL_NAME'] = MODEL_NAME
app.config['JSON_AS_ASCII'] = False

# Register blueprints
app.register_blueprint(process_callback)
app.register_blueprint(research_callback)

flaskPort = 3670

if __name__ == '__main__':
    print(f"Wordbot Python Server Initializing with endpoints...")
    print("--------------------------------------------------")
    print(f"LLM Endpoint:            http://127.0.0.1:{flaskPort}/process")
    print(f"Research Endpoint:       http://127.0.0.1:{flaskPort}/research")
    print("--------------------------------------------------")
    print(f"LLM Server:              {LLM_URL}")
    print(f"Selected LLM Model:      {MODEL_NAME}")
    print("Ready to receive requests from MS Word.")
    
    serve(app, host='127.0.0.1', port=flaskPort)