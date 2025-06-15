import os
import sys
from dotenv import set_key
# Add parent directory to path for imports
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from CacheAPIValues import create_cache_data
from ExcelGenerator import make_xslx

# Create .env file path
env_file = os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(__file__))), '.env')

print("=== API Key Configuration ===")

# Get T212 API key
t212_key = input("Enter your Trading212 API key: ").strip()
if t212_key:
    set_key(env_file, "T212_API_KEY", f'"{t212_key}"')
    print("✓ T212 API key updated")

# Get OpenAI API key
openai_key = input("Enter your OpenAI API key: ").strip()
if openai_key:
    set_key(env_file, "OPENAI_API_KEY", f'"{openai_key}"')
    print("✓ OpenAI API key updated")

print(f"Configuration saved to: {env_file}")

create_cache_data()
make_xslx()