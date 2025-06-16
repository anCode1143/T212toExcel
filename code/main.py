import os
import sys
import shutil
from dotenv import set_key
# Add parent directory to path for imports
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

# Define paths
project_root = os.path.dirname(os.path.dirname(__file__))
env_file = os.path.join(project_root, '.env')
cache_dir = os.path.join(project_root, 'cache')

# Reset cache
if os.path.exists(env_file):
    os.remove(env_file)
if os.path.exists(cache_dir):
    shutil.rmtree(cache_dir)
# Create cache directory
os.makedirs(cache_dir, exist_ok=True)

os.makedirs(os.path.dirname(env_file), exist_ok=True)

# Create .env file
with open(env_file, 'w') as f:
    f.write("# Environment variables\n")

print("=== API Key Configuration ===")


account_type = input("Are you using a demo account? (y/n): ").strip().lower()
is_demo = account_type in ['y', 'yes', '1', 'true']
if is_demo == True:
    print("NOTE: Due to limited features in the demo account, Advanced Account Info will not be available")
# Get T212 API key
t212_key = input("Enter your Trading212 API key: ").strip()
if t212_key:
    try:
        set_key(env_file, "T212_API_KEY", t212_key)
        set_key(env_file, "T212_DEMO", str(is_demo))
        print("✓ T212 API key updated")
        print(f"✓ Account type set to: {'Demo' if is_demo else 'Live'}")
    except Exception as e:
        print(f"Error updating T212 API key: {e}")

# Get OpenAI API key
openai_key = input("Enter your OpenAI API key: ").strip()
if openai_key:
    try:
        set_key(env_file, "OPENAI_API_KEY", openai_key)
        print("✓ OpenAI API key updated")
    except Exception as e:
        print(f"Error updating OpenAI API key: {e}")

print(f"Configuration saved to: {env_file}")

from CacheAPIValues import create_cache_data
from sheet_generators.ExcelGenerator import make_xslx

create_cache_data()
make_xslx()