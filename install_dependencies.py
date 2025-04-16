import subprocess
import sys
import os

def install_dependencies():
    requirements_path = os.path.join(os.path.dirname(__file__), "requirements.txt")
    
    try:
        print("Installing dependencies from requirements.txt...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", "-r", requirements_path])
        print("\nAll dependencies installed successfully.")
    
    except subprocess.CalledProcessError as e:
        print(f"\nFailed to install dependencies. Error: {e}")
        sys.exit(1)
    
    except Exception as e:
        print(f"\nAn unexpected error occurred: {e}")
        sys.exit(1)

if __name__ == "__main__":
    install_dependencies()
    
    if os.name == "nt":
        input("Press Enter to exit...")