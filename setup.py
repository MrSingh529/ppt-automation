#!/usr/bin/env python3
"""
Setup script to fix Python environment issues and install dependencies correctly
"""

import subprocess
import sys
import os

def run_command(cmd, description):
    """Run a command and handle errors"""
    print(f"Running: {description}")
    try:
        result = subprocess.run(cmd, shell=True, check=True, capture_output=True, text=True)
        print(f"Success: {description}")
        return True
    except subprocess.CalledProcessError as e:
        print(f"Error in {description}: {e}")
        print(f"Output: {e.stdout}")
        print(f"Error: {e.stderr}")
        return False

def main():
    print("Setting up PPT Automation environment...")
    
    # Check Python version
    python_version = sys.version_info
    print(f"Python version: {python_version.major}.{python_version.minor}.{python_version.micro}")
    
    if python_version.major < 3 or (python_version.major == 3 and python_version.minor < 8):
        print("ERROR: Python 3.8 or higher is required")
        return False
    
    # Create virtual environment
    if not os.path.exists('venv'):
        if not run_command(f"{sys.executable} -m venv venv", "Creating virtual environment"):
            return False
    
    # Determine activation script based on OS
    if os.name == 'nt':  # Windows
        activate_script = r"venv\Scripts\activate"
        pip_cmd = r"venv\Scripts\pip"
        python_cmd = r"venv\Scripts\python"
    else:  # Unix/Linux/macOS
        activate_script = "source venv/bin/activate"
        pip_cmd = "venv/bin/pip"
        python_cmd = "venv/bin/python"
    
    print(f"Virtual environment created. Activation command: {activate_script}")
    
    # Upgrade pip first
    if not run_command(f"{pip_cmd} install --upgrade pip", "Upgrading pip"):
        return False
    
    # Install requirements
    if os.path.exists('requirements.txt'):
        if not run_command(f"{pip_cmd} install -r requirements.txt", "Installing requirements"):
            return False
    else:
        print("requirements.txt not found, installing individual packages...")
        packages = [
            "Flask==2.3.3",
            "gunicorn==21.2.0",
            "openpyxl==3.1.2",
            "python-pptx==0.6.21",
            "requests==2.31.0",
            "anthropic==0.7.8",
            "Werkzeug==2.3.7",
            "python-dotenv==1.0.0"
        ]
        
        for package in packages:
            if not run_command(f"{pip_cmd} install {package}", f"Installing {package}"):
                print(f"Warning: Failed to install {package}")
    
    # Create .env file if it doesn't exist
    if not os.path.exists('.env'):
        with open('.env', 'w') as f:
            f.write("# Environment variables for PPT Automation\n")
            f.write("ANTHROPIC_API_KEY=your_api_key_here\n")
            f.write("FLASK_SECRET_KEY=your_secret_key_here\n")
        print("Created .env file - please update with your actual API keys")
    
    # Create necessary directories
    os.makedirs('templates', exist_ok=True)
    os.makedirs('uploads', exist_ok=True)
    os.makedirs('outputs', exist_ok=True)
    
    print("\nSetup complete!")
    print("\nNext steps:")
    print(f"1. Activate virtual environment: {activate_script}")
    print("2. Update .env file with your API keys")
    print("3. Ensure you have these files:")
    print("   - app.py (Flask application)")
    print("   - main_script.py (your PowerPoint automation script)")
    print("   - templates/index.html (web interface)")
    print(f"4. Run the application: {python_cmd} app.py")
    
    return True

if __name__ == "__main__":
    success = main()
    if not success:
        sys.exit(1)