#!/usr/bin/env python3
"""
Simplified setup script for PowerPoint Presentation Generator
Handles dependency installation and provides next steps for configuration.
"""

import os
import subprocess
import sys

def install_requirements():
    """Install required Python packages."""
    print("📦 Installing required packages...")
    try:
        # The -m pip ensures that the correct pip for the active interpreter is used
        subprocess.check_call([sys.executable, "-m", "pip", "install", "-r", "requirements.txt"])
        print("✅ All packages installed successfully!")
        return True
    except subprocess.CalledProcessError:
        print("\n❌ Failed to install packages automatically.")
        print("Please ensure you have pip and Python installed, and then run the following command manually:")
        print("pip install -r requirements.txt")
        return False

def main():
    """Main setup function to guide the user."""
    print("🚀 PowerPoint Presentation Generator - Setup")
    print("=" * 50)
    
    # Step 1: Install dependencies
    install_requirements()
   
if __name__ == "__main__":
    main()
