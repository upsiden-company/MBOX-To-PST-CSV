"""Cross-platform script to build a standalone Windows EXE using PyInstaller.

This script builds mbox_converter.py into a single executable file
that can run without Python installed.
"""
import subprocess
import sys


def ensure_pyinstaller() -> None:
    """Ensure PyInstaller is installed."""
    try:
        import PyInstaller  # type: ignore
    except ImportError:
        print("Installing PyInstaller...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", "--upgrade", "pyinstaller"])


def ensure_dependencies() -> None:
    """Ensure all dependencies are installed."""
    print("Installing dependencies...")
    subprocess.check_call([sys.executable, "-m", "pip", "install", "-r", "requirements.txt"])


def build():
    """Build the executable."""
    ensure_dependencies()
    ensure_pyinstaller()
    
    print("Building executable...")
    subprocess.check_call([
        sys.executable, "-m", "PyInstaller",
        "--onefile",
        "--name", "mbox_converter",
        "--add-data", "README.md;.",  # Windows uses semicolon
        "mbox_converter.py"
    ])
    
    print("\n=== Build Complete ===")
    print("Executable created at: dist/mbox_converter.exe")
    print("\nUsage examples:")
    print("  mbox_converter.exe convert inbox.mbox --format csv")
    print("  mbox_converter.exe info inbox.mbox")
    print("  mbox_converter.exe --help")


if __name__ == "__main__":
    build()
