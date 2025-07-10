"""Cross-platform script to build a standalone Windows EXE using PyInstaller."""
import subprocess
import sys


def ensure_pyinstaller() -> None:
    """Ensure PyInstaller is installed."""
    try:
        import PyInstaller  # type: ignore
    except ImportError:
        subprocess.check_call([sys.executable, "-m", "pip", "install", "--upgrade", "pyinstaller"])  # noqa: E501


def build():
    ensure_pyinstaller()
    subprocess.check_call([sys.executable, "-m", "PyInstaller", "--onefile", "mbox_converter.py"])


if __name__ == "__main__":
    build()
