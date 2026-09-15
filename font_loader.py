# font_loader.py
"""
Load bundled Myanmar Unicode fonts at runtime.
Works on Windows, Linux, and macOS.
"""
import os
import sys
import shutil
import subprocess
import tempfile


MYANMAR_FONT_FAMILY = "Noto Sans Myanmar"


# ============================================================
# RESOURCE PATH (works with PyInstaller)
# ============================================================

def get_resource_path(relative_path):
    """Get absolute path to resource (works for dev and PyInstaller)."""
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)


# ============================================================
# PLATFORM-SPECIFIC FONT INSTALLERS
# ============================================================

def _install_font_linux(font_path):
    """Copy font to user font dir and refresh font cache on Linux."""
    user_font_dir = os.path.expanduser("~/.local/share/fonts")
    os.makedirs(user_font_dir, exist_ok=True)

    dest = os.path.join(user_font_dir, os.path.basename(font_path))

    need_copy = True
    if os.path.exists(dest):
        try:
            if os.path.getsize(dest) == os.path.getsize(font_path):
                need_copy = False
        except Exception:
            need_copy = True

    if need_copy:
        shutil.copy2(font_path, dest)
        print(f"📁 Copied to {dest}")

        try:
            subprocess.run(
                ["fc-cache", "-f", user_font_dir],
                check=False,
                capture_output=True,
                timeout=10
            )
            print(f"🔄 Refreshed font cache")
        except Exception as e:
            print(f"⚠️  fc-cache failed: {e}")


def _install_font_windows(font_path):
    """Register font on Windows using ctypes AddFontResourceEx."""
    import ctypes
    from ctypes import wintypes

    safe_name = os.path.basename(font_path)
    temp_dir = os.path.join(tempfile.gettempdir(), "km_pos_fonts")
    os.makedirs(temp_dir, exist_ok=True)
    safe_path = os.path.join(temp_dir, safe_name)

    if not os.path.exists(safe_path) or \
       os.path.getsize(safe_path) != os.path.getsize(font_path):
        shutil.copy2(font_path, safe_path)

    gdi32 = ctypes.WinDLL("gdi32")
    FR_PRIVATE = 0x10

    result = gdi32.AddFontResourceExW(
        wintypes.LPCWSTR(safe_path),
        FR_PRIVATE,
        0
    )

    if result == 0:
        result = gdi32.AddFontResourceW(wintypes.LPCWSTR(safe_path))

    if result == 0:
        raise RuntimeError(f"AddFontResource failed for {safe_path}")

    HWND_BROADCAST = 0xFFFF
    WM_FONTCHANGE = 0x001D
    user32 = ctypes.WinDLL("user32")
    user32.SendMessageW(HWND_BROADCAST, WM_FONTCHANGE, 0, 0)


def _install_font_mac(font_path):
    """Copy font to ~/Library/Fonts on macOS."""
    user_font_dir = os.path.expanduser("~/Library/Fonts")
    os.makedirs(user_font_dir, exist_ok=True)
    dest = os.path.join(user_font_dir, os.path.basename(font_path))
    if not os.path.exists(dest):
        shutil.copy2(font_path, dest)


def _install_font(font_path):
    """Install font file for current platform."""
    if sys.platform == "win32":
        _install_font_windows(font_path)
    elif sys.platform == "darwin":
        _install_font_mac(font_path)
    else:
        _install_font_linux(font_path)


# ============================================================
# CHECK AVAILABLE MYANMAR FONTS IN TK
# ============================================================

def _list_available_myanmar_fonts():
    """List which font families are available to Tk."""
    try:
        import tkinter as tk
        import tkinter.font as tkfont

        root = tk._default_root
        if root is None:
            root = tk.Tk()
            root.withdraw()
            created = True
        else:
            created = False

        families = set(tkfont.families())

        if created:
            root.destroy()

        return families
    except Exception as e:
        print(f"font list error: {e}")
        return set()


# ============================================================
# MAIN LOADER
# ============================================================

def load_bundled_fonts():
    """
    Load Myanmar fonts bundled with the app.
    Returns the font family name if successful, else fallback.
    """
    font_files = [
        "fonts/NotoSansMyanmar-Regular.ttf",
        "fonts/NotoSansMyanmar-Bold.ttf",
        "fonts/Pyidaungsu-Regular.ttf",
        "fonts/Pyidaungsu-Bold.ttf",
    ]

    for font_rel_path in font_files:
        font_path = get_resource_path(font_rel_path)

        if not os.path.exists(font_path):
            print(f"⚠️  Font not found: {font_path}")
            continue

        # Check file size (avoid 0-byte files)
        try:
            size = os.path.getsize(font_path)
            if size < 1000:
                print(f"⚠️  Font file too small ({size} bytes): {font_path}")
                continue
        except Exception:
            pass

        try:
            _install_font(font_path)
            print(f"✅ Loaded: {os.path.basename(font_path)}")
        except Exception as e:
            print(f"❌ Failed to load {font_path}: {e}")

    # Check available fonts in Tk
    families = _list_available_myanmar_fonts()

    preferred_order = [
        "Noto Sans Myanmar",
        "Noto Sans Myanmar UI",
        "Pyidaungsu",
        "Pyidaungsu Book",
        "Myanmar Text",
        "Myanmar3",
        "Padauk",
    ]

    for name in preferred_order:
        if name in families:
            print(f"🇲🇲 Using font family: {name}")
            return name

    myanmar_like = [f for f in families if any(
        k in f.lower() for k in ["myanmar", "pyidaung", "padauk", "noto"]
    )]
    print(f"⚠️  No preferred font. Myanmar-like: {myanmar_like}")
    return "Arial"
