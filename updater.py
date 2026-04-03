"""
GST Suite Updater
=================
Standalone script compiled to updater.exe by PyInstaller.
Called by GST_Suite.exe when an update is available:

    updater.exe --url <download_url> --target <old_exe_path> --restart

Flow:
  1. Download new EXE to a temp folder
  2. Wait for GST_Suite.exe process to fully exit
  3. Replace old EXE with the downloaded one
  4. Optionally restart the new EXE
"""

import sys
import os
import time
import shutil
import tempfile
import argparse
import subprocess
import urllib.request


def _download(url: str, dest: str) -> None:
    """Download url to dest with a simple progress bar."""
    def _hook(count, block, total):
        if total > 0:
            pct = min(int(count * block * 100 / total), 100)
            filled = pct // 2
            bar = "█" * filled + "░" * (50 - filled)
            print(f"\r  [{bar}] {pct:3d}%", end="", flush=True)

    urllib.request.urlretrieve(url, dest, _hook)
    print("\n\n  Download complete.\n")


def _wait_for_exit(exe_name: str, timeout_sec: int = 60) -> bool:
    """Return True when the process is no longer running (or timeout hit)."""
    print(f"  Waiting for application to close ...", end="", flush=True)
    for _ in range(timeout_sec * 2):
        result = subprocess.run(
            ["tasklist", "/FI", f"IMAGENAME eq {exe_name}", "/NH", "/FO", "CSV"],
            capture_output=True, text=True
        )
        if exe_name.lower() not in result.stdout.lower():
            print("  closed.\n")
            return True
        print(".", end="", flush=True)
        time.sleep(0.5)
    print("  timed out.\n")
    return False


def _replace(src: str, target: str) -> None:
    """Move src to target, retrying once on PermissionError."""
    try:
        shutil.move(src, target)
    except PermissionError:
        print("  File still locked, retrying in 3 s …")
        time.sleep(3)
        shutil.move(src, target)


def main() -> None:
    parser = argparse.ArgumentParser(description="GST Suite Updater")
    parser.add_argument("--url",     required=True, help="URL to download the new EXE from")
    parser.add_argument("--target",  required=True, help="Full path of the EXE to replace")
    parser.add_argument("--restart", action="store_true", help="Launch the new EXE after update")
    args = parser.parse_args()

    print()
    print("   ██████╗████████╗██╗   ██╗██████╗ ██╗   ██╗ ██████╗ █████╗ ███████╗███████╗")
    print("  ██╔════╝╚══██╔══╝██║   ██║██╔══██╗╚██╗ ██╔╝██╔════╝██╔══██╗██╔════╝██╔════╝")
    print("   ███████╗  ██║   ██║   ██║██║  ██║ ╚████╔╝ ██║     ███████║█████╗  █████╗  ")
    print("  ╚════██║   ██║   ██║   ██║██║  ██║  ╚██╔╝  ██║     ██╔══██║██╔══╝  ██╔══╝  ")
    print("   ███████║  ██║   ╚██████╔╝██████╔╝   ██║   ╚██████╗██║  ██║██║     ███████╗")
    print("   ╚══════╝  ╚═╝    ╚═════╝ ╚═════╝    ╚═╝    ╚═════╝╚═╝  ╚═╝╚═╝     ╚══════╝")
    print()
    print("  " + "─" * 48)
    print("         Updating to the latest version ...")
    print("  " + "─" * 48)
    print()

    # ── 1. Download ───────────────────────────────────────────────
    tmp_dir = tempfile.mkdtemp(prefix="gst_update_")
    tmp_exe = os.path.join(tmp_dir, "GST_Suite_new.exe")

    try:
        _download(args.url, tmp_exe)
    except Exception as exc:
        print(f"\n  ERROR: Download failed — {exc}")
        input("\n  Press Enter to exit …")
        sys.exit(1)

    # ── 2. Wait for the old app to exit ───────────────────────────
    exe_name = os.path.basename(args.target)
    _wait_for_exit(exe_name, timeout_sec=30)

    # ── 3. Replace ────────────────────────────────────────────────
    print(f"  Installing update ...")
    try:
        _replace(tmp_exe, args.target)
        print("  Update installed successfully!")
        print()
    except Exception as exc:
        print(f"\n  ERROR: Could not replace file — {exc}")
        print(f"  The downloaded file is at:\n  {tmp_exe}")
        input("\n  Press Enter to exit …")
        sys.exit(1)

    # ── 4. Restart ────────────────────────────────────────────────
    if args.restart:
        print("  Restarting application ...")
        time.sleep(1)
        subprocess.Popen([args.target])

    print("  Done. This window will close in 3 seconds.")
    time.sleep(3)


if __name__ == "__main__":
    main()
