import os
import sys
import stat
import shutil
import platform
import subprocess
import urllib.request
from pathlib import Path


BASE_DIR = Path(__file__).resolve().parent
LIBREOFFICE_DIR = BASE_DIR / ".libreoffice"
APPIMAGE_PATH = LIBREOFFICE_DIR / "LibreOffice.AppImage"

# LibreOffice 26.2.5 - Linux x86_64 AppImage
LIBREOFFICE_URL = (
    "https://download.documentfoundation.org/"
    "libreoffice/stable/26.2.5/appimage/x86_64/"
    "LibreOffice-fresh-x86_64.AppImage"
)


def find_soffice():
    """Search for an already installed LibreOffice."""

    possible_paths = [
        shutil.which("soffice"),
        shutil.which("libreoffice"),
        "/usr/bin/soffice",
        "/usr/bin/libreoffice",
        "/usr/local/bin/soffice",
        "/usr/local/bin/libreoffice",
    ]

    for path in possible_paths:
        if path and Path(path).exists():
            return str(Path(path).resolve())

    return None


def download_libreoffice():
    """Download LibreOffice AppImage only when it is not already present."""

    LIBREOFFICE_DIR.mkdir(parents=True, exist_ok=True)

    if APPIMAGE_PATH.exists() and APPIMAGE_PATH.stat().st_size > 50_000_000:
        print("LibreOffice AppImage already exists.")
        return APPIMAGE_PATH

    print("=" * 60)
    print("Downloading LibreOffice...")
    print("Version: 26.2.5")
    print("=" * 60)

    temporary_file = APPIMAGE_PATH.with_suffix(".download")

    try:
        request = urllib.request.Request(
            LIBREOFFICE_URL,
            headers={
                "User-Agent": "Bewerber-Flask/1.0"
            }
        )

        with urllib.request.urlopen(request, timeout=120) as response:
            total = response.headers.get("Content-Length")
            total = int(total) if total else None

            downloaded = 0

            with open(temporary_file, "wb") as output:
                while True:
                    chunk = response.read(1024 * 1024)

                    if not chunk:
                        break

                    output.write(chunk)
                    downloaded += len(chunk)

                    if total:
                        percent = downloaded * 100 / total
                        print(
                            f"\rDownloading: {percent:.1f}%",
                            end="",
                            flush=True
                        )
                    else:
                        print(
                            f"\rDownloaded: {downloaded / 1024 / 1024:.1f} MB",
                            end="",
                            flush=True
                        )

        print()

        temporary_file.replace(APPIMAGE_PATH)

        # chmod +x
        current_mode = APPIMAGE_PATH.stat().st_mode
        APPIMAGE_PATH.chmod(
            current_mode
            | stat.S_IXUSR
            | stat.S_IXGRP
            | stat.S_IXOTH
        )

        print("LibreOffice downloaded successfully.")
        return APPIMAGE_PATH

    except Exception:
        if temporary_file.exists():
            temporary_file.unlink()

        raise


def extract_appimage():
    """
    AppImage normally runs directly.
    If FUSE is unavailable, extract it and use the extracted AppRun.
    """

    extracted_dir = LIBREOFFICE_DIR / "squashfs-root"
    apprun = extracted_dir / "AppRun"

    if apprun.exists():
        return apprun

    print("Trying to extract LibreOffice AppImage...")

    result = subprocess.run(
        [
            str(APPIMAGE_PATH),
            "--appimage-extract"
        ],
        cwd=LIBREOFFICE_DIR,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        text=True,
        timeout=180
    )

    if result.returncode != 0:
        print("AppImage extraction failed.")
        print(result.stdout)
        print(result.stderr)
        return None

    if apprun.exists():
        return apprun

    return None


def get_libreoffice():
    """
    Return the executable that should be used by Flask.
    """

    # 1. Existing system LibreOffice
    system_soffice = find_soffice()

    if system_soffice:
        print("LibreOffice found:")
        print(system_soffice)
        return system_soffice

    # 2. Windows
    if platform.system().lower() == "windows":
        possible_windows = [
            Path(os.environ.get("PROGRAMFILES", "C:/Program Files"))
            / "LibreOffice/program/soffice.exe",

            Path(os.environ.get("PROGRAMFILES(X86)", "C:/Program Files (x86)"))
            / "LibreOffice/program/soffice.exe",
        ]

        for path in possible_windows:
            if path.exists():
                return str(path)

        raise RuntimeError(
            "LibreOffice is not installed on Windows."
        )

    # 3. Linux AppImage
    if platform.system().lower() == "linux":

        appimage = download_libreoffice()

        # First try direct AppImage execution.
        try:
            appimage.chmod(
                appimage.stat().st_mode
                | stat.S_IXUSR
                | stat.S_IXGRP
                | stat.S_IXOTH
            )

            test = subprocess.run(
                [
                    str(appimage),
                    "--headless",
                    "--version"
                ],
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                text=True,
                timeout=30
            )

            if test.returncode == 0:
                print("LibreOffice AppImage works.")
                print(test.stdout.strip())
                return str(appimage)

            print("Direct AppImage execution failed:")
            print(test.stderr)

        except Exception as e:
            print("Direct AppImage execution error:")
            print(e)

        # 4. Try AppImage extraction when FUSE is unavailable.
        extracted = extract_appimage()

        if extracted and extracted.exists():
            print("Using extracted LibreOffice AppImage.")
            return str(extracted)

        raise RuntimeError(
            "LibreOffice AppImage was downloaded, "
            "but could not be executed."
        )

    raise RuntimeError(
        f"Unsupported operating system: {platform.system()}"
    )


if __name__ == "__main__":
    try:
        path = get_libreoffice()

        print()
        print("=" * 60)
        print("LibreOffice READY")
        print(path)
        print("=" * 60)

    except Exception as e:
        print()
        print("=" * 60)
        print("LibreOffice SETUP FAILED")
        print(str(e))
        print("=" * 60)

        sys.exit(1)