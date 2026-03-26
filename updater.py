"""
updater.py
Auto-updater via GitHub Releases.
Chequea en background si hay una versión nueva y notifica a la app.
"""

import os
import sys
import threading
import subprocess
import tempfile
import urllib.request
import json

from version import VERSION

GITHUB_API = "https://api.github.com/repos/lev1ll/APP-01---LAZYTEST/releases/latest"


def _parse_version(v: str) -> tuple:
    """Convierte '1.2.3' → (1, 2, 3) para comparar."""
    try:
        return tuple(int(x) for x in v.lstrip("v").split("."))
    except Exception:
        return (0,)


def chequear_actualizacion(callback):
    """
    Chequea GitHub en un hilo separado.
    Llama callback(version, url) si hay versión nueva, o callback(None, None) si no.
    """
    def _run():
        try:
            req = urllib.request.Request(
                GITHUB_API,
                headers={"User-Agent": "GeneradorAMR-Updater"}
            )
            with urllib.request.urlopen(req, timeout=8) as resp:
                data = json.loads(resp.read().decode())

            latest = data.get("tag_name", "")
            if _parse_version(latest) > _parse_version(VERSION):
                # Buscar el .exe en los assets
                url = None
                for asset in data.get("assets", []):
                    if asset["name"].endswith(".exe"):
                        url = asset["browser_download_url"]
                        break
                if url:
                    callback(latest, url)
                    return
        except Exception:
            pass
        callback(None, None)

    threading.Thread(target=_run, daemon=True).start()


def descargar_e_instalar(url: str, on_progress=None, on_done=None, on_error=None):
    """
    Descarga el nuevo .exe y lo instala reemplazando el actual.
    Usa un .bat para hacer el reemplazo después de que la app cierre.
    """
    def _run():
        try:
            # Ruta del exe actual
            if getattr(sys, "frozen", False):
                exe_actual = sys.executable
            else:
                # En modo desarrollo, simular la ruta
                exe_actual = os.path.join(
                    os.path.dirname(os.path.abspath(__file__)),
                    "GeneradorAMR.exe"
                )

            # Descargar a archivo temporal
            tmp = tempfile.NamedTemporaryFile(
                delete=False, suffix=".exe",
                dir=os.path.dirname(exe_actual)
            )
            tmp_path = tmp.name
            tmp.close()

            req = urllib.request.Request(url, headers={"User-Agent": "GeneradorAMR-Updater"})
            with urllib.request.urlopen(req, timeout=60) as resp:
                total = int(resp.headers.get("Content-Length", 0))
                descargado = 0
                with open(tmp_path, "wb") as f:
                    while True:
                        chunk = resp.read(8192)
                        if not chunk:
                            break
                        f.write(chunk)
                        descargado += len(chunk)
                        if on_progress and total:
                            on_progress(int(descargado / total * 100))

            # Crear script .bat que reemplaza el exe y relanza
            bat_path = tmp_path + "_update.bat"
            bat_content = f"""@echo off
ping 127.0.0.1 -n 3 > nul
move /Y "{tmp_path}" "{exe_actual}"
start "" "{exe_actual}"
del "%~f0"
"""
            with open(bat_path, "w") as f:
                f.write(bat_content)

            if on_done:
                on_done(bat_path, exe_actual)

        except Exception as e:
            if on_error:
                on_error(str(e))

    threading.Thread(target=_run, daemon=True).start()
