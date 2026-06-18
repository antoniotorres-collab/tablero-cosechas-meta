"""
Sincroniza archivos xlsx 2026 desde Google Drive hacia el repo local,
luego hace commit y push a GitHub.

Solo lectura de Drive — nunca modifica los originales.
Solo copia archivos que aún no existen en datos/.
"""

import re
import shutil
import subprocess
import sys
from pathlib import Path

# ── rutas ─────────────────────────────────────────────────────────────────────

DRIVE_DIR = Path(
    "/Users/antoniotorres/Library/CloudStorage"
    "/GoogleDrive-antonio.torres@metafinanciera.com"
    "/Mi unidad/META/CosechasSemanal"
)
REPO_DIR  = Path("/Users/antoniotorres/Desktop/Cosechas_Claude")
DATOS_DIR = REPO_DIR / "datos"

# Patrón: Cosechas<día><mes>26[CH|MC].xlsx
PATRON_2026 = re.compile(r"^Cosechas\d{1,2}[A-Za-z]{2,4}26(?:CH|MC)?\.xlsx$")

# ── colores para terminal ──────────────────────────────────────────────────────

VERDE  = "\033[92m"
ROJO   = "\033[91m"
AMARILLO = "\033[93m"
RESET  = "\033[0m"
NEGRITA = "\033[1m"

def ok(msg):  print(f"{VERDE}✓{RESET} {msg}")
def err(msg): print(f"{ROJO}✗{RESET} {msg}")
def info(msg): print(f"{AMARILLO}→{RESET} {msg}")

# ── lógica principal ───────────────────────────────────────────────────────────

def buscar_archivos_drive():
    """Busca recursivamente todos los xlsx 2026 en DRIVE_DIR."""
    archivos = []
    for f in DRIVE_DIR.rglob("*.xlsx"):
        if f.name.startswith("~$"):
            continue
        if PATRON_2026.match(f.name):
            archivos.append(f)
    return sorted(archivos, key=lambda f: f.name)

def base_fecha(nombre):
    """'Cosechas12Ene26CH.xlsx' → 'Cosechas12Ene26'"""
    return re.sub(r"(CH|MC)?\.xlsx$", "", nombre)

def archivos_en_datos():
    return {f.name for f in DATOS_DIR.glob("*.xlsx") if not f.name.startswith("~$")}

def tiene_ch_en_datos(nombre, en_datos):
    """True si ya existe la versión CH de la misma fecha en datos/."""
    base = base_fecha(nombre)
    return f"{base}CH.xlsx" in en_datos

def git(*args):
    result = subprocess.run(
        ["git", "-C", str(REPO_DIR), *args],
        capture_output=True, text=True
    )
    if result.returncode != 0:
        raise RuntimeError(result.stderr.strip())
    return result.stdout.strip()

def sincronizar(dry_run=False):
    print(f"\n{NEGRITA}Sincronizando datos Drive → GitHub{RESET}")
    print(f"  Origen : {DRIVE_DIR}")
    print(f"  Destino: {DATOS_DIR}\n")

    if not DRIVE_DIR.exists():
        err("No se encontró la carpeta de Drive. ¿Está montada?")
        sys.exit(1)

    en_drive   = buscar_archivos_drive()
    en_datos   = archivos_en_datos()

    nuevos    = []
    omitidos  = []
    ya_existen = []

    for f in en_drive:
        if f.name in en_datos:
            ya_existen.append(f)
        elif tiene_ch_en_datos(f.name, en_datos):
            omitidos.append(f)   # plain ignorado porque ya hay CH
        else:
            nuevos.append(f)

    print(f"  En Drive (2026)  : {len(en_drive)} archivos")
    print(f"  Ya en datos/     : {len(ya_existen)} archivos")
    print(f"  Omitidos (hay CH): {len(omitidos)} archivos")
    print(f"  Nuevos a copiar  : {NEGRITA}{len(nuevos)}{RESET}\n")

    if not nuevos:
        ok("Todo actualizado — no hay archivos nuevos.")
        return

    for f in nuevos:
        info(f"Copiando {f.name}")
        if not dry_run:
            shutil.copy2(f, DATOS_DIR / f.name)
            ok(f"{f.name}")

    if dry_run:
        print(f"\n{AMARILLO}[dry-run] No se copió ni confirmó nada.{RESET}")
        return

    # ── git add + commit + push ────────────────────────────────────────────────
    print()
    nombres = ", ".join(f.name for f in nuevos)
    mensaje = f"datos: agrega {len(nuevos)} archivo(s) semana ({nombres})"

    try:
        git("add", "datos/")
        git("commit", "-m", mensaje)
        ok(f"Commit: {mensaje}")
        git("push")
        ok("Push a GitHub completado.")
    except RuntimeError as e:
        err(f"Error en git: {e}")
        sys.exit(1)

    print(f"\n{VERDE}{NEGRITA}Sincronización completa.{RESET}")

# ── entry point ────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    dry_run = "--dry-run" in sys.argv
    if dry_run:
        print(f"{AMARILLO}[modo dry-run — solo muestra, no copia ni hace push]{RESET}")
    sincronizar(dry_run=dry_run)
