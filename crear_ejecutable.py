import os
import subprocess
import sys
import shutil
from pathlib import Path

def limpiar_carpetas():
    """Borra carpetas de compilaciones anteriores para evitar conflictos."""
    for folder in ['build', 'dist']:
        if os.path.exists(folder):
            print(f"Limpiando carpeta {folder}...")
            shutil.rmtree(folder, ignore_errors=True)

def crear_app():
    print("\n--- Iniciando proceso de empaquetado del Sistema de Inventario ---")
    
    limpiar_carpetas()

    # 1. Verificar PyInstaller
    try:
        import PyInstaller
    except ImportError:
        print("Instalando PyInstaller...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", "pyinstaller"])

    # 2. Carpetas que van DENTRO del ejecutable (Solo lectura)
    # Formato: (origen, destino)
    data_folders = [
        ("app/static", "app/static"),
        ("app/templates", "app/templates"),
        ("database/schema.sql", "database"), # Necesario para inicializar DBs nuevas
    ]

    cmd = [
        "pyinstaller",
        "--noconfirm",
        "--onedir",
        "--windowed",
        "--name", "SistemaInventario",
        "--clean",
    ]

    for src, dest in data_folders:
        if os.path.exists(src):
            cmd.extend(["--add-data", f"{src}{os.pathsep}{dest}"])

    cmd.append("app/app.py")

    # 3. Ejecutar PyInstaller
    print(f"\nEmpaquetando aplicación (esto puede tardar unos minutos)...")
    try:
        subprocess.check_call(cmd)
        
        # 4. Preparar carpeta de salida final
        dist_path = Path("dist/SistemaInventario")
        
        # Crear carpetas necesarias fuera del exe para persistencia
        print("\nConfigurando carpetas de persistencia...")
        (dist_path / "plantillas").mkdir(exist_ok=True)
        (dist_path / "salidas").mkdir(exist_ok=True)
        
        # Copiar plantillas base si existen
        if os.path.exists("plantillas"):
            print("Copiando plantillas base...")
            for item in os.listdir("plantillas"):
                s = os.path.join("plantillas", item)
                d = os.path.join(dist_path / "plantillas", item)
                if os.path.isfile(s):
                    shutil.copy2(s, d)
                elif os.path.isdir(s) and item != "_historial":
                    shutil.copytree(s, d, dirs_exist_ok=True)

        print("\n" + "="*50)
        print("¡PROCESO FINALIZADO CON ÉXITO!")
        print("="*50)
        print(f"\n1. Ve a la carpeta: {dist_path.absolute()}")
        print("2. Ejecuta 'SistemaInventario.exe'")
        print("\nNOTA: Para mover el programa, debes copiar TODA la carpeta 'SistemaInventario',")
        print("no solo el archivo .exe.")
        print("="*50)

    except subprocess.CalledProcessError:
        print("\n--- ERROR: No se pudo crear el ejecutable ---")

if __name__ == "__main__":
    crear_app()
