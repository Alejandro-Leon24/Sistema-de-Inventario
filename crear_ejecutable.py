import os
import subprocess
import sys
import shutil

def crear_app():
    print("--- Iniciando proceso de empaquetado del Sistema de Inventario ---")
    
    # 1. Verificar dependencias necesarias
    try:
        import PyInstaller
    except ImportError:
        print("Error: PyInstaller no está instalado. Instalando...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", "pyinstaller"])

    # 2. Definir carpetas de datos (static, templates, database, plantillas)
    # Formato para PyInstaller: "origen;destino" (en Windows usa ;)
    data_folders = [
        ("app/static", "app/static"),
        ("app/templates", "app/templates"),
        ("database", "database"),
        ("plantillas", "plantillas"),
    ]

    # Construir el comando de PyInstaller
    cmd = [
        "pyinstaller",
        "--noconfirm",
        "--onedir", # Recomendado para apps portables: más rápido al arrancar que --onefile
        "--windowed", # No abre consola de comandos al iniciar
        "--name", "SistemaInventario",
        "--icon", "app/static/img/favicon.ico" if os.path.exists("app/static/img/favicon.ico") else "NONE",
        "--clean",
    ]

    # Añadir las carpetas de datos al comando
    for src, dest in data_folders:
        if os.path.exists(src):
            cmd.extend(["--add-data", f"{src}{os.pathsep}{dest}"])
        else:
            print(f"Advertencia: No se encontró la carpeta {src}, se omitirá.")

    # El script principal de entrada
    cmd.append("app/app.py")

    # 3. Ejecutar PyInstaller
    print(f"Ejecutando: {' '.join(cmd)}")
    try:
        subprocess.check_call(cmd)
        print("\n--- ¡ÉXITO! ---")
        print("La aplicación ha sido creada en la carpeta: dist/SistemaInventario")
        print("Para usarla, solo copia la carpeta 'SistemaInventario' a cualquier lugar (USB, Disco, etc.) y ejecuta 'SistemaInventario.exe'.")
    except subprocess.CalledProcessError:
        print("\n--- ERROR ---")
        print("Hubo un error durante el proceso de empaquetado.")

if __name__ == "__main__":
    crear_app()
