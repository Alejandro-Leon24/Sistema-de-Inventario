import argparse
import os
import sqlite3
from pathlib import Path
from datetime import datetime


PROJECT_ROOT = Path(__file__).resolve().parent.parent
REQUIRED_ENV_FLAG = "INVENTARIO_ALLOW_DESTRUCTIVE"
CONFIRMATION_PHRASE = "DELETE-ALL-INVENTORY"


def _get_database_path() -> Path:
	return PROJECT_ROOT / "inventario.db"


def _create_backup(database_path: Path) -> Path:
	backup_dir = PROJECT_ROOT / "database" / "backups"
	backup_dir.mkdir(parents=True, exist_ok=True)
	stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
	backup_path = backup_dir / f"inventario_pre_reset_{stamp}.db"

	with sqlite3.connect(database_path) as source_conn, sqlite3.connect(backup_path) as backup_conn:
		source_conn.backup(backup_conn)

	return backup_path


def main() -> int:
	parser = argparse.ArgumentParser(
		description="Limpia inventario_items de forma destructiva (solo para entornos controlados)."
	)
	parser.add_argument(
		"--confirm-reset",
		action="store_true",
		help="Habilita modo destructivo (requiere además --confirmation-code).",
	)
	parser.add_argument(
		"--confirmation-code",
		type=str,
		default="",
		help=f"Codigo de seguridad requerido: {CONFIRMATION_PHRASE}",
	)
	parser.add_argument(
		"--yes",
		action="store_true",
		help="Omite confirmación interactiva por consola.",
	)
	parser.add_argument(
		"--skip-backup",
		action="store_true",
		help="No recomendado: omite backup previo de la base.",
	)
	args = parser.parse_args()

	if not args.confirm_reset:
		print("Operacion cancelada. Use --confirm-reset para continuar.")
		return 1

	if args.confirmation_code != CONFIRMATION_PHRASE:
		print(f"Operacion cancelada. Debe usar --confirmation-code {CONFIRMATION_PHRASE}")
		return 1

	if os.environ.get(REQUIRED_ENV_FLAG) != "1":
		print(
			"Operacion cancelada. Defina la variable de entorno "
			f"{REQUIRED_ENV_FLAG}=1 para habilitar acciones destructivas."
		)
		return 1

	database_path = _get_database_path()
	if not database_path.exists():
		print(f"Operacion cancelada. No existe la base de datos: {database_path}")
		return 1

	if not args.yes:
		confirm = input("Esta acción elimina TODO el inventario. Escriba RESET para confirmar: ").strip()
		if confirm != "RESET":
			print("Operacion cancelada por el usuario.")
			return 1

	backup_path = None
	if not args.skip_backup:
		backup_path = _create_backup(database_path)
		print(f"Backup creado en: {backup_path}")

	with sqlite3.connect(database_path) as conn:
		total = conn.execute("SELECT COUNT(1) FROM inventario_items").fetchone()[0]
		conn.execute("DELETE FROM inventario_items")
		conn.execute("DELETE FROM sqlite_sequence WHERE name = 'inventario_items'")
		conn.commit()

	print(
		f"Inventario limpiado: {total} registros eliminados en {database_path}. "
		f"Backup: {backup_path if backup_path else 'omitido por --skip-backup'}."
	)
	return 0


if __name__ == "__main__":
	raise SystemExit(main())