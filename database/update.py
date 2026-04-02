import argparse
import sqlite3
from pathlib import Path


def _get_database_path() -> Path:
	return Path(__file__).resolve().parent.parent / "inventario.db"


def main() -> int:
	parser = argparse.ArgumentParser(
		description="Limpia inventario_items de forma destructiva (requiere confirmacion explicita)."
	)
	parser.add_argument(
		"--confirm-reset",
		action="store_true",
		help="Confirma la eliminacion total de inventario_items.",
	)
	args = parser.parse_args()

	if not args.confirm_reset:
		print("Operacion cancelada. Use --confirm-reset para ejecutar el borrado.")
		return 1

	database_path = _get_database_path()
	with sqlite3.connect(database_path) as conn:
		total = conn.execute("SELECT COUNT(1) FROM inventario_items").fetchone()[0]
		conn.execute("DELETE FROM inventario_items")
		conn.execute("DELETE FROM sqlite_sequence WHERE name = 'inventario_items'")
		conn.commit()

	print(f"Inventario limpiado: {total} registros eliminados en {database_path}.")
	return 0


if __name__ == "__main__":
	raise SystemExit(main())