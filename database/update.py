import sqlite3


with sqlite3.connect("inventario.db") as conn:
	conn.execute("DELETE FROM inventario_items")
	conn.execute("DELETE FROM sqlite_sequence WHERE name='inventario_items'")

print("Registro actualizado")