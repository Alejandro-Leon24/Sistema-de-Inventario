import sqlite3
conn = sqlite3.connect('inventario.db')
conn.execute('UPDATE administradores SET activo = 1 WHERE id = 1')
conn.commit()
conn.close()
print('✓ Registro actualizado')