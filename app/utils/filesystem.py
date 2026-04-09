import os


def cleanup_empty_parent_dirs(path, stop_at):
    """Elimina directorios vacios ascendiendo desde path hasta stop_at (sin borrar stop_at)."""
    current = os.path.abspath(os.path.dirname(path or ""))
    stop_dir = os.path.abspath(stop_at)
    while current.startswith(stop_dir) and current != stop_dir:
        try:
            if os.path.isdir(current) and not os.listdir(current):
                os.rmdir(current)
                current = os.path.abspath(os.path.dirname(current))
                continue
        except Exception:
            pass
        break