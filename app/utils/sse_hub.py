import json
import threading
import uuid
from queue import Empty, Queue

_SUBSCRIBERS = {}
_LOCK = threading.Lock()


def _format_sse(event_name, data):
    payload = json.dumps(data or {}, ensure_ascii=False)
    return f"event: {event_name}\ndata: {payload}\n\n"


def subscribe(max_queue_size=128):
    subscriber_id = str(uuid.uuid4())
    queue = Queue(maxsize=max_queue_size)
    with _LOCK:
        _SUBSCRIBERS[subscriber_id] = queue
    return subscriber_id, queue


def unsubscribe(subscriber_id):
    with _LOCK:
        _SUBSCRIBERS.pop(subscriber_id, None)


def publish_event(event_name, data=None):
    with _LOCK:
        subscribers = list(_SUBSCRIBERS.items())

    for _, queue in subscribers:
        try:
            queue.put_nowait((event_name, data or {}))
        except Exception:
            try:
                queue.get_nowait()
                queue.put_nowait((event_name, data or {}))
            except Exception:
                continue


def stream_generator(subscriber_id, queue, heartbeat_seconds=20):
    # Instruye al navegador para reintentar la conexion SSE si se corta.
    yield "retry: 5000\n\n"
    yield _format_sse("connected", {"subscriber_id": subscriber_id})

    try:
        while True:
            try:
                event_name, data = queue.get(timeout=heartbeat_seconds)
                yield _format_sse(event_name, data)
            except Empty:
                yield ": heartbeat\n\n"
    finally:
        unsubscribe(subscriber_id)
