from flask import Blueprint, Response, stream_with_context

try:
    from app.utils.sse_hub import stream_generator, subscribe
except ModuleNotFoundError:
    from utils.sse_hub import stream_generator, subscribe


realtime_bp = Blueprint("realtime", __name__)


@realtime_bp.route("/api/events", methods=["GET"])
def api_events_stream():
    subscriber_id, queue = subscribe()

    headers = {
        "Cache-Control": "no-cache",
        "Connection": "keep-alive",
        "X-Accel-Buffering": "no",
    }

    return Response(
        stream_with_context(stream_generator(subscriber_id, queue)),
        mimetype="text/event-stream",
        headers=headers,
    )
