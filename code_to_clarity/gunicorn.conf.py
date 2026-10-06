# Auto-loaded by gunicorn from the working directory.
# Two sequential LLM calls take minutes; the default 30s timeout kills the worker.
bind = "0.0.0.0:" + __import__("os").environ.get("PORT", "10000")
workers = 1
threads = 4
worker_class = "gthread"
timeout = 900
graceful_timeout = 30
