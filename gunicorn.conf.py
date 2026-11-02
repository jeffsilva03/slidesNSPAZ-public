import os


bind = f"0.0.0.0:{os.environ.get('PORT', '10000')}"
workers = int(os.environ.get('WEB_CONCURRENCY', '1'))
timeout = int(os.environ.get('WEB_TIMEOUT', '300'))
graceful_timeout = int(os.environ.get('WEB_GRACEFUL_TIMEOUT', '60'))
