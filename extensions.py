import os
from flask_wtf.csrf import CSRFProtect
from flask_limiter import Limiter
from flask_limiter.util import get_remote_address

csrf = CSRFProtect()
_limiter_storage = os.environ.get("REDIS_URL", "memory://")
limiter = Limiter(get_remote_address, default_limits=[], storage_uri=_limiter_storage)
