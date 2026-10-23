"""Logging and Application Insights setup shared by app.py and offline_sync.py.

Call setup() once, as the very first thing a process does. In app.py it must
run before Flask is imported: Azure Monitor instruments Flask by replacing the
Flask class, so an app built from a class imported earlier sends no request
telemetry.

Conventions (please keep to these as the app grows)
---------------------------------------------------
* Never print(). Use a named logger: logging.getLogger("scandroid.<area>").
  The area becomes the "log_area" custom dimension in App Insights.
* Levels mean something:
    DEBUG    per-item detail (each photo, each row). Off unless LOG_LEVEL=DEBUG.
    INFO     one line per meaningful step or outcome.
    WARNING  degraded but carried on (fallback used, lookup failed, row skipped).
    ERROR    the user's action did not do what they asked.
    CRITICAL / logger.exception   crashes.
* Attach identifiers as dimensions, not only in the message text, so they can
  be filtered and grouped in App Insights:
      log.info("...", extra=telemetry.dims(program_id=pid, payment_id=x))
  Requests made by a logged-in user get route, role, actor and program_id
  automatically (see set_context_provider).
* Never log beneficiary data: match values, registration attributes, photo
  URLs or raw 121/Kobo response bodies. Pass any external text through
  scrub() first. Staff usernames are fine - they are the audit trail.
* Lines that exist only to drive the sync progress bar go to stdout, not to
  App Insights: extra=telemetry.dims(stdout_only=True).
"""

import logging
import os
import re
import sys

DEFAULT_FORMAT = "%(asctime)s %(levelname)s [%(name)s] %(message)s"

# Libraries that log every HTTP call or every telemetry export at INFO. Left at
# the root level they would flood the container log, and the Azure SDK's own
# lines would be exported back into App Insights.
_NOISY_LOGGERS = ("azure", "opentelemetry", "urllib3", "werkzeug")

# LogRecord attribute names. A dimension with one of these names would either
# be rejected by logging ("Attempt to overwrite ...") or silently clobber a
# built-in field (e.g. "process" is the OS process id), so dims() renames them.
_RESERVED = set(vars(logging.LogRecord("", 0, "", 0, "", (), None))) | {
    "message",
    "asctime",
}

_state = {
    "configured": False,
    "static": {},
    "provider": None,
}


# ---------------------------------------------------------------------------
# Public helpers
# ---------------------------------------------------------------------------
def dims(**values):
    """Build an `extra=` dict of custom dimensions.

    None and empty values are dropped (OpenTelemetry rejects None attributes),
    everything else is stringified, except the two routing flags, which stay
    booleans so the handler filters can read them.
    """
    out = {}
    for key, value in values.items():
        if value is None or value == "":
            continue
        if key in ("stdout_only", "telemetry_only"):
            out[key] = bool(value)
            continue
        if key in _RESERVED:
            key = "dim_" + key
        out[key] = str(value)
    return out


_EMAIL_RE = re.compile(r"[\w.+-]+@[\w-]+\.[\w.-]+")
_LONG_DIGITS_RE = re.compile(r"\d{6,}")


def scrub(text, limit=300):
    """Make external text (an API error body, an exception message) safe to log.

    Masks email addresses and runs of 6+ digits (phone, card and account
    numbers are the usual match values), flattens newlines and truncates.
    """
    s = "" if text is None else str(text)
    s = _EMAIL_RE.sub("<email>", s)
    s = _LONG_DIGITS_RE.sub(lambda m: "<%d digits>" % len(m.group()), s)
    s = s.replace("\r", " ").replace("\n", " ")
    return s if len(s) <= limit else s[:limit] + "…"


def describe(exc):
    """'ExceptionType: scrubbed message' for log lines."""
    return "%s: %s" % (type(exc).__name__, scrub(exc, 500))


def _is_network_error(exc):
    module = type(exc).__module__ or ""
    return module.startswith(("requests", "urllib3")) or isinstance(
        exc, (ConnectionError, TimeoutError)
    )


def log_failure(logger, message, exc, level=logging.WARNING, **dimensions):
    """Log a caught exception at the right detail.

    Network failures (121 or Kobo unreachable, timeouts) get one line: they are
    already recorded as failed dependencies, and a stack trace adds nothing.
    Anything else is probably a bug (bad JSON shape, KeyError...), so its stack
    trace is attached and it shows up in the App Insights "exceptions" table.
    """
    logger.log(
        level,
        "%s: %s",
        message,
        describe(exc),
        exc_info=None if _is_network_error(exc) else exc,
        extra=dims(error_type=type(exc).__name__, **dimensions),
    )


def set_context_provider(fn):
    """Register fn() -> dict of dimensions for the current context.

    app.py uses this to stamp every log line written during a request with the
    route, role, actor and program. fn must be cheap and must never raise.
    """
    _state["provider"] = fn


def install_excepthook(logger):
    """Send uncaught exceptions to App Insights before the process dies.

    Without this an uncaught exception only reaches stderr, so a crashed sync
    never appears in App Insights. The original hook still runs afterwards, so
    the traceback on stderr - which app.py shows to the FSP user - is
    unchanged.
    """
    previous = sys.excepthook

    def hook(exc_type, exc, tb):
        if not issubclass(exc_type, KeyboardInterrupt):
            try:
                logger.critical(
                    "Unhandled %s: %s",
                    exc_type.__name__,
                    scrub(exc, 1000),
                    exc_info=(exc_type, exc, tb),
                    extra=dims(
                        event="process.crash",
                        error_type=exc_type.__name__,
                        telemetry_only=True,
                    ),
                )
            except Exception:
                pass
        previous(exc_type, exc, tb)

    sys.excepthook = hook


# ---------------------------------------------------------------------------
# Setup
# ---------------------------------------------------------------------------
class _ContextFilter(logging.Filter):
    """Adds the static and per-request dimensions to every record."""

    def filter(self, record):
        values = dict(_state["static"])
        provider = _state["provider"]
        if provider is not None:
            try:
                values.update(provider() or {})
            except Exception:
                pass
        values.setdefault("log_area", record.name)
        for key, value in values.items():
            # Explicit extra= values on the call win over context values.
            if value not in (None, "") and not hasattr(record, key):
                setattr(record, key, value)
        return True


class _Exclude(logging.Filter):
    def __init__(self, flag):
        super().__init__()
        self.flag = flag

    def filter(self, record):
        return not getattr(record, self.flag, False)


def setup(component, stream_format=None, static_dims=None):
    """Configure logging for this process. Safe to call more than once.

    component    "web" or "sync": the "component" dimension in App Insights,
                 so the two processes can be told apart.
    stream_format  format of the stdout lines. offline_sync.py passes
                 "%(message)s" because app.py parses its stdout.
    static_dims  dimensions stamped on every record (e.g. job_id).
    """
    if _state["configured"]:
        return
    _state["configured"] = True
    _state["static"] = dims(component=component, **(static_dims or {}))

    # Load .env before reading the connection string (local development).
    try:
        from dotenv import load_dotenv

        load_dotenv()
    except Exception:
        pass

    root = logging.getLogger()
    # Set explicitly. logging.basicConfig() used to do this, but basicConfig
    # silently does nothing once configure_azure_monitor() has attached its
    # handler to the root logger - leaving the root at WARNING, so every INFO
    # line (including the ones app.py parses for sync progress) was dropped.
    level_name = os.getenv("LOG_LEVEL", "INFO").upper()
    root.setLevel(getattr(logging, level_name, logging.INFO))

    handlers_before = list(root.handlers)
    azure_problem = None
    if os.getenv("APPLICATIONINSIGHTS_CONNECTION_STRING"):
        try:
            from azure.monitor.opentelemetry import configure_azure_monitor

            configure_azure_monitor()
        except ImportError:
            azure_problem = (
                "azure-monitor-opentelemetry is not installed; telemetry is off."
            )
        except Exception as e:  # a bad connection string must not stop the app
            azure_problem = "Azure Monitor setup failed; telemetry is off. " + describe(e)
    azure_handlers = [h for h in root.handlers if h not in handlers_before]

    # stdout: what the App Service log stream shows, and what app.py reads
    # from the sync subprocess. Added only if nothing already writes there.
    has_stream = any(
        type(h) is logging.StreamHandler for h in handlers_before
    )
    stream_handlers = []
    if not has_stream:
        stream = logging.StreamHandler(sys.stdout)
        stream.setFormatter(logging.Formatter(stream_format or DEFAULT_FORMAT))
        root.addHandler(stream)
        stream_handlers.append(stream)

    context = _ContextFilter()
    for h in stream_handlers + azure_handlers:
        h.addFilter(context)
    for h in stream_handlers:
        h.addFilter(_Exclude("telemetry_only"))
    for h in azure_handlers:
        h.addFilter(_Exclude("stdout_only"))

    for name in _NOISY_LOGGERS:
        logging.getLogger(name).setLevel(logging.WARNING)

    log = logging.getLogger("scandroid.telemetry")
    if azure_problem:
        log.warning(azure_problem)
    elif azure_handlers:
        log.info(
            "Application Insights enabled (component=%s, level=%s)",
            component,
            level_name,
            extra=dims(stdout_only=True),
        )