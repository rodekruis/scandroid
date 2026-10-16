# --- Application Insights / Azure Monitor setup -----------------------------
# This script is launched as a subprocess by app.py and inherits the
# APPLICATIONINSIGHTS_CONNECTION_STRING env var. It reports telemetry as its
# own process. Configure before the instrumented "requests" library is used.
import os
import sys
import logging

# Azure Monitor telemetry is only used in the cloud deployment, where the
# connection string is provided via env var. Import it lazily and only when
# that string is present, so local dev runs (which don't have
# azure-monitor-opentelemetry installed) don't crash the sync subprocess at
# import time. If you *do* want telemetry locally: pip install azure-monitor-opentelemetry
if os.getenv("APPLICATIONINSIGHTS_CONNECTION_STRING"):
    try:
        from azure.monitor.opentelemetry import configure_azure_monitor
        configure_azure_monitor()
    except ImportError:
        logging.getLogger(__name__).warning(
            "azure-monitor-opentelemetry not installed; skipping telemetry."
        )

# Log to stdout with a bare "message only" format so the parent process
# (app.py /sync-fsp) can keep scanning this script's stdout for its status
# lines exactly as it did with print(). App Insights export is handled by the
# separate handler that configure_azure_monitor() attaches to the root logger.
logging.basicConfig(level=logging.INFO, stream=sys.stdout, format="%(message)s")
logger = logging.getLogger(__name__)
# ----------------------------------------------------------------------------

import os
import json
import requests
from datetime import datetime, timedelta
from collections import defaultdict
from concurrent.futures import ThreadPoolExecutor, as_completed

from cryptography.fernet import Fernet
from config_loader import load_config, load_display_config


# ----------------------------------------------------------------------
# CONFIG & GLOBALS
# ----------------------------------------------------------------------

config = load_config()
API_BASE = config["url121"] + "/api"
KOBO_TOKEN = config["KOBO_TOKEN"]
KOBO_BASE = config.get("KOBO_SERVER")

program_id = os.environ.get("PROGRAM_ID")
if not program_id:
    raise RuntimeError("PROGRAM_ID not provided to offline_sync.py")

program_id = str(program_id)

programs = config.get("PROGRAMS", [])
program = next(
    (p for p in programs if str(p.get("programId")) == program_id),
    None
)

if not program:
    raise RuntimeError(f"Program not found for programId={program_id}")

ASSET_ID = program["koboAssetId"]

PROGRAM_ID = program_id

ENCRYPTION_KEY = config["ENCRYPTION_KEY"]

display_config = load_display_config()
_prog_config = display_config.get("programs", {}).get(str(program_id), {})
FIELD_KEYS = [field["key"] for field in _prog_config.get("fields", [])]
PHOTO_FIELD_NAME = _prog_config.get("photo", {}).get("field_name", "photo")

# The attribute staff type into the scan box when the QR code cannot be used.
# "referenceId" (the default) needs nothing extra cached — it IS the record key.
# Anything else must be pulled into the batch, or the device has nothing to
# match against and the setting silently does nothing in the field.
LOOKUP_FIELD = str(
    (_prog_config.get("lookup") or {}).get("field") or "referenceId"
).strip() or "referenceId"

logger.info(f"[INFO] Loaded {len(FIELD_KEYS)} field keys for program {program_id}: {FIELD_KEYS}")
logger.info(f"[INFO] Lookup field for program {program_id}: '{LOOKUP_FIELD}'")

try:
    fernet = Fernet(ENCRYPTION_KEY.encode())
except ValueError as e:
    logger.error(f"[!] ENCRYPTION_KEY is invalid (not a valid Fernet key): {e}")
    sys.exit(1)

# Thread pool size (can be overridden by env var)
MAX_WORKERS = int(os.getenv("OFFLINE_SYNC_WORKERS", "8"))

# ---------------------------------------------------------------------------
# COMPLETENESS / PAGINATION SETTINGS
# ---------------------------------------------------------------------------
# GET /programs/{id}/payments/{paymentId}/transactions is a nestjs-paginate
# endpoint. If no `limit` is sent, 121 applies DEFAULT_PAGINATION_LIMIT (20)
# and returns {"data": [...20 items...], "meta": {...}}. Because the old code
# read only the "data" key, every payment silently synced at most 20
# beneficiaries and the rest were simply absent from the offline batch — with
# no error anywhere. Never call a 121 list endpoint without an explicit limit.
#
# nestjs-paginate treats limit=-1 as NO_PAGINATION, and 121's own
# PaginateConfigTransactionView sets maxLimit=-1, so -1 is accepted and
# returns everything. We still verify the returned count against meta.totalItems
# and fall back to explicit page-by-page fetching if a server build ever
# refuses it.
NO_PAGINATION = -1
PAGE_SIZE = int(os.getenv("OFFLINE_SYNC_PAGE_SIZE", "1000"))
MAX_PAGES = int(os.getenv("OFFLINE_SYNC_MAX_PAGES", "10000"))
REQUEST_TIMEOUT = int(os.getenv("OFFLINE_SYNC_TIMEOUT", "120"))

# STRICT mode (default ON): abort the sync rather than write a batch that is
# known to be incomplete. app.py checks the subprocess return code and shows
# the failure to the FSP user. A failed sync is recoverable; a batch that looks
# complete but is missing people is not — those people cannot be verified.
STRICT = os.getenv("OFFLINE_SYNC_STRICT", "1").lower() not in ("0", "false", "no")

COOKIES = None
HEADERS_KOBO = {"Authorization": f"Token {KOBO_TOKEN}"}


class IncompleteSyncError(RuntimeError):
    """Raised when the batch we are about to write is known to be missing
    beneficiaries. Never swallow this."""


def _fail(message):
    """Abort in strict mode; otherwise log loudly and continue."""
    if STRICT:
        raise IncompleteSyncError(message)
    logger.error("[INCOMPLETE] %s (continuing: OFFLINE_SYNC_STRICT is off)", message)


# ----------------------------------------------------------------------
# ENCRYPTION HELPERS
# ----------------------------------------------------------------------

def encrypt_data(data_dict):
    """
    Encrypt all values in a dict with Fernet.
    Values are cast to string; None becomes "".
    """
    encrypted = {}
    for key, value in data_dict.items():
        plain = str(value) if value is not None else ""
        encrypted[key] = fernet.encrypt(plain.encode()).decode()
    return encrypted


def encrypt_photo(photo_bytes):
    return fernet.encrypt(photo_bytes)


# ----------------------------------------------------------------------
# AUTH / SESSION
# ----------------------------------------------------------------------

def login_and_get_token():
    """
    Log in to 121 API and obtain access_token_general.
    """
    global COOKIES
    login_url = f"{API_BASE}/users/login"
    credentials = {
        "username": config["username121"],
        "password": config["password121"],
    }
    response = requests.post(
        login_url,
        headers={"Content-Type": "application/json"},
        json=credentials,
    )
    response.raise_for_status()
    token = response.json().get("access_token_general")
    if not token:
        raise Exception("Login successful but token missing.")
    COOKIES = {"access_token_general": token}
    return token


# Initialise token/cookies at import-time
login_and_get_token()


# ----------------------------------------------------------------------
# 121 API HELPERS
# ----------------------------------------------------------------------

def _unwrap_list(data):
    """121 is inconsistent about envelopes: /payments returns a bare array,
    while /payments/{id}/transactions returns {"data": [...], "meta": {...}}.
    Normalise both so callers always get a list.

    NOTE: this deliberately does NOT look at "meta". Any caller hitting a
    paginated endpoint must use _fetch_all(), which reconciles the returned
    item count against meta.totalItems. Unwrapping alone cannot tell a
    complete response from a truncated one."""
    if isinstance(data, list):
        return data
    if isinstance(data, dict):
        for key in ("data", "transactions", "payments"):
            value = data.get(key)
            if isinstance(value, list):
                return value
        logger.error("[ERROR] Unexpected list payload; keys=%s", list(data.keys()))
    return []


def _meta_of(payload):
    """Return the nestjs-paginate meta block, or {} for bare-array responses."""
    if isinstance(payload, dict):
        meta = payload.get("meta")
        if isinstance(meta, dict):
            return meta
    return {}


def _fetch_all(url, params=None, label=""):
    """GET a (possibly paginated) 121 list endpoint and return ALL items.

    Strategy:
      1. Ask for limit=-1 (NO_PAGINATION). This is what 121's own code uses
         internally when it needs every row.
      2. Compare len(items) against meta.totalItems. If the server paginated us
         anyway, fall back to walking pages explicitly.
      3. If we still end up short, raise — a short read here means missing
         beneficiaries in the field.

    Returns (items, total_items_reported_or_None).
    """
    base_params = dict(params or {})

    # --- 1) single unpaginated request ---
    first_params = dict(base_params)
    first_params["limit"] = NO_PAGINATION
    response = requests.get(url, cookies=COOKIES, params=first_params,
                            timeout=REQUEST_TIMEOUT)
    response.raise_for_status()
    payload = response.json()
    items = _unwrap_list(payload)
    meta = _meta_of(payload)
    total = meta.get("totalItems")

    if total is None or len(items) >= total:
        return items, total

    # --- 2) server enforced pagination: walk the pages ---
    logger.warning(
        "[!] %s: limit=-1 returned %s of %s rows; falling back to paged fetch "
        "(pageSize=%s)", label or url, len(items), total, PAGE_SIZE
    )

    collected = []
    seen_pages = 0
    page = 1
    while page <= MAX_PAGES:
        page_params = dict(base_params)
        page_params["limit"] = PAGE_SIZE
        page_params["page"] = page
        r = requests.get(url, cookies=COOKIES, params=page_params,
                         timeout=REQUEST_TIMEOUT)
        r.raise_for_status()
        page_payload = r.json()
        batch = _unwrap_list(page_payload)
        page_meta = _meta_of(page_payload)
        total = page_meta.get("totalItems", total)
        total_pages = page_meta.get("totalPages")

        collected.extend(batch)
        seen_pages += 1

        if not batch:
            break
        if total_pages is not None and page >= total_pages:
            break
        if total is not None and len(collected) >= total:
            break
        page += 1
    else:
        raise IncompleteSyncError(
            f"{label or url}: exceeded MAX_PAGES={MAX_PAGES} while paging."
        )

    logger.info("[INFO] %s: fetched %s rows across %s page(s)",
                label or url, len(collected), seen_pages)

    # --- 3) final reconciliation ---
    if total is not None and len(collected) < total:
        raise IncompleteSyncError(
            f"{label or url}: fetched {len(collected)} of {total} rows. "
            "Refusing to build an incomplete offline batch."
        )

    return collected, total


def _transfer_value(t):
    """121 transactions carry the money as `transferValue`. Accept `amount` too,
    so a schema difference between endpoints can't silently zero out every
    amount the FSP sees on screen."""
    for key in ("transferValue", "amount"):
        value = t.get(key)
        if value is not None:
            return value
    return 0


def _parse_iso(value):
    """Parse the two ISO shapes 121 emits (with and without fractional
    seconds). Returns None when unparseable rather than raising."""
    if not value:
        return None
    raw = str(value).replace("Z", "").replace("+00:00", "")
    for fmt in ("%Y-%m-%dT%H:%M:%S.%f", "%Y-%m-%dT%H:%M:%S"):
        try:
            return datetime.strptime(raw, fmt)
        except ValueError:
            continue
    return None


def get_payments(program_id):
    """
    GET /api/programs/{id}/payments -> bare array of payment summaries.

    This endpoint is NOT paginated (it maps to getPaymentAggregationsSummaries,
    which returns a plain array), but we route it through _fetch_all anyway so
    that a future change to a paginated shape can't silently truncate it.

    Each item carries paymentId, name, paymentDate and aggregationsPerStatus.
    That last field tells us how many transactions are still `waiting` BEFORE we
    fetch any of them, so we only pull transactions for payments that actually
    have something left to distribute — and it is the yardstick we reconcile
    the fetched transactions against.
    """
    url = f"{API_BASE}/programs/{program_id}/payments"
    payments, _ = _fetch_all(url, label=f"payments(program={program_id})")
    return payments


def get_transactions(program_id, payment_id):
    """
    GET all transactions for one payment.

    THIS ENDPOINT IS PAGINATED. Without an explicit limit 121 returns only
    DEFAULT_PAGINATION_LIMIT (20) rows. _fetch_all sends limit=-1 and verifies
    the count against meta.totalItems, so a truncated read raises instead of
    silently shrinking the batch.
    """
    url = f"{API_BASE}/programs/{program_id}/payments/{payment_id}/transactions"
    transactions, total = _fetch_all(
        url, label=f"transactions(payment={payment_id})"
    )
    if total is not None and len(transactions) != total:
        raise IncompleteSyncError(
            f"payment {payment_id}: got {len(transactions)} transactions, "
            f"API reports {total}."
        )
    logger.info("[INFO] paymentId=%s: %s transaction(s) fetched",
                payment_id, len(transactions))
    return transactions


def get_all_transactions(program_id):
    """
    Get ALL transactions for a program.

    DEPRECATED / unused: the sync is now driven by GET /programs/{id}/payments
    (see select_open_payments). Kept only for ad-hoc debugging.
    """
    url = f"{API_BASE}/programs/{program_id}/transactions"
    response = requests.get(url, cookies=COOKIES, timeout=REQUEST_TIMEOUT)
    response.raise_for_status()
    data = response.json()

    if isinstance(data, dict):
        if "transactions" in data:
            return data["transactions"]
        if "data" in data:
            return data["data"]
        logger.error("[ERROR] Unexpected transaction structure: %s", data.keys())
        return []
    elif isinstance(data, list):
        return data
    else:
        logger.error("[ERROR] Unknown transaction data type")
        return []


def get_registration(program_id, registration_id):
    url = f"{API_BASE}/programs/{program_id}/registrations/{registration_id}"
    response = requests.get(url, cookies=COOKIES, timeout=REQUEST_TIMEOUT)
    response.raise_for_status()
    return response.json()


# In 121, Excel is the only FSP whose integrationType is 'csv', and the only one
# that declares columnToMatch (excel-settings.const.ts, isRequired: true). Every
# other integration is 'api' and is reconciled by 121 itself.
FSP_EXCEL = "Excel"
FSP_TYPE_CSV = "csv"
FSP_TYPE_API = "api"


def build_fsp_column_map(program_id):
    """Returns (col_map, fsp_types) for this program.

        col_map   {programFspConfigurationName: columnToMatch}
        fsp_types {programFspConfigurationName: 'csv' | 'api'}

    INVARIANT 1: keyed on the configuration NAME, never on array position.
    columnToMatch belongs to an FSP configuration, and a program routinely has
    several with different values (program 2 has three: phoneNumber, fullName,
    paID121). The old code returned the first one it happened to see, so the
    answer depended on the order 121 serialised the array in.

    INVARIANT 3: a non-200 raises. Silently treating "unauthorised" as "not
    configured" is what let a wrong column reach the field with no log line.

    INVARIANT 7: whether a payment can be reconciled offline is decided by the
    FSP's integrationType, NEVER by the configuration's name. Names are free
    text: this program has Excel configurations named 'mtn-direct',
    'mtn-nouvel-sim' and 'mtn-ifrc-sim', while the real MTN api integration has
    no columnToMatch at all. A name-based test is wrong in both directions.
    """
    url = f"{API_BASE}/programs/{program_id}/fsp-configurations"
    r = requests.get(url, cookies=COOKIES, timeout=REQUEST_TIMEOUT)
    r.raise_for_status()

    payload = r.json()
    configs = payload if isinstance(payload, list) else payload.get("data", [])

    col_map = {}
    fsp_types = {}

    for fsp in configs:
        if not isinstance(fsp, dict):
            continue
        name = fsp.get("name")
        if not name:
            continue

        # The nested `fsp` block carries the integration settings. 121's own DTO
        # notes it can be undefined when an integration has been removed from
        # the codebase, so fall back to the fspName enum value.
        settings = fsp.get("fsp") or {}
        integration_type = settings.get("integrationType")
        if not integration_type:
            integration_type = (
                FSP_TYPE_CSV if fsp.get("fspName") == FSP_EXCEL else FSP_TYPE_API
            )
        fsp_types[name] = integration_type

        for prop in fsp.get("properties", []):
            if prop.get("name") == "columnToMatch" and prop.get("value"):
                col_map[name] = prop["value"]
                break

    csv_configs = sorted(n for n, t in fsp_types.items() if t == FSP_TYPE_CSV)

    if not csv_configs:
        raise IncompleteSyncError(
            f"Program {program_id} has no Excel FSP configuration. 121 Scan "
            "reconciles payments by Excel upload, so there is nothing here it "
            f"can distribute. Configurations found: {sorted(fsp_types)}"
        )

    # INVARIANT 8. Stricter than the old `if not col_map`, which fired only when
    # EVERY configuration lacked a column. columnToMatch is a REQUIRED property
    # of the Excel FSP, so an Excel configuration without one is a genuine 121
    # misconfiguration and must still stop the sync — even if another Excel
    # configuration on the same program is fine.
    misconfigured = [n for n in csv_configs if n not in col_map]
    if misconfigured:
        raise IncompleteSyncError(
            f"Program {program_id}: Excel FSP configuration(s) {misconfigured} "
            "have no columnToMatch. It is a required property — set 'Field for "
            "identifying registrations' on that configuration in 121."
        )

    for name in sorted(fsp_types):
        if fsp_types[name] == FSP_TYPE_CSV:
            logger.info("[fsp-config] %s (excel) -> columnToMatch '%s'",
                        name, col_map[name])
        else:
            logger.info(
                "[fsp-config] %s (%s) -> reconciled by 121 directly; payments "
                "on it are not distributed by 121 Scan.",
                name, fsp_types[name]
            )

    return col_map, fsp_types


def _is_excel_config(config_name, fsp_types, fsp_name=None):
    """True when this FSP configuration is Excel-based, and therefore
    reconcilable by 121 Scan's CSV upload.

    Two independent signals, most reliable first:
      1. integrationType from /fsp-configurations, keyed on configuration name
      2. fspName carried on the transaction row itself — 121's transaction_view
         exposes programFspConfigurationName AND fspName per row

    An unrecognised configuration is treated as Excel so that a lookup miss can
    never silently drop a real beneficiary. The columnToMatch checks downstream
    are the backstop for that case.
    """
    if config_name and config_name in fsp_types:
        return fsp_types[config_name] == FSP_TYPE_CSV
    if fsp_name:
        return fsp_name == FSP_EXCEL
    return True


def payment_fsp_names(program_id, payment):
    """The programFspConfigurationName(s) a payment covers.

    Prefers the `fsps` array already present on the payments-list payload; falls
    back to GET /programs/{id}/payments/{paymentId} when the list form does not
    carry it. Written to work either way so it does not depend on which shape
    this 121 build returns.
    """
    payment_id = payment.get("paymentId")
    fsps = payment.get("fsps")

    if not isinstance(fsps, list) or not fsps:
        url = f"{API_BASE}/programs/{program_id}/payments/{payment_id}"
        r = requests.get(url, cookies=COOKIES, timeout=REQUEST_TIMEOUT)
        r.raise_for_status()
        fsps = r.json().get("fsps") or []

    names = []
    for entry in fsps:
        if isinstance(entry, dict):
            name = entry.get("programFspConfigurationName")
            if name and name not in names:
                names.append(name)
    return names


def resolve_payment_column(program_id, payment, col_map, names=None):
    """The columnToMatch for ONE payment, or raise.

    INVARIANT 2: resolution is anchored to the payment, which knows its FSP
    configuration. There is deliberately no program-level fallback and no
    global — INVARIANT 5 — because either would reintroduce exactly the class
    of bug this replaces.
    """
    payment_id = payment.get("paymentId")
    if names is None:
        names = payment_fsp_names(program_id, payment)

    if not names:
        _fail(
            f"paymentId={payment_id} \"{payment.get('name')}\" reports no FSP "
            "configuration, so its match column cannot be determined."
        )
        return None

    unknown = [n for n in names if n not in col_map]
    if unknown:
        _fail(
            f"paymentId={payment_id} \"{payment.get('name')}\" uses FSP "
            f"configuration(s) {unknown} which have no columnToMatch in 121. "
            f"Known configurations: {sorted(col_map)}"
        )
        return None

    columns = {col_map[n] for n in names}

    if len(columns) > 1:
        # INVARIANT 4. A single reconciliation CSV has one header, so a payment
        # spanning configurations that disagree cannot be reconciled at all.
        # Surface it here rather than silently half-reconciling in the field.
        _fail(
            f"paymentId={payment_id} \"{payment.get('name')}\" spans FSP "
            f"configurations {names} with conflicting columnToMatch values "
            f"{sorted(columns)}. It cannot be reconciled with one CSV — split "
            "the payment in 121, or align the configurations."
        )
        return None

    column = columns.pop()
    logger.info(
        "[match-column] paymentId=%s \"%s\" fsp=%s -> '%s'",
        payment_id, payment.get("name"), names, column
    )
    return column

def fetch_registrations_bulk(program_id, registration_ids):
    """
    Fetch registrations in parallel for a set of registrationIds.

    Returns (results, failed_ids):
        results:    {registrationId: registration_json}
        failed_ids: [registrationId, ...] that could not be fetched

    A failed fetch used to be a lone warning line, after which the beneficiary
    silently disappeared from the batch. Failures are now returned so the
    caller can refuse to ship an incomplete batch.
    """
    results = {}
    failed_ids = []
    unique_ids = list(set(registration_ids))

    if not unique_ids:
        return results, failed_ids

    def worker(rid):
        last_error = None
        for attempt in range(3):
            try:
                return rid, get_registration(program_id, rid), None
            except Exception as e:
                last_error = e
        return rid, None, last_error

    total = len(unique_ids)
    done = 0
    max_workers = min(MAX_WORKERS, total) or 1
    with ThreadPoolExecutor(max_workers=max_workers) as executor:
        futures = [executor.submit(worker, rid) for rid in unique_ids]
        for fut in as_completed(futures):
            rid, reg, error = fut.result()
            if reg is not None:
                results[rid] = reg
            else:
                failed_ids.append(rid)
                logger.warning(f"[!] Failed to get registration {rid}: {error}")

            # Incremental progress for the /sync-status panel. Without these the
            # UI has nothing to show between the "N beneficiaries" line and the
            # final summary — on a 1,000-person program that is minutes of a
            # frozen progress bar.
            done += 1
            if done % 25 == 0 or done == total:
                logger.info(f"[PROGRESS] registrations {done}/{total}")

    logger.info("[INFO] Registrations: %s requested, %s fetched, %s failed",
                len(unique_ids), len(results), len(failed_ids))

    return results, failed_ids


# ----------------------------------------------------------------------
# KOBO HELPERS
# ----------------------------------------------------------------------

def get_kobo_submission(uuid):
    """
    Fetch a single Kobo submission by _uuid.
    Keeps behaviour identical but now all photo optimisations
    are handled in download_and_encrypt_photo.
    """
    # Keep it simple & safe: full submission (no fields filter),
    # since we rely on photo field, *_URL, _attachments, and _id.
    url = f"{KOBO_BASE}/api/v2/assets/{ASSET_ID}/data.json?query={{\"_uuid\":\"{uuid}\"}}"
    response = requests.get(url, headers=HEADERS_KOBO, timeout=REQUEST_TIMEOUT)
    response.raise_for_status()
    results = response.json().get("results", [])
    return results[0] if results else None


def download_and_encrypt_photo(uuid, save_path):
    """
    Download and encrypt the photo for a given submission UUID.
    Saves the encrypted image bytes to save_path.
    Handles:
    - Kobo's direct *_URL field (photo_URL)
    - Kobo _attachments list
    - IFRC Kobo /attachments/<uid>/ format
    Uses a smaller 'medium' view to speed up sync.

    Returns True on success, False on failure (a beneficiary without a photo
    cannot be visually verified, so the caller counts these).
    """

    # If photo already exists and is non-empty, you *could* skip.
    # For now, we always refresh since each batch dir is unique.

    # --- 1) Fetch Kobo submission ---
    submission = get_kobo_submission(uuid)
    if not submission:
        logger.warning(f"[!] No Kobo submission found for UUID {uuid}")
        return False

    photo_field = PHOTO_FIELD_NAME  # e.g. "photo"
    photo_filename = submission.get(photo_field)

    # --- 2) New Kobo way: direct photo URL (BEST METHOD) ---
    photo_url_field = f"{photo_field}_URL"  # e.g. "photo_URL"
    photo_url = submission.get(photo_url_field)

    if photo_url:
        logger.info(f"[OK] Direct Kobo photo URL found for UUID {uuid}: {photo_url}")

        # Use smaller 'medium' image instead of original
        photo_url = photo_url.replace("/original/", "/medium/")

        res = requests.get(photo_url, headers=HEADERS_KOBO, timeout=REQUEST_TIMEOUT)
        if res.status_code != 200:
            logger.warning(f"[!] Direct photo download failed for UUID {uuid}: {res.status_code}")
            return False

        encrypted_bytes = encrypt_photo(res.content)
        with open(save_path, "wb") as f:
            f.write(encrypted_bytes)

        logger.info(f"[OK] Photo downloaded & encrypted (direct URL) for UUID {uuid}")
        return True

    # --- 3) Fallback: match against _attachments (older Kobo submissions) ---
    if not photo_filename:
        logger.warning(f"[!] No '{photo_field}' value for UUID {uuid}")
        return False

    attachments = submission.get("_attachments", [])
    if not attachments:
        logger.warning(f"[!] No attachments in submission for UUID {uuid}")
        return False

    from urllib.parse import unquote

    # Normalise filename for comparison: URL-decode, underscores=spaces, lowercase
    def norm(s):
        return unquote(str(s)).replace("_", " ").lower()

    photo_base = photo_filename.rsplit(".", 1)[0]

    matching = [
        a for a in attachments
        if norm(photo_base) in norm(a.get("filename", ""))
        or norm(photo_filename) in norm(a.get("filename", ""))
    ]

    if not matching:
        fnames = [a.get("filename", "") for a in attachments]
        logger.warning(f"[!] No matching attachment for '{photo_filename}' (UUID {uuid})")
        logger.debug(f"[DEBUG] Available filenames: {fnames}")
        return False

    att = matching[0]
    attach_uid = att.get("uid")
    submission_id = submission["_id"]

    # --- 4) Try direct download_url first, then construct IFRC URL ---
    direct_url = att.get("download_url") or att.get("download_medium_url")

    if direct_url:
        file_url = direct_url.replace("/original/", "/medium/")
        logger.info(f"[OK] Using direct download_url for UUID {uuid}: {file_url[:80]}")
    elif attach_uid:
        file_url = (
            f"{KOBO_BASE}/api/v2/assets/{ASSET_ID}/data/"
            f"{submission_id}/attachments/{attach_uid}/?view=medium"
        )
        logger.info(f"[OK] Using constructed IFRC URL for UUID {uuid}: {file_url[:80]}")
    else:
        logger.warning(f"[!] No download URL or UID for attachment (UUID {uuid})")
        return False

    # --- 5) Download (retry without medium if it fails) ---
    res = requests.get(file_url, headers=HEADERS_KOBO, timeout=REQUEST_TIMEOUT)
    if res.status_code != 200:
        fallback = file_url.replace("/medium/", "/original/").replace("?view=medium", "")
        logger.warning(f"[!] Medium download failed ({res.status_code}), retrying: {fallback[:80]}")
        res = requests.get(fallback, headers=HEADERS_KOBO, timeout=REQUEST_TIMEOUT)
    if res.status_code != 200:
        logger.warning(f"[!] Failed to download attachment for UUID {uuid}: {res.status_code}")
        return False

    # --- 6) Encrypt & save ---
    encrypted_bytes = encrypt_photo(res.content)
    with open(save_path, "wb") as f:
        f.write(encrypted_bytes)

    logger.info(f"[OK] Photo downloaded & encrypted for UUID {uuid}")
    return True


def download_photos_bulk(records, photos_dir):
    """
    Download & encrypt photos for all records in parallel.
    Each record should have 'uuid' and 'photo_filename'.

    Returns a list of uuids whose photo could NOT be retrieved. These
    beneficiaries are still in the batch (their data is valid) but cannot be
    visually verified, so the count is surfaced in the manifest and the log.
    """
    failed = []
    if not records:
        return failed

    os.makedirs(photos_dir, exist_ok=True)

    def worker(rec):
        uuid = rec["uuid"]
        photo_filename = rec["photo_filename"]
        save_path = os.path.join(photos_dir, photo_filename)
        try:
            ok = download_and_encrypt_photo(uuid, save_path)
        except Exception as e:
            logger.warning(f"[!] Photo error for UUID {uuid}: {e}")
            ok = False
        return uuid, ok

    total = len(records)
    done = 0
    max_workers = min(MAX_WORKERS, total) or 1
    with ThreadPoolExecutor(max_workers=max_workers) as executor:
        futures = [executor.submit(worker, rec) for rec in records]
        for fut in as_completed(futures):
            uuid, ok = fut.result()
            if not ok:
                failed.append(uuid)

            # Photo download is the longest phase by far, so report often —
            # this is what actually drives the progress bar in fsp_admin.html.
            done += 1
            if done % 10 == 0 or done == total:
                logger.info(f"[PROGRESS] photos {done}/{total}")

    if failed:
        logger.error("[!] %s of %s photo(s) could not be downloaded — those "
                     "beneficiaries cannot be visually verified offline.",
                     len(failed), len(records))
    else:
        logger.info("[INFO] All %s photo(s) downloaded & encrypted.", len(records))

    return failed


# ----------------------------------------------------------------------
# BATCH DIRECTORY HELPERS
# ----------------------------------------------------------------------

def get_next_batch_dir(base_path, payment_id):
    """
    Create a unique batch directory for a given payment or "recent",
    with a 'photos' subfolder.
    """
    batch_number = 1
    while True:
        batch_path = os.path.join(base_path, f"payment-{payment_id}-batch-{batch_number}")
        if not os.path.exists(batch_path):
            os.makedirs(os.path.join(batch_path, "photos"), exist_ok=True)
            return batch_path
        batch_number += 1


# ----------------------------------------------------------------------
# MAIN: SPECIFIC PAYMENT BATCH
# ----------------------------------------------------------------------

def download_cache(program_id, payment_id):
    """
    Original behaviour: download cache for a single paymentId.

    - Fetch transactions for this payment (ALL of them — see get_transactions)
    - Fetch registrations (in parallel)
    - Fetch + encrypt photos (in parallel, medium-size)
    - Save registrations_cache.json and transactions.json
    """
    base_path = "offline-cache"
    os.makedirs(base_path, exist_ok=True)
    batch_dir = get_next_batch_dir(base_path, payment_id)
    photos_dir = os.path.join(batch_dir, "photos")

    # Resolve the match column ONCE, live from 121, reused for every record
    # and written into the manifest below.
    col_map, _fsp_types = build_fsp_column_map(program_id)
    match_key = resolve_payment_column(
        program_id, {"paymentId": payment_id, "name": f"Payment {payment_id}"}, col_map
    )
    if not match_key:
        raise IncompleteSyncError(
            f"Could not resolve columnToMatch for paymentId {payment_id}."
        )

    transactions = get_transactions(program_id, payment_id)
    cache_data = []

    # 1) Collect registrationIds & uuids from transactions
    reg_ids = []
    for t in transactions:
        if "registrationId" in t and "registrationReferenceId" in t:
            reg_ids.append(t["registrationId"])

    # 2) Fetch registrations in bulk (parallel)
    registrations_map, failed_reg_ids = fetch_registrations_bulk(program_id, reg_ids)
    if failed_reg_ids:
        _fail(
            f"{len(failed_reg_ids)} registration(s) could not be fetched for "
            f"payment {payment_id}: {failed_reg_ids[:20]}"
        )

    # 3) Build records (encryption, validity checks)
    for t in transactions:
        reg_id = t.get("registrationId")
        uuid = t.get("registrationReferenceId")

        if not reg_id or not uuid:
            logger.info("[SKIP] Missing reg_id or uuid in transaction")
            continue

        reg = registrations_map.get(reg_id)
        if not reg:
            logger.warning(f"[!] No registration data for {reg_id}")
            continue

        filtered_data = {key: reg.get(key) for key in FIELD_KEYS}
        if match_key:
            filtered_data[match_key] = reg.get(match_key)
        if LOOKUP_FIELD != "referenceId":
            filtered_data[LOOKUP_FIELD] = reg.get(LOOKUP_FIELD)

        encrypted_data = encrypt_data(filtered_data)

        photo_filename = f"{uuid}.enc"

        status = (t.get("status") or t.get("transactionStatus") or "").lower()
        deleted = (t.get("registrationStatus") or "").lower() == "deleted"

        is_valid = status == "waiting" and not deleted
        reason = "ok"
        if not is_valid:
            if status != "waiting":
                reason = f"status={status}"
            elif deleted:
                reason = "deleted"

        record = {
            "uuid": uuid,
            "registrationId": reg_id,
            "photo_filename": photo_filename,
            "paymentId": t.get("paymentId"),
            "amount": _transfer_value(t),
            "data": encrypted_data,
            # Explicit encryption metadata — app.py's reconciliation reads
            # record["dataEncrypted"] to decide whether to decrypt, instead of
            # guessing. Do not remove.
            "dataEncrypted": True,
            "encryptionScheme": "fernet-v1",
            "valid": is_valid,
            "reason": reason,
        }

        cache_data.append(record)

    # 4) Download & encrypt all photos in parallel
    photo_failures = download_photos_bulk(cache_data, photos_dir)

    # 5) Save encrypted registration data & transactions
    json_path = os.path.join(batch_dir, "registrations_cache.json")
    with open(json_path, "w", encoding="utf-8") as f:
        json.dump(cache_data, f, indent=2)

    tx_path = os.path.join(batch_dir, "transactions.json")
    with open(tx_path, "w", encoding="utf-8") as f:
        json.dump(transactions, f, indent=2)

    batch_info = {
        "batchType": "payment-single",
        "programId": program_id,
        "paymentId": payment_id,
        "columnToMatch": match_key,
        "dataEncrypted": True,
        "encryptionScheme": "fernet-v1",
        "recordCount": len(cache_data),
        "transactionCount": len(transactions),
        "registrationFetchFailures": len(failed_reg_ids),
        "photoFailures": len(photo_failures),
        "complete": not failed_reg_ids,
        "generatedAt": datetime.utcnow().isoformat() + "Z",
    }
    with open(os.path.join(batch_dir, "batch_info.json"), "w", encoding="utf-8") as f:
        json.dump(batch_info, f, indent=2)

    logger.info(f"\n[OK] Done. Batch saved to: {batch_dir}")
    logger.info(f"{len(cache_data)} beneficiaries ready for offline validation.")
    return len(cache_data)


# ----------------------------------------------------------------------
# MAIN: OPEN PAYMENTS BATCH (multi-payment aware)
# ----------------------------------------------------------------------

# How far back to look. A payment older than this is assumed to be closed out
# even if 121 still reports waiting transactions on it.
# NOTE: this is a second, independent way beneficiaries can go missing from a
# batch. A payment with waiting transactions whose paymentDate is older than
# WINDOW_DAYS is skipped entirely. Raise OFFLINE_SYNC_WINDOW_DAYS if a
# distribution is running long.
WINDOW_DAYS = int(os.getenv("OFFLINE_SYNC_WINDOW_DAYS", "365"))


def select_open_payments(program_id, col_map, fsp_types, window_days=WINDOW_DAYS):
    """
    Decide which payments are worth pulling transactions for.

    Driven by GET /programs/{id}/payments, which reports, per payment, how many
    transactions sit in each status. We keep a payment when it still has waiting
    transactions and its paymentDate falls inside the window. This replaces the
    old approach of fetching every transaction in the program and filtering
    client-side, and it gives us the payment `name` for free.

    The waiting count from this endpoint is also the yardstick: fetch_tranches
    reconciles the transactions it actually retrieves against it, which is what
    catches a truncated/paginated read.

    Returns a list of dicts: paymentId, name, paymentDate, waitingCount.
    """
    payments = get_payments(program_id)
    logger.info(f"[INFO] Program {program_id}: {len(payments)} payment(s) in 121")

    cutoff = datetime.utcnow() - timedelta(days=window_days)
    selected = []
    skipped_no_waiting = 0
    skipped_too_old = 0
    skipped_waiting_total = 0
    skipped_not_excel = []          # payments 121 reconciles itself
    skipped_not_excel_waiting = 0

    for p in payments:
        payment_id = p.get("paymentId", p.get("id"))
        if payment_id is None:
            continue

        aggregations = p.get("aggregationsPerStatus") or {}
        waiting_count = (aggregations.get("waiting") or {}).get("count") or 0
        if not waiting_count:
            skipped_no_waiting += 1
            continue

        payment_date = p.get("paymentDate")
        parsed_date = _parse_iso(payment_date)
        if parsed_date and parsed_date < cutoff:
            skipped_too_old += 1
            skipped_waiting_total += waiting_count
            logger.warning(
                "[!] paymentId=%s \"%s\" skipped as older than %s days "
                "(date=%s) but still has %s waiting transaction(s). Those "
                "beneficiaries will NOT be in this batch.",
                payment_id, p.get("name"), window_days, payment_date, waiting_count
            )
            continue

        fsp_names = payment_fsp_names(program_id, p)
        excel_names = [n for n in fsp_names if _is_excel_config(n, fsp_types)]

        # INVARIANT 8. A payment entirely on api-integrated FSPs (Safaricom,
        # Airtel, MTN, Nedbank, Intersolve, Onafriq, ...) has no columnToMatch
        # and no CSV to upload: 121 reconciles it itself through the FSP's API.
        # 121 Scan cannot reconcile it and must not try. SKIP it — this used to
        # call _fail() and abort the entire sync, taking every Excel payment on
        # the program down with it.
        if fsp_names and not excel_names:
            skipped_not_excel.append({
                "paymentId": payment_id,
                "name": p.get("name"),
                "fspConfigNames": fsp_names,
                "waitingCount": waiting_count,
            })
            skipped_not_excel_waiting += waiting_count
            logger.info(
                "[skip] paymentId=%s \"%s\" fsp=%s — reconciled by 121 "
                "directly, not via 121 Scan. %s waiting transaction(s) "
                "excluded from this batch.",
                payment_id, p.get("name"), fsp_names, waiting_count
            )
            continue

        if len(excel_names) != len(fsp_names):
            # Mixed payment: the Excel side is ours, the api side is 121's.
            # fetch_tranches() filters the individual transactions, because the
            # FSP configuration lives on the transaction (INVARIANT 9).
            logger.warning(
                "[!] paymentId=%s \"%s\" spans Excel and api FSP "
                "configurations %s. Only the Excel one(s) %s are distributed "
                "by 121 Scan; the rest stay with 121.",
                payment_id, p.get("name"), fsp_names, excel_names
            )

        entry = {
            "paymentId": payment_id,
            # Names are auto-generated by 121 and are NOT guaranteed unique —
            # two payments minutes apart can share a name. The paymentId is the
            # only reliable identifier, so it is carried alongside and is what
            # the FSP's selection is keyed on.
            "name": p.get("name") or f"Payment {payment_id}",
            "paymentDate": payment_date,
            "waitingCount": waiting_count,
            "isPaymentApproved": p.get("isPaymentApproved"),
            # Excel configurations ONLY. The api ones are not ours to reconcile
            # and must never reach resolve_payment_column(), which would
            # correctly but uselessly fail on them.
            "fspConfigNames": excel_names,
            "allFspConfigNames": fsp_names,
        }
        # Resolved once, here, and carried everywhere downstream. Nothing later
        # in the pipeline re-derives it (INVARIANT 6).
        entry["columnToMatch"] = resolve_payment_column(
            program_id, p, col_map, names=entry["fspConfigNames"]
        )
        selected.append(entry)

    selected.sort(key=lambda x: (str(x.get("paymentDate") or ""), x["paymentId"]))

    logger.info(
        f"[INFO] {len(selected)} open payment(s) selected "
        f"(skipped: {skipped_no_waiting} with nothing waiting, "
        f"{skipped_too_old} older than {window_days} days "
        f"holding {skipped_waiting_total} waiting transaction(s), "
        f"{len(skipped_not_excel)} on api-integrated FSPs holding "
        f"{skipped_not_excel_waiting} waiting transaction(s))"
    )
    for p in selected:
        logger.info(
            f"  - paymentId={p['paymentId']} \"{p['name']}\" "
            f"waiting={p['waitingCount']} date={p.get('paymentDate')} "
            f"fsp={p.get('fspConfigNames')} column='{p.get('columnToMatch')}'"
        )

    return selected, skipped_not_excel


def fetch_tranches(program_id, payments, col_map, fsp_types):
    """
    Fetch waiting transactions for each open payment (in parallel) and group them
    by beneficiary.

    A beneficiary legitimately appears under several payments at once — that is
    the whole point of this batch. Returns:
        tranches_by_uuid: {registrationReferenceId: [tranche, ...]}
        reg_id_by_uuid:   {registrationReferenceId: registrationId}
        reconciliation:   per-payment expected-vs-seen report

    RECONCILIATION: for every payment we compare the number of `waiting`
    transactions we actually received against the waitingCount 121 reported in
    the payments aggregation. Fewer waiting rows than expected means the read
    was truncated (the pagination bug) or the data shifted mid-sync — either
    way the batch would be missing people, so we refuse to write it.
    """
    tranches_by_uuid = defaultdict(dict)  # uuid -> {paymentId: tranche}
    reg_id_by_uuid = {}
    reconciliation = []

    if not payments:
        return {}, {}, reconciliation

    def worker(payment):
        try:
            return payment, get_transactions(program_id, payment["paymentId"]), None
        except Exception as e:
            logger.warning(
                f"[!] Failed to fetch transactions for paymentId "
                f"{payment['paymentId']}: {e}"
            )
            return payment, [], e

    counts = {"total": 0, "not_waiting": 0, "not_excel": 0, "deleted": 0,
              "missing_ids": 0, "kept": 0}
    fetch_errors = []

    max_workers = min(MAX_WORKERS, len(payments)) or 1
    with ThreadPoolExecutor(max_workers=max_workers) as executor:
        futures = [executor.submit(worker, p) for p in payments]
        for future in as_completed(futures):
            payment, transactions, error = future.result()

            if error is not None:
                fetch_errors.append((payment["paymentId"], error))

            per_payment = {
                "paymentId": payment["paymentId"],
                "name": payment.get("name"),
                "expectedWaiting": payment.get("waitingCount"),
                "fetchedTotal": len(transactions),
                "waitingSeen": 0,
                "notExcel": 0,
                "deleted": 0,
                "missingIds": 0,
                "kept": 0,
                "error": str(error) if error else None,
            }
            for t in transactions:
                if not isinstance(t, dict):
                    continue
                counts["total"] += 1

                status = (t.get("status") or t.get("transactionStatus") or "").lower()
                if status != "waiting":
                    counts["not_waiting"] += 1
                    continue

                per_payment["waitingSeen"] += 1

                # INVARIANT 9. The FSP configuration is a property of the
                # TRANSACTION, not the payment: 121's transaction_view selects
                # programFspConfigurationName and fspName per row, joined via
                # the transaction's last event. A payment can therefore mix
                # Excel and api recipients, and stamping the Excel payment's
                # column onto an api transaction would reconcile the wrong
                # person against the wrong field.
                #
                # waitingSeen is incremented ABOVE this check on purpose: it is
                # reconciled against 121's waiting aggregation, which counts api
                # transactions too. Filtering before counting would turn every
                # mixed payment into a spurious shortfall.
                if not _is_excel_config(
                    t.get("programFspConfigurationName"),
                    fsp_types,
                    t.get("fspName"),
                ):
                    counts["not_excel"] += 1
                    per_payment["notExcel"] += 1
                    continue

                if (t.get("registrationStatus") or "").lower() == "deleted":
                    counts["deleted"] += 1
                    per_payment["deleted"] += 1
                    continue

                uuid = t.get("registrationReferenceId")
                reg_id = t.get("registrationId")
                if not uuid or not reg_id:
                    counts["missing_ids"] += 1
                    per_payment["missingIds"] += 1
                    continue

                payment_id = t.get("paymentId", payment["paymentId"])

                tranche = {
                    "uuid": uuid,
                    # Kept so the client importer can read either shape.
                    "registrationReferenceId": uuid,
                    "registrationId": reg_id,
                    "transactionId": t.get("id"),
                    "paymentId": payment_id,
                    "paymentName": payment["name"],
                    "paymentDate": payment.get("paymentDate"),
                    "amount": _transfer_value(t),
                    "created": t.get("created"),
                    "status": status,
                    # INVARIANT 6: self-describing. The device reads the column
                    # off the tranche the FSP selected — it never has to infer
                    # one for the program.
                    #
                    # Resolved from THIS transaction's own FSP configuration
                    # (INVARIANT 2, one scope finer than before), falling back
                    # to the payment-level value when 121 reports no
                    # configuration on the row — possible when the config was
                    # deleted after the payment was made.
                    "columnToMatch": (
                        col_map.get(t.get("programFspConfigurationName"))
                        or payment.get("columnToMatch")
                    ),
                    "fspConfigName": t.get("programFspConfigurationName"),
                    "fspConfigNames": payment.get("fspConfigNames"),
                }
                # One tranche per (uuid, paymentId) — matches the compound
                # IndexedDB key on the client. If 121 ever returns more than one
                # transaction for the same pair, keep the most recent.
                existing = tranches_by_uuid[uuid].get(payment_id)
                if existing and str(existing.get("created") or "") > str(t.get("created") or ""):
                    continue

                if not existing:
                    counts["kept"] += 1
                    per_payment["kept"] += 1

                tranches_by_uuid[uuid][payment_id] = tranche
                reg_id_by_uuid[uuid] = reg_id

            reconciliation.append(per_payment)

    logger.info("[INFO] Transaction scan: " + ", ".join(f"{k}={v}" for k, v in counts.items()))

    # --- per-payment reconciliation against the 121 aggregation ---
    reconciliation.sort(key=lambda r: r["paymentId"])
    shortfalls = []
    for r in reconciliation:
        expected = r["expectedWaiting"]
        logger.info(
            "[INFO] paymentId=%s: expected waiting=%s, waiting received=%s, "
            "kept=%s (notExcel=%s, deleted=%s, missingIds=%s)",
            r["paymentId"], expected, r["waitingSeen"], r["kept"],
            r["notExcel"], r["deleted"], r["missingIds"],
        )
        if r["error"]:
            shortfalls.append(
                f"paymentId={r['paymentId']}: transaction fetch failed ({r['error']})"
            )
        elif expected is not None and r["waitingSeen"] < expected:
            shortfalls.append(
                f"paymentId={r['paymentId']}: received {r['waitingSeen']} waiting "
                f"transaction(s) but 121 reports {expected}"
            )

    # Backstop for a tranche that passed the Excel filter but still has no
    # column — e.g. its FSP configuration was deleted in 121 after the payment
    # was made, so neither the transaction nor the payment could supply one.
    # Reconciling it is impossible, so refuse rather than ship it.
    uncolumned = [
        (tr["uuid"], tr["paymentId"], tr.get("fspConfigName"))
        for by_payment in tranches_by_uuid.values()
        for tr in by_payment.values()
        if not tr.get("columnToMatch")
    ]
    if uncolumned:
        shortfalls.append(
            f"{len(uncolumned)} tranche(s) have no columnToMatch and cannot be "
            f"reconciled, e.g. {uncolumned[:5]}"
        )

    if shortfalls:
        _fail(
            "Transaction reconciliation failed — the batch would be missing "
            "beneficiaries:\n  - " + "\n  - ".join(shortfalls)
        )

    # Flatten to {uuid: [tranche, ...]} sorted oldest payment first, so the
    # order the FSP sees on screen is stable between syncs.
    flattened = {}
    for uuid, by_payment in tranches_by_uuid.items():
        tranches = sorted(
            by_payment.values(),
            key=lambda x: (str(x.get("paymentDate") or x.get("created") or ""), x["paymentId"]),
        )
        flattened[uuid] = tranches

    multi = sum(1 for v in flattened.values() if len(v) > 1)
    logger.info(
        f"[INFO] {len(flattened)} beneficiar(y/ies) with open payments; "
        f"{multi} of them have more than one"
    )

    return flattened, reg_id_by_uuid, reconciliation


def download_open_payments_cache(program_id):
    """
    Build the offline batch, multi-payment aware.

    1) Ask 121 which payments still have waiting transactions (and their names)
    2) Fetch those payments' transactions in parallel — ALL pages of them
    3) Group into a list of tranches per beneficiary
    4) Fetch each registration and photo ONCE per beneficiary, not per tranche
    5) Reconcile what we built against what 121 said existed
    6) Write registrations_cache.json, transactions.json, batch_info.json

    The batch directory keeps the historical "payment-recent-batch-N" name
    because /submit-payments and /api/offline/latest.zip select batches by that
    prefix.
    """
    base_path = "offline-cache"
    os.makedirs(base_path, exist_ok=True)
    batch_dir = get_next_batch_dir(base_path, "recent")
    photos_dir = os.path.join(batch_dir, "photos")

    # columnToMatch is per FSP configuration, reached via the payment — never
    # per program. Build the name->column map once, then resolve each payment
    # against it inside select_open_payments().
    col_map, fsp_types = build_fsp_column_map(program_id)

    # 1) + 2) + 3)
    open_payments, skipped_non_excel = select_open_payments(
        program_id, col_map, fsp_types
    )
    if not open_payments:
        logger.warning("[!] No open payments for program %s — nothing to cache.", program_id)

    tranches_by_uuid, reg_id_by_uuid, reconciliation = fetch_tranches(
        program_id, open_payments, col_map, fsp_types
    )

    if not tranches_by_uuid:
        logger.warning("[!] No waiting transactions found — writing an empty batch.")

    # 4) Fetch registrations once per beneficiary
    registrations_map, failed_reg_ids = fetch_registrations_bulk(
        program_id, list(reg_id_by_uuid.values())
    )

    cache_data = []
    all_tranches = []
    missing_registration_uuids = []
    missing_match_values = []   # [(uuid, column)]
    lookup_values = defaultdict(list)   # normalised value -> [uuid, ...]
    lookup_blank = []                   # uuids with no lookup value at all

    for uuid, tranches in tranches_by_uuid.items():
        reg_id = reg_id_by_uuid.get(uuid)
        reg = registrations_map.get(reg_id)
        if not reg:
            logger.warning(f"[!] Failed registration fetch for {reg_id} (uuid {uuid})")
            missing_registration_uuids.append(uuid)
            continue

        filtered_data = {key: reg.get(key) for key in FIELD_KEYS}

        # The configured lookup attribute. Cached even when it is not one of the
        # displayed fields — the scan box needs it, the beneficiary screen only
        # renders what display config lists, so nothing extra appears on screen.
        if LOOKUP_FIELD != "referenceId":
            lookup_raw = reg.get(LOOKUP_FIELD)
            filtered_data[LOOKUP_FIELD] = lookup_raw
            lookup_clean = str(lookup_raw or "").strip()
            if lookup_clean:
                lookup_values[lookup_clean.lower()].append(uuid)
            else:
                lookup_blank.append(uuid)

        # The union of columns this beneficiary's tranches need — they may sit
        # on different FSP configurations within the same program.
        match_cols = {t["columnToMatch"] for t in tranches if t.get("columnToMatch")}
        for col in match_cols:
            raw = reg.get(col)
            filtered_data[col] = raw
            # encrypt_data() casts None to "" and encrypts it, producing a valid
            # Fernet token that decrypts to empty. Once encrypted, "no value" is
            # indistinguishable from a real one and the device only finds out
            # mid-distribution. Catch it here, while it is still fixable.
            if not str(raw or "").strip():
                missing_match_values.append((uuid, col))

        encrypted_data = encrypt_data(filtered_data)

        first = tranches[0]

        record = {
            "uuid": uuid,
            "registrationId": reg_id,
            "photo_filename": f"{uuid}.enc",
            # Full tranche list — this is the authoritative payment data now.
            "payments": tranches,
            "paymentCount": len(tranches),
            # Legacy scalars, kept so an older client build (or the
            # single-tranche fallback in beneficiary_offline.html) still renders
            # something sane. They describe the FIRST open payment only and must
            # not be treated as "the" payment when paymentCount > 1.
            "paymentId": first["paymentId"],
            "amount": first["amount"],
            # Every record here has at least one waiting, non-deleted tranche,
            # so it is valid by construction.
            "valid": True,
            "reason": "ok",
            "data": encrypted_data,
            # Explicit encryption metadata — app.py's reconciliation reads
            # record["dataEncrypted"] to decide whether to decrypt, instead of
            # guessing. Do not remove.
            "dataEncrypted": True,
            "encryptionScheme": "fernet-v1",
        }

        cache_data.append(record)
        all_tranches.extend(tranches)

    # 5) Reconcile the finished batch against 121 before anyone relies on it.
    expected_beneficiaries = len(tranches_by_uuid)
    expected_tranches = sum(len(v) for v in tranches_by_uuid.values())

    if missing_registration_uuids:
        _fail(
            f"{len(missing_registration_uuids)} beneficiar(y/ies) were dropped "
            f"because their registration could not be fetched "
            f"(registrationIds: {failed_reg_ids[:20]}). Batch would be incomplete."
        )

    if len(cache_data) != expected_beneficiaries:
        _fail(
            f"Built {len(cache_data)} record(s) but expected "
            f"{expected_beneficiaries}."
        )

    if missing_match_values:
        by_col = defaultdict(list)
        for uuid, col in missing_match_values:
            by_col[col].append(uuid)
        detail = "; ".join(
            f"'{col}': {len(ids)} beneficiar(y/ies) e.g. {ids[:5]}"
            for col, ids in sorted(by_col.items())
        )
        _fail(
            "Some beneficiaries have no value for their payment's match column "
            f"in 121, so their payments could not be reconciled — {detail}"
        )
    # --- lookup-field health check -------------------------------------
    # Deliberately NOT _fail(): every person here is still findable by their
    # reference ID, and scan.html refuses to resolve an ambiguous lookup value
    # rather than guessing. This is a data-quality signal, not a broken batch.
    lookup_duplicates = {v: ids for v, ids in lookup_values.items() if len(ids) > 1}
    if LOOKUP_FIELD != "referenceId":
        if lookup_duplicates:
            affected = sum(len(ids) for ids in lookup_duplicates.values())
            logger.warning(
                "[!] Lookup field '%s' is NOT unique in this batch: %s value(s) "
                "shared by %s beneficiar(y/ies). Those people can only be found "
                "by reference ID. Fix the duplicates in 121.",
                LOOKUP_FIELD, len(lookup_duplicates), affected,
            )
        if lookup_blank:
            logger.warning(
                "[!] %s beneficiar(y/ies) have no value for lookup field '%s' "
                "and cannot be found by it (e.g. %s).",
                len(lookup_blank), LOOKUP_FIELD, lookup_blank[:5],
            )
        if not lookup_duplicates and not lookup_blank:
            logger.info(
                "[INFO] Lookup field '%s': unique and populated across all %s "
                "beneficiar(y/ies).", LOOKUP_FIELD, len(cache_data),
            )
    # 6) Download & encrypt photos in parallel (one per beneficiary)
    photo_failures = download_photos_bulk(cache_data, photos_dir)

    json_path = os.path.join(batch_dir, "registrations_cache.json")
    with open(json_path, "w", encoding="utf-8") as f:
        json.dump(cache_data, f, indent=2)

    # Flat list of every open tranche — the client imports this into its
    # 'transaction' store keyed on [uuid, paymentId].
    tx_path = os.path.join(batch_dir, "transactions.json")
    with open(tx_path, "w", encoding="utf-8") as f:
        json.dump(all_tranches, f, indent=2)

    # Per-payment map is the authoritative artifact now. The scalar
    # "columnToMatch" is kept ONLY so a device still running the old client can
    # read something sane; it is populated only when every payment in the batch
    # agrees, and is otherwise null rather than an arbitrary pick.
    column_by_payment = {
        str(p["paymentId"]): p.get("columnToMatch")
        for p in open_payments
        if p.get("columnToMatch")
    }
    _distinct_columns = set(column_by_payment.values())
    _legacy_column = (
        next(iter(_distinct_columns)) if len(_distinct_columns) == 1 else None
    )

    batch_info = {
        "batchType": "payment-recent",
        "multiPayment": True,
        "programId": program_id,
        "columnToMatch": _legacy_column,
        "columnToMatchByPayment": column_by_payment,
        "fspColumnMap": col_map,
        # INVARIANT 6, applied to the skip: whoever reads this manifest can see
        # exactly which payments 121 Scan declined to distribute, and why,
        # without re-deriving it from 121.
        "fspTypes": fsp_types,
        "skippedNonExcelPayments": skipped_non_excel,
        "nonExcelTransactionsExcluded": sum(
            r.get("notExcel", 0) for r in reconciliation
        ),
        "dataEncrypted": True,
        "encryptionScheme": "fernet-v1",
        "windowDays": WINDOW_DAYS,
        "recordCount": len(cache_data),
        "trancheCount": len(all_tranches),
        # Completeness evidence — a batch that cannot prove it is whole should
        # not be trusted in the field.
        "expectedBeneficiaryCount": expected_beneficiaries,
        "expectedTrancheCount": expected_tranches,
        "registrationFetchFailures": len(failed_reg_ids),
        "photoFailures": len(photo_failures),
        "missingMatchValues": len(missing_match_values),
        # Lookup-field evidence. Deliberately NOT part of "complete" below: a
        # non-unique or partly-blank lookup field does not make the batch
        # incomplete — everyone in it is still findable by reference ID.
        "lookupField": LOOKUP_FIELD,
        "lookupBlankCount": len(lookup_blank),
        "lookupDuplicateValueCount": len(lookup_duplicates),
        "lookupDuplicateAffected": sum(len(ids) for ids in lookup_duplicates.values()),
        "complete": (
            len(cache_data) == expected_beneficiaries
            and len(all_tranches) == expected_tranches
            and not failed_reg_ids
            and not missing_match_values
            and all(p.get("columnToMatch") for p in open_payments)
        ),
        "reconciliation": reconciliation,
        "payments": [
            {
                "paymentId": p["paymentId"],
                "name": p["name"],
                "paymentDate": p.get("paymentDate"),
                "waitingCount": p.get("waitingCount"),
            }
            for p in open_payments
        ],
        "generatedAt": datetime.utcnow().isoformat() + "Z",
    }
    with open(os.path.join(batch_dir, "batch_info.json"), "w", encoding="utf-8") as f:
        json.dump(batch_info, f, indent=2)

    logger.info(f"\n[OK] Batch saved to: {batch_dir}")
    if photo_failures:
        logger.warning("[!] %s beneficiar(y/ies) have no photo in this batch.",
                       len(photo_failures))
    # Keep this exact phrasing: app.py /sync-fsp scans stdout for a line
    # containing "beneficiaries" and shows it to the FSP user.
    logger.info(
        f"{len(cache_data)} beneficiaries ready for offline validation "
        f"({len(all_tranches)} open payments in total)."
    )
    return len(cache_data)


# Backwards-compatible alias: anything still calling the old name gets the new,
# multi-payment-aware behaviour.
download_recent_payments_cache = download_open_payments_cache


# ----------------------------------------------------------------------
# CLI ENTRY
# ----------------------------------------------------------------------

if __name__ == "__main__":
    # Default behaviour: generate the multi-payment "recent" batch
    download_open_payments_cache(PROGRAM_ID)