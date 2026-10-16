import json
import os

from cryptography.fernet import Fernet, InvalidToken

# Optional: load a local .env file during development so the env-managed fields
# below can be set without exporting them in your shell. On Azure these values
# come from App Settings, so this block is a harmless no-op there (and silently
# skips if python-dotenv isn't installed).
try:
    from dotenv import load_dotenv
    load_dotenv()
except Exception:
    pass


# ---------------------------------------------------------------------------
# Fields read from environment variables instead of system_config.json.
#
#       config key      ->   ENV VAR NAME
#
# Only list fields that are actually set in Azure App Settings. save_config()
# strips every listed key from system_config.json whether or not the env var
# exists, so listing a field Azure does not set silently deletes it from the
# JSON on the next save (this previously wiped programTitle, programCurrency
# and nationalSociety).
#
# Anything NOT listed here stays in system_config.json: KOBO_SERVER,
# KOBO_TOKEN, PROGRAMS, programTitle, programCurrency, nationalSociety, and
# COLUMN_TO_MATCH_PER_PROGRAM.
#
# To move a field in or out of env management, edit this dict AND add or
# remove the matching App Setting on every instance.
# ---------------------------------------------------------------------------
ENV_MANAGED_FIELDS = {
    "url121":         "URL_121",
    "username121":    "USERNAME_121",
    "password121":    "PASSWORD_121",
    "ENCRYPTION_KEY": "ENCRYPTION_KEY",
}


# ---------------------------------------------------------------------------
# Secrets encrypted at rest inside system_config.json.
#
# These are values a user sets through the System Config page, so they cannot
# live in env vars — but they should not sit in plaintext on the App Service
# volume either. They are Fernet-encrypted on save and decrypted on load.
#
# The key (ENCRYPTION_KEY) is env-managed and stripped by save_config(), so it
# is never written to disk beside the ciphertext it protects.
#
# Decryption happens in load_config() — ONE choke point. Every consumer
# (app.py's Kobo calls, offline_sync.py, anything added later) receives
# plaintext without needing to know encryption exists. Do not decrypt at the
# call sites: that is how a value ends up encrypted in one code path and
# plaintext in another.
# ---------------------------------------------------------------------------
SECRET_FIELDS = {"KOBO_TOKEN"}

# Fields that an earlier release (main, 2026-07-14 "encrypt kobo credentials
# before saving") encrypted on save but which are no longer secret. They are
# DECRYPTED on load so instances that saved under that release keep working,
# but never re-encrypted: the next save_config() writes them as plaintext,
# which retires the legacy ciphertext. Once every instance has re-saved its
# System Config, this set can be emptied.
LEGACY_ENCRYPTED_FIELDS = {"KOBO_SERVER"}

# Fernet tokens always begin with this. Used to tell an encrypted value from a
# plaintext one so that both keep working: a token typed straight into the JSON
# by hand, or saved before encryption existed, must not be mangled — and a
# re-save must never double-encrypt.
_FERNET_PREFIX = "gAAAA"


def _get_paths():
    env = os.getenv("SCANDROID_ENV", "local")
    context = os.getenv("SCANDROID_CONTEXT", "local")

    if env == "azure":
        base = f"/home/site/configs/{context}"
    else:
        base = os.path.join(os.path.dirname(__file__), "configs", context)

    os.makedirs(base, exist_ok=True)
    return (
        os.path.join(base, "system_config.json"),
        os.path.join(base, "display_config.json"),
    )


def _fernet():
    """Cipher built from the env-managed ENCRYPTION_KEY, or None if unusable."""
    key = os.getenv("ENCRYPTION_KEY")
    if not key:
        return None
    try:
        return Fernet(key.encode())
    except Exception:
        return None


def _looks_encrypted(value):
    return isinstance(value, str) and value.startswith(_FERNET_PREFIX)


def _decrypt_secrets(data):
    """Replace encrypted secrets with their plaintext, in place.

    Raises rather than swallowing. A silently-undecrypted secret is exactly the
    failure that took photo sync down: every Kobo request went out with
    "Authorization: Token gAAAAAB..." and came back 401, while the sync reported
    success and the UI showed a green tick. Fail loudly and name the field.

    Also decrypts LEGACY_ENCRYPTED_FIELDS, with the same fail-loud rule: an
    undecrypted KOBO_SERVER would send every Kobo call to "gAAAA.../api/v2/...".
    """
    cipher = None
    for key in SECRET_FIELDS | LEGACY_ENCRYPTED_FIELDS:
        value = data.get(key)
        if not _looks_encrypted(value):
            continue  # plaintext (hand-edited, or pre-encryption) — leave as-is

        if cipher is None:
            cipher = _fernet()
        if cipher is None:
            raise RuntimeError(
                f"{key} is encrypted in system_config.json but ENCRYPTION_KEY is "
                "missing or invalid, so it cannot be decrypted. Set ENCRYPTION_KEY "
                "in the app settings."
            )

        try:
            data[key] = cipher.decrypt(value.encode()).decode()
        except InvalidToken:
            raise RuntimeError(
                f"{key} could not be decrypted with the current ENCRYPTION_KEY. "
                "The key was probably rotated after this value was saved. "
                "Re-enter the value on the System Config page to store it under "
                "the new key."
            )
    return data


def _encrypt_secrets(data):
    """Encrypt secret fields prior to writing. Never double-encrypts.

    With no usable key the values are left as they are — better a readable
    config than one encrypted under a key nobody has.

    Only SECRET_FIELDS are encrypted. LEGACY_ENCRYPTED_FIELDS are deliberately
    written back as plaintext.
    """
    cipher = None
    for key in SECRET_FIELDS:
        value = data.get(key)
        if not isinstance(value, str) or not value or _looks_encrypted(value):
            continue

        if cipher is None:
            cipher = _fernet()
        if cipher is None:
            continue

        data[key] = cipher.encrypt(value.encode()).decode()
    return data


def _apply_env_overrides(data):
    """Overlay env-managed fields onto the JSON data.

    The environment variable wins whenever it is set; otherwise any existing
    value already in the JSON is kept as a fallback. This lets Azure drive
    everything from App Settings while local dev can rely on a .env file (or,
    until the next save, the values still sitting in the JSON on disk).
    """
    for key, env_name in ENV_MANAGED_FIELDS.items():
        value = os.getenv(env_name)
        if value is not None:
            data[key] = value
    return data


def load_config():
    system_path, _ = _get_paths()
    data = {}
    if os.path.exists(system_path):
        with open(system_path, "r", encoding="utf-8") as f:
            data = json.load(f)
    # Env overrides FIRST: ENCRYPTION_KEY has to be in place before anything
    # can be decrypted with it.
    return _decrypt_secrets(_apply_env_overrides(data))


def save_config(data):
    """Persist only the JSON-resident fields.

    Env-managed fields are stripped before writing, so values that come from
    the environment (including secrets like password121 and ENCRYPTION_KEY) are
    never baked back into system_config.json by a UI save or a runtime write.

    Fields in SECRET_FIELDS are encrypted on the way out. Callers pass plaintext
    (that is what load_config gave them) and never need to think about it.
    """
    system_path, _ = _get_paths()
    to_save = {k: v for k, v in data.items() if k not in ENV_MANAGED_FIELDS}
    to_save = _encrypt_secrets(dict(to_save))
    with open(system_path, "w", encoding="utf-8") as f:
        json.dump(to_save, f, indent=2, ensure_ascii=False)


def load_display_config():
    _, display_path = _get_paths()
    if not os.path.exists(display_path):
        return {}
    with open(display_path, "r", encoding="utf-8") as f:
        return json.load(f)


def save_display_config(data):
    _, display_path = _get_paths()
    with open(display_path, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2, ensure_ascii=False)
