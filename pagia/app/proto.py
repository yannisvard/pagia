"""Επικοινωνία με την εφαρμογή Ηλεκτρονικού Πρωτοκόλλου.

Όλες οι κλήσεις είναι σύγχρονες και με ρητό timeout: αν το πρωτόκολλο
δεν απαντήσει, η Παγία δεν κρεμάει και ΔΕΝ αποθηκεύει μισοτελειωμένη
κατάσταση. Καλύτερα καμία πρωτοκόλληση παρά αριθμός στο κενό.
"""

import os

import httpx

# Οι τιμές έρχονται από το .env στην εκκίνηση, αλλά μπορούν να αλλάξουν
# από τις Ρυθμίσεις χωρίς Build — γι' αυτό ζουν σε μεταβλητή, όχι σε
# σταθερά. Το `configure()` καλείται από το main.
PROTO_URL = os.environ.get("PROTO_URL", "http://host.docker.internal:5500").rstrip("/")
PROTO_KEY = os.environ.get("PROTO_API_KEY", "").strip()
TIMEOUT = float(os.environ.get("PROTO_TIMEOUT", "15"))
RECIPIENT = os.environ.get("PROTO_RECIPIENT", "Ι.ΝΕ.ΔΙ.ΒΙ.Μ.")


def configure(url: str = None, key: str = None, recipient: str = None):
    global PROTO_URL, PROTO_KEY, RECIPIENT
    if url:
        PROTO_URL = url.strip().rstrip("/")
    if key:
        PROTO_KEY = key.strip()
    if recipient:
        RECIPIENT = recipient.strip()


class ProtoError(Exception):
    """Σφάλμα επικοινωνίας ή απόρριψης από το πρωτόκολλο."""


def _headers():
    if not PROTO_KEY:
        raise ProtoError("Δεν έχει οριστεί PROTO_API_KEY στις ρυθμίσεις.")
    return {"X-API-Key": PROTO_KEY}


def _net_error(e) -> str:
    """Μετάφραση σφαλμάτων δικτύου σε κάτι που εξηγεί τι να δοκιμάσεις."""
    txt = str(e)
    if "No route to host" in txt or "Network is unreachable" in txt:
        return (f"Ο container δεν φτάνει στο {PROTO_URL}. Αν έβαλες "
                f"διεύθυνση Tailscale (100.x.x.x), δεν δουλεύει μέσα από "
                f"container — δοκίμασε http://host.docker.internal:5500 ή "
                f"την τοπική IP του NAS.")
    if "Connection refused" in txt:
        return (f"Η διεύθυνση {PROTO_URL} απαντά αλλά η θύρα είναι κλειστή. "
                f"Τρέχει το πρωτόκολλο; Είναι σωστή η θύρα;")
    if "Name or service not known" in txt or "getaddrinfo" in txt:
        return (f"Το όνομα στο {PROTO_URL} δεν αναγνωρίζεται.")
    if "timed out" in txt.lower():
        return f"Λήξη χρόνου αναμονής από το {PROTO_URL}."
    return f"Δεν υπάρχει επικοινωνία με το πρωτόκολλο: {e}"


def _check(r):
    if r.status_code == 401:
        raise ProtoError("Το πρωτόκολλο απέρριψε το API key.")
    if r.status_code == 404:
        raise ProtoError("Δεν βρέθηκε η εγγραφή στο πρωτόκολλο.")
    if r.status_code >= 400:
        raise ProtoError(f"Το πρωτόκολλο απάντησε {r.status_code}: {r.text[:200]}")
    return r.json()


def register(subject: str, ext_id: str, party: str = None) -> dict:
    """Ζητά αριθμό πρωτοκόλλου. Ιδεμποτεντικό μέσω του ext_id."""
    try:
        with httpx.Client(timeout=TIMEOUT) as c:
            r = c.post(f"{PROTO_URL}/api/protokollo", headers=_headers(),
                       data={"subject": subject,
                             "party": party or RECIPIENT,
                             "direction": "out",
                             "ext_source": "pagia",
                             "ext_id": ext_id,
                             "notify": "false"})
    except httpx.RequestError as e:
        raise ProtoError(_net_error(e))
    return _check(r)


def fetch(year: int, number: int) -> dict:
    try:
        with httpx.Client(timeout=TIMEOUT) as c:
            r = c.get(f"{PROTO_URL}/api/entries/{year}/{number}",
                      headers=_headers())
    except httpx.RequestError as e:
        raise ProtoError(_net_error(e))
    return _check(r)


def attach(year: int, number: int, filename: str, content: bytes) -> dict:
    try:
        with httpx.Client(timeout=TIMEOUT * 2) as c:
            r = c.post(f"{PROTO_URL}/api/entries/{year}/{number}/files",
                       headers=_headers(),
                       files={"files": (filename, content)})
    except httpx.RequestError as e:
        raise ProtoError(_net_error(e))
    return _check(r)


def cancel(year: int, number: int, reason: str) -> dict:
    try:
        with httpx.Client(timeout=TIMEOUT) as c:
            r = c.post(f"{PROTO_URL}/api/entries/{year}/{number}/cancel",
                       headers=_headers(), data={"reason": reason})
    except httpx.RequestError as e:
        raise ProtoError(_net_error(e))
    return _check(r)


def status() -> tuple[bool, str]:
    """Γρήγορος έλεγχος για την οθόνη ρυθμίσεων."""
    if not PROTO_KEY:
        return False, "Δεν έχει οριστεί PROTO_API_KEY"
    try:
        with httpx.Client(timeout=5) as c:
            r = c.get(f"{PROTO_URL}/api/entries/1900/1", headers=_headers())
        if r.status_code == 401:
            return False, "Λάθος API key"
        if r.status_code in (200, 404):
            return True, "Σύνδεση εντάξει"
        return False, f"Απάντησε {r.status_code}"
    except httpx.RequestError as e:
        return False, _net_error(e)
