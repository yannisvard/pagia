"""Παγία Προκαταβολή — 2η Σ.Α.Ε.Κ. Ηρακλείου.

Διαχείριση παραστατικών, προμηθευτών και μηνιαίων εγγράφων, με αυτόματη
πρωτοκόλληση μέσω της εφαρμογής Ηλεκτρονικού Πρωτοκόλλου.
"""

import csv
import io
import os
import zipfile
from datetime import date, datetime
from urllib.parse import quote

from fastapi import Depends, FastAPI, File, Form, HTTPException, Request, UploadFile
from fastapi.responses import HTMLResponse, RedirectResponse, Response
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
from passlib.context import CryptContext
from sqlalchemy import func, or_
from starlette.middleware.sessions import SessionMiddleware

import docgen
import proto
from db import (DOY_LIST, MONTHS_NOM, OVERRIDE_AUTO, OVERRIDE_NONE,
                PRAKTIKO_TYPES,
                Department, DocType, ExpenseCode,
                Invoice, InvoiceFile, MonthlyBatch, PaymentMethod, Setting,
                Supplier, User, get_db, get_setting, init_db, set_setting)

app = FastAPI(title="Παγία Προκαταβολή")
app.add_middleware(SessionMiddleware,
                   secret_key=os.environ.get("SECRET_KEY", "change-me"),
                   max_age=60 * 60 * 12)
app.mount("/static", StaticFiles(directory="static"), name="static")
templates = Jinja2Templates(directory="templates")
pwd = CryptContext(schemes=["bcrypt"], deprecated="auto")

SCHOOL = os.environ.get("SCHOOL_NAME", "2η Σ.Α.Ε.Κ. Ηρακλείου")
FILES_DIR = os.path.join(os.environ.get("DATA_DIR", "/data"), "files")
os.makedirs(FILES_DIR, exist_ok=True)


def save_upload(inv, up) -> InvoiceFile:
    """Αποθήκευση συνημμένου με ονομασία που δείχνει σε ποιο παραστατικό
    ανήκει, ώστε ο φάκελος να διαβάζεται και χωρίς τη βάση."""
    import re as _re
    import uuid
    year = (inv.doc_date or date.today()).year
    folder = os.path.join(FILES_DIR, str(year))
    os.makedirs(folder, exist_ok=True)
    ext = os.path.splitext(up.filename)[1][:10]
    safe_num = _re.sub(r"[^\w\-.]", "_", inv.number or "χωρίς-αριθμό")
    safe_sup = _re.sub(r"[^\w\-. ]", "", inv.supplier.name)[:35].strip()
    stem = f"{safe_num}_{safe_sup}_{uuid.uuid4().hex[:6]}"
    stored = stem + ext
    content = up.file.read() if hasattr(up, "file") else b""
    with open(os.path.join(folder, stored), "wb") as fh:
        fh.write(content)
    return InvoiceFile(invoice_id=inv.id, stored_name=f"{year}/{stored}",
                       original_name=up.filename, size=len(content))


# ------------------------------------------------------------------ βοηθητικά

DOCX_MIME = ("application/vnd.openxmlformats-officedocument."
             "wordprocessingml.document")
XLSX_MIME = ("application/vnd.openxmlformats-officedocument."
             "spreadsheetml.sheet")


def download(data: bytes, filename: str, mime: str) -> Response:
    """Απάντηση λήψης αρχείου.

    Τα ελληνικά ονόματα δεν χωρούν σε HTTP header (latin-1 μόνο), οπότε
    κωδικοποιούνται κατά RFC 5987 και δίνεται και απλό ASCII εφεδρικό.
    """
    ascii_name = filename.encode("ascii", "ignore").decode() or "document"
    return Response(data, media_type=mime, headers={
        "Content-Disposition":
            f'attachment; filename="{ascii_name}"; '
            f"filename*=UTF-8''{quote(filename)}"})


def current_user(request: Request, db):
    uid = request.session.get("uid")
    return db.get(User, uid) if uid else None


def require_login(request: Request, db):
    u = current_user(request, db)
    if not u:
        raise HTTPException(303, headers={"Location": "/login"})
    return u


def render(request, name, db, **ctx):
    ctx.setdefault("user", current_user(request, db))
    ctx.setdefault("school", SCHOOL)
    ctx.setdefault("months", MONTHS_NOM)
    ctx.setdefault("flash", request.session.pop("flash", None))
    ctx.setdefault("error", request.session.pop("error", None))
    return templates.TemplateResponse(request, name, ctx)


def flash(request, msg, error=False):
    request.session["error" if error else "flash"] = msg


def parse_date(s):
    if not s:
        return None
    for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y"):
        try:
            return datetime.strptime(s.strip(), fmt).date()
        except ValueError:
            continue
    return None


def as_int(v, default=None):
    """Μετατροπή παραμέτρου διεύθυνσης σε ακέραιο.

    Οι λίστες επιλογής στέλνουν κενό όταν διαλέξεις «Όλοι» ή «Όλο το
    έτος». Αν το πεδίο δηλωθεί ως int, το FastAPI απαντά 422 αντί να το
    θεωρήσει «χωρίς φίλτρο».
    """
    try:
        return int(str(v).strip())
    except (TypeError, ValueError):
        return default


def parse_float(s):
    if s is None or str(s).strip() == "":
        return 0.0
    return float(str(s).replace("€", "").replace(".", "").replace(",", ".").strip()) \
        if "," in str(s) else float(str(s).replace("€", "").strip())


ORG_FIELDS = [
    ("org_school", "ΣΧΟΛΗ", "Επωνυμία σχολής", "2η Σ.Α.Ε.Κ. Ηρακλείου"),
    ("org_school_gen", "ΣΧΟΛΗ_ΓΕΝΙΚΗ", "Επωνυμία σε γενική (π.χ. «της …»)",
     "της 2ης Σ.Α.Ε.Κ. Ηρακλείου"),
    ("org_place", "ΤΟΠΟΣ", "Τόπος (για την ημερομηνία)", "Ηράκλειο Κρήτης"),
    ("org_decision", "ΑΠΟΦΑΣΗ", "Αριθμός απόφασης ορισμού επιτροπής", ""),
    ("org_c1", "ΕΠΙΤΡΟΠΗ_Α", "Επιτροπή — μέλος Α", ""),
    ("org_c2", "ΕΠΙΤΡΟΠΗ_Β", "Επιτροπή — μέλος Β", ""),
    ("org_c3", "ΕΠΙΤΡΟΠΗ_Γ", "Επιτροπή — μέλος Γ", ""),
    ("org_time", "ΩΡΑ", "Ώρα συνεδρίασης επιτροπής", "15:00"),
    ("org_director", "ΔΙΕΥΘΥΝΤΗΣ", "Ονοματεπώνυμο διευθυντή", ""),
    ("org_director_title", "ΤΙΤΛΟΣ_ΔΙΕΥΘΥΝΤΗ", "Ιδιότητα διευθυντή",
     "Ο ΔΙΕΥΘΥΝΤΗΣ ΚΑΙ ΟΙΚΟΝΟΜΙΚΟΣ ΔΙΑΧΕΙΡΙΣΤΗΣ"),
]


def org_context(db) -> dict:
    """Τα στοιχεία του φορέα, όπως μπαίνουν στα πρότυπα.

    Ζουν στη βάση και όχι μέσα στα .docx, ώστε η αλλαγή επιτροπής ή
    απόφασης να μη χρειάζεται επεξεργασία εγγράφου — και ώστε η ίδια
    εφαρμογή να στηθεί σε άλλη σχολή χωρίς να αλλάξει τίποτα στον κώδικα.
    """
    return {ph: get_setting(db, key, default)
            for key, ph, _label, default in ORG_FIELDS}


def month_invoices(db, year, month, only_active=True):
    q = db.query(Invoice).filter(
        func.strftime("%Y", Invoice.doc_date) == str(year),
        func.strftime("%m", Invoice.doc_date) == f"{month:02d}")
    if only_active:
        q = q.filter(Invoice.status == "active")
    return q.order_by(Invoice.doc_date, Invoice.id).all()


def subject_for(inv):
    label = inv.praktiko_label
    return (f"Πρακτικό {label} — {inv.supplier.name} — "
            f"{inv.number}/{inv.doc_date.strftime('%d-%m-%Y')}")


def guess_praktiko_type(label: str) -> str:
    """Πρώτη εκτίμηση τύπου πρακτικού από το λεκτικό της δαπάνης.
    Είναι μόνο αρχική τιμή — διορθώνεται από τις Ρυθμίσεις."""
    l = label.lower()
    if "τραπέζ" in l or "προμήθειες τραπ" in l:
        return ""
    if "εκπαιδευτ" in l or "εκπαίδευσ" in l:
        return "ekpaideysis"
    if "υπηρεσί" in l or "επισκευ" in l or "συντήρησ" in l:
        return "kalis"
    return "ylikwn"


def seed(db):
    """Αρχικά δεδομένα, μόνο σε άδεια βάση."""
    import json
    base = os.environ.get("DATA_DIR", "/data")
    path = os.path.join(base, "seed.json")
    if not db.query(ExpenseCode).count() and os.path.exists(path):
        data = json.load(open(path, encoding="utf-8"))
        for entry in data.get("codes", []):
            label, code = entry[0], entry[1]
            # Τρίτο στοιχείο = ρητός τύπος πρακτικού. Αν λείπει, εικασία.
            ptype = entry[2] if len(entry) > 2 else guess_praktiko_type(label)
            if not db.query(ExpenseCode).filter_by(label=label).first():
                db.add(ExpenseCode(label=label, code=code,
                                   praktiko_type=ptype))
        for entry in data.get("methods", []):
            # Μορφή: [όνομα, χρειάζεται ημερομηνία, χρειάζεται αποδεικτικό]
            if isinstance(entry, str):
                entry = [entry, True, True]
            db.add(PaymentMethod(name=entry[0],
                                 needs_date=bool(entry[1]),
                                 needs_proof=bool(entry[2])))
        for n in data.get("doctypes", []):
            db.add(DocType(name=n))
        db.commit()

    for fname, kind in (("promitheftes.csv", "sup"), ("tmimata.csv", "dep")):
        p = os.path.join(base, fname)
        if not os.path.exists(p):
            continue
        if kind == "sup" and db.query(Supplier).count():
            continue
        if kind == "dep" and db.query(Department).count():
            continue
        with open(p, encoding="utf-8-sig") as f:
            seen = set()
            for row in csv.DictReader(f):
                row = {k.strip(): (v or "").strip() for k, v in row.items()}
                if kind == "sup" and row.get("ΑΦΜ"):
                    # Το ίδιο ΑΦΜ μπορεί να εμφανίζεται δύο φορές με
                    # διαφορετική γραφή επωνυμίας. Κρατάμε την πρώτη.
                    if row["ΑΦΜ"] in seen:
                        continue
                    seen.add(row["ΑΦΜ"])
                    db.add(Supplier(afm=row["ΑΦΜ"], name=row["ΕΠΩΝΥΜΙΑ"],
                                    bank=row.get("ΤΡΑΠΕΖΑ", ""),
                                    iban=row.get("IBAN", ""),
                                    phone=row.get("ΤΗΛΕΦΩΝΟ", ""),
                                    email=row.get("EMAIL", ""),
                                    address=row.get("ΔΙΕΥΘΥΝΣΗ", "")))
                elif kind == "dep" and row.get("ΕΙΔΙΚΟΤΗΤΑ"):
                    key = (row["ΕΙΔΙΚΟΤΗΤΑ"], row.get("ΕΞΑΜΗΝΟ", ""),
                           row.get("ΤΜΗΜΑ", "1"))
                    if key in seen:
                        continue
                    seen.add(key)
                    db.add(Department(specialty=key[0], semester=key[1],
                                      section=key[2]))
        db.commit()


@app.on_event("startup")
def startup():
    init_db()
    db = next(get_db())
    if not db.query(User).count():
        db.add(User(username=os.environ.get("ADMIN_USERNAME", "admin"),
                    password_hash=pwd.hash(
                        os.environ.get("ADMIN_PASSWORD", "allaxeme")),
                    full_name="Διαχειριστής", is_admin=True))
        db.commit()
    seed(db)
    load_proto_config(db)
    db.close()


def load_proto_config(db):
    """Οι ρυθμίσεις της βάσης υπερισχύουν του .env, ώστε να αλλάζουν
    χωρίς Build."""
    proto.configure(get_setting(db, "proto_url"),
                    get_setting(db, "proto_key"),
                    get_setting(db, "proto_recipient"))


# --------------------------------------------------------------------- είσοδος

@app.get("/login", response_class=HTMLResponse)
def login_page(request: Request, db=Depends(get_db)):
    return render(request, "login.html", db)


@app.post("/login")
def login(request: Request, username: str = Form(...),
          password: str = Form(...), db=Depends(get_db)):
    u = db.query(User).filter(User.username == username.strip()).first()
    if not u or not pwd.verify(password, u.password_hash):
        flash(request, "Λάθος όνομα χρήστη ή κωδικός.", error=True)
        return RedirectResponse("/login", status_code=303)
    request.session["uid"] = u.id
    return RedirectResponse("/", status_code=303)


@app.get("/logout")
def logout(request: Request):
    request.session.clear()
    return RedirectResponse("/login", status_code=303)


# ------------------------------------------------------------------- επισκόπηση

@app.get("/")
def index(request: Request, db=Depends(get_db)):
    """Η επισκόπηση καταργήθηκε: ό,τι έδειχνε υπάρχει ήδη στα
    Παραστατικά και στα Έγγραφα μήνα."""
    require_login(request, db)
    return RedirectResponse("/parastatika", status_code=303)


# ----------------------------------------------------------------- παραστατικά

def filtered_invoices(db, year, month, apo, eos, q=""):
    """Το φίλτρο περιόδου, κοινό για τη λίστα και τα συγκεντρωτικά.

    Δύο τρόποι: έτος/μήνας, ή ελεύθερο διάστημα. Αν δοθεί έστω μία
    ημερομηνία, υπερισχύει το διάστημα.
    """
    today = date.today()
    d_apo, d_eos = parse_date(apo), parse_date(eos)
    use_range = bool(d_apo or d_eos)
    query = db.query(Invoice)
    if use_range:
        if d_apo:
            query = query.filter(Invoice.doc_date >= d_apo)
        if d_eos:
            query = query.filter(Invoice.doc_date <= d_eos)
    else:
        year = year or today.year
        query = query.filter(func.strftime("%Y", Invoice.doc_date) == str(year))
        if month:
            query = query.filter(
                func.strftime("%m", Invoice.doc_date) == f"{month:02d}")
    if q:
        query = query.join(Supplier).filter(or_(
            Supplier.name.ilike(f"%{q}%"), Supplier.afm.ilike(f"%{q}%"),
            Invoice.number.ilike(f"%{q}%")))
    invs = query.order_by(Invoice.doc_date.desc(), Invoice.id.desc()).all()
    return invs, use_range, year or today.year


def period_label(year, month, apo, eos, use_range):
    if use_range:
        a, e = parse_date(apo), parse_date(eos)
        if a and e:
            return f"{a.strftime('%d/%m/%Y')} – {e.strftime('%d/%m/%Y')}"
        if a:
            return f"από {a.strftime('%d/%m/%Y')}"
        return f"έως {e.strftime('%d/%m/%Y')}"
    return f"{MONTHS_NOM[month - 1]} {year}" if month else str(year)


@app.get("/synkentrotika", response_class=HTMLResponse)
def summary(request: Request, year: str = "", month: str = "",
            apo: str = "", eos: str = "", db=Depends(get_db)):
    """Σύνολα ανά κωδικό δαπάνης και, για τα υλικά εκπαίδευσης, ανά τμήμα."""
    require_login(request, db)
    year, month = as_int(year), as_int(month)
    invs, use_range, year = filtered_invoices(db, year, month, apo, eos)
    active = [i for i in invs if i.status == "active"]

    by_code, by_dept = {}, {}
    for i in active:
        key = i.expense_code.label if i.expense_code else "— χωρίς κωδικό —"
        row = by_code.setdefault(key, {"n": 0, "net": 0.0, "vat": 0.0,
                                       "gross": 0.0, "code": ""})
        row["n"] += 1
        row["net"] += i.net
        row["vat"] += i.vat
        row["gross"] += i.gross
        if i.expense_code:
            row["code"] = i.expense_code.code

        if i.expense_code and i.expense_code.praktiko_type == "ekpaideysis":
            dk = i.department.label if i.department else "— χωρίς τμήμα —"
            d = by_dept.setdefault(dk, {"n": 0, "net": 0.0, "vat": 0.0,
                                        "gross": 0.0})
            d["n"] += 1
            d["net"] += i.net
            d["vat"] += i.vat
            d["gross"] += i.gross

    total = sum(i.gross for i in active)
    return render(request, "synkentrotika.html", db,
                  year=year, month=month, apo=apo, eos=eos,
                  use_range=use_range,
                  period=period_label(year, month, apo, eos, use_range),
                  by_code=sorted(by_code.items(),
                                 key=lambda kv: -kv[1]["gross"]),
                  by_dept=sorted(by_dept.items(),
                                 key=lambda kv: -kv[1]["gross"]),
                  count=len(active), total=total,
                  net_total=sum(i.net for i in active),
                  vat_total=sum(i.vat for i in active),
                  edu_total=sum(d["gross"] for d in by_dept.values()))


@app.get("/synkentrotika/excel")
def summary_excel(request: Request, year: str = "", month: str = "",
                  apo: str = "", eos: str = "", db=Depends(get_db)):
    require_login(request, db)
    year, month = as_int(year), as_int(month)
    invs, use_range, year = filtered_invoices(db, year, month, apo, eos)
    active = [i for i in invs if i.status == "active"]
    label = period_label(year, month, apo, eos, use_range)
    rows_code, rows_dept = {}, {}
    for i in active:
        k = i.expense_code.label if i.expense_code else "— χωρίς κωδικό —"
        c = i.expense_code.code if i.expense_code else ""
        r = rows_code.setdefault(k, [c, 0, 0.0, 0.0, 0.0])
        r[1] += 1
        r[2] += i.net
        r[3] += i.vat
        r[4] += i.gross
        if i.expense_code and i.expense_code.praktiko_type == "ekpaideysis":
            dk = i.department.label if i.department else "— χωρίς τμήμα —"
            d = rows_dept.setdefault(dk, [0, 0.0, 0.0, 0.0])
            d[0] += 1
            d[1] += i.net
            d[2] += i.vat
            d[3] += i.gross
    fname, data = docgen.build_summary(label, rows_code, rows_dept)
    return download(data, fname, XLSX_MIME)


@app.get("/parastatika", response_class=HTMLResponse)
def invoices_list(request: Request, year: str = "", month: str = "",
                  apo: str = "", eos: str = "", q: str = "",
                  db=Depends(get_db)):
    require_login(request, db)
    year, month = as_int(year), as_int(month)
    invs, use_range, year = filtered_invoices(db, year, month, apo, eos, q)
    years = [r[0] for r in db.query(
        func.strftime("%Y", Invoice.doc_date)).distinct().all() if r[0]]
    active = [i for i in invs if i.status == "active"]
    return render(request, "parastatika.html", db, invoices=invs,
                  year=year, month=month, q=q,
                  apo=apo, eos=eos, use_range=use_range,
                  years=sorted(years, reverse=True),
                  active_count=len(active),
                  total=sum(i.gross for i in active),
                  net_total=sum(i.net for i in active),
                  vat_total=sum(i.vat for i in active),
                  pending=sum(1 for i in active
                              if i.needs_praktiko and not i.protocol_number))


def _form_lists(db):
    return dict(
        codes=db.query(ExpenseCode).filter_by(active=True)
                .order_by(ExpenseCode.label).all(),
        doctypes=db.query(DocType).filter_by(active=True)
                   .order_by(DocType.name).all(),
        methods=db.query(PaymentMethod).filter_by(active=True)
                  .order_by(PaymentMethod.name).all(),
        departments=db.query(Department).filter_by(active=True)
                      .order_by(Department.specialty).all(),
        suppliers=db.query(Supplier).filter_by(active=True)
                    .order_by(Supplier.name).all(),
    )


@app.get("/parastatiko/neo", response_class=HTMLResponse)
def invoice_new_page(request: Request, promitheftis: int = 0,
                     db=Depends(get_db)):
    """Νέο παραστατικό. Με ?promitheftis=<id> ανοίγει με τον προμηθευτή
    ήδη επιλεγμένο — έτσι λειτουργεί το κουμπί από την καρτέλα του."""
    require_login(request, db)
    preset = db.get(Supplier, promitheftis) if promitheftis else None
    return render(request, "parastatiko.html", db, inv=None, preset=preset,
                  praktiko_types=PRAKTIKO_TYPES, **_form_lists(db))


@app.get("/parastatiko/{inv_id}", response_class=HTMLResponse)
def invoice_edit_page(inv_id: int, request: Request, db=Depends(get_db)):
    require_login(request, db)
    inv = db.get(Invoice, inv_id)
    if not inv:
        raise HTTPException(404, "Δεν βρέθηκε το παραστατικό")
    return render(request, "parastatiko.html", db, inv=inv, preset=None,
                  praktiko_types=PRAKTIKO_TYPES, **_form_lists(db))


def _save_invoice(inv, f, db):
    inv.supplier_id = int(f["supplier_id"])
    inv.expense_code_id = int(f["expense_code_id"]) if f.get("expense_code_id") else None
    inv.doc_type_id = int(f["doc_type_id"]) if f.get("doc_type_id") else None
    inv.department_id = int(f["department_id"]) if f.get("department_id") else None
    # Το τμήμα έχει νόημα μόνο στα υλικά εκπαίδευσης. Ο έλεγχος γίνεται και
    # εδώ, όχι μόνο στη φόρμα, ώστε να μη μείνει τμήμα από προηγούμενη
    # επιλογή αν αλλάξει ο κωδικός δαπάνης.
    code = db.get(ExpenseCode, inv.expense_code_id) if inv.expense_code_id else None
    if not code or code.praktiko_type != "ekpaideysis":
        inv.department_id = None
    inv.payment_method_id = int(f["payment_method_id"]) if f.get("payment_method_id") else None
    inv.number = f["number"].strip()
    inv.doc_date = parse_date(f["doc_date"])
    inv.net = parse_float(f.get("net"))
    inv.vat = parse_float(f.get("vat"))
    inv.gross = parse_float(f.get("gross")) or (inv.net + inv.vat)
    inv.payment_date = parse_date(f.get("payment_date"))
    # Όταν ο τρόπος πληρωμής είναι ταυτόχρονος (κάρτα, IRIS), η εξόφληση
    # συμπίπτει με το παραστατικό και το πεδίο ούτε εμφανίζεται.
    pm = db.get(PaymentMethod, inv.payment_method_id) if inv.payment_method_id else None
    if pm and not pm.needs_date:
        inv.payment_date = inv.doc_date
    inv.notes = (f.get("notes") or "").strip()
    inv.updated_at = datetime.utcnow()


REQUIRED = [("supplier_id", "προμηθευτής"),
            ("expense_code_id", "κωδικός δαπάνης"),
            ("doc_type_id", "είδος παραστατικού"),
            ("number", "αριθμός παραστατικού"),
            ("doc_date", "ημερομηνία παραστατικού"),
            ("payment_method_id", "τρόπος πληρωμής"),
            ("gross", "μεικτή αξία")]


def missing_fields(f: dict, db) -> list:
    """Έλεγχος και στον server, όχι μόνο στη φόρμα: ο browser μπορεί να
    παρακαμφθεί, η βάση όχι."""
    missing = [label for key, label in REQUIRED if not (f.get(key) or "").strip()]
    # Η μεικτή αξία είναι το ποσό του πρακτικού — μηδέν δεν στέκει
    if "μεικτή αξία" not in missing and parse_float(f.get("gross")) <= 0:
        missing.append("μεικτή αξία (μεγαλύτερη του μηδενός)")
    pm = db.get(PaymentMethod, int(f["payment_method_id"])) \
        if (f.get("payment_method_id") or "").strip() else None
    if pm and pm.needs_date and not (f.get("payment_date") or "").strip():
        missing.append("ημερομηνία εξόφλησης")
    code = db.get(ExpenseCode, int(f["expense_code_id"])) \
        if (f.get("expense_code_id") or "").strip() else None
    if code and code.praktiko_type == "ekpaideysis" \
            and not (f.get("department_id") or "").strip():
        missing.append("τμήμα")
    return missing


def _attach_uploads(form, inv, db) -> int:
    """Αποθηκεύει όσα αρχεία ήρθαν στο πεδίο `files`.

    Διαβάζουμε από τη φόρμα με έλεγχο χαρακτηριστικών και όχι isinstance,
    γιατί το request.form() επιστρέφει το UploadFile του Starlette."""
    n = 0
    for up in form.getlist("files"):
        if hasattr(up, "file") and getattr(up, "filename", ""):
            db.add(save_upload(inv, up))
            n += 1
    return n


@app.post("/parastatiko/neo")
async def invoice_create(request: Request, db=Depends(get_db)):
    require_login(request, db)
    form = await request.form()
    f = dict(form)
    gaps = missing_fields(f, db)
    if gaps:
        flash(request, "Λείπουν υποχρεωτικά πεδία: " + ", ".join(gaps),
              error=True)
        return RedirectResponse("/parastatiko/neo", status_code=303)
    inv = Invoice()
    _save_invoice(inv, f, db)
    db.add(inv)
    try:
        db.commit()
    except Exception:
        db.rollback()
        flash(request, "Υπάρχει ήδη παραστατικό με ίδιο προμηθευτή, "
                       "αριθμό και ημερομηνία.", error=True)
        return RedirectResponse("/parastatiko/neo", status_code=303)
    db.refresh(inv)
    n = _attach_uploads(form, inv, db)
    db.commit()
    flash(request, f"Καταχωρίστηκε το παραστατικό {inv.number}."
                   + (f" Επισυνάφθηκαν {n} αρχεία." if n else ""))
    if f.get("exit") == "1":
        return RedirectResponse("/parastatika", status_code=303)
    return RedirectResponse(f"/parastatiko/{inv.id}", status_code=303)


@app.post("/parastatiko/{inv_id}")
async def invoice_update(inv_id: int, request: Request, db=Depends(get_db)):
    require_login(request, db)
    inv = db.get(Invoice, inv_id)
    if not inv:
        raise HTTPException(404)
    form = await request.form()
    gaps = missing_fields(dict(form), db)
    if gaps:
        flash(request, "Λείπουν υποχρεωτικά πεδία: " + ", ".join(gaps),
              error=True)
        return RedirectResponse(f"/parastatiko/{inv_id}", status_code=303)
    _save_invoice(inv, dict(form), db)
    db.commit()
    n = _attach_uploads(form, inv, db)
    db.commit()
    flash(request, "Αποθηκεύτηκε." + (f" Επισυνάφθηκαν {n} αρχεία." if n else ""))
    if form.get("exit") == "1":
        return RedirectResponse("/parastatika", status_code=303)
    return RedirectResponse(f"/parastatiko/{inv.id}", status_code=303)


@app.post("/parastatiko/{inv_id}/cancel")
def invoice_cancel(inv_id: int, request: Request, reason: str = Form(...),
                   db=Depends(get_db)):
    require_login(request, db)
    inv = db.get(Invoice, inv_id)
    inv.status = "cancelled"
    inv.cancel_reason = reason.strip()
    inv.updated_at = datetime.utcnow()
    db.commit()
    # Αν είχε πρωτοκολληθεί, ακυρώνεται και η εγγραφή στο πρωτόκολλο.
    if inv.protocol_linked and inv.protocol_number:
        try:
            proto.cancel(inv.protocol_year, inv.protocol_number,
                         f"Ακύρωση παραστατικού: {reason.strip()}")
            flash(request, "Ακυρώθηκε, μαζί με την εγγραφή πρωτοκόλλου.")
        except proto.ProtoError as e:
            flash(request, f"Το παραστατικό ακυρώθηκε, αλλά το πρωτόκολλο "
                           f"δεν ενημερώθηκε: {e}", error=True)
    else:
        flash(request, "Το παραστατικό ακυρώθηκε.")
    return RedirectResponse(f"/parastatiko/{inv_id}", status_code=303)


def wipe_invoice_files(inv, db):
    """Σβήνει και τα αρχεία από τον δίσκο, όχι μόνο τις εγγραφές."""
    for ef in list(inv.files):
        try:
            os.remove(os.path.join(FILES_DIR, ef.stored_name))
        except OSError:
            pass
        db.delete(ef)


@app.post("/parastatiko/{inv_id}/delete")
def invoice_delete(inv_id: int, request: Request, db=Depends(get_db)):
    """Οριστική διαγραφή.

    Επιτρέπεται μόνο όσο δεν έχει αποδοθεί αριθμός πρωτοκόλλου. Μετά την
    πρωτοκόλληση ο αριθμός έχει ήδη δεσμευτεί στο βιβλίο, οπότε η σωστή
    ενέργεια είναι ακύρωση — αλλιώς μένει ορφανή εγγραφή.
    """
    require_login(request, db)
    inv = db.get(Invoice, inv_id)
    if not inv:
        raise HTTPException(404)
    if inv.protocol_number:
        flash(request, "Έχει αριθμό πρωτοκόλλου — χρησιμοποίησε ακύρωση, "
                       "ώστε να μη μείνει ορφανή εγγραφή στο βιβλίο.",
              error=True)
        return RedirectResponse(f"/parastatiko/{inv_id}", status_code=303)
    num = inv.number
    wipe_invoice_files(inv, db)
    db.delete(inv)
    db.commit()
    flash(request, f"Το παραστατικό {num} διαγράφηκε οριστικά.")
    return RedirectResponse("/parastatika", status_code=303)


@app.post("/parastatiko/{inv_id}/restore")
def invoice_restore(inv_id: int, request: Request, db=Depends(get_db)):
    require_login(request, db)
    inv = db.get(Invoice, inv_id)
    inv.status = "active"
    inv.cancel_reason = ""
    db.commit()
    flash(request, "Επαναφέρθηκε.")
    return RedirectResponse(f"/parastatiko/{inv_id}", status_code=303)


# ---------------------------------------------------------------- πρωτοκόλληση

@app.post("/parastatiko/{inv_id}/protokollo")
def invoice_protocol(inv_id: int, request: Request, db=Depends(get_db)):
    """Αυτόματη πρωτοκόλληση: ζητά αριθμό από την εφαρμογή πρωτοκόλλου."""
    require_login(request, db)
    inv = db.get(Invoice, inv_id)
    if not inv.needs_praktiko:
        flash(request, "Η δαπάνη αυτή δεν παίρνει πρακτικό.", error=True)
        return RedirectResponse(f"/parastatiko/{inv_id}", status_code=303)
    if inv.protocol_number:
        flash(request, "Έχει ήδη αριθμό πρωτοκόλλου.", error=True)
        return RedirectResponse(f"/parastatiko/{inv_id}", status_code=303)
    try:
        res = proto.register(subject_for(inv), inv.ext_id)
    except proto.ProtoError as e:
        flash(request, str(e), error=True)
        return RedirectResponse(f"/parastatiko/{inv_id}", status_code=303)
    inv.protocol_year = res["year"]
    inv.protocol_number = res["number"]
    inv.protocol_date = parse_date(res.get("entry_date")) or date.today()
    inv.protocol_linked = True
    db.commit()
    flash(request, f"Αποδόθηκε αριθμός πρωτοκόλλου {res['code']}.")
    return RedirectResponse(f"/parastatiko/{inv_id}", status_code=303)


@app.post("/parastatiko/{inv_id}/protokollo/link")
def invoice_protocol_link(inv_id: int, request: Request,
                          code: str = Form(...), db=Depends(get_db)):
    """Χειροκίνητη καταχώριση αριθμού. Αν η εγγραφή υπάρχει στο πρωτόκολλο,
    συνδέεται· αλλιώς αποθηκεύεται ως απλή αναφορά."""
    require_login(request, db)
    inv = db.get(Invoice, inv_id)
    raw = code.strip().replace("\\", "/")
    try:
        if "/" in raw:
            y, n = raw.split("/", 1)
            year, number = int(y), int(n)
        else:
            year, number = date.today().year, int(raw)
    except ValueError:
        flash(request, "Δώσε αριθμό στη μορφή 2026/153 ή σκέτο 153.",
              error=True)
        return RedirectResponse(f"/parastatiko/{inv_id}", status_code=303)

    inv.protocol_year, inv.protocol_number = year, number
    try:
        res = proto.fetch(year, number)
        inv.protocol_date = parse_date(res.get("entry_date")) or date.today()
        inv.protocol_linked = True
        msg = f"Συνδέθηκε με την εγγραφή {res['code']} ({res['subject'][:60]})."
        if res.get("status") == "cancelled":
            msg += " Προσοχή: η εγγραφή είναι ακυρωμένη στο πρωτόκολλο."
    except proto.ProtoError as e:
        inv.protocol_date = inv.protocol_date or date.today()
        inv.protocol_linked = False
        msg = f"Αποθηκεύτηκε ως αναφορά, χωρίς σύνδεση ({e})."
    db.commit()
    flash(request, msg)
    return RedirectResponse(f"/parastatiko/{inv_id}", status_code=303)


@app.post("/parastatiko/{inv_id}/protokollo/clear")
def invoice_protocol_clear(inv_id: int, request: Request, db=Depends(get_db)):
    require_login(request, db)
    inv = db.get(Invoice, inv_id)
    inv.protocol_year = inv.protocol_number = inv.protocol_date = None
    inv.protocol_linked = inv.protocol_uploaded = False
    db.commit()
    flash(request, "Αφαιρέθηκε η αναφορά πρωτοκόλλου. Η εγγραφή στο "
                   "πρωτόκολλο ΔΕΝ ακυρώθηκε — κάν' το χειροκίνητα αν χρειάζεται.")
    return RedirectResponse(f"/parastatiko/{inv_id}", status_code=303)


@app.get("/arxeio/{file_id}")
def invoice_file_download(file_id: int, request: Request, db=Depends(get_db)):
    require_login(request, db)
    ef = db.get(InvoiceFile, file_id)
    if not ef:
        raise HTTPException(404, "Δεν βρέθηκε το αρχείο")
    path = os.path.join(FILES_DIR, ef.stored_name)
    if not os.path.exists(path):
        raise HTTPException(404, "Το αρχείο λείπει από τον δίσκο")
    with open(path, "rb") as fh:
        return download(fh.read(), ef.original_name,
                        "application/octet-stream")


@app.post("/arxeio/{file_id}/delete")
def invoice_file_delete(file_id: int, request: Request, db=Depends(get_db)):
    require_login(request, db)
    ef = db.get(InvoiceFile, file_id)
    inv_id = ef.invoice_id
    try:
        os.remove(os.path.join(FILES_DIR, ef.stored_name))
    except OSError:
        pass
    db.delete(ef)
    db.commit()
    flash(request, "Το συνημμένο διαγράφηκε.")
    return RedirectResponse(f"/parastatiko/{inv_id}", status_code=303)


@app.get("/parastatiko/{inv_id}/praktiko")
def invoice_praktiko(inv_id: int, request: Request, db=Depends(get_db)):
    """Λήψη πρακτικού.

    Δεν αναρτάται στο πρωτόκολλο: το παραγόμενο αρχείο είναι ανυπόγραφο
    και το ΙΝΕΔΙΒΙΜ τα ζητά σε φυσική μορφή. Στο πρωτόκολλο μένει μόνο
    η εγγραφή με τον αριθμό."""
    require_login(request, db)
    inv = db.get(Invoice, inv_id)
    try:
        fname, data = docgen.build_praktiko(inv, org_context(db))
    except (ValueError, FileNotFoundError) as e:
        flash(request, str(e), error=True)
        return RedirectResponse(f"/parastatiko/{inv_id}", status_code=303)
    return download(data, fname, DOCX_MIME)


@app.post("/parastatiko/{inv_id}/praktiko-typos")
def invoice_praktiko_type(inv_id: int, request: Request,
                          praktiko_override: str = Form(OVERRIDE_AUTO),
                          db=Depends(get_db)):
    """Χειροκίνητη αλλαγή του τύπου πρακτικού για ένα παραστατικό.

    Ο κανόνας καλύπτει τη συντριπτική πλειονότητα, αλλά όχι τα πάντα —
    εδώ διορθώνεται η εξαίρεση χωρίς να πειραχτεί ο κανόνας."""
    require_login(request, db)
    inv = db.get(Invoice, inv_id)
    if not inv:
        return RedirectResponse("/parastatika", status_code=303)
    valid = {OVERRIDE_AUTO, OVERRIDE_NONE} | set(PRAKTIKO_TYPES)
    inv.praktiko_override = praktiko_override \
        if praktiko_override in valid else OVERRIDE_AUTO
    inv.updated_at = datetime.utcnow()
    db.commit()
    flash(request, f"Τύπος πρακτικού: {inv.praktiko_label}."
          if inv.needs_praktiko else "Το παραστατικό δεν παίρνει πλέον πρακτικό.")
    return RedirectResponse(f"/parastatiko/{inv_id}", status_code=303)


# -------------------------------------------------------------- μηνιαία έγγραφα

@app.get("/eggrafa", response_class=HTMLResponse)
def documents_page(request: Request, year: str = "", month: str = "",
                   db=Depends(get_db)):
    require_login(request, db)
    today = date.today()
    year = as_int(year, today.year)
    month = as_int(month, today.month)
    invs = month_invoices(db, year, month)
    batch = db.query(MonthlyBatch).filter_by(year=year, month=month).first()
    return render(request, "eggrafa.html", db, year=year, month=month,
                  invoices=invs, batch=batch,
                  with_praktiko=[i for i in invs if i.needs_praktiko],
                  total=sum(i.gross for i in invs))


@app.post("/eggrafa/protokollisi")
def documents_protocol_all(request: Request, year: int = Form(...),
                           month: int = Form(...), db=Depends(get_db)):
    """Μαζική πρωτοκόλληση όσων παραστατικών του μήνα δεν έχουν αριθμό."""
    require_login(request, db)
    done, failed = 0, []
    for inv in month_invoices(db, year, month):
        if not inv.needs_praktiko or inv.protocol_number:
            continue
        try:
            res = proto.register(subject_for(inv), inv.ext_id)
        except proto.ProtoError as e:
            failed.append(f"{inv.number}: {e}")
            break          # σταματάμε στο πρώτο σφάλμα, δεν επιμένουμε
        inv.protocol_year, inv.protocol_number = res["year"], res["number"]
        inv.protocol_date = parse_date(res.get("entry_date")) or date.today()
        inv.protocol_linked = True
        done += 1
    db.commit()
    if failed:
        flash(request, f"Πρωτοκολλήθηκαν {done}. Διακοπή: {failed[0]}",
              error=True)
    else:
        flash(request, f"Πρωτοκολλήθηκαν {done} παραστατικά.")
    return RedirectResponse(f"/eggrafa?year={year}&month={month}",
                            status_code=303)


@app.get("/eggrafa/praktika")
def documents_praktika(request: Request, year: str = "", month: str = "",
                       db=Depends(get_db)):
    """Όλα τα πρακτικά του μήνα σε ένα zip, για εκτύπωση και υπογραφή."""
    require_login(request, db)
    today = date.today()
    year, month = as_int(year, today.year), as_int(month, today.month)
    buf = io.BytesIO()
    n = 0
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        for inv in month_invoices(db, year, month):
            if not inv.needs_praktiko:
                continue
            try:
                fname, data = docgen.build_praktiko(inv, org_context(db))
            except (ValueError, FileNotFoundError):
                continue
            z.writestr(fname, data)
            n += 1
    if not n:
        flash(request, "Δεν υπάρχουν πρακτικά για τον μήνα.", error=True)
        return RedirectResponse(f"/eggrafa?year={year}&month={month}",
                                status_code=303)
    name = f"Πρακτικά_{MONTHS_NOM[month - 1]}_{year}.zip"
    return download(buf.getvalue(), name, "application/zip")


@app.post("/eggrafa/diavivastiko")
def documents_diavivastiko(request: Request, year: int = Form(...),
                           month: int = Form(...), db=Depends(get_db)):
    """Πρωτοκολλεί (αν χρειάζεται) και παράγει το διαβιβαστικό."""
    require_login(request, db)
    batch = db.query(MonthlyBatch).filter_by(year=year, month=month).first()
    if not batch:
        batch = MonthlyBatch(year=year, month=month)
        db.add(batch)
        db.commit()
        db.refresh(batch)
    if not batch.protocol_number:
        try:
            res = proto.register(
                f"Διαβιβαστικό αποστολής τιμολογίων "
                f"{MONTHS_NOM[month - 1]} {year}",
                f"diavivastiko:{year}-{month:02d}")
            batch.protocol_year, batch.protocol_number = res["year"], res["number"]
            batch.protocol_date = parse_date(res.get("entry_date")) or date.today()
            batch.protocol_linked = True
            db.commit()
        except proto.ProtoError as e:
            flash(request, f"Δεν αποδόθηκε αριθμός: {e}", error=True)
            return RedirectResponse(f"/eggrafa?year={year}&month={month}",
                                    status_code=303)
    invs = month_invoices(db, year, month)
    try:
        fname, data = docgen.build_diavivastiko(batch, invs, org_context(db))
    except FileNotFoundError as e:
        flash(request, str(e), error=True)
        return RedirectResponse(f"/eggrafa?year={year}&month={month}",
                                status_code=303)
    return download(data, fname, DOCX_MIME)


@app.post("/eggrafa/diavivastiko/link")
def documents_diavivastiko_link(request: Request, year: int = Form(...),
                                month: int = Form(...), code: str = Form(...),
                                db=Depends(get_db)):
    """Σύνδεση της μηνιαίας αποστολής με αριθμό που κρατήθηκε από πριν.

    Ίδια λογική με το παραστατικό: αν η εγγραφή βρεθεί στο πρωτόκολλο
    συνδέεται, αλλιώς μένει ως απλή αναφορά."""
    require_login(request, db)
    back = f"/eggrafa?year={year}&month={month}"
    raw = code.strip().replace("\\", "/")
    try:
        if "/" in raw:
            y, n = raw.split("/", 1)
            p_year, p_number = int(y), int(n)
        else:
            p_year, p_number = date.today().year, int(raw)
    except ValueError:
        flash(request, "Δώσε αριθμό στη μορφή 2026/153 ή σκέτο 153.",
              error=True)
        return RedirectResponse(back, status_code=303)

    batch = db.query(MonthlyBatch).filter_by(year=year, month=month).first()
    if not batch:
        batch = MonthlyBatch(year=year, month=month)
        db.add(batch)
    batch.protocol_year, batch.protocol_number = p_year, p_number
    try:
        res = proto.fetch(p_year, p_number)
        batch.protocol_date = parse_date(res.get("entry_date")) or date.today()
        batch.protocol_linked = True
        msg = f"Συνδέθηκε με την εγγραφή {res['code']}."
        if res.get("status") == "cancelled":
            msg += " Προσοχή: η εγγραφή είναι ακυρωμένη στο πρωτόκολλο."
    except proto.ProtoError as e:
        batch.protocol_date = batch.protocol_date or date.today()
        batch.protocol_linked = False
        msg = f"Αποθηκεύτηκε ως αναφορά, χωρίς σύνδεση ({e})."
    db.commit()
    flash(request, msg)
    return RedirectResponse(back, status_code=303)


@app.post("/eggrafa/diavivastiko/clear")
def documents_diavivastiko_clear(request: Request, year: int = Form(...),
                                 month: int = Form(...), db=Depends(get_db)):
    require_login(request, db)
    batch = db.query(MonthlyBatch).filter_by(year=year, month=month).first()
    if batch:
        batch.protocol_year = batch.protocol_number = None
        batch.protocol_date = None
        batch.protocol_linked = False
        db.commit()
        flash(request, "Αφαιρέθηκε η αναφορά πρωτοκόλλου.")
    return RedirectResponse(f"/eggrafa?year={year}&month={month}",
                            status_code=303)


@app.get("/eggrafa/imerologio")
def documents_imerologio(request: Request, year: str = "", month: str = "",
                         db=Depends(get_db)):
    require_login(request, db)
    today = date.today()
    year, month = as_int(year, today.year), as_int(month, today.month)
    invs = month_invoices(db, year, month)
    fname, data = docgen.build_imerologio(year, month, invs,
                                          org_context(db))
    return download(data, fname, XLSX_MIME)


# ------------------------------------------------------------------ προμηθευτές

@app.get("/promitheftes", response_class=HTMLResponse)
def suppliers_list(request: Request, q: str = "", db=Depends(get_db)):
    require_login(request, db)
    query = db.query(Supplier)
    if q:
        query = query.filter(or_(Supplier.name.ilike(f"%{q}%"),
                                 Supplier.afm.ilike(f"%{q}%")))
    return render(request, "promitheftes.html", db, q=q,
                  suppliers=query.order_by(Supplier.name).all())


@app.get("/promitheftis/neos", response_class=HTMLResponse)
def supplier_new_page(request: Request, db=Depends(get_db)):
    require_login(request, db)
    return render(request, "promitheftis.html", db, s=None, invoices=[],
                  doy_list=DOY_LIST,
                  apo="", eos="", use_range=False, count=0, total=0,
                  net_total=0, vat_total=0, all_total=0)


@app.get("/promitheftis/{sid}", response_class=HTMLResponse)
def supplier_page(sid: int, request: Request, apo: str = "", eos: str = "",
                  db=Depends(get_db)):
    require_login(request, db)
    s = db.get(Supplier, sid)
    if not s:
        raise HTTPException(404)
    d_apo, d_eos = parse_date(apo), parse_date(eos)
    invs = sorted(s.invoices, key=lambda i: (i.doc_date or date.min, i.id),
                  reverse=True)
    if d_apo:
        invs = [i for i in invs if i.doc_date and i.doc_date >= d_apo]
    if d_eos:
        invs = [i for i in invs if i.doc_date and i.doc_date <= d_eos]
    active = [i for i in invs if i.status == "active"]
    return render(request, "promitheftis.html", db, s=s, invoices=invs,
                  doy_list=DOY_LIST,
                  apo=apo, eos=eos, use_range=bool(d_apo or d_eos),
                  count=len(active),
                  total=sum(i.gross for i in active),
                  net_total=sum(i.net for i in active),
                  vat_total=sum(i.vat for i in active),
                  all_total=sum(i.gross for i in s.invoices
                                if i.status == "active"))


def _save_supplier(s, f):
    s.afm = f["afm"].strip()
    s.name = f["name"].strip()
    for k in ("doy", "address", "phone", "email", "bank", "iban", "rf",
              "notes"):
        setattr(s, k, (f.get(k) or "").strip())
    s.active = f.get("active") == "on"


@app.post("/promitheftis/neos")
async def supplier_create(request: Request, db=Depends(get_db)):
    require_login(request, db)
    f = dict(await request.form())
    s = Supplier()
    _save_supplier(s, f)
    db.add(s)
    try:
        db.commit()
    except Exception:
        db.rollback()
        flash(request, f"Υπάρχει ήδη προμηθευτής με ΑΦΜ {f['afm']}.",
              error=True)
        return RedirectResponse("/promitheftis/neos", status_code=303)
    flash(request, "Καταχωρίστηκε ο προμηθευτής.")
    return RedirectResponse(f"/promitheftis/{s.id}", status_code=303)


@app.post("/promitheftis/{sid}")
async def supplier_update(sid: int, request: Request, db=Depends(get_db)):
    require_login(request, db)
    s = db.get(Supplier, sid)
    f = dict(await request.form())
    _save_supplier(s, f)
    db.commit()
    flash(request, "Αποθηκεύτηκε.")
    return RedirectResponse(f"/promitheftis/{sid}", status_code=303)


@app.get("/api/promitheftes")
def supplier_search(q: str = "", db=Depends(get_db)):
    """Προτάσεις προμηθευτών καθώς πληκτρολογείς ΑΦΜ ή επωνυμία."""
    q = q.strip()
    if len(q) < 2:
        return {"results": []}
    rows = (db.query(Supplier)
            .filter(Supplier.active == True,  # noqa: E712
                    or_(Supplier.afm.like(f"{q}%"),
                        Supplier.name.ilike(f"%{q}%")))
            .order_by(Supplier.name).limit(8).all())
    return {"results": [{"id": s.id, "afm": s.afm, "name": s.name,
                         "iban": s.iban} for s in rows]}


@app.get("/api/promitheftis/{sid}/parastatika")
def supplier_invoices(sid: int, exclude: int = 0, db=Depends(get_db)):
    """Τα τελευταία παραστατικά ενός προμηθευτή — εμφανίζονται στη φόρμα
    καταχώρισης, ώστε να φαίνεται αμέσως τυχόν διπλοεγγραφή."""
    rows = (db.query(Invoice).filter(Invoice.supplier_id == sid,
                                     Invoice.id != exclude)
            .order_by(Invoice.doc_date.desc(), Invoice.id.desc())
            .limit(10).all())
    total = sum(i.gross for i in
                db.query(Invoice).filter(Invoice.supplier_id == sid,
                                         Invoice.status == "active").all())
    return {"count": len(rows), "total": round(total, 2),
            "results": [{"id": i.id, "number": i.number,
                         "date": i.doc_date.strftime("%d/%m/%Y") if i.doc_date else "",
                         "gross": round(i.gross, 2),
                         "protocol": i.protocol_code,
                         "cancelled": i.status == "cancelled"} for i in rows]}


@app.get("/api/promitheftis")
def supplier_lookup(afm: str = "", db=Depends(get_db)):
    """Ανάκληση στοιχείων με ΑΦΜ — για τη φόρμα καταχώρισης."""
    s = db.query(Supplier).filter(Supplier.afm == afm.strip()).first()
    if not s:
        return {"found": False}
    return {"found": True, "id": s.id, "afm": s.afm, "name": s.name,
            "doy": s.doy, "address": s.address, "phone": s.phone,
            "email": s.email, "iban": s.iban}


# --------------------------------------------------------------------- ρυθμίσεις

def template_status(db) -> list:
    """Κατάσταση κάθε προτύπου: αν υπάρχει, πότε άλλαξε, τι μεταβλητές έχει."""
    out = []
    for key, (fname, label) in docgen.TEMPLATES.items():
        path = os.path.join(docgen.TEMPLATE_DIR, fname)
        exists = os.path.exists(path)
        found = docgen.placeholders_in(path) if exists else set()
        known = set(docgen.known_fields(key))
        out.append({
            "key": key, "label": label, "file": fname, "exists": exists,
            "modified": (datetime.fromtimestamp(os.path.getmtime(path))
                         .strftime("%d/%m/%Y %H:%M") if exists else ""),
            "used": sorted(found & known),
            "unknown": sorted(found - known),
            "available": sorted(known - found),
        })
    return out


@app.get("/rythmiseis", response_class=HTMLResponse)
def settings_page(request: Request, db=Depends(get_db)):
    require_login(request, db)
    ok, msg = proto.status()
    return render(request, "rythmiseis.html", db,
                  proto_key_set=bool(proto.PROTO_KEY),
                  proto_recipient=proto.RECIPIENT,
                  codes=db.query(ExpenseCode).order_by(ExpenseCode.label).all(),
                  methods=db.query(PaymentMethod).order_by(PaymentMethod.name).all(),
                  doctypes=db.query(DocType).order_by(DocType.name).all(),
                  departments=db.query(Department).order_by(
                      Department.specialty, Department.semester).all(),
                  users=db.query(User).order_by(User.username).all(),
                  praktiko_types=PRAKTIKO_TYPES,
                  proto_ok=ok, proto_msg=msg, proto_url=proto.PROTO_URL,
                  org_fields=ORG_FIELDS, org=org_context(db),
                  templates=template_status(db))


@app.post("/rythmiseis/protokollo")
def settings_proto(request: Request, url: str = Form(""), key: str = Form(""),
                   recipient: str = Form(""), db=Depends(get_db)):
    require_login(request, db)
    if url.strip():
        set_setting(db, "proto_url", url.strip())
    if key.strip():
        set_setting(db, "proto_key", key.strip())
    if recipient.strip():
        set_setting(db, "proto_recipient", recipient.strip())
    load_proto_config(db)
    ok, msg = proto.status()
    flash(request, f"Αποθηκεύτηκε. Δοκιμή: {msg}", error=not ok)
    return RedirectResponse("/rythmiseis", status_code=303)


@app.post("/rythmiseis/protokollo/dokimi")
def settings_proto_test(request: Request, db=Depends(get_db)):
    require_login(request, db)
    ok, msg = proto.status()
    flash(request, msg, error=not ok)
    return RedirectResponse("/rythmiseis", status_code=303)


DEFAULT_METHODS = [("Τραπεζική μεταφορά", True, True),
                   ("Χρεωστική Κάρτα", False, False),
                   ("IRIS", False, False)]


@app.post("/rythmiseis/katharismos")
def settings_wipe(request: Request, confirm: str = Form(""),
                  reset_methods: str = Form(""), db=Depends(get_db)):
    """Διαγραφή όλων των παραστατικών — για το τέλος των δοκιμών.

    Οι εγγραφές που είχαν πρωτοκολληθεί ακυρώνονται στο πρωτόκολλο, δεν
    σβήνονται: οι αριθμοί έχουν ήδη δοθεί και η σειρά δεν σπάει."""
    require_login(request, db)
    # gr_upper, όχι .upper(): το «διαγραφή».upper() δίνει «ΔΙΑΓΡΑΦΉ»
    # με τόνο και η σύγκριση αποτυγχάνει.
    if docgen.gr_upper(confirm.strip()) != "ΔΙΑΓΡΑΦΗ":
        flash(request, "Για επιβεβαίωση γράψε ΔΙΑΓΡΑΦΗ στο πεδίο.",
              error=True)
        return RedirectResponse("/rythmiseis", status_code=303)

    cancelled, failed, n = 0, 0, 0
    for inv in db.query(Invoice).all():
        if inv.protocol_linked and inv.protocol_number:
            try:
                proto.cancel(inv.protocol_year, inv.protocol_number,
                             "Διαγραφή δοκιμαστικών δεδομένων")
                cancelled += 1
            except proto.ProtoError:
                failed += 1
        wipe_invoice_files(inv, db)
        db.delete(inv)
        n += 1
    for b in db.query(MonthlyBatch).all():
        if b.protocol_linked and b.protocol_number:
            try:
                proto.cancel(b.protocol_year, b.protocol_number,
                             "Διαγραφή δοκιμαστικών δεδομένων")
                cancelled += 1
            except proto.ProtoError:
                failed += 1
        db.delete(b)
    db.commit()

    msg = f"Διαγράφηκαν {n} παραστατικά."
    if cancelled:
        msg += f" Ακυρώθηκαν {cancelled} εγγραφές πρωτοκόλλου."
    if failed:
        msg += (f" ΠΡΟΣΟΧΗ: {failed} εγγραφές πρωτοκόλλου δεν ακυρώθηκαν "
                f"— ακύρωσέ τις χειροκίνητα.")

    if reset_methods == "on":
        for m in db.query(PaymentMethod).all():
            db.delete(m)
        # Ξεχωριστό commit: αλλιώς το SQLAlchemy εκτελεί τις εισαγωγές
        # πριν τις διαγραφές και σκάει στο μοναδικό όνομα.
        db.commit()
        for name, nd, np_ in DEFAULT_METHODS:
            db.add(PaymentMethod(name=name, needs_date=nd, needs_proof=np_))
        db.commit()
        msg += " Οι τρόποι πληρωμής επανήλθαν στους τρεις προεπιλεγμένους."
    flash(request, msg, error=bool(failed))
    return RedirectResponse("/rythmiseis", status_code=303)


@app.post("/rythmiseis/stoixeia")
async def settings_org(request: Request, db=Depends(get_db)):
    require_login(request, db)
    f = dict(await request.form())
    for key, _ph, _label, _default in ORG_FIELDS:
        if key in f:
            set_setting(db, key, f[key].strip())
    flash(request, "Τα στοιχεία αποθηκεύτηκαν.")
    return RedirectResponse("/rythmiseis", status_code=303)


@app.get("/rythmiseis/protypo/{key}")
def template_download(key: str, request: Request, db=Depends(get_db)):
    require_login(request, db)
    if key not in docgen.TEMPLATES:
        raise HTTPException(404)
    fname = docgen.TEMPLATES[key][0]
    path = os.path.join(docgen.TEMPLATE_DIR, fname)
    if not os.path.exists(path):
        raise HTTPException(404, "Το πρότυπο δεν υπάρχει")
    with open(path, "rb") as fh:
        return download(fh.read(), fname, DOCX_MIME)


@app.post("/rythmiseis/protypo/{key}")
async def template_upload(key: str, request: Request,
                          file: UploadFile = File(...), db=Depends(get_db)):
    """Αντικατάσταση προτύπου από την εφαρμογή.

    Το αρχείο ελέγχεται ότι είναι έγκυρο .docx πριν γραφτεί, και το
    προηγούμενο κρατιέται ως αντίγραφο ασφαλείας."""
    require_login(request, db)
    if key not in docgen.TEMPLATES:
        raise HTTPException(404)
    fname = docgen.TEMPLATES[key][0]
    raw = await file.read()
    try:
        found = docgen.placeholders_in(raw)
    except Exception:
        flash(request, "Το αρχείο δεν είναι έγκυρο .docx. "
                       "Αποθήκευσε από το Word ως «Έγγραφο Word (.docx)».",
              error=True)
        return RedirectResponse("/rythmiseis", status_code=303)

    path = os.path.join(docgen.TEMPLATE_DIR, fname)
    if os.path.exists(path):
        import shutil
        shutil.copy2(path, path + ".bak")
    with open(path, "wb") as fh:
        fh.write(raw)

    unknown = sorted(found - set(docgen.known_fields(key)))
    msg = f"Το πρότυπο «{docgen.TEMPLATES[key][1]}» ανέβηκε."
    if unknown:
        msg += (" Προσοχή: οι μεταβλητές " +
                ", ".join(f"«{u}»" for u in unknown) +
                " δεν αναγνωρίζονται και θα μείνουν ως έχουν.")
    flash(request, msg, error=bool(unknown))
    return RedirectResponse("/rythmiseis", status_code=303)


@app.post("/rythmiseis/metavoli")
async def settings_edit(request: Request, db=Depends(get_db)):
    """Επεξεργασία υπάρχοντος στοιχείου ρυθμίσεων (όνομα, κωδικός, τύπος)."""
    require_login(request, db)
    f = dict(await request.form())
    kind, item_id = f.get("kind"), int(f.get("item_id"))
    if kind == "code":
        c = db.get(ExpenseCode, item_id)
        c.label = f.get("label", c.label).strip() or c.label
        c.code = f.get("code", "").strip()
        c.praktiko_type = f.get("praktiko_type", "")
        c.active = True
    elif kind in ("method", "doctype"):
        obj = db.get(PaymentMethod if kind == "method" else DocType, item_id)
        obj.name = f.get("name", obj.name).strip() or obj.name
        if kind == "method":
            obj.needs_proof = f.get("needs_proof") == "on"
            obj.needs_date = f.get("needs_date") == "on"
    elif kind == "dept":
        d = db.get(Department, item_id)
        d.specialty = f.get("specialty", d.specialty).strip().upper()
        d.semester = f.get("semester", d.semester).strip().upper()
        d.section = f.get("section", d.section).strip()
    db.commit()
    flash(request, "Αποθηκεύτηκε.")
    return RedirectResponse("/rythmiseis", status_code=303)


@app.post("/rythmiseis/diagrafi")
def settings_delete(request: Request, kind: str = Form(...),
                    item_id: int = Form(...), db=Depends(get_db)):
    """Διαγραφή στοιχείου ρυθμίσεων. Αν χρησιμοποιείται ήδη σε παραστατικό,
    απενεργοποιείται αντί να διαγραφεί, ώστε να μη χαθεί το ιστορικό."""
    model, field = {"code": (ExpenseCode, Invoice.expense_code_id),
                    "method": (PaymentMethod, Invoice.payment_method_id),
                    "doctype": (DocType, Invoice.doc_type_id),
                    "dept": (Department, Invoice.department_id)}[kind]
    obj = db.get(model, item_id)
    if not obj:
        return RedirectResponse("/rythmiseis", status_code=303)
    used = db.query(Invoice).filter(field == item_id).count()
    if used:
        obj.active = False
        flash(request, f"Χρησιμοποιείται σε {used} παραστατικά — "
                       f"απενεργοποιήθηκε αντί να διαγραφεί.")
    else:
        db.delete(obj)
        flash(request, "Διαγράφηκε.")
    db.commit()
    return RedirectResponse("/rythmiseis", status_code=303)


@app.post("/rythmiseis/apli")
def settings_add_simple(request: Request, kind: str = Form(...),
                        name: str = Form(...), needs_proof: str = Form(""),
                        needs_date: str = Form(""), db=Depends(get_db)):
    require_login(request, db)
    if kind == "method":
        db.add(PaymentMethod(name=name.strip(),
                             needs_proof=needs_proof == "on",
                             needs_date=needs_date == "on"))
        db.commit()
        flash(request, "Προστέθηκε.")
        return RedirectResponse("/rythmiseis", status_code=303)
    model = {"doctype": DocType}[kind]
    db.add(model(name=name.strip()))
    db.commit()
    flash(request, "Προστέθηκε.")
    return RedirectResponse("/rythmiseis", status_code=303)


@app.post("/rythmiseis/tmima")
def settings_add_dept(request: Request, specialty: str = Form(...),
                      semester: str = Form(...), section: str = Form(...),
                      db=Depends(get_db)):
    require_login(request, db)
    db.add(Department(specialty=specialty.strip().upper(),
                      semester=semester.strip().upper(),
                      section=section.strip()))
    db.commit()
    flash(request, "Προστέθηκε τμήμα.")
    return RedirectResponse("/rythmiseis", status_code=303)


@app.post("/rythmiseis/xristis")
def settings_add_user(request: Request, username: str = Form(...),
                      password: str = Form(...), full_name: str = Form(""),
                      db=Depends(get_db)):
    require_login(request, db)
    db.add(User(username=username.strip(), password_hash=pwd.hash(password),
                full_name=full_name.strip()))
    db.commit()
    flash(request, "Προστέθηκε χρήστης.")
    return RedirectResponse("/rythmiseis", status_code=303)


@app.post("/rythmiseis/kodikos_prosvasis")
def settings_change_password(request: Request, password: str = Form(...),
                             db=Depends(get_db)):
    u = require_login(request, db)
    u.password_hash = pwd.hash(password)
    db.commit()
    flash(request, "Ο κωδικός σου άλλαξε.")
    return RedirectResponse("/rythmiseis", status_code=303)


@app.post("/rythmiseis/eisagogi")
async def settings_import(request: Request, kind: str = Form(...),
                          file: UploadFile = File(...), db=Depends(get_db)):
    """Εισαγωγή προμηθευτών ή τμημάτων από CSV (κωδικοποίηση UTF-8)."""
    require_login(request, db)
    raw = (await file.read()).decode("utf-8-sig", errors="replace")
    reader = csv.DictReader(io.StringIO(raw))
    added = updated = 0
    for row in reader:
        row = {(k or "").strip().lower(): (v or "").strip()
               for k, v in row.items()}
        if kind == "suppliers":
            afm = row.get("αφμ") or row.get("afm")
            if not afm:
                continue
            s = db.query(Supplier).filter(Supplier.afm == afm).first()
            if s:
                updated += 1
            else:
                s = Supplier(afm=afm)
                db.add(s)
                added += 1
            s.name = row.get("επωνυμία") or row.get("name") or s.name or afm
            for src, dst in [("τράπεζα", "bank"), ("ιβαν", "iban"),
                             ("iban", "iban"), ("δου", "doy"),
                             ("διεύθυνση", "address"), ("τηλέφωνο", "phone"),
                             ("email", "email")]:
                if row.get(src):
                    setattr(s, dst, row[src])
        elif kind == "departments":
            sp = row.get("ειδικότητα") or row.get("specialty")
            if not sp:
                continue
            sem = row.get("εξάμηνο") or row.get("semester") or ""
            sec = row.get("τμήμα") or row.get("section") or "1"
            exists = db.query(Department).filter_by(
                specialty=sp, semester=sem, section=sec).first()
            if not exists:
                db.add(Department(specialty=sp, semester=sem, section=sec))
                added += 1
    db.commit()
    flash(request, f"Εισαγωγή: {added} νέες, {updated} ενημερώσεις.")
    return RedirectResponse("/rythmiseis", status_code=303)
