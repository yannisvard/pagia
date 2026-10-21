"""Μοντέλο δεδομένων εφαρμογής Παγίας Προκαταβολής.

Σχεδιαστικές αρχές:
  * Ο προμηθευτής είναι οντότητα, όχι κείμενο. Το ΑΦΜ είναι το κλειδί.
  * Το τμήμα κρατιέται σπασμένο σε ειδικότητα/εξάμηνο/τμήμα, ώστε τα
    πρότυπα να μη χρειάζονται παρσάρισμα κειμένου.
  * Ο τύπος πρακτικού καθορίζεται από τον κωδικό δαπάνης. Όποια δαπάνη
    δεν παίρνει πρακτικό (π.χ. προμήθειες τραπέζης) δηλώνεται με "".
"""

import os
from datetime import date, datetime

from sqlalchemy import (Boolean, Column, Date, DateTime, Float, ForeignKey,
                        Integer, String, Text, UniqueConstraint, create_engine,
                        inspect, text)
from sqlalchemy.orm import declarative_base, relationship, sessionmaker

DATA_DIR = os.environ.get("DATA_DIR", "/data")
os.makedirs(DATA_DIR, exist_ok=True)

engine = create_engine(f"sqlite:///{DATA_DIR}/pagia.db",
                       connect_args={"check_same_thread": False})
SessionLocal = sessionmaker(bind=engine, autoflush=False)
Base = declarative_base()

# Τύποι πρακτικού. Κάθε τιμή:
#   0 ετικέτα στην οθόνη
#   1 αρχείο προτύπου
#   2 κατηγορία όπως γράφεται στο διαβιβαστικό
#   3 η γραμμή «Β.» του διαβιβαστικού
PRAKTIKO_TYPES = {
    "ylikwn": ("Παραλαβής Υλικών",
               "ΠΡΟΤΥΠΟ_-_Πρακτικό_Παραλαβής_Υλικών.docx",
               "Υλικών",
               "Πρακτικό Παραλαβής Υλικών"),
    "ekpaideysis": ("Παραλαβής Υλικών Εκπαίδευσης",
                    "ΠΡΟΤΥΠΟ_-_Πρακτικό_Παραλαβής_Υλικών_Εκπαίδευσης.docx",
                    "Υλικών Εκπαίδευσης",
                    "Πρακτικό Παραλαβής Υλικών Εκπαίδευσης"),
    "kalis": ("Καλής Εκτέλεσης",
              "ΠΡΟΤΥΠΟ_-_Πρακτικό_Καλής_Εκτέλεσης.docx",
              "Υπηρεσιών",
              "Πρακτικό Καλής Εκτέλεσης"),
}
# Δαπάνες χωρίς πρακτικό (π.χ. προμήθειες τραπέζης) γράφονται ως «Υλικών»
DEFAULT_CATEGORY = "Υλικών"

# Είδη παραστατικού που καθορίζουν από μόνα τους τον τύπο πρακτικού.
# Ζουν στη βάση, οπότε ταιριάζουμε στο λεκτικό — κρατά τον κανόνα
# ορατό σε ένα σημείο αντί για σκόρπιους ελέγχους.
DOCTYPE_YPIRESIES = "Τιμολόγιο Παροχής Υπηρεσιών"
DOCTYPE_TRAPEZA = "Προμήθεια Τράπεζας"

# Τιμές της χειροκίνητης παράκαμψης, πέρα από τα κλειδιά του PRAKTIKO_TYPES
OVERRIDE_AUTO = "auto"     # ό,τι λέει ο κανόνας
OVERRIDE_NONE = "none"     # ρητά χωρίς πρακτικό


def praktiko_auto_type(doc_type_name: str, expense_code) -> str:
    """Ο τύπος πρακτικού όπως προκύπτει από το είδος παραστατικού.

    1. Προμήθεια τράπεζας            -> κανένα πρακτικό
    2. Τιμολόγιο παροχής υπηρεσιών   -> καλής εκτέλεσης
    3. Δαπάνη υλικών εκπαίδευσης     -> παραλαβής υλικών εκπαίδευσης
    4. Οτιδήποτε άλλο                -> παραλαβής υλικών

    Επιστρέφει "" όταν δεν παράγεται πρακτικό.
    """
    name = (doc_type_name or "").strip()
    if name == DOCTYPE_TRAPEZA:
        return ""
    if name == DOCTYPE_YPIRESIES:
        return "kalis"
    if expense_code is not None and expense_code.praktiko_type == "ekpaideysis":
        return "ekpaideysis"
    return "ylikwn"

# Δ.Ο.Υ. όπως τις τηρεί η ΑΑΔΕ (λίστα υπηρεσιών, έκδοση Σεπτεμβρίου 2026).
# Οι Δ.Ο.Υ. συγχωνεύονται κατά καιρούς — γι' αυτό το πεδίο δέχεται και
# ελεύθερο κείμενο, ώστε να μπαίνει παλιά ονομασία όπου χρειάζεται.
DOY_LIST = [
    "ΔΟΥ ΑΓΙΟΥ ΝΙΚΟΛΑΟΥ", "ΔΟΥ ΑΓΡΙΝΙΟΥ", "ΔΟΥ ΑΙΓΙΟΥ",
    "ΔΟΥ ΑΛΕΞΑΝΔΡΟΥΠΟΛΗΣ", "ΔΟΥ ΑΜΑΛΙΑΔΑΣ", "ΔΟΥ ΑΜΦΙΣΣΑΣ",
    "ΔΟΥ ΑΡΓΟΣΤΟΛΙΟΥ", "ΔΟΥ ΑΡΤΑΣ", "ΔΟΥ ΒΕΡΟΙΑΣ", "ΔΟΥ ΒΟΛΟΥ",
    "ΔΟΥ ΓΙΑΝΝΙΤΣΩΝ", "ΔΟΥ ΓΡΕΒΕΝΩΝ", "ΔΟΥ ΔΡΑΜΑΣ", "ΔΟΥ ΕΔΕΣΣΑΣ",
    "ΔΟΥ ΖΑΚΥΝΘΟΥ", "ΔΟΥ ΗΓΟΥΜΕΝΙΤΣΑΣ", "ΔΟΥ ΗΡΑΚΛΕΙΟΥ", "ΔΟΥ ΘΗΒΩΝ",
    "ΔΟΥ ΘΗΡΑΣ", "ΔΟΥ ΙΩΑΝΝΙΝΩΝ", "ΔΟΥ ΚΑΒΑΛΑΣ", "ΔΟΥ ΚΑΛΑΜΑΤΑΣ",
    "ΔΟΥ ΚΑΡΔΙΤΣΑΣ", "ΔΟΥ ΚΑΡΠΕΝΗΣΙΟΥ", "ΔΟΥ ΚΑΣΤΟΡΙΑΣ", "ΔΟΥ ΚΑΤΕΡΙΝΗΣ",
    "ΔΟΥ ΚΕΡΚΥΡΑΣ", "ΔΟΥ ΚΙΛΚΙΣ", "ΔΟΥ ΚΟΖΑΝΗΣ", "ΔΟΥ ΚΟΜΟΤΗΝΗΣ",
    "ΔΟΥ ΚΟΡΙΝΘΟΥ", "ΔΟΥ ΚΥΜΗΣ", "ΔΟΥ ΚΩ", "ΔΟΥ ΛΑΜΙΑΣ", "ΔΟΥ ΛΑΡΙΣΑΣ",
    "ΔΟΥ ΛΕΥΚΑΔΑΣ", "ΔΟΥ ΛΙΒΑΔΕΙΑΣ", "ΔΟΥ ΜΕΣΟΛΟΓΓΙΟΥ", "ΔΟΥ ΜΥΚΟΝΟΥ",
    "ΔΟΥ ΜΥΤΙΛΗΝΗΣ", "ΔΟΥ ΝΑΞΟΥ", "ΔΟΥ ΝΑΥΠΛΙΟΥ", "ΔΟΥ ΝΕΩΝ ΜΟΥΔΑΝΙΩΝ",
    "ΔΟΥ ΞΑΝΘΗΣ", "ΔΟΥ ΟΡΕΣΤΙΑΔΑΣ", "ΔΟΥ ΠΑΡΟΥ", "ΔΟΥ ΠΑΤΡΩΝ",
    "ΔΟΥ ΠΟΛΥΓΥΡΟΥ", "ΔΟΥ ΠΡΕΒΕΖΑΣ", "ΔΟΥ ΠΤΟΛΕΜΑΙΔΑΣ", "ΔΟΥ ΠΥΡΓΟΥ",
    "ΔΟΥ ΡΕΘΥΜΝΟΥ", "ΔΟΥ ΡΟΔΟΥ", "ΔΟΥ ΣΑΜΟΥ", "ΔΟΥ ΣΕΡΡΩΝ", "ΔΟΥ ΣΠΑΡΤΗΣ",
    "ΔΟΥ ΣΥΡΟΥ", "ΔΟΥ ΤΡΙΚΑΛΩΝ", "ΔΟΥ ΤΡΙΠΟΛΗΣ", "ΔΟΥ ΦΛΩΡΙΝΑΣ",
    "ΔΟΥ ΧΑΛΚΙΔΑΣ", "ΔΟΥ ΧΑΝΙΩΝ", "ΔΟΥ ΧΙΟΥ", "ΚΕ.ΦΟ.Δ.Ε. ΑΤΤΙΚΗΣ",
    "ΚΕ.ΦΟ.Δ.Ε. ΘΕΣΣΑΛΟΝΙΚΗΣ"
]

MONTHS_GEN = ["Ιανουαρίου", "Φεβρουαρίου", "Μαρτίου", "Απριλίου", "Μαΐου",
              "Ιουνίου", "Ιουλίου", "Αυγούστου", "Σεπτεμβρίου", "Οκτωβρίου",
              "Νοεμβρίου", "Δεκεμβρίου"]
MONTHS_NOM = ["Ιανουάριος", "Φεβρουάριος", "Μάρτιος", "Απρίλιος", "Μάιος",
              "Ιούνιος", "Ιούλιος", "Αύγουστος", "Σεπτέμβριος", "Οκτώβριος",
              "Νοέμβριος", "Δεκέμβριος"]
MONTHS_ACC = ["Ιανουάριο", "Φεβρουάριο", "Μάρτιο", "Απρίλιο", "Μάιο",
              "Ιούνιο", "Ιούλιο", "Αύγουστο", "Σεπτέμβριο", "Οκτώβριο",
              "Νοέμβριο", "Δεκέμβριο"]
DAYS_GR = ["Δευτέρα", "Τρίτη", "Τετάρτη", "Πέμπτη", "Παρασκευή",
           "Σάββατο", "Κυριακή"]


class User(Base):
    __tablename__ = "users"
    id = Column(Integer, primary_key=True)
    username = Column(String, unique=True, nullable=False)
    password_hash = Column(String, nullable=False)
    full_name = Column(String, default="")
    is_admin = Column(Boolean, default=False)
    created_at = Column(DateTime, default=datetime.utcnow)


class Supplier(Base):
    __tablename__ = "suppliers"
    id = Column(Integer, primary_key=True)
    afm = Column(String, unique=True, nullable=False, index=True)
    name = Column(String, nullable=False, index=True)
    doy = Column(String, default="")
    address = Column(String, default="")
    phone = Column(String, default="")
    email = Column(String, default="")
    bank = Column(String, default="")
    iban = Column(String, default="")
    rf = Column(String, default="")          # κωδικός πληρωμής RF
    notes = Column(Text, default="")
    active = Column(Boolean, default=True)
    created_at = Column(DateTime, default=datetime.utcnow)

    invoices = relationship("Invoice", back_populates="supplier")

    @property
    def total_spent(self):
        return sum(i.gross for i in self.invoices if i.status == "active")


class ExpenseCode(Base):
    """Κωδικός δαπάνης. Το `label` είναι αυτό που εμφανίζει η πλατφόρμα."""
    __tablename__ = "expense_codes"
    id = Column(Integer, primary_key=True)
    label = Column(String, unique=True, nullable=False)
    code = Column(String, default="")  # δεν είναι μοναδικός
    full_name = Column(String, default="")
    # "" = δεν παράγεται πρακτικό (π.χ. προμήθειες τραπέζης)
    praktiko_type = Column(String, default="ylikwn")
    active = Column(Boolean, default=True)

    @property
    def praktiko_label(self):
        return PRAKTIKO_TYPES.get(self.praktiko_type, ("—",))[0]

    @property
    def category(self):
        """Πώς γράφεται η κατηγορία στο διαβιβαστικό."""
        t = PRAKTIKO_TYPES.get(self.praktiko_type)
        return t[2] if t else DEFAULT_CATEGORY

    @property
    def praktiko_line(self):
        t = PRAKTIKO_TYPES.get(self.praktiko_type)
        return t[3] if t else ""


class PaymentMethod(Base):
    __tablename__ = "payment_methods"
    id = Column(Integer, primary_key=True)
    name = Column(String, unique=True, nullable=False)
    # Στο διαβιβαστικό μπαίνει γραμμή «Αποδεικτικό εξόφλησης μέσω
    # τραπεζικής μεταφοράς». Δεν την απαιτούν όλοι οι τρόποι.
    needs_proof = Column(Boolean, default=True)
    # Χρειάζεται ξεχωριστή ημερομηνία εξόφλησης; Με κάρτα ή IRIS η
    # πληρωμή είναι ταυτόχρονη με την αγορά, οπότε όχι — και το πεδίο
    # ούτε εμφανίζεται ούτε ζητείται.
    needs_date = Column(Boolean, default=True)
    active = Column(Boolean, default=True)


class DocType(Base):
    __tablename__ = "doc_types"
    id = Column(Integer, primary_key=True)
    name = Column(String, unique=True, nullable=False)
    active = Column(Boolean, default=True)


class Department(Base):
    """Τμήμα κατάρτισης, σπασμένο στα τρία που ζητούν τα πρότυπα."""
    __tablename__ = "departments"
    id = Column(Integer, primary_key=True)
    specialty = Column(String, nullable=False)   # ΤΕΧΝΙΚΟΣ ΑΜΠΕΛΟΥΡΓΙΑΣ...
    semester = Column(String, nullable=False)    # Α, Β, Γ, Δ
    section = Column(String, nullable=False)     # 1, 2
    active = Column(Boolean, default=True)
    __table_args__ = (UniqueConstraint("specialty", "semester", "section"),)

    @property
    def label(self):
        return f"{self.specialty} {self.semester}{self.section}"


class Invoice(Base):
    """Παραστατικό δαπάνης."""
    __tablename__ = "invoices"
    id = Column(Integer, primary_key=True)

    supplier_id = Column(Integer, ForeignKey("suppliers.id"), nullable=False)
    expense_code_id = Column(Integer, ForeignKey("expense_codes.id"))
    doc_type_id = Column(Integer, ForeignKey("doc_types.id"))
    department_id = Column(Integer, ForeignKey("departments.id"))
    payment_method_id = Column(Integer, ForeignKey("payment_methods.id"))

    number = Column(String, nullable=False)          # ΤΙΠ 31678
    doc_date = Column(Date, nullable=False)
    net = Column(Float, default=0.0)
    vat = Column(Float, default=0.0)
    gross = Column(Float, default=0.0)
    payment_date = Column(Date)

    # Πρωτόκολλο πρακτικού
    protocol_year = Column(Integer)
    protocol_number = Column(Integer)
    protocol_date = Column(Date)
    protocol_linked = Column(Boolean, default=False)  # συνδεδεμένο με το API
    protocol_uploaded = Column(Boolean, default=False)  # ανέβηκε το docx

    # Χειροκίνητη παράκαμψη του κανόνα, ανά παραστατικό.
    # "auto" = ό,τι λέει ο κανόνας, "none" = ρητά χωρίς πρακτικό.
    praktiko_override = Column(String, default=OVERRIDE_AUTO)

    status = Column(String, default="active")        # active | cancelled
    cancel_reason = Column(String, default="")
    notes = Column(Text, default="")
    created_at = Column(DateTime, default=datetime.utcnow)
    updated_at = Column(DateTime, default=datetime.utcnow)

    supplier = relationship("Supplier", back_populates="invoices")
    files = relationship("InvoiceFile", back_populates="invoice",
                         cascade="all, delete-orphan")
    expense_code = relationship("ExpenseCode")
    doc_type = relationship("DocType")
    department = relationship("Department")
    payment_method = relationship("PaymentMethod")

    __table_args__ = (UniqueConstraint("supplier_id", "number", "doc_date"),)

    @property
    def praktiko_auto(self):
        """Ο τύπος που προτείνει ο κανόνας, αγνοώντας την παράκαμψη."""
        return praktiko_auto_type(
            self.doc_type.name if self.doc_type else "", self.expense_code)

    @property
    def praktiko_type(self):
        """Ο τύπος που ισχύει τελικά για αυτό το παραστατικό."""
        ov = self.praktiko_override or OVERRIDE_AUTO
        if ov == OVERRIDE_AUTO:
            return self.praktiko_auto
        if ov == OVERRIDE_NONE:
            return ""
        return ov if ov in PRAKTIKO_TYPES else self.praktiko_auto

    @property
    def praktiko_label(self):
        return PRAKTIKO_TYPES.get(self.praktiko_type, ("—",))[0]

    @property
    def praktiko_category(self):
        """Πώς γράφεται η κατηγορία στο διαβιβαστικό."""
        t = PRAKTIKO_TYPES.get(self.praktiko_type)
        return t[2] if t else DEFAULT_CATEGORY

    @property
    def praktiko_line(self):
        """Η γραμμή «Β.» του διαβιβαστικού."""
        t = PRAKTIKO_TYPES.get(self.praktiko_type)
        return t[3] if t else ""

    @property
    def needs_praktiko(self):
        return bool(self.praktiko_type)

    @property
    def protocol_code(self):
        if self.protocol_year and self.protocol_number:
            return f"{self.protocol_year}/{self.protocol_number}"
        return ""

    @property
    def protocol_ref(self):
        """Ο αριθμός όπως γράφεται στα έγγραφα: 573/10-06-26."""
        return protocol_ref(self.protocol_number, self.protocol_date)

    @property
    def ext_id(self):
        return f"parastatiko:{self.id}"


class InvoiceFile(Base):
    """Συνημμένο παραστατικού (σαρωμένο τιμολόγιο, απόδειξη κ.λπ.).

    Τα αρχεία αποθηκεύονται στον δίσκο, όχι στη βάση: ο φάκελος `files`
    μέσα στο DATA_DIR αντιγράφεται μαζί με το backup."""
    __tablename__ = "invoice_files"
    id = Column(Integer, primary_key=True)
    invoice_id = Column(Integer, ForeignKey("invoices.id"), nullable=False)
    stored_name = Column(String, nullable=False)
    original_name = Column(String, nullable=False)
    size = Column(Integer, default=0)
    uploaded_at = Column(DateTime, default=datetime.utcnow)

    invoice = relationship("Invoice", back_populates="files")


def protocol_ref(number, d) -> str:
    """Αριθμός πρωτοκόλλου με την ημερομηνία του, όπως μπαίνει στα έγγραφα.

    Μορφή 573/10-06-26 — διετές έτος, όπως γράφεται στα εξερχόμενα."""
    if not number:
        return ""
    return f"{number}/{d.strftime('%d-%m-%y')}" if d else str(number)


class MonthlyBatch(Base):
    """Μηνιαία αποστολή: το διαβιβαστικό και το πρωτόκολλό του."""
    __tablename__ = "monthly_batches"
    id = Column(Integer, primary_key=True)
    year = Column(Integer, nullable=False)
    month = Column(Integer, nullable=False)
    protocol_year = Column(Integer)
    protocol_number = Column(Integer)
    protocol_date = Column(Date)
    protocol_linked = Column(Boolean, default=False)
    created_at = Column(DateTime, default=datetime.utcnow)
    __table_args__ = (UniqueConstraint("year", "month"),)

    @property
    def protocol_code(self):
        if self.protocol_year and self.protocol_number:
            return f"{self.protocol_year}/{self.protocol_number}"
        return ""

    @property
    def protocol_ref(self):
        return protocol_ref(self.protocol_number, self.protocol_date)


class Setting(Base):
    __tablename__ = "settings"
    key = Column(String, primary_key=True)
    value = Column(Text, default="")


def get_db():
    db = SessionLocal()
    try:
        yield db
    finally:
        db.close()


def get_setting(db, key, default=""):
    s = db.get(Setting, key)
    return s.value if s else default


def set_setting(db, key, value):
    s = db.get(Setting, key)
    if s:
        s.value = value
    else:
        db.add(Setting(key=key, value=value))
    db.commit()


def init_db():
    Base.metadata.create_all(engine)
    # Ήπιες μεταβάσεις: προσθήκη στηλών που μπήκαν σε νεότερη έκδοση.
    insp = inspect(engine)
    with engine.connect() as conn:
        pm = {c["name"] for c in insp.get_columns("payment_methods")}
        if "needs_date" not in pm:
            conn.execute(text("ALTER TABLE payment_methods "
                              "ADD COLUMN needs_date BOOLEAN DEFAULT 1"))
            conn.execute(text("UPDATE payment_methods SET needs_date = 0 "
                              "WHERE name LIKE '%ΚΑΡΤΑ%' "
                              "OR name LIKE '%ΜΕΤΡΗΤ%' "
                              "OR name LIKE '%IRIS%'"))
        if "needs_proof" not in pm:
            conn.execute(text("ALTER TABLE payment_methods "
                              "ADD COLUMN needs_proof BOOLEAN DEFAULT 1"))
            # Κάρτα και μετρητά δεν έχουν αποδεικτικό μεταφοράς
            conn.execute(text("UPDATE payment_methods SET needs_proof = 0 "
                              "WHERE name LIKE '%ΚΑΡΤΑ%' "
                              "OR name LIKE '%ΜΕΤΡΗΤ%'"))
        sup = {c["name"] for c in insp.get_columns("suppliers")}
        if "rf" not in sup:
            conn.execute(text(
                "ALTER TABLE suppliers ADD COLUMN rf VARCHAR DEFAULT ''"))
        cols = {c["name"] for c in insp.get_columns("invoices")}
        for name, ddl in [
            ("protocol_uploaded",
             "ALTER TABLE invoices ADD COLUMN protocol_uploaded BOOLEAN DEFAULT 0"),
            ("praktiko_override",
             "ALTER TABLE invoices ADD COLUMN praktiko_override "
             "VARCHAR DEFAULT 'auto'"),
        ]:
            if name not in cols:
                conn.execute(text(ddl))
        conn.commit()
