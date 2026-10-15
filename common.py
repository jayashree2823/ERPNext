"""Shared loading / normalisation / de-duplication for the CRM Lead & Deal migration.
Reads the export files only; never modifies them."""
import collections, csv, datetime as dt, hashlib, json, re
import openpyxl

DL = "/mnt/c/Users/JAYASHREE/Downloads/"
FILES = {
    "lead_xlsx": DL + "CRM Lead.xlsx",          # ERPNext export 15-Oct (has Created On)
    "deal_xlsx_all": DL + "CRM Deal (1).xlsx",  # ERPNext export 29-Sep, all 39 deals, no Created On
    "deal_xlsx_new": DL + "CRM Deal.xlsx",      # ERPNext export 15-Oct, 10 deals with Created On
    "zoho_leads": DL + "Leads_2026_10_15.csv",
    "zoho_deals": DL + "Deals_2026_10_15.csv",
}
EXPECTED_SHA256 = {
    "lead_xlsx": "c6d1aa5bf92df9abbc262bed1071393a363692120c02985034b2319656073396",
    "deal_xlsx_all": "2d213a101de1e94e6fcdb0308e54bf2d8d12f3a8f4ea4db3aa030826fd1b4e70",
    "deal_xlsx_new": "15abbff433d88bd9f2f1002aa20ff072539d372d1ea273780c515c887f515c96",
    "zoho_leads": "d646041d47e99477f958970e6f30bfefe7987a83241122c77847298c45fdd90e",
    "zoho_deals": "1f338beaa55cedcf55af31d01de1503b5f58e759ffabc6db340c5b5e6d8144ed",
}


def sha256(path):
    return hashlib.sha256(open(path, "rb").read()).hexdigest()


def check_files():
    bad = {k: sha256(p) for k, p in FILES.items() if sha256(p) != EXPECTED_SHA256[k]}
    if bad:
        raise SystemExit(f"Export files changed since the dry run was approved: {list(bad)}")


def read_xlsx(path):
    rows = list(openpyxl.load_workbook(path, read_only=True).worksheets[0].iter_rows(values_only=True))
    head = [str(h).strip() for h in rows[0]]
    return [dict(zip(head, r), _row=i + 2) for i, r in enumerate(rows[1:]) if any(v is not None for v in r)]


def read_csv(path):
    with open(path, encoding="utf-8-sig", newline="") as f:
        return list(csv.DictReader(f))


def norm_phone(x):
    d = re.sub(r"\D", "", str(x or ""))
    return d[-10:] if len(d) >= 10 else None


def norm_email(x):
    x = str(x or "").strip().lower()
    return x or None


def norm_name(x):
    return re.sub(r"[^a-z0-9]", "", str(x or "").lower()) or None


SALUTATIONS = {"mr", "mrs", "ms", "miss", "dr", "prof", "sir", "smt", "shri"}


def person_key(x):
    """Name without salutations, punctuation or repeated words: 'Mr.Naveed Mr.Naveed' -> 'naveed'."""
    toks = [t for t in re.findall(r"[a-z0-9]+", str(x or "").lower()) if t not in SALUTATIONS]
    return " ".join(dict.fromkeys(toks)) or None


def clean(x):
    if x is None:
        return None
    if isinstance(x, str):
        x = x.strip()
        return x or None
    return x


def fmt_ts(x):
    return x.strftime("%Y-%m-%d %H:%M:%S.%f") if isinstance(x, dt.datetime) else (str(x) if x else "")


def parse_assign(x):
    try:
        return [u.strip() for u in json.loads(x or "[]") if u and u.strip()]
    except ValueError:
        return []


LEAD_EXACT = ["Full Name", "Organization", "Status", "Email", "Mobile No.", "Source", "Lead Owner",
              "Description", "Zoho Description", "Zoho Created Time"]


def dedupe(rows, rules, order_key):
    """rules: [(label, keyfunc)]. Rows sharing any non-empty key are grouped (transitively).
    The row with the smallest order_key in each group is kept.
    Returns ({row_index: kept_row_index}, groups, {row_index: set(labels)})."""
    parent = list(range(len(rows)))

    def find(i):
        while parent[i] != i:
            parent[i] = parent[parent[i]]
            i = parent[i]
        return i

    reasons = collections.defaultdict(set)
    for label, keyf in rules:
        first = {}
        for i, r in enumerate(rows):
            k = keyf(r)
            if not k:
                continue
            if k in first:
                ra, rb = find(first[k]), find(i)
                if ra != rb:
                    parent[rb] = ra
                reasons[i].add(label)
                reasons[first[k]].add(label)
            else:
                first[k] = i
    groups = collections.defaultdict(list)
    for i in range(len(rows)):
        groups[find(i)].append(i)
    keep = {}
    for members in groups.values():
        winner = min(members, key=lambda i: order_key(rows[i]))
        for i in members:
            keep[i] = winner
    return keep, [m for m in groups.values() if len(m) > 1], reasons


def _pair(a, b):
    return (a, b) if a and b else None


LEAD_DEDUPE_RULES = [
    ("exact duplicate", lambda r: tuple(str(r.get(c)) for c in LEAD_EXACT)),
    ("same person + email", lambda r: _pair(person_key(r["Full Name"]), norm_email(r["Email"]))),
    ("same person + mobile", lambda r: _pair(person_key(r["Full Name"]), norm_phone(r["Mobile No."]))),
]


def lead_original_order(r):
    """Earliest real-world date wins inside a duplicate group."""
    return (r["Zoho Created Time"] or r["Created On"], r["Created On"])


def lead_creation(r):
    """Agreed rule: creation = Zoho Created Time when present, else ERPNext Created On."""
    return r["Zoho Created Time"] or r["Created On"]


def lead_source_key(r):
    """Stable per-row key (ERPNext Created On is unique in the export)."""
    return "erp-lead:" + r["Created On"].strftime("%Y-%m-%d %H:%M:%S.%f")
