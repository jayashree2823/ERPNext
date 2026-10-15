"""Stage C: import CRM Leads and Deals into erpnext.local, preserving source timestamps.

Run from the bench `sites` folder (only after the dry run is approved and a backup exists):
    ../env/bin/python ../crm_migration/import_crm.py --execute [--stage setup|leads|deals|all] [--limit N]

How timestamps are preserved
----------------------------
Frappe always overwrites `creation`, `owner`, `modified`, `modified_by` on insert
(frappe/model/document.py: set_user_and_timestamp). So each record is:
  1. inserted with the normal `doc.insert()` - all validations and CRM hooks run;
  2. then, in the same transaction, a narrowly scoped UPDATE sets only
     creation / modified / owner (and the same on its own child rows) for that one record;
  3. re-read from the database and compared exactly; any mismatch rolls the record back.
Re-running is safe: records whose migrated_*_key already exists are skipped.
"""
import argparse, collections, csv, datetime as dt, glob, os, sys, time

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from common import *  # noqa
from dry_run import DEAL_STATUS_TYPE, LEAD_STATUS_TYPE

import frappe
from frappe.custom.doctype.custom_field.custom_field import create_custom_fields
from frappe.desk.form.assign_to import _add as assign_add

SITE = "erpnext.local"
HERE = os.path.dirname(os.path.abspath(__file__))
LOST_REASON, LOST_NOTE = "Other", "Migrated – no lost reason recorded in source"
COMMIT_EVERY = 200

CUSTOM_FIELDS = {
    "CRM Lead": [
        dict(fieldname="migration_section", label="Migrated Data", fieldtype="Section Break", insert_after="lost_notes", collapsible=1),
        dict(fieldname="description", label="Description", fieldtype="Text", insert_after="migration_section"),
        dict(fieldname="zoho_description", label="Zoho Description", fieldtype="Text", insert_after="description"),
        dict(fieldname="zoho_created_time", label="Zoho Created Time", fieldtype="Datetime", read_only=1, insert_after="zoho_description"),
        dict(fieldname="erpnext_created_on", label="Old ERPNext Created On", fieldtype="Datetime", read_only=1, insert_after="zoho_created_time"),
        dict(fieldname="migrated_lead_key", label="Migration Source Key", fieldtype="Data", read_only=1, hidden=1, no_copy=1, search_index=1, insert_after="erpnext_created_on"),
    ],
    "CRM Deal": [
        dict(fieldname="migration_section", label="Migrated Data", fieldtype="Section Break", insert_after="lost_notes", collapsible=1),
        dict(fieldname="zoho_created_time", label="Zoho Created Time", fieldtype="Datetime", read_only=1, insert_after="migration_section"),
        dict(fieldname="erpnext_created_on", label="Old ERPNext Created On", fieldtype="Datetime", read_only=1, insert_after="zoho_created_time"),
        dict(fieldname="migrated_deal_key", label="Migration Source Key", fieldtype="Data", read_only=1, hidden=1, no_copy=1, search_index=1, insert_after="erpnext_created_on"),
    ],
}


class Log:
    def __init__(self, stage):
        os.makedirs(os.path.join(HERE, "logs"), exist_ok=True)
        self.path = os.path.join(HERE, "logs", f"import_{stage}_{dt.datetime.now():%Y%m%d_%H%M%S}.csv")
        self.f = open(self.path, "w", newline="", encoding="utf-8")
        self.w = csv.writer(self.f)
        self.w.writerow(["time", "doctype", "source_key", "source_excel_row", "target_name", "result", "detail"])
        self.counts = collections.Counter()

    def __call__(self, doctype, key, row, name, result, detail=""):
        self.counts[(doctype, result)] += 1
        self.w.writerow([dt.datetime.now().isoformat(timespec="seconds"), doctype, key, row, name, result, detail])
        self.f.flush()


def require_recent_backup():
    files = glob.glob(f"{SITE}/private/backups/*-database.sql.gz")
    newest = max(files, key=os.path.getmtime) if files else None
    if not newest or time.time() - os.path.getmtime(newest) > 2 * 3600:
        raise SystemExit("No database backup from the last 2 hours. Run: bench --site erpnext.local backup")
    print("Backup found:", os.path.basename(newest))


def as_text(x):
    x = clean(x)
    if isinstance(x, float) and x.is_integer():
        x = int(x)
    return None if x is None else str(x)


def set_source_timestamps(doctype, name, creation, modified, owner, child_tables, status_log_from=None):
    """The only direct SQL writes: one record and its own child rows, three columns each."""
    frappe.db.sql(f"update `tab{doctype}` set creation=%s, modified=%s, owner=%s where name=%s",
                  (creation, modified, owner, name))
    for child in child_tables:
        frappe.db.sql(f"update `tab{child}` set creation=%s, modified=%s, owner=%s "
                      "where parent=%s and parenttype=%s", (creation, modified, owner, name, doctype))
    if status_log_from:
        # CRM writes "status changed now"; the status really dates from the record's creation
        frappe.db.sql("update `tabCRM Status Change Log` set from_date=%s where parent=%s and parenttype=%s",
                      (status_log_from, name, doctype))
    frappe.clear_document_cache(doctype, name)


def verify_timestamps(doctype, name, creation, modified, owner):
    got = frappe.db.sql(f"select creation, modified, owner from `tab{doctype}` where name=%s", name)[0]
    if (got[0], got[1], got[2]) != (creation, modified, owner):
        raise ValueError(f"timestamp verification failed: expected {(creation, modified, owner)} got {got}")


def apply_owner_and_assignments(doc, owner_field, owner, wanted):
    """CRM couples owner and assignment (crm/api/todo.py): adding an assignee makes them the
    owner, cancelling one clears the owner. The source often has owner != assignee, so the
    record is inserted WITHOUT an owner (no auto-assignment), the source assignees are added,
    and then the source owner is written back - the same set_value the CRM hook itself uses."""
    if owner and owner != frappe.session.user:
        doc.share_with_agent(owner)  # what CRM after_insert does for a new owner
    for u in wanted:
        assign_add({"assign_to": [u], "doctype": doc.doctype, "name": doc.name}, ignore_permissions=True)
    frappe.db.set_value(doc.doctype, doc.name, owner_field, owner, update_modified=False)
    frappe.local.message_log = []


# ---------------------------------------------------------------- setup
def stage_setup(log):
    leads = read_xlsx(FILES["lead_xlsx"])
    deals = read_xlsx(FILES["deal_xlsx_all"]) + read_xlsx(FILES["deal_xlsx_new"])
    create_custom_fields(CUSTOM_FIELDS, update=True)
    log("Custom Field", "", "", ", ".join(f"{d}.{f['fieldname']}" for d in CUSTOM_FIELDS for f in CUSTOM_FIELDS[d]), "ensured")

    users = set()
    for r in leads:
        users |= {r["Lead Owner"], r["Created By"], *parse_assign(r["Assign"])}
    for r in deals:
        users |= {r["Deal Owner"], r["Created By"], *parse_assign(r["Assign"])}
    for u in sorted(u for u in users if u and not frappe.db.exists("User", u)):
        frappe.get_doc({"doctype": "User", "email": u, "first_name": u.split("@")[0].replace(".", " ").title(),
                        "enabled": 0, "send_welcome_email": 0, "user_type": "System User",
                        "roles": [{"role": "Sales User"}]}).insert(ignore_permissions=True)
        log("User", u, "", u, "created", "disabled, Sales User, no welcome email")

    pos = (frappe.db.sql("select max(position) from `tabCRM Lead Status`")[0][0] or 0)
    for s in sorted({r["Status"] for r in leads}):
        if not frappe.db.exists("CRM Lead Status", s):
            pos += 1
            frappe.get_doc({"doctype": "CRM Lead Status", "lead_status": s, "type": LEAD_STATUS_TYPE[s], "position": pos}).insert()
            log("CRM Lead Status", s, "", s, "created", LEAD_STATUS_TYPE[s])
    pos = (frappe.db.sql("select max(position) from `tabCRM Deal Status`")[0][0] or 0)
    for s in sorted({r["Status"] for r in deals}):
        if not frappe.db.exists("CRM Deal Status", s):
            pos += 1; typ, prob = DEAL_STATUS_TYPE[s]
            frappe.get_doc({"doctype": "CRM Deal Status", "deal_status": s, "type": typ, "probability": prob, "position": pos}).insert()
            log("CRM Deal Status", s, "", s, "created", typ)
    for s in sorted({clean(r["Source"]) for r in leads + deals} - {None}):
        if not frappe.db.exists("CRM Lead Source", s):
            frappe.get_doc({"doctype": "CRM Lead Source", "source_name": s}).insert()
            log("CRM Lead Source", s, "", s, "created")
    for o in sorted({clean(r["Organization"]) for r in deals} - {None}):
        if not frappe.db.exists("CRM Organization", o):
            frappe.get_doc({"doctype": "CRM Organization", "organization_name": o}).insert()
            log("CRM Organization", o, "", o, "created")
    frappe.db.commit()


# ---------------------------------------------------------------- leads
def lead_status_is_lost(status):
    return frappe.get_cached_value("CRM Lead Status", status, "type") == "Lost"


def stage_leads(log, limit=None):
    leads = read_xlsx(FILES["lead_xlsx"])
    keep, _, reasons = dedupe(leads, LEAD_DEDUPE_RULES, lead_original_order)
    for i in range(len(leads)):
        if keep[i] != i:
            log("CRM Lead", lead_source_key(leads[i]), leads[i]["_row"], "", "skipped-duplicate",
                f"duplicate of excel row {leads[keep[i]]['_row']} ({'; '.join(sorted(reasons[i]))})")
    todo = sorted((leads[i] for i in range(len(leads)) if keep[i] == i), key=lead_creation)
    done = 0
    for r in todo[:limit] if limit else todo:
        key = lead_source_key(r)
        existing = frappe.db.get_value("CRM Lead", {"migrated_lead_key": key})
        if existing:
            log("CRM Lead", key, r["_row"], existing, "skipped-already-imported"); continue
        frappe.db.savepoint("rec")
        try:
            doc = frappe.get_doc({
                "doctype": "CRM Lead", "first_name": as_text(r["Full Name"]), "organization": as_text(r["Organization"]),
                "status": r["Status"], "email": as_text(r["Email"]), "mobile_no": as_text(r["Mobile No."]),
                "source": as_text(r["Source"]), "converted": int(r["Converted"] or 0),
                "description": clean(r["Description"]), "zoho_description": clean(r["Zoho Description"]),
                "zoho_created_time": r["Zoho Created Time"], "erpnext_created_on": r["Created On"],
                "migrated_lead_key": key,
            })
            if lead_status_is_lost(r["Status"]):
                doc.lost_reason, doc.lost_notes = LOST_REASON, LOST_NOTE
            doc.insert()
            apply_owner_and_assignments(doc, "lead_owner", clean(r["Lead Owner"]), parse_assign(r["Assign"]))
            creation = lead_creation(r)
            set_source_timestamps("CRM Lead", doc.name, creation, r["Last Updated On"], r["Created By"],
                                  ["CRM Status Change Log"], status_log_from=creation)
            verify_timestamps("CRM Lead", doc.name, creation, r["Last Updated On"], r["Created By"])
            frappe.db.release_savepoint("rec")
            log("CRM Lead", key, r["_row"], doc.name, "created",
                "creation=Zoho Created Time" if r["Zoho Created Time"] else "creation=ERPNext Created On")
        except Exception as e:
            frappe.db.rollback(save_point="rec")
            log("CRM Lead", key, r["_row"], "", "FAILED", f"{type(e).__name__}: {e}"[:500])
        done += 1
        if done % COMMIT_EVERY == 0:
            frappe.db.commit(); print(f"  leads processed: {done}/{len(todo)}")
    frappe.db.commit()


# ---------------------------------------------------------------- deals
def get_or_create_contact(cname, email, mobile, log):
    if frappe.db.exists("Contact", cname):
        return cname
    head, _, tail = cname.rpartition("-")
    first, company = (head, None if tail.isdigit() else tail) if head else (cname, None)
    c = frappe.get_doc({"doctype": "Contact", "first_name": first, "company_name": company,
                        "email_ids": [{"email_id": email, "is_primary": 1}] if email else [],
                        "phone_nos": [{"phone": mobile, "is_primary_mobile_no": 1}] if mobile else []})
    c.insert(set_name=cname)  # keep the source contact ID exactly
    log("Contact", cname, "", c.name, "created")
    return c.name


def find_lead_for_deal(lead_ref, mobile, email):
    if not lead_ref:
        return None, ""
    f = {"mobile_no": mobile} if mobile else ({"email": email} if email else None)
    hits = frappe.get_all("CRM Lead", filters={**f, "migrated_lead_key": ["is", "set"]}, pluck="name") if f else []
    if len(hits) == 1:
        return hits[0], f"source lead {lead_ref} matched by {'mobile' if mobile else 'email'}"
    return None, f"source lead {lead_ref} NOT linked ({len(hits)} candidates)"


def load_deal_sources():
    """Match each deal in the full (29-Sep) export to its Zoho row and, by NAME ONLY, to the 15-Oct export.
    Unnamed deals are never matched by name (an earlier version wrongly paired two unnamed deals).
    Returns (plan, new_only): new_only = deals that exist only in the 15-Oct export."""
    deals = read_xlsx(FILES["deal_xlsx_all"]); newer = read_xlsx(FILES["deal_xlsx_new"])
    newmap = {norm_name(r["Lead Name"]): r for r in newer if norm_name(r["Lead Name"])}
    zoho = read_csv(FILES["zoho_deals"]); zmap = {norm_name(z["Deal Name"]): z for z in zoho}
    named = {norm_name(r["Lead Name"]) for r in deals if r["Lead Name"]}
    unmatched_z = [z for z in zoho if norm_name(z["Deal Name"]) not in named]
    plan = []
    for r in deals:
        z = zmap.get(norm_name(r["Lead Name"])) if r["Lead Name"] else (unmatched_z[0] if len(unmatched_z) == 1 else None)
        plan.append((r, z, newmap.get(norm_name(r["Lead Name"])) if r["Lead Name"] else None))
    old_unnamed_orgs = {clean(r["Organization"]) for r in deals if not r["Lead Name"]}
    new_only = [r for r in newer if (norm_name(r["Lead Name"]) not in named) if r["Lead Name"]
                or clean(r["Organization"]) not in old_unnamed_orgs]
    return plan, new_only


def import_one_deal(log, r, key, creation, zoho_time, erp_created_on, probability, closing_date,
                    contact, lead_ref, deal_name):
    existing = frappe.db.get_value("CRM Deal", {"migrated_deal_key": key})
    if existing:
        log("CRM Deal", key, r["_row"], existing, "skipped-already-imported"); return
    frappe.db.savepoint("rec")
    try:
        lead, lead_note = find_lead_for_deal(lead_ref, as_text(r["Primary mobile no"]), as_text(r["Primary email"]))
        status_type, _ = DEAL_STATUS_TYPE.get(r["Status"], (frappe.get_cached_value("CRM Deal Status", r["Status"], "type"), 0))
        doc = frappe.get_doc({
            "doctype": "CRM Deal", "organization": clean(r["Organization"]), "status": r["Status"],
            "annual_revenue": r["Annual Revenue"] or 0,
            "expected_deal_value": r["Expected Deal Value"] or 0, "deal_value": r["Deal Value"] or 0,
            "source": as_text(r["Source"]), "lead": lead, "lead_name": deal_name, "probability": probability,
            "contacts": [{"contact": contact, "is_primary": 1}] if contact else [],
            "zoho_created_time": zoho_time, "erpnext_created_on": erp_created_on, "migrated_deal_key": key,
        })
        if status_type == "Lost":
            doc.lost_reason, doc.lost_notes = LOST_REASON, LOST_NOTE
        doc.insert()
        apply_owner_and_assignments(doc, "deal_owner", clean(r["Deal Owner"]), parse_assign(r["Assign"]))
        # CRM stamps today's date as closed_date for Won deals; use the source Closing Date instead
        closed = closing_date if status_type == "Won" and closing_date else None
        frappe.db.sql("update `tabCRM Deal` set closed_date=%s where name=%s", (closed, doc.name))
        set_source_timestamps("CRM Deal", doc.name, creation, r["Last Updated On"], r["Created By"],
                              ["CRM Status Change Log", "CRM Contacts"], status_log_from=creation)
        verify_timestamps("CRM Deal", doc.name, creation, r["Last Updated On"], r["Created By"])
        frappe.db.release_savepoint("rec")
        log("CRM Deal", key, r["_row"], doc.name, "created", lead_note)
    except Exception as e:
        frappe.db.rollback(save_point="rec")
        log("CRM Deal", key, r["_row"], "", "FAILED", f"{type(e).__name__}: {e}"[:500])


def stage_deals(log, limit=None):
    plan, new_only = load_deal_sources()
    for r, z, n in plan[:limit] if limit else plan:
        if not z:
            log("CRM Deal", "", r["_row"], "", "FAILED", "no Zoho match -> no creation timestamp"); continue
        creation = dt.datetime.strptime(z["Created Time"], "%Y-%m-%d %H:%M:%S")
        try:
            contact = get_or_create_contact(r["Contact"], as_text(r["Primary email"]), as_text(r["Primary mobile no"]), log) if r["Contact"] else None
        except Exception as e:
            log("CRM Deal", z["Record Id"], r["_row"], "", "FAILED", f"contact: {e}"[:500]); continue
        import_one_deal(log, r, "zoho-deal:" + z["Record Id"], creation, creation, (n or {}).get("Created On"),
                        float(z["Probability (%)"]) if z["Probability (%)"] else None, z["Closing Date"],
                        contact, (n or {}).get("Lead"), as_text(r["Lead Name"]) or z["Deal Name"])
    # deals created after the 29-Sep export: only in the 15-Oct export, no Zoho record
    for r in new_only:
        lead_ref = clean(r["Lead"])
        lead, _ = find_lead_for_deal(lead_ref, as_text(r["Primary mobile no"]), as_text(r["Primary email"]))
        contact = clean(r["Contact"])
        if not contact and lead:  # no contact in source: build it from the linked lead, as CRM conversion would
            ld = frappe.db.get_value("CRM Lead", lead, ["lead_name", "email", "mobile_no"], as_dict=True)
            cname = ld.lead_name + ("-" + clean(r["Organization"]) if clean(r["Organization"]) else "")
            contact = get_or_create_contact(cname, ld.email, ld.mobile_no, log)
        elif contact:
            contact = get_or_create_contact(contact, as_text(r["Primary email"]), as_text(r["Primary mobile no"]), log)
        import_one_deal(log, r, "erp-deal:" + r["Created On"].strftime("%Y-%m-%d %H:%M:%S.%f"), r["Created On"], None,
                        r["Created On"], None, None, contact, lead_ref, as_text(r["Lead Name"]))
    frappe.db.commit()


def stage_fix_deal_dates(log):
    """Idempotent: make erpnext_created_on equal the corrected name-only match."""
    plan, _ = load_deal_sources()
    for r, z, n in plan:
        if not z:
            continue
        name, cur = frappe.db.get_value("CRM Deal", {"migrated_deal_key": "zoho-deal:" + z["Record Id"]},
                                        ["name", "erpnext_created_on"]) or (None, None)
        want = (n or {}).get("Created On")
        if name and cur != want:
            frappe.db.set_value("CRM Deal", name, "erpnext_created_on", want, update_modified=False)
            log("CRM Deal", z["Record Id"], r["_row"], name, "erpnext_created_on-fixed", f"{cur} -> {want}")
    frappe.db.commit()


# ---------------------------------------------------------------- repairs
def stage_fix_owners(log):
    """Idempotent: write back the source owner on migrated leads where it differs."""
    for r in read_xlsx(FILES["lead_xlsx"]):
        name, cur = frappe.db.get_value("CRM Lead", {"migrated_lead_key": lead_source_key(r)}, ["name", "lead_owner"]) or (None, None)
        if name and cur != clean(r["Lead Owner"]):
            frappe.db.set_value("CRM Lead", name, "lead_owner", clean(r["Lead Owner"]), update_modified=False)
            log("CRM Lead", lead_source_key(r), r["_row"], name, "owner-restored", f"{cur} -> {r['Lead Owner']}")
    frappe.db.commit()


def stage_mark_read(log):
    """Mark CRM assignment notifications about migrated records as read (nothing deleted)."""
    for doctype, keyfield in (("CRM Lead", "migrated_lead_key"), ("CRM Deal", "migrated_deal_key")):
        n = frappe.db.sql(f"""update `tabCRM Notification` n join `tab{doctype}` d on d.name = n.notification_type_doc
            set n.`read` = 1 where n.notification_type_doctype = %s and n.type = 'Assignment'
            and n.from_user = 'Administrator' and n.`read` = 0 and ifnull(d.{keyfield}, '') != ''""", doctype)
        log("CRM Notification", doctype, "", "", "marked-read", str(frappe.db._cursor.rowcount))
    frappe.db.commit()


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--execute", action="store_true", help="required: actually write data")
    ap.add_argument("--stage", choices=["setup", "leads", "deals", "all", "fix-owners", "mark-read", "fix-deal-dates"], default="all")
    ap.add_argument("--limit", type=int, help="import only the first N records (trial run)")
    a = ap.parse_args()
    if not a.execute:
        raise SystemExit("Refusing to write without --execute. Use dry_run.py to preview.")
    check_files()
    frappe.init(site=SITE, sites_path="."); frappe.connect()
    try:
        require_recent_backup()
        frappe.set_user("Administrator")
        frappe.flags.mute_emails = True  # this process only; no system setting is changed
        log = Log(a.stage)
        if a.stage in ("setup", "all"):
            stage_setup(log)
        if a.stage in ("leads", "all"):
            stage_leads(log, a.limit)
        if a.stage in ("deals", "all"):
            stage_deals(log, a.limit)
        if a.stage == "fix-owners":
            stage_fix_owners(log)
        if a.stage == "fix-deal-dates":
            stage_fix_deal_dates(log)
        if a.stage == "mark-read":
            stage_mark_read(log)
        print("Log:", log.path)
        for (d, res), n in sorted(log.counts.items()):
            print(f"  {d:18} {res:26} {n}")
    except Exception:
        frappe.db.rollback(); raise
    finally:
        frappe.destroy()


if __name__ == "__main__":
    main()
