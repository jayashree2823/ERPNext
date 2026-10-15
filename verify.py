"""Stage D: compare the source files with what is stored in the database. READ-ONLY.
Run from the bench `sites` folder:  ../env/bin/python ../crm_migration/verify.py"""
import collections, csv, datetime as dt, os, sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from common import *  # noqa
from import_crm import as_text, load_deal_sources
import frappe

OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "reports")


def same(a, b):
    if isinstance(a, (int, float)) or isinstance(b, (int, float)):
        try:
            return float(a or 0) == float(b or 0)
        except (TypeError, ValueError):
            return False
    return (a if a not in ("", None) else None) == (b if b not in ("", None) else None)


def assigned(doctype, name):
    return sorted(frappe.get_all("ToDo", {"reference_type": doctype, "reference_name": name, "status": "Open"},
                                 pluck="allocated_to"))


def main():
    frappe.init(site="erpnext.local", sites_path="."); frappe.connect()
    try:
        run()
    finally:
        frappe.db.rollback(); frappe.destroy()


def run():
    rows, disc = [], []
    cnt = collections.Counter()
    # ---- leads
    leads = read_xlsx(FILES["lead_xlsx"])
    keep, _, _ = dedupe(leads, LEAD_DEDUPE_RULES, lead_original_order)
    cnt["lead_source_rows"] = len(leads)
    cnt["lead_skipped_duplicates"] = sum(1 for i in keep if keep[i] != i)
    for i, r in enumerate(leads):
        if keep[i] != i:
            continue
        key = lead_source_key(r)
        d = frappe.db.get_value("CRM Lead", {"migrated_lead_key": key}, "*", as_dict=True)
        exp = lead_creation(r)
        if not d:
            cnt["lead_missing"] += 1
            rows.append(["CRM Lead", r["_row"], key, "", fmt_ts(exp), "", "NOT IMPORTED", ""]); continue
        cnt["lead_imported"] += 1
        ok = d.creation == exp
        cnt["lead_creation_match" if ok else "lead_creation_MISMATCH"] += 1
        checks = {
            "modified": (r["Last Updated On"], d.modified), "owner": (r["Created By"], d.owner),
            "lead_name": (as_text(r["Full Name"]), d.lead_name), "organization": (as_text(r["Organization"]), d.organization),
            "status": (r["Status"], d.status), "email": (as_text(r["Email"]), d.email),
            "mobile_no": (as_text(r["Mobile No."]), d.mobile_no), "source": (as_text(r["Source"]), d.source),
            "lead_owner": (clean(r["Lead Owner"]), d.lead_owner), "converted": (int(r["Converted"] or 0), d.converted),
            "description": (clean(r["Description"]), d.description), "zoho_description": (clean(r["Zoho Description"]), d.zoho_description),
            "zoho_created_time": (r["Zoho Created Time"], d.zoho_created_time), "erpnext_created_on": (r["Created On"], d.erpnext_created_on),
            "assignments": (sorted(parse_assign(r["Assign"])), assigned("CRM Lead", d.name)),
        }
        for f, (a, b) in checks.items():
            if not same(a, b):
                disc.append(["CRM Lead", d.name, r["_row"], f, str(a)[:200], str(b)[:200]])
        rows.append(["CRM Lead", r["_row"], key, d.name, fmt_ts(exp), fmt_ts(d.creation), "MATCH" if ok else "MISMATCH",
                     "n/a"])
    # ---- deals
    plan, new_only = load_deal_sources()
    cnt["deal_source_rows"] = len(plan) + len(new_only)
    items = [(r, z, n, "zoho-deal:" + z["Record Id"] if z else "",
              dt.datetime.strptime(z["Created Time"], "%Y-%m-%d %H:%M:%S") if z else None) for r, z, n in plan]
    items += [(r, None, r, "erp-deal:" + r["Created On"].strftime("%Y-%m-%d %H:%M:%S.%f"), r["Created On"]) for r in new_only]
    for r, z, n, key, exp in items:
        d = frappe.db.get_value("CRM Deal", {"migrated_deal_key": key}, "*", as_dict=True) if key else None
        if not d:
            cnt["deal_missing"] += 1
            rows.append(["CRM Deal", r["_row"], key, "", fmt_ts(exp), "", "NOT IMPORTED", ""]); continue
        cnt["deal_imported"] += 1
        ok = d.creation == exp
        cnt["deal_creation_match" if ok else "deal_creation_MISMATCH"] += 1
        contact = frappe.get_all("CRM Contacts", {"parent": d.name, "parenttype": "CRM Deal"}, pluck="contact")
        rel = []
        if r["Organization"] and not frappe.db.exists("CRM Organization", d.organization or ""):
            rel.append("organization missing")
        if not contact and (r["Primary email"] or r["Primary mobile no"]):
            rel.append("no contact")
        if r["Contact"] and contact != [r["Contact"]]:
            rel.append(f"contact {contact} != {r['Contact']}")
        if d.lead and not frappe.db.exists("CRM Lead", d.lead):
            rel.append("lead link broken")
        if (n or {}).get("Lead") and not d.lead:
            rel.append(f"source lead {n['Lead']} not linked")
        checks = {
            "modified": (r["Last Updated On"], d.modified), "owner": (r["Created By"], d.owner),
            "status": (r["Status"], d.status), "organization": (clean(r["Organization"]), d.organization),
            "deal_owner": (clean(r["Deal Owner"]), d.deal_owner), "deal_value": (r["Deal Value"], d.deal_value),
            "expected_deal_value": (r["Expected Deal Value"], d.expected_deal_value),
            "annual_revenue": (r["Annual Revenue"], d.annual_revenue), "source": (as_text(r["Source"]), d.source),
            "email": (as_text(r["Primary email"]), d.email), "mobile_no": (as_text(r["Primary mobile no"]), d.mobile_no),
            "deal_name": (as_text(r["Lead Name"]) or (z["Deal Name"] if z else None), d.lead_name),
            "erpnext_created_on": ((n or {}).get("Created On"), d.erpnext_created_on),
            "zoho_created_time": (exp if z else None, d.zoho_created_time),
            "assignments": (sorted(parse_assign(r["Assign"])), assigned("CRM Deal", d.name)),
        }
        for f, (a, b) in checks.items():
            if not same(a, b):
                disc.append(["CRM Deal", d.name, r["_row"], f, str(a)[:200], str(b)[:200]])
        rows.append(["CRM Deal", r["_row"], key, d.name, fmt_ts(exp), fmt_ts(d.creation), "MATCH" if ok else "MISMATCH",
                     "; ".join(rel) or "OK"])
    with open(os.path.join(OUT, "verification_records.csv"), "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f); w.writerow(["doctype", "source_excel_row", "source_key", "target_name", "source_creation",
                                       "target_creation", "timestamp", "relationships"]); w.writerows(rows)
    with open(os.path.join(OUT, "verification_discrepancies.csv"), "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f); w.writerow(["doctype", "target_name", "source_excel_row", "field", "source_value", "target_value"]); w.writerows(disc)
    extra = {dtp: frappe.db.count(dtp, {k: ["is", "not set"]}) for dtp, k in
             (("CRM Lead", "migrated_lead_key"), ("CRM Deal", "migrated_deal_key"))}
    print("Summary:", dict(cnt)); print("Field discrepancies by field:", collections.Counter((d[0], d[3]) for d in disc))
    print("Records in target NOT from this migration:", extra)
    print("Reports: reports/verification_records.csv, reports/verification_discrepancies.csv")


if __name__ == "__main__":
    main()
