"""Stage A+B: audit + dry run. READ-ONLY against the database; writes CSV/markdown reports only.
Run from the bench `sites` folder:  ../env/bin/python ../crm_migration/dry_run.py"""
import collections, csv, datetime as dt, os, sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from common import *  # noqa
import frappe
from frappe.utils import validate_email_address

OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "reports")
C = collections.Counter

# Proposed status types for statuses missing in the target (exact source text is kept)
LEAD_STATUS_TYPE = {
    "Not Contacted": "Open", "Attempt to Contact": "Ongoing", "Attempted to Contact": "Ongoing",
    "Attempted to call": "Ongoing", "Attempt to Call": "Ongoing", "Contact in Future": "Ongoing",
    "Connected": "Ongoing", "Pre-Qualified": "Ongoing", "Hot Prospect": "Ongoing",
    "Qualified Prospect": "Won", "Lost Lead": "Lost", "Junk Lead": "Lost", "Not Qualified": "Lost",
}
DEAL_STATUS_TYPE = {
    "Need Analysis / Meeting Done": ("Ongoing", 25), "Proposal or Price Quote Submitted": ("Ongoing", 50),
    "Negotiation / Review": ("Ongoing", 70), "Closed Won": ("Won", 100), "Closed Lost": ("Lost", 0),
}


def w(name, header, rows):
    with open(os.path.join(OUT, name), "w", newline="", encoding="utf-8") as f:
        wr = csv.writer(f); wr.writerow(header); wr.writerows(rows)


def main():
    frappe.init(site="erpnext.local", sites_path="."); frappe.connect()
    try:
        run()
    finally:
        frappe.db.rollback(); frappe.destroy()


def run():
    md = []; p = md.append
    leads = read_xlsx(FILES["lead_xlsx"]); deals_all = read_xlsx(FILES["deal_xlsx_all"])
    deals_new = read_xlsx(FILES["deal_xlsx_new"]); zdeals = read_csv(FILES["zoho_deals"])
    users = set(frappe.get_all("User", pluck="name"))
    lead_status = set(frappe.get_all("CRM Lead Status", pluck="name"))
    deal_status = set(frappe.get_all("CRM Deal Status", pluck="name"))
    sources = set(frappe.get_all("CRM Lead Source", pluck="name"))
    orgs = set(frappe.get_all("CRM Organization", pluck="name"))
    p(f"# CRM migration dry run — {dt.datetime.now():%Y-%m-%d %H:%M}\n\nNothing was written to the database.\n")
    p("## Source files\n\n| file | rows | sha256 |\n|---|---|---|")
    for k, path in FILES.items():
        n = len(read_csv(path)) if path.endswith(".csv") else len(read_xlsx(path))
        p(f"| {os.path.basename(path)} | {n} | `{sha256(path)[:16]}…` |")

    # ---------------- LEADS ----------------
    orig = lead_original_order
    keep, groups, reasons = dedupe(leads, LEAD_DEDUPE_RULES, orig)
    to_import = [i for i in range(len(leads)) if keep[i] == i]
    skipped = [i for i in range(len(leads)) if keep[i] != i]
    rows = []
    for g in sorted(groups, key=lambda g: -len(g)):
        for i in sorted(g, key=lambda i: orig(leads[i])):
            r = leads[i]
            rows.append([f"G{groups.index(g)+1}", "KEEP" if keep[i] == i else "SKIP", r["_row"], leads[keep[i]]["_row"],
                         "; ".join(sorted(reasons[i])), r["Full Name"], r["Organization"], r["Email"], r["Mobile No."],
                         r["Status"], fmt_ts(r["Created On"]), fmt_ts(r["Zoho Created Time"]), (r["Description"] or "")[:120]])
    w("lead_duplicates.csv", ["group", "action", "excel_row", "kept_excel_row", "matched_on", "full_name", "organization",
                              "email", "mobile", "status", "created_on", "zoho_created_time", "description(120)"], rows)
    by_reason = C()
    for i in skipped:
        by_reason["; ".join(sorted(reasons[i]))] += 1

    issues = []; owners = C(); assignees = C(); creators = C()
    for i in to_import:
        r = leads[i]
        for col, req in (("Full Name", True), ("Status", True), ("Created On", True)):
            if req and clean(r[col]) is None:
                issues.append((r["_row"], "error", f"missing {col}"))
        if r["Email"] and not validate_email_address(str(r["Email"]).strip()):
            issues.append((r["_row"], "warning", f"invalid email {r['Email']!r} -> imported blank, original kept in log"))
        if r["Email"] and r["Lead Owner"] and str(r["Email"]).strip() == r["Lead Owner"]:
            issues.append((r["_row"], "error", "email equals lead owner (CRM rejects this)"))
        for col, lim in (("Full Name", 140), ("Organization", 140), ("Email", 140), ("Mobile No.", 140)):
            if r[col] and len(str(r[col])) > lim:
                issues.append((r["_row"], "error", f"{col} longer than {lim} chars"))
        owners[r["Lead Owner"]] += 1; creators[r["Created By"]] += 1
        for u in parse_assign(r["Assign"]):
            assignees[u] += 1
    w("lead_issues.csv", ["excel_row", "severity", "issue"], issues)
    ts = [lead_creation(leads[i]) for i in to_import]
    erp_ts = [leads[i]["Created On"] for i in to_import]
    p(f"\n## Leads\n\n- Rows in `CRM Lead.xlsx`: **{len(leads)}**\n- Duplicate groups (exact copy, OR same person name + same email, OR same person name + same mobile): **{len(groups)}**"
      f"\n- Rows skipped as duplicates: **{len(skipped)}** — " + ", ".join(f"{k}: {v}" for k, v in by_reason.most_common()) +
      f"\n- **Leads to import: {len(to_import)}** (kept row = earliest original date in each group)"
      f"\n- `creation` source: Zoho Created Time **{sum(1 for i in to_import if leads[i]['Zoho Created Time'])}**, ERPNext Created On **{sum(1 for i in to_import if not leads[i]['Zoho Created Time'])}**"
      f"\n- `creation` range: {min(ts)} → {max(ts)}"
      f"\n- ERPNext Created On (kept in custom field) all unique: {len(set(erp_ts)) == len(erp_ts)}"
      f"\n- Validation issues: {C(s for _, s, _ in issues)} (see lead_issues.csv)")

    # ---------------- DEALS ----------------
    newmap = {norm_name(r["Lead Name"]): r for r in deals_new}
    zmap = {norm_name(z["Deal Name"]): z for z in zdeals}
    dplan = []
    for r in deals_all:
        n = newmap.get(norm_name(r["Lead Name"]))
        z = zmap.get(norm_name(r["Lead Name"]))
        how = "name"
        if z is None and not r["Lead Name"]:  # unnamed deal: match Zoho deal created ~same time & same creator
            cand = [x for x in zdeals if norm_name(x["Deal Name"]) not in {norm_name(d["Lead Name"]) for d in deals_all}]
            if len(cand) == 1:
                z, how = cand[0], "only unmatched Zoho deal (time check below)"
        created = n["Created On"] if n else None
        ztime = dt.datetime.strptime(z["Created Time"], "%Y-%m-%d %H:%M:%S") if z else None
        src = "Zoho Created Time" if ztime else ("ERPNext Created On" if created else "NONE")
        dplan.append(dict(r=r, n=n, z=z, how=how, creation=ztime or created, creation_src=src, ztime=ztime))
    # agreed: deals sharing an email are separate projects -> only exact duplicate deals would be skipped
    dkeep, dgroups, dreasons = dedupe([d["r"] for d in dplan], [
        ("exact duplicate", lambda r: tuple(str(r.get(c)) for c in r if c not in ("Sr", "_row")))], lambda r: r["_row"])
    drows = []
    for i, d in enumerate(dplan):
        r, z = d["r"], d["z"]
        missing_users = [u for u in [r["Deal Owner"], r["Created By"], *parse_assign(r["Assign"])] if u and u not in users]
        drows.append([r["_row"], "IMPORT" if dkeep[i] == i else f"DUPLICATE of row {dplan[dkeep[i]]['r']['_row']}",
                      "; ".join(sorted(dreasons[i])), r["Lead Name"], r["Organization"], r["Status"], r["Deal Value"],
                      r["Expected Deal Value"], r["Primary email"], r["Primary mobile no"], r["Contact"],
                      fmt_ts(d["creation"]), d["creation_src"], z["Record Id"] if z else "", d["how"] if z else "no Zoho match",
                      z["Probability (%)"] if z else "", z["Closing Date"] if z else "", (n or {}).get("Lead") or "",
                      ", ".join(missing_users)])
        if r["Deal Owner"]: owners[r["Deal Owner"]] += 1
        creators[r["Created By"]] += 1
        for u in parse_assign(r["Assign"]): assignees[u] += 1
    w("deal_plan.csv", ["excel_row", "action", "duplicate_reason", "deal_name", "organization", "status", "deal_value",
                        "expected_deal_value", "email", "mobile", "contact", "creation_to_set", "creation_source",
                        "zoho_record_id", "zoho_match", "zoho_probability", "zoho_closing_date", "lead_link", "missing_users"], drows)
    p(f"\n## Deals\n\n- Deals in `CRM Deal (1).xlsx`: **{len(deals_all)}**; exact duplicates skipped: **{sum(len(g)-1 for g in dgroups)}**; **to import: {sum(1 for i in range(len(dplan)) if dkeep[i]==i)}**"
      f"\n- `creation` from Zoho Created Time: **{sum(1 for d in dplan if d['creation_src']=='Zoho Created Time')}**; from ERPNext Created On: **{sum(1 for d in dplan if d['creation_src']=='ERPNext Created On')}**; none: **{sum(1 for d in dplan if d['creation_src']=='NONE')}**"
      f"\n- ERPNext Created On available (kept in custom field): **{sum(1 for d in dplan if d['n'])}**")
    unnamed = [d for d in dplan if not d["r"]["Lead Name"]]
    for d in unnamed:
        p(f"- Unnamed deal (row {d['r']['_row']}, last updated {d['r']['Last Updated On']}) matched to Zoho "
          f"'{d['z']['Deal Name'] if d['z'] else '-'}' created {d['z']['Created Time'] if d['z'] else '-'} — **please confirm**")

    # ---------------- MASTER DATA ----------------
    all_users = (set(owners) | set(assignees) | set(creators)) - {None, "", "Administrator"}
    new_users = sorted(u for u in all_users if u not in users)
    w("users_to_create.csv", ["email", "as_owner", "as_assignee", "as_creator"],
      [[u, owners[u], assignees[u], creators[u]] for u in new_users])
    ls_new = sorted({leads[i]["Status"] for i in to_import} - lead_status)
    ds_new = sorted({d["r"]["Status"] for d in dplan} - deal_status)
    src_new = sorted(({clean(leads[i]["Source"]) for i in to_import} | {clean(d["r"]["Source"]) for d in dplan}) - sources - {None})
    org_new = sorted({clean(d["r"]["Organization"]) for d in dplan} - orgs - {None})
    lost_leads = sum(1 for i in to_import if LEAD_STATUS_TYPE.get(leads[i]["Status"]) == "Lost")
    lost_deals = sum(1 for d in dplan if DEAL_STATUS_TYPE.get(d["r"]["Status"], ("",))[0] == "Lost")
    p(f"\n## Master data that must be created first\n\n- **Users ({len(new_users)})**: " + ", ".join(new_users))
    p(f"- **Lead statuses ({len(ls_new)})**: " + ", ".join(f"{s} → type {LEAD_STATUS_TYPE.get(s,'?')}" for s in ls_new))
    p(f"- **Deal statuses ({len(ds_new)})**: " + ", ".join(f"{s} → type {DEAL_STATUS_TYPE.get(s,('?',))[0]}" for s in ds_new))
    p(f"- **Lead sources ({len(src_new)})**: " + ", ".join(src_new))
    p(f"- **CRM Organizations ({len(org_new)})**: " + ", ".join(org_new))
    p(f"- Leads with a 'Lost'-type status: {lost_leads}; deals: {lost_deals} → CRM requires a lost reason; proposed: reason **Other**, note *'Migrated – no lost reason recorded in source'*")

    # ---------------- FIELD MAPPING ----------------
    lead_map = [
        ("Sr", "(not imported)", "row number added by the export"),
        ("Full Name", "first_name (→ lead_name)", "kept exactly as one string, incl. salutation"),
        ("Organization", "organization", ""), ("Status", "status", "exact text; missing statuses created"),
        ("Email", "email", "invalid emails left blank, logged"), ("Mobile No.", "mobile_no", ""),
        ("Assign", "assignments (ToDo)", "exactly the source list; auto-added owner removed if not in source"),
        ("Last Updated On", "modified", "set after insert"), ("Source", "source", "missing sources created"),
        ("Lead Owner", "lead_owner", ""), ("Converted", "converted", ""),
        ("Description", "description (NEW custom field, Text)", ""),
        ("Zoho Description", "zoho_description (NEW custom field, Text)", ""),
        ("Zoho Created Time", "creation  (+ zoho_created_time NEW custom field)", "agreed: real original date; set after insert, verified exactly"),
        ("Created On", "erpnext_created_on (NEW custom field, Datetime)  — and creation when no Zoho time", "old ERPNext creation kept exactly"),
        ("Created By", "owner", "set after insert"),
        ("(none)", "modified_by", "not in export → stays the migration user (Administrator)"),
        ("(none)", "migration_source_key (NEW hidden custom field)", "stable key 'erp-lead:<ERPNext Created On>' to stop double imports"),
    ]
    deal_map = [
        ("Organization", "organization (Link)", "CRM Organization created if missing"),
        ("Annual Revenue", "annual_revenue", ""), ("Status", "status", "exact text; missing statuses created"),
        ("Primary email / mobile", "contacts table → email / mobile_no", "CRM clears these unless a Contact is linked → Contact created"),
        ("Contact", "Contact record + contacts table", "29 contacts to create"),
        ("Assign", "assignments (ToDo)", "as for leads"), ("Last Updated On", "modified", ""),
        ("Deal Owner", "deal_owner", ""), ("Expected Deal Value", "expected_deal_value", ""), ("Deal Value", "deal_value", ""),
        ("Source", "source", ""), ("Lead", "lead", "1 deal; linked to the imported lead found by mobile"),
        ("Lead Name", "lead_name", "this is the deal's title in CRM"),
        ("Zoho CSV: Created Time", "creation  (+ zoho_created_time NEW custom field)", "agreed: real original date for all deals"),
        ("Created On", "erpnext_created_on (NEW custom field)", "only 10 deals have it"),
        ("Created By", "owner", ""),
        ("Zoho: Probability (%)", "probability", "else CRM fills a status default"),
        ("Zoho: Closing Date", "closed_date for Won deals", "else CRM would write today's date"),
        ("Zoho: Record Id", "migration_source_key", ""),
    ]
    p("\n## Field mapping — CRM Lead\n\n| source column | target field | note |\n|---|---|---|")
    md.extend(f"| {a} | {b} | {c} |" for a, b, c in lead_map)
    p("\n## Field mapping — CRM Deal\n\n| source column | target field | note |\n|---|---|---|")
    md.extend(f"| {a} | {b} | {c} |" for a, b, c in deal_map)
    open(os.path.join(OUT, "dry_run_summary.md"), "w", encoding="utf-8").write("\n".join(md) + "\n")
    print("\n".join(md))


if __name__ == "__main__":
    main()
