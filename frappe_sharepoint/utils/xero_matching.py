"""
Match attachments from a Xero data export to SmartOps documents.

SmartOps did not keep the Xero IDs when the ledgers were migrated, so files
are matched on what both sides share: date, amount and company, with the
contact name and reference text as tie-breakers.
"""

import csv
import re
import sys
from collections import defaultdict
from datetime import datetime, timedelta

import frappe
from frappe import _
from frappe.utils import getdate

csv.field_size_limit(sys.maxsize)

MULTI_FILE_SEPARATOR = " : "
XERO_ID_COLUMNS = ("InvoiceID", "BankTransactionID", "ManualJournalID", "CreditNoteID")

CONFIDENCE_EXACT = "Exact"
CONFIDENCE_EXACT_NAME_DIFFERS = "Exact (name differs)"
CONFIDENCE_EXACT_SHARED = "Exact (shared bill)"
CONFIDENCE_REVIEW = "Review"
CONFIDENCE_NO_MATCH = "No match"
CONFIDENCE_CONFIRMED = "Confirmed"
IMPORTABLE_CONFIDENCE = ("", CONFIDENCE_EXACT, CONFIDENCE_EXACT_NAME_DIFFERS, CONFIDENCE_EXACT_SHARED, CONFIDENCE_CONFIRMED)

# Per document type: date field, amount fields, name/reference fields, child
# tables whose amounts also count (an expense claim line often equals one bill)
DOCTYPE_FIELDS = {
	"Expense Claim": {"date": "posting_date", "amounts": ["total_claimed_amount", "total_sanctioned_amount", "grand_total"],
		"names": ["employee_name", "remark"], "children": [("Expense Claim Detail", "expenses", ["amount", "sanctioned_amount"])]},
	"Journal Entry": {"date": "posting_date", "amounts": ["total_debit"],
		"names": ["pay_to_recd_from", "cheque_no", "bill_no", "user_remark", "title"], "children": [("Journal Entry Account", "accounts", ["debit_in_account_currency", "credit_in_account_currency"])]},
	"Purchase Invoice": {"date": "posting_date", "amounts": ["grand_total", "base_grand_total", "rounded_total"],
		"names": ["supplier_name", "bill_no", "remarks"], "children": []},
	"Sales Invoice": {"date": "posting_date", "amounts": ["grand_total", "base_grand_total", "rounded_total"],
		"names": ["customer_name", "po_no", "remarks"], "children": []},
	"Payment Entry": {"date": "posting_date", "amounts": ["paid_amount", "received_amount", "base_paid_amount"],
		"names": ["party_name", "reference_no", "remarks"], "children": []},
	"Employee Advance": {"date": "posting_date", "amounts": ["advance_amount", "paid_amount"],
		"names": ["employee_name", "purpose"], "children": []},
}

OUTPUT_COLUMNS = ["Document Type", "ID", "AttachmentFolder", "AttachmentFilenames", "Confidence",
	"Xero Date", "Xero Contact", "Xero Total", "Xero Reference", "SmartOps Date", "SmartOps Name", "Note"]


def normalize_name(value):
	"""'David Nkrumah-Boateng (david@peas.org.uk)' -> 'davidnkrumahboateng'"""
	return re.sub(r"[^a-z]", "", (value or "").lower().split("(")[0])


def names_agree(xero_name, smartops_names):
	a = normalize_name(xero_name)
	if not a:
		return False
	for candidate in smartops_names:
		b = normalize_name(candidate)
		if b and (a in b or b in a):
			return True
	return False


def read_xero_export(path):
	"""
	One record per Xero document that has attachments, keyed by its Xero ID.
	The export is line-level, so several rows collapse into one document
	"""
	docs = {}
	with open(path, newline="", encoding="utf-8-sig") as f:
		reader = csv.DictReader(f)
		required = {"Date", "Total", "AttachmentFolder", "AttachmentFilenames"} | set(XERO_ID_COLUMNS)
		missing = required - set(reader.fieldnames or [])
		if missing:
			frappe.throw(_("The Xero export is missing these columns: {0}").format(", ".join(sorted(missing))))

		for row in reader:
			files = [p.strip() for p in (row.get("AttachmentFilenames") or "").split(MULTI_FILE_SEPARATOR) if p.strip()]
			if not files or (row.get("Status") or "").upper() == "DELETED":
				continue
			xero_id = next((row[c] for c in XERO_ID_COLUMNS if row.get(c)), None)
			if not xero_id:
				continue
			try:
				date = datetime.strptime(row["Date"], "%Y/%m/%d").date()
			except (ValueError, TypeError):
				continue

			doc = docs.setdefault(xero_id, {
				"xero_id": xero_id, "type": row.get("Type", ""), "date": date, "contact": (row.get("Contact.Name") or "").strip(),
				"total": _float(row.get("Total")), "reference": (row.get("Reference") or row.get("InvoiceNumber") or "").strip(),
				"folder": (row.get("AttachmentFolder") or "").strip("/"), "files": [], "amounts": set(), "lines": [],
			})
			for name in files:
				if name not in doc["files"]:
					doc["files"].append(name)
			doc["amounts"].add(round(doc["total"], 2))
			line = _float(row.get("LineItem.LineAmount"))
			if line:
				doc["amounts"].add(round(line, 2))
				doc["lines"].append(round(line, 2))
	return docs


def _float(value):
	try:
		return float(value or 0)
	except (ValueError, TypeError):
		return 0.0


def load_smartops_documents(doctype, company, date_from, date_to):
	"""Documents of one type in the date window, with their amounts and names"""
	spec = DOCTYPE_FIELDS.get(doctype)
	if not spec:
		frappe.throw(_("Matching is not supported for {0}. Supported: {1}").format(doctype, ", ".join(DOCTYPE_FIELDS)))

	meta = frappe.get_meta(doctype)
	amount_fields = [f for f in spec["amounts"] if meta.has_field(f)]
	name_fields = [f for f in spec["names"] if meta.has_field(f)]
	filters = {spec["date"]: ("between", [date_from, date_to]), "docstatus": ("<", 2)}
	if company and meta.has_field("company"):
		filters["company"] = company

	rows = frappe.get_all(doctype, filters=filters, fields=["name", spec["date"]] + amount_fields + name_fields, limit_page_length=0)
	docs = {}
	for row in rows:
		docs[row.name] = {
			"doctype": doctype, "name": row.name, "date": getdate(row[spec["date"]]),
			"amounts": {round(_float(row[f]), 2) for f in amount_fields if row[f]},
			"names": [row[f] for f in name_fields if row[f]],
		}

	for child_doctype, parentfield, fields in spec["children"]:
		if not docs or not frappe.db.exists("DocType", child_doctype):
			continue
		child_fields = [f for f in fields if frappe.get_meta(child_doctype).has_field(f)]
		if not child_fields:
			continue
		for child in frappe.get_all(child_doctype, filters={"parenttype": doctype, "parentfield": parentfield, "parent": ("in", list(docs))},
				fields=["parent"] + child_fields, limit_page_length=0):
			docs[child.parent]["amounts"].update(round(_float(child[f]), 2) for f in child_fields if child[f])

	return list(docs.values())


def match(xero_docs, smartops_docs, date_tolerance=0):
	"""
	Returns output rows (one per file). Each Xero document is placed with
	exactly one SmartOps document when date and amount single it out, listed
	under Review when several documents fit, and under No match otherwise
	"""
	by_date = defaultdict(list)
	for doc in smartops_docs:
		by_date[doc["date"]].append(doc)

	rows = []
	for xero in sorted(xero_docs.values(), key=lambda x: (x["date"], x["contact"])):
		candidates = []
		for offset in range(-date_tolerance, date_tolerance + 1):
			for doc in by_date.get(xero["date"] + timedelta(days=offset), []):
				if xero["amounts"] & doc["amounts"]:
					candidates.append((doc, names_agree(xero["contact"], doc["names"])))

		with_name = [doc for doc, agrees in candidates if agrees]
		if len(with_name) == 1:
			rows += output_rows(xero, with_name[0], CONFIDENCE_EXACT)
		elif with_name and is_split_bill(xero, with_name):
			# One Xero bill became one SmartOps document per line item, the
			# file belongs on each of them
			for doc in with_name:
				rows += output_rows(xero, doc, CONFIDENCE_EXACT_SHARED, f"Bill with {len(xero['lines'])} lines split over {len(with_name)} documents")
		elif len(candidates) == 1:
			rows += output_rows(xero, candidates[0][0], CONFIDENCE_EXACT_NAME_DIFFERS, "Contact name differs from the SmartOps document")
		elif with_name or candidates:
			pool = with_name or [doc for doc, _ in candidates]
			for doc in pool:
				rows += output_rows(xero, doc, CONFIDENCE_REVIEW, f"{len(pool)} documents fit this date and amount, keep one")
		else:
			rows += output_rows(xero, None, CONFIDENCE_NO_MATCH, "No document with this date and amount")

	# A SmartOps document claimed by several Xero documents needs a look too
	claimed = defaultdict(set)
	for r in rows:
		if r["Confidence"] in (CONFIDENCE_EXACT, CONFIDENCE_EXACT_NAME_DIFFERS, CONFIDENCE_EXACT_SHARED):
			claimed[(r["Document Type"], r["ID"])].add(r["_xero_id"])
	for r in rows:
		if len(claimed.get((r["Document Type"], r["ID"]), ())) > 1 and r["Confidence"] != CONFIDENCE_REVIEW:
			r["Confidence"] = CONFIDENCE_REVIEW
			r["Note"] = "Several Xero documents point at this SmartOps document"
	return rows


def is_split_bill(xero, docs):
	"""
	True when the bill has at least as many line items as candidate documents
	and every candidate's amount is one of those line amounts (not the total)
	"""
	lines = list(xero["lines"])
	if len(docs) > len(lines):
		return False
	for doc in docs:
		hit = next((l for l in lines if l in doc["amounts"]), None)
		if hit is None:
			return False
		lines.remove(hit)
	return True


def output_rows(xero, doc, confidence, note=""):
	return [{
		"Document Type": doc["doctype"] if doc else "", "ID": doc["name"] if doc else "",
		"AttachmentFolder": "/" + xero["folder"], "AttachmentFilenames": file_name, "Confidence": confidence,
		"Xero Date": xero["date"].isoformat(), "Xero Contact": xero["contact"], "Xero Total": xero["total"],
		"Xero Reference": xero["reference"], "SmartOps Date": doc["date"].isoformat() if doc else "",
		"SmartOps Name": "; ".join(str(n) for n in doc["names"][:2]) if doc else "", "Note": note, "_xero_id": xero["xero_id"],
	} for file_name in xero["files"]]


def summarize(rows):
	counts = defaultdict(int)
	for r in rows:
		counts[r["Confidence"]] += 1
	return counts


def write_csv(rows, path):
	with open(path, "w", newline="", encoding="utf-8") as f:
		writer = csv.DictWriter(f, fieldnames=OUTPUT_COLUMNS, extrasaction="ignore")
		writer.writeheader()
		writer.writerows(rows)
