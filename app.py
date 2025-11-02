#!/usr/bin/env python3
# -*- coding: utf-8 -*-

from flask import Flask, request, jsonify
import odoorpc
import csv
import io
from datetime import datetime
import os
import sys
from pathlib import Path
from dotenv import load_dotenv
import logging
from logging.handlers import RotatingFileHandler

# google sheets (optional)
try:
    import gspread
    from google.oauth2.service_account import Credentials
except Exception:
    gspread = None
    Credentials = None

# ---------------------------------------------------------------------
# ENV / CONFIG
# ---------------------------------------------------------------------
BASE_DIR = Path(__file__).parent
ENV_PATH = BASE_DIR / ".env"
load_dotenv(dotenv_path=ENV_PATH, override=False)

LOG_PATH = os.getenv("LOG_PATH", "shopify_payout.log")

# Odoo
ODOO_URL = os.getenv("ODOO_URL", "https://your-odoo-host.com")
ODOO_DB = os.getenv("ODOO_DB", "your_db")
ODOO_USER = os.getenv("ODOO_USER", "admin@example.com")
ODOO_PASSWORD = os.getenv("ODOO_PASSWORD", "your_password")
ODOO_PORT = int(os.getenv("ODOO_PORT", "8073"))
ODOO_PROTOCOL = os.getenv("ODOO_PROTOCOL", "jsonrpc+ssl")

# journals / accounts (MAKE THESE IN .env)
JOURNAL_PAYMENT_ID = int(os.getenv("JOURNAL_PAYMENT_ID", "53"))  # Shopify Payments Clearing Account B2C
JOURNAL_BANK_ID = int(os.getenv("JOURNAL_BANK_ID", "88"))        # CBA Business Transaction
JOURNAL_MISC_ID = int(os.getenv("JOURNAL_MISC_ID", "3"))         # Miscellaneous Operations
ACCOUNT_CLEARING_ID = int(os.getenv("ACCOUNT_CLEARING_ID", "522"))  # 102400 - Shopify Payments Clearing Account B2C (AU)
ACCOUNT_BANK_ID = int(os.getenv("ACCOUNT_BANK_ID", "633"))          # 101100 - CBA Business Transaction (AU)
ACCOUNT_FEE_ID = int(os.getenv("ACCOUNT_FEE_ID", "720"))            # 592000 - Merchant Account Fees (AU)

# GSheets
ENABLE_GSHEETS = os.getenv("ENABLE_GSHEETS", "false").lower() in ("1", "true", "yes", "y")
GSHEETS_SA_FILE = os.getenv("GSHEETS_SA_FILE", "./sa.json")
GSHEETS_SPREADSHEET_ID = os.getenv("GSHEETS_SPREADSHEET_ID", "")
GSHEETS_TAB = os.getenv("GSHEETS_TAB", "Logs")
GSHEETS_BATCH_SIZE = int(os.getenv("GSHEETS_BATCH_SIZE", "1"))

# Flask app
app = Flask(__name__)

# ---------------------------------------------------------------------
# LOGGING
# ---------------------------------------------------------------------
def setup_logging() -> logging.Logger:
    logger = logging.getLogger("shopify_payout")
    logger.setLevel(logging.INFO)
    logger.propagate = False

    for h in list(logger.handlers):
        logger.removeHandler(h)

    formatter = logging.Formatter(
        "%(asctime)sZ %(levelname)s %(message)s",
        "%Y-%m-%dT%H:%M:%S",
    )

    sh = logging.StreamHandler(sys.stdout)
    sh.setLevel(logging.INFO)
    sh.setFormatter(formatter)

    fh = RotatingFileHandler(LOG_PATH, maxBytes=2_000_000, backupCount=5, encoding="utf-8")
    fh.setLevel(logging.INFO)
    fh.setFormatter(formatter)

    logger.addHandler(sh)
    logger.addHandler(fh)
    return logger

logger = setup_logging()

# ---------------------------------------------------------------------
# GOOGLE SHEETS SINK
# ---------------------------------------------------------------------
class GSheetsSink:
    def __init__(self, sa_file: str, spreadsheet_id: str, tab_name: str = "Logs", batch_size: int = 1):
        if gspread is None or Credentials is None:
            raise RuntimeError("gspread/google-auth not installed. Run: pip install gspread google-auth")

        sa_path = (BASE_DIR / sa_file).resolve()
        scopes = ["https://www.googleapis.com/auth/spreadsheets"]
        self.creds = Credentials.from_service_account_file(str(sa_path), scopes=scopes)
        self.gc = gspread.authorize(self.creds)
        self.sh = self.gc.open_by_key(spreadsheet_id)
        self.ws = self._ensure_worksheet(tab_name)
        self.batch = []
        self.batch_size = max(1, int(batch_size))

        if len(self.ws.get_all_values()) == 0:
            self.ws.append_row(["ts_utc", "status", "payout_date", "order", "note"])

    def _ensure_worksheet(self, name: str):
        try:
            return self.sh.worksheet(name)
        except gspread.WorksheetNotFound:
            return self.sh.add_worksheet(title=name, rows=1000, cols=12)

    def log(self, *, status: str = "", payout_date: str = "", order: str = "", note: str = ""):
        ts = datetime.utcnow().isoformat(timespec="seconds")
        row = [ts, status, payout_date, order, note]
        self.batch.append(row)
        if len(self.batch) >= self.batch_size:
            self.flush()

    def flush(self):
        if not self.batch:
            return
        self.ws.append_rows(self.batch, value_input_option="RAW")
        self.batch.clear()

gsheets_sink = None
if ENABLE_GSHEETS:
    try:
        gsheets_sink = GSheetsSink(
            sa_file=GSHEETS_SA_FILE,
            spreadsheet_id=GSHEETS_SPREADSHEET_ID,
            tab_name=GSHEETS_TAB,
            batch_size=GSHEETS_BATCH_SIZE,
        )
        logger.info("Google Sheets logging enabled.")
    except Exception as e:
        logger.error(f"Google Sheets sink init failed: {e}")
        gsheets_sink = None

def log_to_sheets(status: str, payout_date: str = "", order: str = "", note: str = ""):
    if not gsheets_sink:
        return
    try:
        gsheets_sink.log(status=status, payout_date=payout_date, order=order, note=note)
    except Exception as e:
        logger.warning(f"GSheets log failed: {e}")

# ---------------------------------------------------------------------
# ODOO CONNECTION
# ---------------------------------------------------------------------
odoo = odoorpc.ODOO(ODOO_URL, port=ODOO_PORT, protocol=ODOO_PROTOCOL)
odoo.login(ODOO_DB, ODOO_USER, ODOO_PASSWORD)
logger.info("Connected to Odoo")

# ---------------------------------------------------------------------
# HELPERS
# ---------------------------------------------------------------------
def parse_payout_date(raw_date: str) -> str:
    """
    Try multiple formats:
    - 08/12/2025  (US style from Shopify exports)
    - 2025-08-12  (ISO, sometimes when user edits in Sheets/Excel)
    - 08/12/25    (short year)
    Always return 'YYYY-MM-DD'
    """
    if not raw_date:
        raise ValueError("empty date")

    raw_date = raw_date.strip()

    fmts = [
        "%m/%d/%Y",   # 08/12/2025
        "%Y-%m-%d",   # 2025-08-12
        "%d/%m/%Y",   # 12/08/2025 (just in case your export flips)
        "%m/%d/%y",   # 08/12/25
    ]

    for fmt in fmts:
        try:
            return datetime.strptime(raw_date, fmt).strftime("%Y-%m-%d")
        except ValueError:
            continue

    # last resort: try dateutil if you want, or raise
    raise ValueError(f"Cannot parse payout date: {raw_date}")

def find_sale_order_by_name(odoo_env, order_name: str, company_id: int | None = None):
    if not order_name:
        return None
    domain = [("name", "=", order_name)]
    if company_id:
        domain.append(("company_id", "=", company_id))  # <-- company scoped
    ids = odoo_env["sale.order"].search(domain, limit=1)
    return odoo_env["sale.order"].browse(ids[0]) if ids else None

def find_invoices_for_so(odoo_env, so, move_type, company_id: int | None = None):
    domain = [
        ("invoice_origin", "=", so.name),
        ("move_type", "=", move_type),
        ("state", "=", "posted"),
    ]
    if company_id:
        domain.append(("company_id", "=", company_id))  # <-- company scoped
    return odoo_env["account.move"].search(domain)

def find_payments_for_invoice(odoo_env, invoice_id, company_id: int | None = None):
    invoice = odoo_env["account.move"].browse(invoice_id)
    payment_moves = set()
    for line in invoice.line_ids:
        for p in line.matched_debit_ids:
            payment_moves.add(p.debit_move_id.move_id.id)
        for p in line.matched_credit_ids:
            payment_moves.add(p.credit_move_id.move_id.id)
    if not payment_moves:
        return []
    domain = [
        ("move_id", "in", list(payment_moves)),
        ("state", "in", ["posted"]),
    ]
    if company_id:
        domain.append(("company_id", "=", company_id))  # <-- company scoped
    return odoo_env["account.payment"].search(domain)


def append_internal_note_to_clearing_lines(
    odoo_env,
    *,
    payment_ids,
    clearing_account_id: int,
    payout_date_raw: str,
    order_num: str,
    company_id: int | None = None,
):
    """
    For each payment, find its move lines on the clearing account and
    append text to internal_notes.
    """
    if not payment_ids:
        return

    MoveLine = odoo_env["account.move.line"]
    Payment = odoo_env["account.payment"]

    for pay in Payment.browse(payment_ids):
        move_id = pay.move_id.id
        if not move_id:
            continue

        domain = [
            ("move_id", "=", move_id),
            ("account_id", "=", clearing_account_id),
        ]
        if company_id:
            domain.append(("company_id", "=", company_id))

        ml_ids = MoveLine.search(domain)
        if not ml_ids:
            continue

        for ml in MoveLine.browse(ml_ids):
            old = ml.internal_note or ""
            addition = f"{payout_date_raw} | {order_num}"
            if old:
                new_val = old + "\n" + addition
            else:
                new_val = addition
            MoveLine.write([ml.id], {"internal_note": new_val})

        return ml_ids

# ---------------------------------------------------------------------
# MAIN ROUTE
# ---------------------------------------------------------------------
@app.route("/upload_payout", methods=["POST"])
def upload_payout():
    if "file" not in request.files:
        return jsonify({"error": "No file uploaded"}), 400

    file = request.files["file"]

    # company_id from form
    company_id_raw = request.form.get("company_id")
    company_id = int(company_id_raw) if company_id_raw else None
    company_name = None
    if company_id:
        try:
            company = odoo.env["res.company"].browse(company_id)
            if company.exists():
                company_name = company.name
                logger.info(f"Processing payout for company: {company_name} (ID: {company_id})")
            else:
                logger.warning(f"Company with ID {company_id} not found in Odoo.")
        except Exception as e:
            logger.error(f"Failed to fetch company name for ID {company_id}: {e}")
    else:
        logger.warning("No company_id provided in request form.")

    try:
        stream = io.StringIO(file.stream.read().decode("utf-8"))
    except UnicodeDecodeError:
        stream = io.StringIO(file.stream.read().decode("latin-1"))

    reader = csv.DictReader(stream)

    AccountPayment = odoo.env["account.payment"]
    Move = odoo.env["account.move"]
    MoveLine = odoo.env["account.move.line"]

    total_amount = 0.0
    total_fee = 0.0
    total_net = 0.0
    total_refund = 0.0
    has_not_found = False
    refund_move_id = None
    odoo_date = None
    missing_rows = []
    all_move_line_ids = []
    for row in reader:
        payout_date_raw = (row.get("Payout Date") or "").strip()
        if not payout_date_raw:
            continue
        odoo_date = parse_payout_date(payout_date_raw)

        trx_type = (row.get("Type") or "").strip().lower()
        amount = float(row.get("Amount") or 0.0)
        fee = float(row.get("Fee") or 0.0)
        net = float(row.get("Net") or 0.0)
        order_num = (row.get("Order") or "").strip()
        
        if trx_type == "charge":
            found_payment = False

            # 1) SO first (scoped by company)
            so = find_sale_order_by_name(odoo.env, order_num, company_id=company_id)

            if so:
                invoice_ids = find_invoices_for_so(odoo.env, so, "out_invoice", company_id=company_id)

                payment_ids = []
                for inv_id in invoice_ids:
                    pay_ids = find_payments_for_invoice(odoo.env, inv_id, company_id=company_id)
                    if pay_ids:
                        payment_ids = pay_ids
                        break

                if payment_ids:
                    found_payment = True
                    ml_ids = append_internal_note_to_clearing_lines(
                        odoo.env,
                        payment_ids=payment_ids,
                        clearing_account_id=ACCOUNT_CLEARING_ID,
                        payout_date_raw=payout_date_raw,
                        order_num=order_num,
                        company_id=company_id,
                    )
                    if ml_ids:
                        all_move_line_ids = all_move_line_ids + ml_ids
                else:
                    logger.warning(f"SO {so.name} found but no payment found for its invoices (company={company_name})")

            # else:
                # 2) Fallback: date + amount + company
                # domain = [
                #     ("journal_id", "=", JOURNAL_PAYMENT_ID),
                #     ("date", "=", odoo_date),
                #     ("amount", "=", amount),
                #     ("state", "in", ["posted"]),
                # ]
                # if company_id:
                #     domain.append(("company_id", "=", company_id))  # <-- company scoped
                # payment_ids = AccountPayment.search(domain, limit=1)
                # if payment_ids:
                #     found_payment = True

            if not found_payment:
                has_not_found = True
                missing_rows.append(row)
                logger.warning(f"Payment not found for order={order_num} amount={amount} company={company_name}")
                log_to_sheets(
                    status="ORDER_NOT_FOUND",
                    payout_date=payout_date_raw,
                    order=order_num,
                    note=f"Transaction not found in {company_name}",
                )
            else:
                total_amount += amount
                total_net += net
                total_fee += fee

        elif trx_type == "refund":
            found_refund = False
            so = find_sale_order_by_name(odoo.env, order_num, company_id=company_id)

            if so:
                # 1️⃣ find credit notes for this SO
                refund_inv_ids = find_invoices_for_so(odoo.env, so, "out_refund", company_id=company_id)
                refund_ids = []
                if refund_inv_ids:
                    # 2️⃣ check payments for these credit notes
                    for refund_id in refund_inv_ids:
                        pay_ids = find_payments_for_invoice(odoo.env, refund_id, company_id=company_id)
                        if pay_ids:
                            refund_ids = pay_ids
                            found_refund = True
                            break

                    if found_refund:
                        total_refund += abs(net)
                        ml_ids = append_internal_note_to_clearing_lines(
                            odoo.env,
                            payment_ids=refund_ids,
                            clearing_account_id=ACCOUNT_CLEARING_ID,
                            payout_date_raw=payout_date_raw,
                            order_num=order_num,
                            company_id=company_id,
                        )
                        if ml_ids:
                            all_move_line_ids = all_move_line_ids + ml_ids
                    else:
                        logger.warning(f"SO {so.name} refund found but no payment matched (company={company_name})")
                        has_not_found = True
                        missing_rows.append(row)
                        log_to_sheets(
                            status="NOT_FOUND_REFUND",
                            payout_date=payout_date_raw,
                            order=order_num,
                            note=f"Refund credit note found but no payment for company {company_name}"
                        )
                else:
                    logger.warning(f"No credit note found for refund SO {so.name} (company={company_name})")
                    has_not_found = True
                    missing_rows.append(row)
                    log_to_sheets(
                        status="NOT_FOUND_REFUND",
                        payout_date=payout_date_raw,
                        order=order_num,
                        note=f"No credit note found for refund in company {company_name}"
                    )
            else:
                logger.warning(f"Refund row with order {order_num} has no matching SO (company={company_name})")
                has_not_found = True
                missing_rows.append(row)
                log_to_sheets(
                    status="NOT_FOUND_REFUND",
                    payout_date=payout_date_raw,
                    order=order_num,
                    note=f"No SO found for refund in company {company_name}"
                )

    # 1) create refund move if needed
    # if total_refund > 0 and odoo_date:
    #     move_vals = {
    #         "move_type": "entry",
    #         "journal_id": JOURNAL_MISC_ID,
    #         "date": odoo_date,
    #         "ref": f"Shopify Refund {odoo_date}",
    #         "line_ids": [
    #             (0, 0, {
    #                 "name": f"Shopify Refund {odoo_date}",
    #                 "account_id": ACCOUNT_CLEARING_ID,
    #                 "debit": total_refund,
    #                 "credit": 0.0,
    #             }),
    #             (0, 0, {
    #                 "name": f"Shopify Refund {odoo_date}",
    #                 "account_id": ACCOUNT_BANK_ID,
    #                 "debit": 0.0,
    #                 "credit": total_refund,
    #             }),
    #         ],
    #     }
    #     if company_id:
    #         move_vals["company_id"] = company_id  # <-- company scoped create

    #     refund_move_id = Move.create(move_vals)
    #     refund_move = Move.browse(refund_move_id)
    #     refund_move.action_post()
    #     logger.info(f"Created refund entry {refund_move_id} for total_refund={total_refund} company_id={company_id}")

    if has_not_found:
        return jsonify({
            "error": "Some payments were not found in Odoo. Check sheets/log.",
            "not_found_count": len(missing_rows),
            "refund_move_id": refund_move_id,
        }), 400

    if not odoo_date:
        return jsonify({"error": "No valid payout date found in CSV"}), 400

    # find clearing lines to credit (unreconciled)
    domain_pay_lines = [
        ("journal_id", "=", JOURNAL_PAYMENT_ID),
        ("account_id", "=", ACCOUNT_CLEARING_ID),
        ("reconciled", "=", False),
        ("debit", ">", 0),
        ("parent_state", "=", "posted"),
    ]
    if company_id:
        domain_pay_lines.append(("company_id", "=", company_id))  # <-- company scoped

    # payment_lines = MoveLine.search(domain_pay_lines)
    payments = MoveLine.browse(all_move_line_ids)
    matched_lines = []
    total_credit = 0.0
    tolerance = 0.01

    for line in payments:
        line_amt = float(line.debit or 0.0)
        if total_credit + line_amt <= total_amount + tolerance:
            total_credit += line_amt
            matched_lines.append(line)
        # if total_credit >= total_amount - tolerance:
        #     break

    if not matched_lines:
        logger.warning("No matching clearing payment lines found for reconciliation.")
        resp = {
            "error": "No matching payment lines found for reconciliation.",
            "refund_move_id": refund_move_id,
        }
        return jsonify(resp), 400

    move_vals = {
        "move_type": "entry",
        "journal_id": JOURNAL_BANK_ID,
        "date": odoo_date,
        "ref": f"Shopify Payout {odoo_date}",
        "line_ids": [],
    }
    if company_id:
        move_vals["company_id"] = company_id  # <-- company scoped create

    # debit bank with NET
    move_vals["line_ids"].append(
        (0, 0, {
            "name": f"Shopify Payout {odoo_date}",
            "account_id": ACCOUNT_BANK_ID,
            "debit": total_net,
            "credit": 0.0,
        })
    )

    if total_fee > 0:
        move_vals["line_ids"].append(
            (0, 0, {
                "name": f"Shopify Fee {odoo_date}",
                "account_id": ACCOUNT_FEE_ID,
                "debit": total_fee,
                "credit": 0.0,
            })
        )

    for line in matched_lines:
        move_vals["line_ids"].append(
            (0, 0, {
                "name": line.name or "Shopify Payment",
                "account_id": line.account_id.id,
                "debit": float(line.credit or 0.0),
                "credit": float(line.debit or 0.0),
                "partner_id": line.partner_id.id if line.partner_id else False,
            })
        )

    if total_refund > 0:
        move_vals["line_ids"].append(
            (0, 0, {
                    'name': f'Shopify Refund {odoo_date}',
                    'account_id': ACCOUNT_BANK_ID,
                    'debit': 0.0,
                    'credit': total_refund,
                })
        )

    logger.info("Creating payout move...")
    logger.info(move_vals)
    try :
        move_id = Move.create(move_vals)
        move = Move.browse(move_id)
        move.action_post()

        # reconcile clearing lines
        debit_lines_domain = [
            ("move_id", "=", move.id),
            # ("credit", ">", 0),
            ("account_id.reconcile", "=", True),
            ("account_id", "=", ACCOUNT_CLEARING_ID),
            ("reconciled", "=", False),      
        ]
        # refund_lines_domain = [
        #     ("move_id", "=", move.id),
        #     ("name", "=", f"Shopify Refund {odoo_date}"),
        #     ("reconciled", "=", False),
        # ]
        # company filter not needed here, move_id is unique
    
        debit_lines = MoveLine.search(debit_lines_domain)
        # refund_lines = MoveLine.search(refund_lines_domain)
        move_line_ids = [l.id for l in matched_lines if not l.reconciled] + debit_lines 
        move_line_ids = list(set(move_line_ids))
        reconcile_lines = MoveLine.browse(move_line_ids)
        reconcile_lines.reconcile()
    except Exception as e:
        log_to_sheets(
                            status="ERORR RECONCILIATION",
                            payout_date=odoo_date,
                            order=f"",
                            note=f"{e}"
                        )
        logger.error(f"Error reconciling lines: {e}")
        return jsonify({"error": f"Error reconciling lines: {e}"}), 500

    resp = {
        "success": True,
        "message": f"Reconciliation done for {odoo_date}",
        "total_net": total_net,
        "matched_credit": total_credit,
        "move_id": move.id,
        "refund_move_id": refund_move_id,
        "company_id": company_id,
    }
    return jsonify(resp), 200


# ---------------------------------------------------------------------
# ENTRYPOINT
# ---------------------------------------------------------------------
if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5100, debug=True)