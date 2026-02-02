from flask import Flask, request, jsonify, send_from_directory
from flask_cors import CORS
import pandas as pd
import traceback
import re
from datetime import datetime
import webbrowser
import threading
import time
import os

app = Flask(__name__)
CORS(app)

# =============================================================
# CONFIG (UPDATED FOR RENDER)
# =============================================================

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

EXCEL_PATH = os.path.join(BASE_DIR, "rga.xlsx.xlsx")
PINCODE_CSV = os.path.join(BASE_DIR, "pincode.csv")
SHEET_NAME = "RGA status"
LOG_FILE = os.path.join(BASE_DIR, "chatbot.log")

# os.makedirs(os.path.dirname(LOG_FILE), exist_ok=True)

COL_ADDRESS_1 = "Address 1"
COL_ADDRESS_2 = "Address 2"
COL_ADDRESS_3 = "Address 3"
COL_ADDRESS_4 = "Address 4"
COL_PINCODE = "Chatbot Pincode"
COL_UPDATED = "Updated On"

DF = None
PIN_DF = None

# =============================================================
# CHAT STATE
# =============================================================

MODE = None
WAIT_ACCOUNT = False
WAIT_UPDATE_CHOICE = False
WAIT_FIELD = None
LAST_ACC = None

TEMP_ADDR = {"A1": "", "A2": "", "A3": "", "A4": "", "PIN": ""}

# =============================================================
# RESET STATE
# =============================================================

def reset_chat_state():
    global MODE, WAIT_ACCOUNT, WAIT_UPDATE_CHOICE, WAIT_FIELD, LAST_ACC, TEMP_ADDR
    MODE = None
    WAIT_ACCOUNT = False
    WAIT_UPDATE_CHOICE = False
    WAIT_FIELD = None
    LAST_ACC = None
    TEMP_ADDR = {"A1": "", "A2": "", "A3": "", "A4": "", "PIN": ""}

# =============================================================
# LOGGING
# =============================================================

def write_log(msg, error=False):
    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    lvl = "ERROR" if error else "INFO"
    line = f"[{ts}] [{lvl}] {msg}"
    print(line)
    with open(LOG_FILE, "a", encoding="utf-8") as f:
        f.write(line + "\n")

def log_sop_response(text):
    clean = text.replace("<br>", " | ")
    write_log(f"BOT Response: {clean}")

# =============================================================
# LOAD EXCEL
# =============================================================

def load_excel():
    global DF
    DF = pd.read_excel(EXCEL_PATH, sheet_name=SHEET_NAME, engine="openpyxl")
    DF.columns = [c.strip() for c in DF.columns]
    DF = DF[DF["Account No"].notna()]
    DF["Account No"] = DF["Account No"].astype(int)

    for c in [
        COL_ADDRESS_1,
        COL_ADDRESS_2,
        COL_ADDRESS_3,
        COL_ADDRESS_4,
        COL_PINCODE,
        COL_UPDATED
    ]:
        if c not in DF.columns:
            DF[c] = ""

    write_log("Excel loaded successfully.")

# =============================================================
# LOAD PINCODE CSV
# =============================================================

def load_pincode_csv():
    global PIN_DF
    PIN_DF = pd.read_csv(PINCODE_CSV)
    PIN_DF.columns = [c.strip().lower() for c in PIN_DF.columns]

    required = {"state", "city", "pincode"}
    if not required.issubset(PIN_DF.columns):
        raise Exception("Pincode CSV must contain state, city, pincode")

    PIN_DF["state"] = PIN_DF["state"].astype(str).str.lower().str.strip()
    PIN_DF["city"] = PIN_DF["city"].astype(str).str.lower().str.strip()
    PIN_DF["pincode"] = PIN_DF["pincode"].astype(str).str.strip()

    write_log("Pincode CSV loaded successfully.")

# =============================================================
# HELPERS
# =============================================================

def get_row(acc):
    r = DF[DF["Account No"] == int(acc)]
    return None if r.empty else r.iloc[0]

def normalize(v):
    return str(v).replace("–", "-").replace("—", "-").strip().lower()

def is_kyc_case(status):
    return "kyc" in normalize(status)

def courier_link(name):
    if not name:
        return ""
    n = str(name).lower()
    if "blue" in n:
        return '<br><a href="https://bluedart.com" target="_blank">Track on Bluedart</a>'
    if "sequel" in n:
        return '<br><a href="https://sequelglobal.com" target="_blank">Track on Sequel</a>'
    return ""

# =============================================================
# SOP RESPONSE
# =============================================================

def build_sop(row):
    status_raw = row.get("Chat Bot-Status", "")
    status = normalize(status_raw)
    acc = int(row["Account No"])

    refund = str(row.get("Refund Date", "")).split(" ")[0]
    inv = str(row.get("Invoice Created Date", "")).split(" ")[0]
    disp = str(row.get("Dispatch Date", "")).split(" ")[0]
    delivered = str(row.get("Delivered Date", "")).split(" ")[0]

    logistics = row.get("Logistics", "")
    docket = row.get("Docket No", "")
    logistics1 = row.get("Logistics1", "")
    new_docket = row.get("New Docket No", "")
    returned = str(row.get("Returned Date", "")).split(" ")[0]
    redispatch = str(row.get("Redispatched Date", "")).split(" ")[0]
    remarks = row.get("Remarks", "")

    if status in ["dispatched", "dispatched from warehouse"]:
        return (
            f"Account: {acc}<br>"
            f"Refund Date: {refund}<br>"
            f"Invoice Created: {inv}<br>"
            f"Dispatched: {disp}<br>"
            f"Courier: {logistics}<br>"
            f"Docket: {docket}"
            f"{courier_link(logistics)}"
        )

    if status in ["redispatched", "re-dispatched"]:
        return (
            f"Account: {acc}<br>"
            f"Returned: {returned}<br>"
            f"Remarks: {remarks}<br>"
            f"Re-dispatched: {redispatch}<br>"
            f"Courier: {logistics1}<br>"
            f"New Docket: {new_docket}"
            f"{courier_link(logistics1)}"
        )

    EXTRA_LIST = [
        "redeemed at btq",
        "pincode - no services",
        "invoice yet to prepare",
        "ready for dispatch",
        "due to cash memo",
        "from store",
        "out of india",
        "unable to reach, due to incorrect address",
        "nap - issue"
    ]

    if status == "delivered to the customer":
        return (
            f"Account: {acc}<br>"
            f"Status: {status_raw}<br>"
            f"Delivered Date: {delivered}"
        )

    if status in EXTRA_LIST or is_kyc_case(status):
        return f"Account: {acc}<br>Status: {status_raw}"

    return f"Account: {acc}<br>Status: {status_raw}"

# =============================================================
# ADDRESS VALIDATION
# =============================================================

def validate_field(field, val):
    if re.search(r"[\U0001F600-\U0001FAFF]", val):
        return False, "Emojis are not allowed."

    if field in ["A1", "A2"]:
        if not re.fullmatch(r"[A-Za-z0-9 /-]+", val):
            return False, "Special characters are not allowed."

    if field in ["A3", "A4"]:
        if not re.fullmatch(r"[A-Za-z ]+", val):
            return False, "Only letters allowed."

    if field == "PIN":
        if not re.fullmatch(r"\d{6}", val):
            return False, "Pincode must be 6 digits."

        city = TEMP_ADDR["A3"].lower().strip()
        state = TEMP_ADDR["A4"].lower().strip()

        match = PIN_DF[
            (PIN_DF["pincode"] == val) &
            (PIN_DF["state"] == state) &
            (PIN_DF["city"].str.contains(city))
        ]

        if match.empty:
            return False, "Pincode does not match City & State."

    return True, ""

# =============================================================
# UPDATE ADDRESS
# =============================================================

def update_address():
    idx = DF.index[DF["Account No"] == LAST_ACC][0]

    DF.loc[idx, COL_ADDRESS_1] = TEMP_ADDR["A1"]
    DF.loc[idx, COL_ADDRESS_2] = TEMP_ADDR["A2"]
    DF.loc[idx, COL_ADDRESS_3] = TEMP_ADDR["A3"]
    DF.loc[idx, COL_ADDRESS_4] = TEMP_ADDR["A4"]
    DF.loc[idx, COL_PINCODE] = TEMP_ADDR["PIN"]
    DF.loc[idx, COL_UPDATED] = datetime.now().strftime("%d-%m-%Y %H:%M")

    DF.to_excel(EXCEL_PATH, sheet_name=SHEET_NAME, index=False, engine="openpyxl")
    write_log(f"Address updated for Account {LAST_ACC}")
    return "Address updated successfully."

# =============================================================
# CHATBOT API
# =============================================================

@app.route("/chatbot", methods=["POST"])
def chatbot():
    global MODE, WAIT_ACCOUNT, WAIT_UPDATE_CHOICE, WAIT_FIELD, LAST_ACC, TEMP_ADDR

    try:
        msg = request.json.get("message", "").strip()
        write_log(f"User message: {msg}")

        if msg.upper() in ["RGA", "ECOM"]:
            reset_chat_state()
            MODE = msg.upper()
            WAIT_ACCOUNT = True
            return jsonify({"reply": f"{MODE} selected. Enter Account Number."})

        if WAIT_FIELD:
            ok, err = validate_field(WAIT_FIELD, msg)
            if not ok:
                return jsonify({"reply": err})

            TEMP_ADDR[WAIT_FIELD] = msg
            order = ["A1", "A2", "A3", "A4", "PIN"]
            prompts = {
                "A1": "Enter Plot / Door:",
                "A2": "Enter Street:",
                "A3": "Enter City:",
                "A4": "Enter State:",
                "PIN": "Enter Pincode:"
            }

            i = order.index(WAIT_FIELD)
            if i < 4:
                WAIT_FIELD = order[i + 1]
                return jsonify({"reply": prompts[WAIT_FIELD]})

            WAIT_FIELD = None
            return jsonify({"reply": update_address()})

        if WAIT_UPDATE_CHOICE:
            if msg.lower() == "yes":
                WAIT_UPDATE_CHOICE = False
                WAIT_FIELD = "A1"
                TEMP_ADDR = {"A1": "", "A2": "", "A3": "", "A4": "", "PIN": ""}
                return jsonify({"reply": "Enter Plot / Door:"})

            if msg.lower() == "no":
                WAIT_UPDATE_CHOICE = False
                return jsonify({"reply": "Okay. Thank you. Have a great day."})

            return jsonify({"reply": "Please reply Yes or No."})

        if WAIT_ACCOUNT:
            if not msg.isdigit():
                return jsonify({"reply": "Invalid account number."})

            row = get_row(msg)
            if row is None:
                return jsonify({"reply": "Record not found."})

            LAST_ACC = int(msg)
            WAIT_ACCOUNT = False

            sop = build_sop(row)
            log_sop_response(sop)

            if MODE == "RGA" and is_kyc_case(row["Chat Bot-Status"]):
                WAIT_UPDATE_CHOICE = True
                followup = sop + "<br><br>Would you like to update the address?"
                log_sop_response(followup)
                return jsonify({
                    "reply": followup,
                    "ask_update": True
                })

            return jsonify({"reply": sop})

        return jsonify({"reply": "Please start by selecting RGA or ECOM."})

    except Exception:
        write_log(traceback.format_exc(), True)
        return jsonify({"reply": "Internal error occurred."})

@app.route("/", methods=["GET"])
def home():
    return send_from_directory(BASE_DIR, "web1.html")

# =============================================================
# SERVE STATIC FILES FROM ROOT
# =============================================================
# Serve any file from root
@app.route("/<path:filename>")
def serve_file(filename):
    return send_from_directory(BASE_DIR, filename)

@app.route("/")
def home():
    return send_from_directory(BASE_DIR, "web1.html")

@app.route("/chatbot", methods=["POST"])
def chatbot():
    try:
        msg = request.json.get("message", "")
        # Temporary demo response, replace with your actual chatbot logic
        return jsonify({"reply": f"You said: {msg}"})
    except Exception:
        traceback.print_exc()
        return jsonify({"reply": "Internal error occurred."})

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5050)
