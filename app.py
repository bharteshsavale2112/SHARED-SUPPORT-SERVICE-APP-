from flask import Flask, request, jsonify, send_from_directory, session
from functools import wraps
from datetime import datetime, timezone, timedelta
from werkzeug.security import generate_password_hash, check_password_hash
from flask import request, jsonify
import openpyxl
import pandas as pd
import json
import os

app = Flask(__name__)

# ==========================
# SECURITY CONFIG
# ==========================
# Secret key used to sign session cookies. Set SECRET_KEY as an
# environment variable on Render (Settings -> Environment) in production.
app.secret_key = os.environ.get("SECRET_KEY", "dev-secret-change-me")

# ---- Cookie / Session hardening ----
# HTTPONLY: JavaScript can't read the session cookie (protects against XSS
#           stealing the login session)
# SAMESITE: cookie is not sent on cross-site requests (protects against CSRF)
# SECURE:   cookie is only sent over HTTPS. Render serves the site over
#           HTTPS, so this is safe to enable in production. If you ever
#           test locally over plain http://127.0.0.1, set
#           FLASK_DEBUG=True (which also flips this off automatically).
app.config["SESSION_COOKIE_HTTPONLY"] = True
app.config["SESSION_COOKIE_SAMESITE"] = "Lax"
app.config["SESSION_COOKIE_SECURE"] = os.environ.get("FLASK_DEBUG", "False") != "True"

# Auto logout after 2 hours of inactivity
app.config["PERMANENT_SESSION_LIFETIME"] = timedelta(hours=2)

# Admin credentials now come from environment variables instead of being
# hardcoded in the source code (which is visible to anyone with repo access).
# Set ADMIN_EMAIL and ADMIN_PASSWORD as environment variables on Render.
ADMIN_EMAIL = os.environ.get("ADMIN_EMAIL", "admin@bajaj.com")
ADMIN_PASSWORD_HASH = os.environ.get("ADMIN_PASSWORD_HASH") or generate_password_hash(
    os.environ.get("ADMIN_PASSWORD", "admin123")
)


def login_required(role=None):
    """Blocks access to a route unless the user is logged in
    (and, if `role` is given, logged in as that specific role)."""
    def decorator(f):
        @wraps(f)
        def wrapper(*args, **kwargs):
            if "role" not in session:
                return jsonify({"status": "error", "message": "Login Required"}), 401
            if role and session.get("role") != role:
                return jsonify({"status": "error", "message": "Not Authorized"}), 403
            return f(*args, **kwargs)
        return wrapper
    return decorator


EMPLOYEE_FILE = "employees.xlsx"

# ==========================
# CREATE EXCEL FILE SAFELY
# ==========================

def create_excel_if_needed():

    cols = [
    "employeeCode",
    "employeeName",
    "mobile",
    "department",
    "email",
    "shift",
    "busStop",
    "routeNumber",
    "address",
    "password",
    "route"
]

    if not os.path.exists(EMPLOYEE_FILE):
        pd.DataFrame(columns=cols).to_excel(EMPLOYEE_FILE, index=False)
        return

    try:
        df = pd.read_excel(EMPLOYEE_FILE)
        df.columns = df.columns.str.strip()

        for col in cols:
            if col not in df.columns:
                df[col] = ""

        df = df[cols]
        df.to_excel(EMPLOYEE_FILE, index=False)

    except Exception:
        pd.DataFrame(columns=cols).to_excel(EMPLOYEE_FILE, index=False)

create_excel_if_needed()



# ==========================
# LOAD EXCEL SAFELY
# ==========================

def load_df():
    create_excel_if_needed()   # <-- ही नवीन line add कर

    df = pd.read_excel(EMPLOYEE_FILE)
    df.columns = df.columns.str.strip()

    return df

# ==========================
# HOME
# ==========================

@app.route("/")
def home():
    return """
    <h2>🚍 Bajaj Transport Backend Running</h2>
    <a href="/employeeregistration.html">Employee Registration</a><br>
    <a href="/employeelogin.html">Employee Login</a><br>
    <a href="/adminlogin.html">Admin Login</a>
    """

# ==========================
# SERVE HTML FILES
# ==========================

@app.route("/<path:path>")
def serve(path):
    return send_from_directory(".", path)








# ==========================
# EMPLOYEE REGISTRATION
# ==========================

@app.route("/employee-registration", methods=["POST"])
def register():

    try:

        df = load_df()

        data = request.form.to_dict()

        # Required Fields
        required = [
            "employeeCode",
            "employeeName",
            "mobile",
            "department",
            "email",
            "shift",
            "busStop",
            "routeNumber",
            "address",
            "password"
        ]

        for field in required:

            if field not in data or str(data[field]).strip() == "":
                return f"<script>alert('{field} is Required');history.back()</script>"

        # Duplicate Employee Code
        if str(data["employeeCode"]).strip() in df["employeeCode"].astype(str).str.strip().values:

            return "<script>alert('Employee Code Already Registered');history.back()</script>"

        # Duplicate Email
        if str(data["email"]).strip().lower() in df["email"].astype(str).str.strip().str.lower().values:

            return "<script>alert('Email Already Registered');history.back()</script>"

        # Duplicate Mobile
        if str(data["mobile"]).strip() in df["mobile"].astype(str).str.strip().values:

            return "<script>alert('Mobile Number Already Registered');history.back()</script>"

        # Default Route
        data["route"] = ""

        # Hash the password before saving so it's never stored in plain text
        data["password"] = generate_password_hash(str(data["password"]))

        # Save Employee
        df = pd.concat(
            [df, pd.DataFrame([data])],
            ignore_index=True
        )

        df.to_excel(EMPLOYEE_FILE, index=False)

        return """
        <script>
        alert('Registration Successful');
        window.location.href='/employeelogin.html';
        </script>
        """

    except Exception as e:

        print(e)

        return f"""
        <script>
        alert('{str(e)}');
        history.back();
        </script>
        """
# ==========================
# LOGIN
# ==========================

@app.route("/employee-login", methods=["POST"])
def employee_login():

    try:

        data = request.get_json()

        if not data:
            return jsonify({
                "status": "error",
                "message": "Invalid Request"
            })

        employee_id = str(data.get("employeeId", "")).strip()
        password = str(data.get("password", "")).strip()

        if employee_id == "" or password == "":
            return jsonify({
                "status": "error",
                "message": "Employee ID and Password Required"
            })

        # Load Excel Data
        df = load_df()

        # Find Employee
        user = df[
            df["employeeCode"].astype(str).str.strip() == employee_id
        ]

        if user.empty:
            return jsonify({
                "status": "error",
                "message": "Employee ID Not Found"
            })

        # Employee Record
        row = user.iloc[0]

        # Password Check (supports old plain-text passwords too, and
        # upgrades them to a secure hash automatically on next login)
        db_password = str(row["password"]).strip()

        if db_password.startswith(("pbkdf2:", "scrypt:")):
            password_ok = check_password_hash(db_password, password)
        else:
            password_ok = (db_password == password)
            if password_ok:
                # Legacy plain-text password found; upgrade it to a hash now
                df.loc[user.index, "password"] = generate_password_hash(password)
                df.to_excel(EMPLOYEE_FILE, index=False)

        if not password_ok:
            return jsonify({
                "status": "error",
                "message": "Wrong Password"
            })

        # Login Success - mark this browser session as logged in
        session.permanent = True
        session["role"] = "employee"
        session["employeeCode"] = employee_id

        return jsonify({

            "status": "success",
            "message": "Login Successful",

            "passNumber": str(row.get("passNumber", "")),
            "employeeCode": str(row.get("employeeCode", "")).strip(),
            "employeeName": str(row.get("employeeName", "")),
            "department": str(row.get("department", "")),
            "mobile": str(row.get("mobile", "")),
            "shift": str(row.get("shift", "")),
            "routeNumber": str(row.get("routeNumber", "")),
            "busStop": str(row.get("busStop", "")),
            "issueDate": str(row.get("issueDate", "")),
            "passStatus": str(row.get("status", ""))

        })

    except Exception as e:

        return jsonify({
            "status": "error",
            "message": str(e)
        })
            
# ==========================
# ADMIN LOGIN
# ==========================

@app.route("/admin-login", methods=["POST"])
def admin_login():

    data = request.get_json()

    email = str(data.get("email", "")).strip()
    password = str(data.get("password", "")).strip()

    if email == ADMIN_EMAIL and check_password_hash(ADMIN_PASSWORD_HASH, password):
        session.permanent = True
        session["role"] = "admin"
        return jsonify({
            "status": "success",
            "message": "Login Successful"
        })

    return jsonify({
        "status": "error",
        "message": "Wrong Credentials"
    })


@app.route("/logout", methods=["POST"])
def logout():
    session.clear()
    return jsonify({"status": "success", "message": "Logged Out"})


@app.route("/check-session", methods=["GET"])
def check_session():
    """Used by pages to verify the login session is still valid
    (e.g. after using the browser Back/Forward buttons post-logout)."""
    if "role" not in session:
        return jsonify({"status": "error", "message": "Not Logged In"}), 401
    return jsonify({
        "status": "success",
        "role": session.get("role"),
        "employeeCode": session.get("employeeCode", "")
    })


@app.after_request
def add_no_cache_headers(response):
    """Stops the browser from serving a cached copy of protected pages
    from its Back/Forward cache after the user has logged out."""
    response.headers["Cache-Control"] = "no-store, no-cache, must-revalidate, max-age=0"
    response.headers["Pragma"] = "no-cache"
    return response





# ==========================
# ADMIN DASHBOARD STATS
# ==========================

@app.route("/dashboard-stats")
@login_required("admin")
def dashboard_stats():

    try:

        emp = load_df()

        totalEmployees = len(emp)

        totalRoutes = emp["route"].replace("", pd.NA).dropna().nunique()

        pendingRequests = 0
        activePasses = 0

        if os.path.exists(PASS_REQUEST_FILE):

            req = pd.read_excel(PASS_REQUEST_FILE)

            pendingRequests = len(
                req[req["status"] == "Pending"]
            )

        if os.path.exists(PASS_FILE):

            passes = pd.read_excel(PASS_FILE)

            activePasses = len(passes)

        return jsonify({

            "totalEmployees": totalEmployees,

            "totalRoutes": totalRoutes,

            "pendingRequests": pendingRequests,

            "activePasses": activePasses

        })

    except Exception as e:

        return jsonify({

            "status": "error",

            "message": str(e)

        })
    

# ==========================
# RECENT EMPLOYEES API
# ==========================

@app.route("/recent-employees")
@login_required("admin")
def recent_employees():

    try:

        df = load_df()

        # शेवटचे 5 Employees
        df = df.tail(5)

        employees = []

        for _, row in df.iterrows():

            employees.append({

                "employeeCode": str(row["employeeCode"]),

                "employeeName": str(row["employeeName"]),

                "department": str(row["department"]),

                "route": str(row["route"])

            })

        return jsonify(employees)

    except Exception as e:

        return jsonify({
            "status":"error",
            "message":str(e)
        })


# ==========================
# VIEW EMPLOYEES
# ==========================

@app.route("/employees")
@login_required("admin")
def employees():

    try:

        df = load_df()

        if "password" in df.columns:
            df = df.drop(columns=["password"])

        return df.to_html(index=False)

    except Exception as e:

        return f"<h3>Error : {str(e)}</h3>"
    

# ==========================
# ASSIGN ROUTE
# ==========================

@app.route("/assign-route", methods=["POST"])
@login_required("admin")
def assign():

    try:

        df = load_df()

        code = str(request.form.get("employeeCode", "")).strip()
        route = str(request.form.get("route", "")).strip()

        # Empty Validation
        if code == "" or route == "":
            return jsonify({
                "success": False,
                "message": "Employee Code and Route Required"
            })

        # Employee Exists?
        if code not in df["employeeCode"].astype(str).str.strip().values:
            return jsonify({
                "success": False,
                "message": "Employee Not Found"
            })

        # Assign Route
        df.loc[
            df["employeeCode"].astype(str).str.strip() == code,
            "route"
        ] = route

        df.to_excel(EMPLOYEE_FILE, index=False)

        return jsonify({
            "success": True,
            "message": "Route Assigned Successfully"
        })

    except Exception as e:

        return jsonify({
            "success": False,
            "message": str(e)
        })
    
# ==========================
# GPS STORAGE (LIVE TRACKING)
# --------------------------------------------------------------------
# Admin dashboard pushes lat/lng here (POST /update-live-location).
# The employee's live map (viewroots.html) polls
# GET /live-location/<shift>/<route> every ~2 seconds and moves the
# bus icon when a new position comes in.
#
# Stored in a JSON file (not just an in-memory dict) so it survives a
# server restart, and keyed by shift+route together because the same
# route name (e.g. "Route 1 (SINHGAD)") is reused across different
# shifts.
# ==========================

LIVE_LOCATIONS_FILE = "live_locations.json"

# A ping older than this is treated as stale and NOT shown as "live"
# (stops yesterday's last position from looking like a live bus today).
STALE_AFTER_SECONDS = 15 * 60  # 15 minutes


def _live_key(shift, route):
    return f"{shift}||{route}"


def _load_live_locations():
    if os.path.exists(LIVE_LOCATIONS_FILE):
        try:
            with open(LIVE_LOCATIONS_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except (json.JSONDecodeError, OSError):
            return {}
    return {}


def _save_live_locations(data):
    with open(LIVE_LOCATIONS_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


def _location_is_stale(updated_at_str):
    try:
        updated_at = datetime.fromisoformat(updated_at_str)
        age_seconds = (datetime.now(timezone.utc) - updated_at).total_seconds()
        return age_seconds > STALE_AFTER_SECONDS
    except Exception:
        return False


@app.route("/update-live-location", methods=["POST"])
@login_required("admin")
def update_live_location():

    try:

        data = request.get_json(silent=True)

        if not data:
            return jsonify({"success": False, "message": "Invalid Data"})

        shift = str(data.get("shift", "")).strip()
        route = str(data.get("route", "")).strip()

        if not shift or not route:
            return jsonify({"success": False, "message": "Shift and Route are required"})

        try:
            lat = float(data.get("lat"))
            lng = float(data.get("lng"))
        except (TypeError, ValueError):
            return jsonify({"success": False, "message": "lat/lng must be numbers"})

        locations = _load_live_locations()

        locations[_live_key(shift, route)] = {
            "lat": lat,
            "lng": lng,
            "updatedAt": datetime.now(timezone.utc).isoformat()
        }

        _save_live_locations(locations)

        return jsonify({
            "success": True,
            "message": "Location Updated"
        })

    except Exception as e:

        return jsonify({
            "success": False,
            "message": str(e)
        })


# ---- Employee side: shift + route qualified (preferred) ----
@app.route("/live-location/<shift>/<route>")
def get_live_location(shift, route):

    locations = _load_live_locations()
    entry = locations.get(_live_key(shift, route))

    if not entry:
        return jsonify({"success": False, "message": "Location Not Found"})

    if _location_is_stale(entry.get("updatedAt", "")):
        return jsonify({"success": False, "message": "Location is stale"})

    return jsonify({
        "success": True,
        "lat": entry["lat"],
        "lng": entry["lng"],
        "updatedAt": entry.get("updatedAt", "")
    })


# ---- Backward-compatible: route only, no shift ----
# Only kept in case an older page still calls it this way. Prefer the
# shift-qualified route above since route names repeat across shifts.
@app.route("/live-location/<route>")
def get_live_location_legacy(route):

    locations = _load_live_locations()

    matches = [
        (key, entry) for key, entry in locations.items()
        if key.split("||", 1)[-1] == route
    ]

    if len(matches) != 1:
        return jsonify({"success": False, "message": "Route Not Found or Ambiguous Across Shifts"})

    _, entry = matches[0]

    if _location_is_stale(entry.get("updatedAt", "")):
        return jsonify({"success": False, "message": "Location is stale"})

    return jsonify({
        "success": True,
        "lat": entry["lat"],
        "lng": entry["lng"],
        "updatedAt": entry.get("updatedAt", "")
    })







# ==========================
# PASS FILE
# ==========================

PASS_FILE = "passes.xlsx"

# ==========================
# CREATE PASSES EXCEL
# ==========================

def create_pass_file():

    cols=[

"passNumber",

"employeeCode",

"employeeName",

"department",

"mobile",

"shift",

"routeNumber",

"busStop",

"issueDate",

"status"

]
    if os.path.exists(PASS_FILE):

        try:

            df = pd.read_excel(PASS_FILE)

            df.columns = df.columns.str.strip()

            for col in cols:

                if col not in df.columns:

                    df[col] = ""

            df = df[cols]

            df.to_excel(
                PASS_FILE,
                index=False
            )

        except Exception:

            pd.DataFrame(columns=cols).to_excel(
                PASS_FILE,
                index=False
            )

    else:

        pd.DataFrame(columns=cols).to_excel(
            PASS_FILE,
            index=False
        )

create_pass_file()

# ==========================
# BUS PASS API
# ==========================

@app.route("/bus-pass", methods=["POST"])
@login_required("employee")
def bus_pass():

    try:

        data = request.get_json()

        employeeCode = data.get("employeeCode")

        if str(session.get("employeeCode", "")).strip() != str(employeeCode).strip():
            return jsonify({"status": "error", "message": "Not Authorized"}), 403

        df = load_df()

        user = df[
            df["employeeCode"].astype(str) == str(employeeCode)
        ]

        if user.empty:

            return jsonify({
                "status":"error",
                "message":"Employee Not Found"
            })

        pass_df = pd.read_excel(PASS_FILE)

        month = datetime.now().strftime("%B")
        year = datetime.now().year

        already = pass_df[
            (pass_df["employeeCode"].astype(str) == str(employeeCode)) &
            (pass_df["month"] == month) &
            (pass_df["year"] == year)
        ]

        if not already.empty:

            return jsonify({
                "status":"success",
                "message":"Pass Already Generated"
            })

        new_pass = {
            "employeeCode": user.iloc[0]["employeeCode"],
            "employeeName": user.iloc[0]["employeeName"],
            "route": user.iloc[0]["route"],
            "busStop": user.iloc[0]["busStop"],
            "month": month,
            "year": year,
            "created_at": datetime.now().strftime("%d-%m-%Y %H:%M")
        }

        pass_df = pd.concat(
            [pass_df, pd.DataFrame([new_pass])],
            ignore_index=True
        )

        pass_df.to_excel(PASS_FILE, index=False)

        return jsonify({
            "status":"success",
            "message":"Bus Pass Generated Successfully"
        })

    except Exception as e:

        return jsonify({
            "status":"error",
            "message":str(e)
        })

# ==========================
# CHECK PASS API
# ==========================


@app.route("/check-pass/<employeeCode>")
@login_required("employee")
def check_pass(employeeCode):

    if str(session.get("employeeCode", "")).strip() != str(employeeCode).strip():
        return jsonify({"status": "error", "message": "Not Authorized"}), 403

    df = pd.read_excel(PASS_FILE)

    user = df[
        df["employeeCode"].astype(str) == str(employeeCode)
    ]

    if user.empty:

        return jsonify({
            "hasPass": False
        })

    return jsonify({

        "hasPass": True,

        "pass": user.iloc[0].to_dict()

    })


# ==========================
#pp request
# ==========================


import uuid
PASS_REQUEST_FILE = "pass_requests.xlsx"

# ==========================
# CREATE PASS REQUEST FILE
# ==========================

# ==========================
# APPROVE PASS
# ==========================

@app.route("/approve-pass", methods=["POST"])
@login_required("admin")
def approve_pass():

    try:
        data = request.get_json()
        requestId = str(data.get("requestId", "")).strip()

        # Load request file
        req = pd.read_excel(PASS_REQUEST_FILE)
        req = req.fillna("")

        req["requestId"] = req["requestId"].astype(str)

        row = req[req["requestId"] == requestId]

        if row.empty:
            return jsonify({
                "status": "error",
                "message": "Request Not Found"
            })

        index = row.index[0]

        # Update request
        req.loc[index, "status"] = "Approved"
        req.loc[index, "approvedBy"] = "Admin"
        req.loc[index, "approvedDate"] = datetime.now().strftime("%d-%m-%Y %H:%M")

        req.to_excel(PASS_REQUEST_FILE, index=False)

        # Load pass file
        passes = pd.read_excel(PASS_FILE)
        passes = passes.fillna("")

        # FIXED COLUMN NAME (IMPORTANT)
        newPass = {

    "passNumber":"PP"+datetime.now().strftime("%Y%m%d%H%M%S"),

    "employeeCode":row.iloc[0]["employeeCode"],

    "employeeName":row.iloc[0]["employeeName"],

    "department":row.iloc[0]["department"],

    "mobile":row.iloc[0]["mobile"],

    "shift":row.iloc[0]["shift"],

    "routeNumber":row.iloc[0]["routeNumber"],

    "busStop":row.iloc[0]["busStop"],

    "issueDate":datetime.now().strftime("%d-%m-%Y"),

    "status":"ACTIVE"

}
        passes = pd.concat([passes, pd.DataFrame([newPass])], ignore_index=True)
        passes.to_excel(PASS_FILE, index=False)

        return jsonify({
            "status": "success",
            "message": "Pass Generated Successfully"
        })

    except Exception as e:

        print("APPROVE PASS ERROR:", e)

        return jsonify({

            "status":"error",

            "message":str(e)

        })

 # ==========================
# GET MY PASS
# ==========================
@app.route("/my-pass/<employeeCode>", methods=["GET"])
@login_required("employee")
def my_pass(employeeCode):

    if str(session.get("employeeCode", "")).strip() != str(employeeCode).strip():
        return jsonify({"status": "error", "message": "Not Authorized"}), 403

    try:

        print("Employee Code from URL:", employeeCode)

        df = pd.read_excel(PASS_FILE)
        df.columns = df.columns.str.strip()

        print(df[["employeeCode"]])

        user = df[
            df["employeeCode"].astype(str).str.strip() ==
            str(employeeCode).strip()
        ]

        print("Matched Rows:", len(user))

        if user.empty:
            return jsonify({
                "status": "error",
                "message": "Pass Not Found"
            })

        row = user.iloc[0]

        return jsonify({
            "status": "success",
            "employeeCode": str(row["employeeCode"]),
            "employeeName": str(row["employeeName"]),
            "department": str(row["department"]),
            "mobile": str(row["mobile"]),
            "shift": str(row["shift"]),
            "routeNumber": str(row["routeNumber"]),
            "route": str(row["routeNumber"]),
            "busStop": str(row["busStop"]),
            "issueDate": str(row["issueDate"]),
            "passStatus": str(row["status"])
        })

    except Exception as e:
        return jsonify({
            "status": "error",
            "message": str(e)
        })
        
# ==========================
# CREATE PASS FILE
# ==========================

def create_pass_file():

    if not os.path.exists(PASS_FILE):

        pd.DataFrame(columns=[

            "passNumber",

            "employeeCode",

            "employeeName",

            "department",

            "mobile",

            "shift",

            "routeNumber",

            "busStop",

            "issueDate",

            "status"

        ]).to_excel(
            PASS_FILE,
            index=False
        )


create_pass_file()




# ==========================
# REQUEST PERMANENT PASS
# ==========================

@app.route("/request-pass", methods=["POST"])
@login_required("employee")
def request_pass():

    try:

        data = request.get_json()

        employeeCode = str(
            data.get("employeeCode", "")
        ).strip()

        if employeeCode == "":

            return jsonify({
                "status": "error",
                "message": "Employee Code Required"
            })

        if str(session.get("employeeCode", "")).strip() != str(employeeCode).strip():
            return jsonify({"status": "error", "message": "Not Authorized"}), 403

        # Employee Data
        df = load_df()

        user = df[
            df["employeeCode"].astype(str).str.strip()
            == employeeCode
        ]

        if user.empty:

            return jsonify({

                "status": "error",

                "message": "Employee Not Found"

            })

        # Pass Request File
        req = pd.read_excel(PASS_REQUEST_FILE)
        req.columns = req.columns.str.strip()

        # Check Pending Request
        old = req[

            (req["employeeCode"].astype(str).str.strip() == employeeCode)

            &

            (req["status"].astype(str).str.strip() == "Pending")

        ]

        if not old.empty:

            return jsonify({

                "status": "error",

                "message": "Request Already Pending"

            })

        # Generate Request ID
        requestId = "REQ" + datetime.now().strftime("%Y%m%d%H%M%S")

        # New Request
        newRow = {

            "requestId": requestId,

            "employeeCode": str(user.iloc[0]["employeeCode"]),

            "employeeName": str(user.iloc[0]["employeeName"]),

            "department": str(user.iloc[0]["department"]),

            "mobile": str(user.iloc[0]["mobile"]),

            "email": str(user.iloc[0]["email"]),

            "shift": str(user.iloc[0]["shift"]),

            "routeNumber": str(user.iloc[0]["routeNumber"]),

            "busStop": str(user.iloc[0]["busStop"]),

            "requestDate": datetime.now().strftime("%d-%m-%Y %H:%M"),

            "status": "Pending",

            "approvedBy": "",

            "approvedDate": ""

        }

        req = pd.concat(

            [req, pd.DataFrame([newRow])],

            ignore_index=True

        )

        req.to_excel(

            PASS_REQUEST_FILE,

            index=False

        )

        return jsonify({

            "status": "success",

            "message": "Permanent Pass Request Submitted Successfully"

        })

    except Exception as e:

        return jsonify({

            "status": "error",

            "message": str(e)

        })


# ==========================
# PENDING PASS REQUESTS API
# ==========================

@app.route("/pending-pass-requests")
@login_required("admin")
def pending_pass_requests():

    try:

        if not os.path.exists(PASS_REQUEST_FILE):

            return jsonify([])

        df = pd.read_excel(PASS_REQUEST_FILE)

        df = df[df["status"] == "Pending"]

        requests = []

        for _, row in df.iterrows():

            requests.append({

                "requestId": str(row["requestId"]),

                "employeeCode": str(row["employeeCode"]),

                "employeeName": str(row["employeeName"]),

                "route": str(row["routeNumber"]),

                "requestDate": str(row["requestDate"])

            })

        return jsonify(requests)

    except Exception as e:

        return jsonify({
            "status":"error",
            "message":str(e)
        })


# ==========================
# REJECT PASS REQUEST
# ==========================

@app.route("/reject-pass", methods=["POST"])
@login_required("admin")
def reject_pass():

    try:

        data = request.get_json()

        requestId = str(data.get("requestId"))

        req = pd.read_excel(PASS_REQUEST_FILE)

        index = req[
            req["requestId"].astype(str) == requestId
        ].index

        if len(index) == 0:

            return jsonify({

                "status":"error",

                "message":"Request Not Found"

            })

        req.loc[index, "status"] = "Rejected"

        req.loc[index, "approvedBy"] = "Admin"

        req.loc[index, "approvedDate"] = datetime.now().strftime("%d-%m-%Y")

        req.to_excel(
            PASS_REQUEST_FILE,
            index=False
        )

        return jsonify({

            "status":"success",

            "message":"Request Rejected"

        })

    except Exception as e:

        return jsonify({

            "status":"error",

            "message":str(e)

        })
    


# ==========================
# DASHBOARD SUMMARY
# ==========================

@app.route("/dashboard-summary")
@login_required("admin")
def dashboard_summary():

    try:

        emp = load_df()

        totalEmployees = len(emp)

        totalRoutes = emp["route"].fillna("").astype(str).str.strip().replace("", pd.NA).dropna().nunique()

        pendingRequests = 0
        activePasses = 0

        if os.path.exists(PASS_REQUEST_FILE):

            req = pd.read_excel(PASS_REQUEST_FILE)

            pendingRequests = len(
                req[req["status"] == "Pending"]
            )

        if os.path.exists(PASS_FILE):

            passes = pd.read_excel(PASS_FILE)

            activePasses = len(passes)

        return jsonify({

            "totalEmployees": totalEmployees,

            "totalRoutes": totalRoutes,

            "pendingRequests": pendingRequests,

            "activePasses": activePasses

        })

    except Exception as e:

        return jsonify({

            "status":"error",

            "message":str(e)

        })


# ==========================
# EMPLOYEE SUMMARY (for employee-list.html)
# ==========================

@app.route("/employee-summary")
@login_required("admin")
def employee_summary():

    try:

        emp = load_df()

        totalEmployees = len(emp)

        totalDepartments = emp["department"].fillna("").astype(str).str.strip().replace("", pd.NA).dropna().nunique()

        totalRoutes = emp["route"].fillna("").astype(str).str.strip().replace("", pd.NA).dropna().nunique()

        return jsonify({

            "totalEmployees": totalEmployees,

            "totalDepartments": totalDepartments,

            "totalRoutes": totalRoutes

        })

    except Exception as e:

        return jsonify({

            "status":"error",

            "message":str(e)

        })



# ==========================
# GET ALL PASS REQUESTS
# ==========================

@app.route("/admin/pass-requests", methods=["GET"])
@login_required("admin")
def admin_pass_requests():

    try:

        if not os.path.exists(PASS_REQUEST_FILE):

            return jsonify([])

        df = pd.read_excel(PASS_REQUEST_FILE)

        df = df.fillna("")

        return jsonify(
            df.to_dict(orient="records")
        )

    except Exception as e:

        return jsonify({
            "status":"error",
            "message":str(e)
        })
    

    # ==========================
# GENERATED PASSES
# ==========================

@app.route("/generated-passes")
@login_required("admin")
def generated_passes():

    try:

        if not os.path.exists(PASS_FILE):

            return jsonify([])

        passes = pd.read_excel(PASS_FILE)

        return jsonify(

            passes.fillna("").to_dict(
                orient="records"
            )

        )

    except Exception as e:

        return jsonify({

            "status":"error",

            "message":str(e)

        })
    


# ==========================
# GET ALL EMPLOYEES API
# Admin Dashboard → Employee List
# ==========================

@app.route("/employees-json")
@login_required("admin")
def employees_json():

    try:

        # employees.xlsx मधील सर्व employees load करा
        df = load_df()

        # रिकामे values "" करा
        df = df.fillna("")

        # password column कधीही API response मध्ये पाठवायचा नाही
        if "password" in df.columns:
            df = df.drop(columns=["password"])

        # JSON मध्ये return करा
        return jsonify(
            df.to_dict(orient="records")
        )

    except Exception as e:

        return jsonify({
            "status": "error",
            "message": str(e)
        })
    


# ==========================
# TEMPORARY PASS SYSTEM
# ==========================

TEMP_PASS_FILE = "temp_pass_requests.xlsx"


def create_temp_pass_file():
    if not os.path.exists(TEMP_PASS_FILE):
        df = pd.DataFrame(columns=[
            "requestId", "employeeCode", "employeeName", "department",
            "mobile", "reason", "pickupLocation", "dropLocation",
            "travelDateTime", "requestDate", "status",
            "tempPassId", "validUntil"
        ])
        df.to_excel(TEMP_PASS_FILE, index=False)


create_temp_pass_file()


# ---- Employee: submit a temporary pass request ----
@app.route("/request-temp-pass", methods=["POST"])
@login_required("employee")
def request_temp_pass():

    try:
        data = request.get_json()

        employeeCode = str(data.get("employeeCode", "")).strip()

        if str(session.get("employeeCode", "")).strip() != str(employeeCode).strip():
            return jsonify({"status": "error", "message": "Not Authorized"}), 403

        if employeeCode == "":
            return jsonify({"status": "error", "message": "Employee Code Required"})

        df = pd.read_excel(TEMP_PASS_FILE)

        request_id = "TMP" + str(int(datetime.now().timestamp()))

        new_row = {
            "requestId": request_id,
            "employeeCode": employeeCode,
            "employeeName": data.get("employeeName", ""),
            "department": data.get("department", ""),
            "mobile": data.get("mobile", ""),
            "reason": data.get("reason", ""),
            "pickupLocation": data.get("pickupLocation", ""),
            "dropLocation": data.get("dropLocation", "Bajaj Chakan Plant 1"),
            "travelDateTime": data.get("travelDateTime", ""),
            "requestDate": datetime.now().strftime("%Y-%m-%d %H:%M"),
            "status": "Pending",
            "tempPassId": "",
            "validUntil": ""
        }

        df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
        df.to_excel(TEMP_PASS_FILE, index=False)

        return jsonify({
            "status": "success",
            "message": "Temporary Pass Request Submitted Successfully",
            "requestId": request_id
        })

    except Exception as e:
        return jsonify({"status": "error", "message": str(e)})


# ---- Admin: view all temporary pass requests ----
@app.route("/admin/temp-pass-requests", methods=["GET"])
@login_required("admin")
def admin_temp_pass_requests():

    try:
        df = pd.read_excel(TEMP_PASS_FILE)
        df = df.fillna("")
        return jsonify(df.to_dict(orient="records"))

    except Exception as e:
        return jsonify({"status": "error", "message": str(e)})


# ---- Admin: approve a temporary pass request ----
@app.route("/approve-temp-pass", methods=["POST"])
@login_required("admin")
def approve_temp_pass():

    try:
        data = request.get_json()
        request_id = str(data.get("requestId", ""))

        df = pd.read_excel(TEMP_PASS_FILE)
        match = df["requestId"].astype(str) == request_id

        if not match.any():
            return jsonify({"status": "error", "message": "Request Not Found"})

        temp_pass_id = "TP-" + str(int(datetime.now().timestamp()))
        valid_until = (datetime.now() + timedelta(hours=12)).strftime("%Y-%m-%d %H:%M")

        df.loc[match, "status"] = "Approved"
        df.loc[match, "tempPassId"] = temp_pass_id
        df.loc[match, "validUntil"] = valid_until

        df.to_excel(TEMP_PASS_FILE, index=False)

        return jsonify({
            "status": "success",
            "message": "Temporary Pass Approved"
        })

    except Exception as e:
        return jsonify({"status": "error", "message": str(e)})


# ---- Admin: reject a temporary pass request ----
@app.route("/reject-temp-pass", methods=["POST"])
@login_required("admin")
def reject_temp_pass():

    try:
        data = request.get_json()
        request_id = str(data.get("requestId", ""))

        df = pd.read_excel(TEMP_PASS_FILE)
        match = df["requestId"].astype(str) == request_id

        if not match.any():
            return jsonify({"status": "error", "message": "Request Not Found"})

        df.loc[match, "status"] = "Rejected"
        df.to_excel(TEMP_PASS_FILE, index=False)

        return jsonify({
            "status": "success",
            "message": "Temporary Pass Rejected"
        })

    except Exception as e:
        return jsonify({"status": "error", "message": str(e)})


# ---- Employee: view / download their approved temporary pass ----
@app.route("/my-temp-pass/<employeeCode>", methods=["GET"])
@login_required("employee")
def my_temp_pass(employeeCode):

    if str(session.get("employeeCode", "")).strip() != str(employeeCode).strip():
        return jsonify({"status": "error", "message": "Not Authorized"}), 403

    try:
        df = pd.read_excel(TEMP_PASS_FILE)

        user_passes = df[
            (df["employeeCode"].astype(str) == str(employeeCode)) &
            (df["status"] == "Approved")
        ]

        if user_passes.empty:
            return jsonify({
                "status": "error",
                "message": "No Approved Temporary Pass Found"
            })

        latest = user_passes.iloc[-1]

        valid_until_str = str(latest.get("validUntil", ""))
        is_expired = False

        try:
            vu = datetime.strptime(valid_until_str, "%Y-%m-%d %H:%M")
            if datetime.now() > vu:
                is_expired = True
        except Exception:
            pass

        return jsonify({
            "status": "success",
            "employeeCode": str(latest["employeeCode"]),
            "employeeName": str(latest["employeeName"]),
            "department": str(latest["department"]),
            "mobile": str(latest["mobile"]),
            "reason": str(latest["reason"]),
            "pickupLocation": str(latest["pickupLocation"]),
            "dropLocation": str(latest["dropLocation"]),
            "travelDateTime": str(latest["travelDateTime"]),
            "tempPassId": str(latest["tempPassId"]),
            "validUntil": valid_until_str,
            "isExpired": is_expired
        })

    except Exception as e:
        return jsonify({"status": "error", "message": str(e)})


# ==========================
# OT MANPOWER / TRANSPORT REQUEST
# --------------------------------------------------------------------
# Every "Save & Send Request" submit from the OT Manpower Approval page
# appends rows here: one row per Manpower entry + one row per Transport
# route, all tagged with the same Date/Time/Submitted By, so the sheet
# can be filtered/pivoted by date in Excel. File is created
# automatically on first submit — no manual setup needed.
# ==========================

OT_FILE = "OT_Requests.xlsx"

OT_COLUMNS = [
    "Date",
    "Time",
    "Submitted By",
    "Department",
    "Shift",
    "Emergency",
    "Entry Type",
    "Category / Route",
    "Bus Stops",
    "2 Hours",
    "3 Hours"
]


def create_ot_file_if_needed():

    if not os.path.exists(OT_FILE):
        pd.DataFrame(columns=OT_COLUMNS).to_excel(OT_FILE, index=False)
        return

    try:
        df = pd.read_excel(OT_FILE)
        df.columns = df.columns.str.strip()

        for col in OT_COLUMNS:
            if col not in df.columns:
                df[col] = ""

        df = df[OT_COLUMNS]
        df.to_excel(OT_FILE, index=False)

    except Exception:
        pd.DataFrame(columns=OT_COLUMNS).to_excel(OT_FILE, index=False)


create_ot_file_if_needed()


@app.route("/save-ot-request", methods=["POST"])
def save_ot_request():

    try:

        data = request.get_json(silent=True)

        submitted_by = str((data or {}).get("submittedBy", "")).strip()
        department = str((data or {}).get("department", "")).strip()

        if not data or not submitted_by or not department:
            return jsonify({
                "status": "error",
                "message": "Submitted By / Department Required"
            }), 400

        create_ot_file_if_needed()

        df = pd.read_excel(OT_FILE)
        df.columns = df.columns.str.strip()

        now = datetime.now()
        date_str = now.strftime("%d-%m-%Y")
        time_str = now.strftime("%I:%M:%S %p")

        shift = str(data.get("shift", "")).strip()
        emergency = "Yes" if data.get("emergency") else "No"

        manpower = data.get("manpower") or []
        transport = data.get("transport") or []

        new_rows = []

        for row in manpower:
            new_rows.append({
                "Date": date_str,
                "Time": time_str,
                "Submitted By": submitted_by,
                "Department": department,
                "Shift": shift,
                "Emergency": emergency,
                "Entry Type": "Manpower",
                "Category / Route": row.get("provider", ""),
                "Bus Stops": "",
                "2 Hours": row.get("twoHours", 0),
                "3 Hours": row.get("threeHours", 0)
            })

        for row in transport:
            new_rows.append({
                "Date": date_str,
                "Time": time_str,
                "Submitted By": submitted_by,
                "Department": department,
                "Shift": shift,
                "Emergency": emergency,
                "Entry Type": "Transport",
                "Category / Route": row.get("route", ""),
                "Bus Stops": row.get("stops", ""),
                "2 Hours": row.get("twoHours", 0),
                "3 Hours": row.get("threeHours", 0)
            })

        # Don't silently lose a submission that had no rows at all
        if not new_rows:
            new_rows.append({
                "Date": date_str,
                "Time": time_str,
                "Submitted By": submitted_by,
                "Department": department,
                "Shift": shift,
                "Emergency": emergency,
                "Entry Type": "Info",
                "Category / Route": "No manpower/transport rows submitted",
                "Bus Stops": "",
                "2 Hours": 0,
                "3 Hours": 0
            })

        df = pd.concat([df, pd.DataFrame(new_rows)], ignore_index=True)
        df.to_excel(OT_FILE, index=False)

        return jsonify({
            "status": "success",
            "message": "Saved Successfully to Excel Sheet"
        })

    except Exception as e:

        print("SAVE OT REQUEST ERROR:", e)

        return jsonify({
            "status": "error",
            "message": str(e)
        })


# ==========================
# RUN SERVER
# ==========================

if __name__ == "__main__":
    # debug mode should NEVER be on when the site is public (it leaks
    # internal code/data on errors). Set FLASK_DEBUG=True locally if needed.
    app.run(debug=os.environ.get("FLASK_DEBUG", "False") == "True")