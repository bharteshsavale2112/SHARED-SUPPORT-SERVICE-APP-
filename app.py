from flask import Flask, request, jsonify, send_from_directory, session, send_file
from functools import wraps
from werkzeug.security import generate_password_hash, check_password_hash
import pandas as pd
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
from datetime import timedelta as _timedelta
app.config["PERMANENT_SESSION_LIFETIME"] = _timedelta(hours=2)

# Admin credentials now come from environment variables instead of being
# hardcoded in the source code (which is visible to anyone with repo access).
# Set ADMIN_EMAIL and ADMIN_PASSWORD as environment variables on Render.
ADMIN_EMAIL = os.environ.get("ADMIN_EMAIL", "admin@bajaj.com")
ADMIN_PASSWORD_HASH = os.environ.get("ADMIN_PASSWORD_HASH") or generate_password_hash(
    os.environ.get("ADMIN_PASSWORD", "admin123")
)


# ---- Brute-force login protection ----
# Blocks an IP address after too many rapid login attempts, so someone
# can't sit there guessing passwords over and over.
from flask_limiter import Limiter
from flask_limiter.util import get_remote_address



# ---- Extra security response headers ----
@app.after_request
def add_security_headers(response):
    response.headers["X-Content-Type-Options"] = "nosniff"
    response.headers["X-Frame-Options"] = "DENY"
    response.headers["Referrer-Policy"] = "strict-origin-when-cross-origin"
    return response


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
# DATABASE (Azure Database for PostgreSQL / Neon) + LIVE EXCEL MIRROR
# ==========================
# Set DATABASE_URL as an environment variable (Azure App Service ->
# Configuration -> Environment variables, or Render -> Environment) to
# make Postgres the permanent source of truth. Every time data is saved,
# it is written to Postgres AND to a local .xlsx file (so you always also
# have an Excel copy). If DATABASE_URL is not set, the app falls back to
# using ONLY the local Excel files (old behaviour, useful for local
# testing).

from sqlalchemy import create_engine, inspect as _sa_inspect

DATABASE_URL = os.environ.get("DATABASE_URL", "")

db_engine = None

if DATABASE_URL:
    try:
        # Some providers give a URL starting with "postgres://";
        # SQLAlchemy needs "postgresql://"
        _url = DATABASE_URL.replace("postgres://", "postgresql://", 1)

        # Azure Database for PostgreSQL requires an SSL connection.
        # If the connection string doesn't already specify sslmode,
        # add it automatically so it works out of the box on Azure.
        if "sslmode" not in _url:
            _sep = "&" if "?" in _url else "?"
            _url = _url + _sep + "sslmode=require"

        db_engine = create_engine(_url, pool_pre_ping=True, pool_recycle=300)
    except Exception as _e:
        print("DATABASE CONNECTION ERROR:", _e)
        db_engine = None


def db_table_exists(table_name):
    if db_engine is None:
        return False
    try:
        return _sa_inspect(db_engine).has_table(table_name)
    except Exception as e:
        print("DB CHECK ERROR:", e)
        return False


def init_table_if_needed(table_name, xlsx_path, cols, default_df=None):
    """Creates the table in Postgres (with the right columns, optionally
    pre-populated with default_df rows) only if it doesn't already exist
    there. Never overwrites existing data.
    Also makes sure a local Excel mirror file exists."""

    starter_df = default_df if default_df is not None else pd.DataFrame(columns=cols)

    if db_engine is not None:
        if not db_table_exists(table_name):
            try:
                starter_df.to_sql(
                    table_name, db_engine, index=False
                )
            except Exception as e:
                print("DB INIT ERROR (" + table_name + "):", e)
        # Refresh the local Excel mirror from the database (source of truth)
        try:
            df = pd.read_sql_table(table_name, db_engine)
            df.to_excel(xlsx_path, index=False)
        except Exception as e:
            print("EXCEL MIRROR ERROR (" + table_name + "):", e)
        return

    # No database configured -> old local-Excel-only behaviour
    if not os.path.exists(xlsx_path):
        starter_df.to_excel(xlsx_path, index=False)


def read_table(table_name, xlsx_path):
    """Reads a table from Postgres if a database is configured,
    otherwise reads the local Excel file.

    IMPORTANT: dtype=str is used for the Excel read so that
    numeric-looking text (employee codes, mobile numbers, route
    numbers) is never silently turned into a float (e.g. "9876543210"
    becoming "9876543210.0")."""

    if db_engine is not None:
        try:
            return pd.read_sql_table(table_name, db_engine)
        except Exception as e:
            print("DB READ ERROR (" + table_name + "):", e)

    if os.path.exists(xlsx_path):
        df = pd.read_excel(xlsx_path, dtype=str)
        return df.fillna("")

    return pd.DataFrame()


def save_table(df, table_name, xlsx_path):
    """Saves a table to Postgres (source of truth, if configured) AND
    always also writes a local Excel copy (as requested), so an
    up-to-date .xlsx file is available any time the app is running."""

    if db_engine is not None:
        try:
            df.to_sql(table_name, db_engine, if_exists="replace", index=False)
        except Exception as e:
            print("DB WRITE ERROR (" + table_name + "):", e)

    try:
        df.to_excel(xlsx_path, index=False)
    except Exception as e:
        print("EXCEL WRITE ERROR (" + table_name + "):", e)


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

    init_table_if_needed("employees", EMPLOYEE_FILE, cols)

create_excel_if_needed()



# ==========================
# LOAD EXCEL SAFELY
# ==========================

def load_df():
    create_excel_if_needed()   # <-- ही नवीन line add कर

    df = read_table("employees", EMPLOYEE_FILE)
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
@limiter.limit("10 per hour")
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

        save_table(df, "employees", EMPLOYEE_FILE)

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
@limiter.limit("6 per minute")
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
                save_table(df, "employees", EMPLOYEE_FILE)

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
@limiter.limit("6 per minute")
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

            req = read_pass_requests()

            pendingRequests = len(
                req[req["status"] == "Pending"]
            )

        if os.path.exists(PASS_FILE):

            passes = read_passes()

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

        save_table(df, "employees", EMPLOYEE_FILE)

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
# ==========================

live = {}

@app.route("/update-location", methods=["POST"])
def update():

    try:

        data = request.get_json()

        if (
            not data or
            "route" not in data or
            "lat" not in data or
            "lng" not in data
        ):
            return jsonify({
                "success": False,
                "message": "Invalid Data"
            })

        route = str(data["route"]).strip()

        live[route] = {
            "lat": float(data["lat"]),
            "lng": float(data["lng"])
        }

        return jsonify({
            "success": True,
            "message": "Location Updated"
        })

    except Exception as e:

        return jsonify({
            "success": False,
            "message": str(e)
        })

#=========================
# live location route
#=========================

@app.route("/live-location/<route>")
def get(route):

    route = str(route).strip()

    if route not in live:
        return jsonify({
            "success": False,
            "message": "Location Not Found"
        })

    return jsonify({
        "success": True,
        "lat": live[route]["lat"],
        "lng": live[route]["lng"]
    })







# ==========================
# PASS FILE
# ==========================

from datetime import datetime

PASS_FILE = "passes.xlsx"

# ==========================
# CREATE PASSES TABLE (Postgres + Excel mirror)
# ==========================

PASS_COLS = [
    "passNumber", "employeeCode", "employeeName", "department",
    "mobile", "shift", "routeNumber", "busStop", "issueDate", "status"
]


def create_pass_file():
    init_table_if_needed("passes", PASS_FILE, PASS_COLS)


def read_passes():
    df = read_table("passes", PASS_FILE)
    df.columns = df.columns.str.strip()
    return df


def save_passes(df):
    save_table(df, "passes", PASS_FILE)


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

        pass_df = read_passes()

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

        save_passes(pass_df)

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

    df = read_passes()

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

PASS_REQUEST_COLS = [
    "requestId", "employeeCode", "employeeName", "department", "mobile",
    "routeNumber", "busStop", "requestDate", "status",
    "approvedBy", "approvedDate"
]

init_table_if_needed("pass_requests", PASS_REQUEST_FILE, PASS_REQUEST_COLS)


def read_pass_requests():
    df = read_table("pass_requests", PASS_REQUEST_FILE)
    df.columns = df.columns.str.strip()
    return df


def save_pass_requests(df):
    save_table(df, "pass_requests", PASS_REQUEST_FILE)

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
        req = read_pass_requests()
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

        save_pass_requests(req)

        # Load pass file
        passes = read_passes()
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
        save_passes(passes)

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

        df = read_passes()
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
        req = read_pass_requests()
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

        save_pass_requests(req)

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

        df = read_pass_requests()

        df = df[df["status"] == "Pending"]

        requests = []

        for _, row in df.iterrows():

            requests.append({

                "requestId": str(row["requestId"]),

                "employeeCode": str(row["employeeCode"]),

                "employeeName": str(row["employeeName"]),

                "route": str(row["route"]),

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

        req = read_pass_requests()

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

        save_pass_requests(req)

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

            req = read_pass_requests()

            pendingRequests = len(
                req[req["status"] == "Pending"]
            )

        if os.path.exists(PASS_FILE):

            passes = read_passes()

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
# GET ALL PASS REQUESTS
# ==========================

@app.route("/admin/pass-requests", methods=["GET"])
@login_required("admin")
def admin_pass_requests():

    try:

        if not os.path.exists(PASS_REQUEST_FILE):

            return jsonify([])

        df = read_pass_requests()

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

        passes = read_passes()

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

from datetime import timedelta

TEMP_PASS_FILE = "temp_pass_requests.xlsx"

TEMP_PASS_COLS = [
    "requestId", "employeeCode", "employeeName", "department",
    "mobile", "reason", "pickupLocation", "dropLocation",
    "travelDateTime", "requestDate", "status",
    "tempPassId", "validUntil"
]


def create_temp_pass_file():
    init_table_if_needed("temp_pass_requests", TEMP_PASS_FILE, TEMP_PASS_COLS)


def read_temp_passes():
    df = read_table("temp_pass_requests", TEMP_PASS_FILE)
    df.columns = df.columns.str.strip()
    return df


def save_temp_passes(df):
    save_table(df, "temp_pass_requests", TEMP_PASS_FILE)


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

        df = read_temp_passes()

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
        save_temp_passes(df)

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
        df = read_temp_passes()
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

        df = read_temp_passes()
        match = df["requestId"].astype(str) == request_id

        if not match.any():
            return jsonify({"status": "error", "message": "Request Not Found"})

        temp_pass_id = "TP-" + str(int(datetime.now().timestamp()))
        valid_until = (datetime.now() + timedelta(hours=12)).strftime("%Y-%m-%d %H:%M")

        df.loc[match, "status"] = "Approved"
        df.loc[match, "tempPassId"] = temp_pass_id
        df.loc[match, "validUntil"] = valid_until

        save_temp_passes(df)

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

        df = read_temp_passes()
        match = df["requestId"].astype(str) == request_id

        if not match.any():
            return jsonify({"status": "error", "message": "Request Not Found"})

        df.loc[match, "status"] = "Rejected"
        save_temp_passes(df)

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
        df = read_temp_passes()

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
# ADMIN: DOWNLOAD LIVE EXCEL BACKUP
# ==========================
# These always read straight from Postgres (if DATABASE_URL is set) so the
# downloaded file is guaranteed to reflect the current live data, even if
# the local .xlsx mirror on this server instance was reset by a restart.

import io


def _excel_download(df, filename):
    buffer = io.BytesIO()
    df.to_excel(buffer, index=False)
    buffer.seek(0)
    return send_file(
        buffer,
        as_attachment=True,
        download_name=filename,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )


@app.route("/admin/export/employees")
@login_required("admin")
def export_employees():
    df = load_df()
    if "password" in df.columns:
        df = df.drop(columns=["password"])
    return _excel_download(df, "employees.xlsx")


@app.route("/admin/export/passes")
@login_required("admin")
def export_passes():
    return _excel_download(read_passes(), "passes.xlsx")


@app.route("/admin/export/pass-requests")
@login_required("admin")
def export_pass_requests():
    return _excel_download(read_pass_requests(), "pass_requests.xlsx")


@app.route("/admin/export/temp-pass-requests")
@login_required("admin")
def export_temp_pass_requests():
    return _excel_download(read_temp_passes(), "temp_pass_requests.xlsx")


# ==========================
# SHIFT -> ROUTE -> BUS STOPS (real data)
# ==========================
# Single source of truth used by both employeeregistration.html and
# ot-approval.html for the Route Number -> Bus Stop cascading dropdown.
#
# NOTE: Only "1st Shift" and "2nd Shift" data has been provided so far.
# "3rd Shift" and "General Shift" are empty until that data is supplied.

SHIFT_ROUTES = {

    "1st Shift": {
        "Route 1 (LONI KALBHOR)": ["Loni Gaon", "HP Pump, Loni Gaon", "Wakwasti", "Manjre Farm", "Shawalwadi Depot", "Bhekarai Nagar", "Bhaji Mandai, Hadapsar", "Pune Station", "Maldhakka Ambedkar Bhavan", "Dapodi"],
        "Route 2 (NARAYANGAON)": ["Hotel Shravani", "Narayangaon Bus Stop", "Police Station", "Manchar Market Yard", "Mulewadi Road", "Lohakare Hospita", "Rakshewade"],
        "Route 3 (SINHAGAD ROAD)": ["Kolhewadi", "Sun City", "Warje Bridge", "CNG Petrol Pump, Bavdhan", "Balewadi", "Dange Chowk", "Chinchwad Nagar"],
        "Route 4 (WARJE)": ["NDA", "Mayur Colony", "Ashish Garden, Kothrud"],
        "Route 5 (WAGHOLI)": ["Wageshwar Mandir", "Kharadi Bypass", "Wadgaon Sheri, Tingre Nagar", "Bhopkhel", "Parande Nagar", "Datta Nagar"],
        "Route 6 (KATRAJ)": ["Santosh Nagar", "Bharati Vidyapeeth", "Hatti Chowk", "Padmavati", "Bibwewadi", "Shaniwar Wada", "Shivajinagar", "Khondba Chowk"],
        "Route 7 (MENGDEWADI)": ["Mengdewadi", "Hingewasti", "Ausari BK Bus Stand", "Pandurang Mandir", "Temkar Wasti", "Ausari KH Water Tank", "Ranwara Hotel", "Shivdhan Plaza", "Marathi Shala", "Market Yard", "Khed Bus Stand", "Shiroli Phata", "Rajratna Hotel"],
        "Route 8 (PABAL)": ["Pabal", "Petrol Pump Pabal", "Tamane Mala", "Kanersir", "Nimgaon", "Chavan Mala", "Maharaja Chowk"],
        "Route 9 (VISHRANTWADI)": ["Vishrantwadi Chowk", "Dhanori", "Kotak Mahindra", "Charholi Gaon", "Bank of Maharashtra", "Charholi Phata", "Kate Colony", "Kale Colony"],
        "Route 10 (OLD SANGAVI)": ["Sangavi Bus Stop", "Ganpati Mandir", "Kranti Chowk", "Katepuram Chowk", "Radha Krishna Mangal Karyalaya", "Dehukar Park", "Tulja Bhavani Mandir", "Kalpataru Chowk", "Nashik Phata", "Empire Bridge", "Akurdi Bajaj Gate"],
        "Route 11 (ALANDI)": ["Cosmos Bank", "Dehu Phata", "Dudulgaon", "Dudulgaon 1", "Moshi", "Bharat Mata Chowk", "Tupe Wasti", "Moshi Toll Naka", "Chimbli Phata", "Kuruli Phata"],
        "Route 12 (RAVET)": ["Kiwale", "Mukai Chowk", "Adarsh Nagar", "Soni Ghar", "Shinde Wasti", "CNG Pump"],
        "Route 13 (KALE WADI)": ["Thergaon", "16 No. Bus Stop", "Kalewadi Bridge", "Rahatani Phata", "Dhangar Baba", "Tapkir Chowk", "Keshav Nagar Corner", "Keshav Nagar Bus Stop", "Talera Hospital", "Chafekar Chowk", "SKF Company", "Prem Lok Park", "Delvi Nagar"],
        "Route 14 (SHASTRI CHOWK)": ["Shastri Chowk", "Dighi Road", "Panjar Pol", "Godam Chowk", "Gandharva Nagari", "Sainath Hospital", "Borade Wasti"],
        "Route 15 (DEHU ROAD)": ["Mamurdi (Sai Nagar)", "Vikas Nagar", "SBI Bank", "Dehu Road", "Chincholi", "Abhilekha Park", "Sangurdi Phata", "Yelwadi"],
        "Route 16 (SHEETAL BAGH)": ["Landewadi", "Spine Road", "Jadhavwadi", "Chikhli"],
        "Route 17 (WADGAON)": ["Wadgaon", "Matoshree", "Jambhul Phata", "Wadgaon Court", "Chhatrapati Shivaji Chowk", "Lotus", "Wadgaon Phata", "Murlesh Hotel", "Aarth Hospital", "Seva Dham", "Machhi Market", "Indrayani College"],
        "Route 18 (BHOSARI EXTRA)": ["PMT Chowk, Bhosari", "Dhawade Wasti", "Satguru Nagar", "Jadhavwadi", "Laxmi Chowk", "D Mart, Moshi", "Silver 9", "Rustic"],
        "Route 19 (SHARAD NAGAR)": ["Sane Chowk", "Sharad Nagar", "Gharkul Chowk", "Kaseriyo Society", "Nevale Wasti", "Ganesh Mandir", "Navale Wasti Corner", "Imperial Hospital", "Talawade Chowk"],
        "Route 20 (RAJGURU NAGAR - KHED)": ["Sangam Garden", "Sangam Classic", "Panyachi Taki", "Panchayat Samiti", "Chandoli Phata", "Vishwakalyan"],
        "Route 21 (YCM)": ["Mahesh Nagar Chowk", "Nehru Nagar", "Vitthal Mandir", "Ajmera 1", "Ajmera 2", "Ajmera 3", "Bajaj Colony - Amruteshwar Colony", "Old RTO", "Petrol Pump", "Kudalwadi 1", "Patil Nagar, Chikhli", "Bajaj Auto"],
        "Route 22 (LINK ROAD)": ["Gawade Petrol Pump", "Chafekar Chowk", "Bijli Nagar", "Ankush Chowk", "Triveni Nagar", "Tawre Line", "Ganesh Nagar"],
        "Route 23 (TALEGAON)": ["McDonald's", "Bagicha Hotel", "Nim Phata", "Bhandari Hospital", "Kesar Hotel", "Jijamata Chowk", "Khadge Pump", "Nagar Palika", "Kaka Halwai", "BSNL", "Siddhi Khed", "Machhi Market", "Induri Gaon", "Khalumbre"],
        "Route 24 (SHAHU NAGAR)": ["Atal Bihari Udyan Corner", "Rasrang", "Ganesh Mandir", "Bajaj School", "Someshwar Mandir", "Chintamani Corner", "Kasturi Market", "Polite Harmony", "Ganesh International School"],
        "Route 25 (AUTO GAS TALEGAON)": ["Varale Phata", "Shivaji Chowk", "Sindket Bank", "Bullet Showroom", "Balaji Marble", "Aishwarya Hotel", "Indori Gaon"],
        "Route 26 (WALHEKARWADI)": ["Jakat Naka", "Walhekarwadi Chowk", "Aaher Garden", "Ganpati Mandir", "Gurudwara Chowk", "Railway Station", "Appu Ghar"],
        "Route 27 (MOHAN NAGAR)": ["Pimple Saudagar", "Pimpri Gaon", "Kalewadi", "Vijay Nagar Bus Stop", "Mohan Nagar", "Huma Bakery"],
        "Route 28 (BAJAJ SCHOOL)": ["Vinayak Sweets", "Shivalkar Chowk", "Hanuman Mandir", "Mehtre Garden", "More Wasti", "Ashtavinayak Chowk", "Vande Mataram Chowk", "Waghu Sane Chowk", "Aishwaryam Society"],
        "Route 29 (DEHUGAON)": ["Krida Sankul", "Omi Home", "Gatha Mandir", "Arogya Hospital", "Parandwal Chowk", "Parishri Hotel", "Omkar Society", "V-Mart", "Banner Bank", "Vitthalwadi"],
        "Route 30 (WAKI PHATA)": ["Sumbare Nagar", "Ekta Nagar", "Swapna Nagari", "Chakan Market Yard", "Yeshwant Nagar", "Dnyanvardhani School", "Jhipre Mala", "Biradwadi"],
        "Route 31 (MEDANKARWADI)": ["Kadachiwadi", "Kalpataru Society", "Medankarwadi Phata", "Vishal Garden"],
        "Route 32 (BALAJI NAGAR CHAKAN)": ["Chakan Chowk", "Mutkewadi", "Balaji Nagar", "Premacha Chaha", "IAI Company"],
        "Route 33 (CHAKAN CHOWK)": ["Ambethan Chowk", "Ghadge Mala", "Unicare Hospital", "Chakan Chowk", "Arogyam Hospital", "Sahara City", "Kharabwadi", "Mahalunge"],
        "Route 34 (DANGE CHOWK)": ["Dange Chowk", "Bijli Nagar", "Big India", "Sambhaji Chowk", "Mhalsakant Chowk", "TJSB Bank", "Datta Wadi", "Nigdi", "LIC Corner", "Hatti Chowk"],
        "Route 35 (NIGDI)": ["Pawale Bridge", "Swanand Dairy", "LIC Corner", "Hatti Chowk", "Ganesh Nagar", "Jyotiba Nagar", "Talawade (Only Ladies)"],
        "Route 36 (MEDANKARWADI EXTRA)": ["Kadachiwadi", "Kalpataru Society", "Medankarwadi Phata", "Vishal Garden"],
        "Route 37 (IAI COMPANY)": ["IAI Company"],
        "Route 38 (KHANDO MANDIR CHAKAN)": ["Khandoba Mandir", "Vishal Garden", "Manik Chowk", "Premacha Chaha"],
        "Route 40 (CHAKAN CHOWK)": ["Ambethan Chowk", "Ghadge Mala", "Unicare Hospital", "Chakan Chowk", "Arogyam Hospital", "Sahara City", "Kharabwadi", "Mahalunge"]
    },

    "2nd Shift": {
        "Route 1 (SINHGAD)": ["Kolhewadi", "Sinhagad Road", "Dahyri Gaon", "Navale Bridge", "Warje", "Bawdhan", "Balewadi", "Dange Chowk", "Aditya Birla", "Chinchwade Nagar", "Kachghar Chowk", "Ankush Chowk", "Chakan Chowk", "Bajaj Chakan"],
        "Route 2 (WAGHOLI)": ["Wagholi", "Vishrantwadi", "Dighi", "Sai Mandir", "Kale Colony", "Kate Colony", "Chikhali", "Bajaj Chakan"],
        "Route 3 (DEHU GAON)": ["Krida Sankul", "Gatha Mandir", "V Mart", "Baner Bank", "Bypass Chowk", "Devi Indrani", "Bajaj Chakan"],
        "Route 4 (MEDANKARWADI)": ["Kadachiwadi", "Medankarwadi Gate", "Vishal Garden", "Manik Chowk", "Indian Oil Pump", "Balaji Nagar (Premacha Chaha)", "IAI Company", "Bajaj Chakan"],
        "Route 5 (WAKI PHATA)": ["Biradwadi", "Waki Phata", "Ekta Nagar", "Swapna Nagari", "Market Yard", "Ambethan Chowk", "Zitrai Mala", "Mahalunge", "Bajaj Chakan"],
        "Route 6 (RAVET-WALHEKARWADI)": ["Mukai Chowk", "CNG Pump", "Shinde Wasti", "Walhekarwadi", "Chintamani Chowk", "Ganesh Mandir", "Athurwa Park", "Gurudwara Chowk", "Tower Line", "Huma Bakery", "Bajaj Chakan"],
        "Route 7 (ALANDI)": ["Charoli Phata", "Dehu Phata", "Dudulgaon", "Moshi", "Bharat Mata Chowk", "Chikhali Rustic Paradise", "Bajaj Chakan"],
        "Route 8 (SHASTRI CHOWK)": ["Shastri Chowk", "PMT Chowk", "Dhawde Wasti", "Godown Chowk", "Panjarpol", "Gandharv Nagari", "Borhade Wasti", "Tupe Wasti", "Toll Naka", "Chimbli Phata", "Kurali Phata", "Bajaj Chakan"],
        "Route 9 (DEHUROAD)": ["Sai Nagar Mamurdi", "Shinde Petrol Pump Dehuroad", "Vikas Nagar Dehuroad", "Bank of India Dehuroad", "Mali Nagar (Malwadi)", "Abhilasha Society", "Mangalkaryalaya", "Parandwal Chowk", "Yelwadi", "Bajaj Chakan"],
        "Route 10 (PIMPLE GURAV)": ["Dapodi", "Kate Puram Chowk", "Bhau Nagar", "Pimpri", "Akurdi Main Gate", "LIC Corner", "Triveni Nagar", "Bajaj Chakan"],
        "Route 11 (JADHAVWADI)": ["Jadhavwadi", "RTO", "Sharad Nagar", "Newale Wasti", "Ganpati Mandir", "Newale Wasti Corner", "Patil Nagar", "Rustic Plaza", "Talwade Chowk", "Mahindra Gate", "Bajaj Chakan"],
        "Route 12 (TALEGAON)": ["Bagicha Hotel", "Tukaram Nagar", "Bhandari Hospital", "Bhosle Chaha", "Jijamata Chowk", "Khandge Petrol Pump", "Kaka Halwai", "BSNL", "Balaji Marble", "Sadumbre", "Khalumbre", "Bajaj Chakan"],
        "Route 13 (WADGAON)": ["Chhatrapati Shivaji Chowk", "Lotus", "Rupesh Hotel", "Athurwa Hospital", "Sevadham", "Indrayani College Talegaon", "Bullet Showroom", "Induri", "Khalumbre", "Bajaj Chakan"],
        "Route 14 (KHED)": ["Kardewasti", "Sangam Garden", "Sangam Classic", "Water Tank", "Pabal Road", "Dhadge Mala", "Unicare Hospital", "Arogyam Hospital", "IFL City", "Kharabwadi", "Sara City", "Mahalunge"],
        "Route 15 (KALEWADI)": ["16 No Bus Stop", "Kalewadi Phata", "Rahatani Phata", "Laxman Nagar Thergaon", "Old Jakat Naka", "Gawde Petrol Pump Link Road", "Chafekar Chowk", "Chaitanya Hall", "Giriraj Housing Society", "Hanuman Sweet", "Mhalsakant Chowk", "TJSB Bank", "Axis Bank", "Dattawadi", "Bajaj Chakan"],
        "Route 16 (SHAHU NAGAR)": ["Ajmera", "Amruteshwar Colony", "Shahu Garden", "Aryan Residency", "Ganpati Mandir", "Bajaj School", "Rameshwar Mandir", "Sambhaji Nagar Chowk", "Rajdeep Society", "Kasturi Market", "Shivarkar Chowk", "Polite Harmony", "Ganesh International School", "Shine City", "Bajaj Chakan"],
        "Route 17 (CHAKAN CHOWK)": ["Balaji Nagar", "Chakan Chowk", "Ghanvat Plaza", "Arogyam Hospital", "Nanekarwadi", "Kharabwadi", "Sara City", "Mahalunge", "Hotel Karwa"],
        "Route 18 (MOHAN NAGAR)": ["Mohan Nagar", "Morewasti", "Chikhali Rustic Paradise", "Talwade Chowk", "Mahindra Gate"],
        "Route 19 (ANKUSH CHOWK)": ["Ankush Chowk", "Triveni Nagar", "Tower Line", "Huma Bakery", "Ganesh Nagar", "Jotiba Nagar", "Talwade Chowk"],
        "Route 20 (NARAYANGAON)": ["Raut Hospital Narayangaon", "Narayangaon ST Stand", "Manchar Market Yard", "Manchar College Road", "Maxcare Hospital Manchar", "Paragaon Phata", "Market Yard", "Ambethan Chowk"],
        "Route 21 (NIGDI)": ["Big India", "Dattawadi", "LIC Corner", "Hatti Chowk"]
    },

    "3rd Shift": {
        "Route 1 (NIGDI)": ["Nigdi Bridge", "LIC Chowk"],
        "Route 2 (TALEGAON)": ["Talegaon", "Khalumbre", "Manohar Nagar Talegaon"],
        "Route 3 (DEHU GAON)": ["Dehugaon", "Parishree Hospital"],
        "Route 4 (SHASTRI CHOWK)": ["Shastri Chowk", "Bhosari", "Alandi"],
        "Route 5 (RAJGURU NAGAR)": ["Rajgurunagar", "Khed Toll Naka", "Ambethan Chowk", "Biradwadi"],
        "Route 6 (MEDANKARWADI)": ["Medankarwadi"],
        "Route 7 (BALAJI NAGAR)": ["Balaji Nagar", "Chakan", "Eiffel City", "Karwa Hotel"],
        "Route 8 (CHIKHALI)": ["Chikhali", "Sane Chowk"],
        "Route 9 (DEHUROAD)": ["Dehuroad"],
        "Route 10 (PCMC)": ["Mhalsakant Chowk"]
    },

    "General Shift": {
        "Route 1 (SINHGAD)": ["Kolhewadi", "Sinhagad Road", "Dahyri Gaon", "Navale Bridge", "Warje", "Bawdhan", "Balewadi", "Dange Chowk", "Aditya Birla", "Chinchwade Nagar", "Kachghar Chowk", "Ankush Chowk", "Chakan Chowk", "Bajaj Chakan"],
        "Route 2 (WAGHOLI)": ["Wagholi", "Vishrantwadi", "Dighi", "Sai Mandir", "Kale Colony", "Kate Colony", "Chikhali", "Bajaj Chakan"],
        "Route 3 (DEHU GAON)": ["Krida Sankul", "Gatha Mandir", "V Mart", "Baner Bank", "Bypass Chowk", "Devi Indrani", "Bajaj Chakan"],
        "Route 4 (MEDANKARWADI)": ["Kadachiwadi", "Medankarwadi Gate", "Vishal Garden", "Manik Chowk", "Indian Oil Pump", "Balaji Nagar (Premacha Chaha)", "IAI Company", "Bajaj Chakan"],
        "Route 5 (WAKI PHATA)": ["Biradwadi", "Waki Phata", "Ekta Nagar", "Swapna Nagari", "Market Yard", "Ambethan Chowk", "Zitrai Mala", "Mahalunge", "Bajaj Chakan"],
        "Route 6 (RAVET-WALHEKARWADI)": ["Mukai Chowk", "CNG Pump", "Shinde Wasti", "Walhekarwadi", "Chintamani Chowk", "Ganesh Mandir", "Athurwa Park", "Gurudwara Chowk", "Tower Line", "Huma Bakery", "Bajaj Chakan"],
        "Route 7 (ALANDI)": ["Charoli Phata", "Dehu Phata", "Dudulgaon", "Moshi", "Bharat Mata Chowk", "Chikhali Rustic Paradise", "Bajaj Chakan"],
        "Route 8 (SHASTRI CHOWK)": ["Shastri Chowk", "PMT Chowk", "Dhawde Wasti", "Godown Chowk", "Panjarpol", "Gandharv Nagari", "Borhade Wasti", "Tupe Wasti", "Toll Naka", "Chimbli Phata", "Kurali Phata", "Bajaj Chakan"],
        "Route 9 (DEHUROAD)": ["Sai Nagar Mamurdi", "Shinde Petrol Pump Dehuroad", "Vikas Nagar Dehuroad", "Bank of India Dehuroad", "Mali Nagar (Malwadi)", "Abhilasha Society", "Mangalkaryalaya", "Parandwal Chowk", "Yelwadi", "Bajaj Chakan"],
        "Route 10 (PIMPLE GURAV)": ["Dapodi", "Kate Puram Chowk", "Bhau Nagar", "Pimpri", "Akurdi Main Gate", "LIC Corner", "Triveni Nagar", "Bajaj Chakan"],
        "Route 11 (JADHAVWADI)": ["Jadhavwadi", "RTO", "Sharad Nagar", "Newale Wasti", "Ganpati Mandir", "Newale Wasti Corner", "Patil Nagar", "Rustic Plaza", "Talwade Chowk", "Mahindra Gate", "Bajaj Chakan"],
        "Route 12 (TALEGAON)": ["Bagicha Hotel", "Tukaram Nagar", "Bhandari Hospital", "Bhosle Chaha", "Jijamata Chowk", "Khandge Petrol Pump", "Kaka Halwai", "BSNL", "Balaji Marble", "Sadumbre", "Khalumbre", "Bajaj Chakan"],
        "Route 13 (WADGAON)": ["Chhatrapati Shivaji Chowk", "Lotus", "Rupesh Hotel", "Athurwa Hospital", "Sevadham", "Indrayani College Talegaon", "Bullet Showroom", "Induri", "Khalumbre", "Bajaj Chakan"],
        "Route 14 (KHED)": ["Kardewasti", "Sangam Garden", "Sangam Classic", "Water Tank", "Pabal Road", "Dhadge Mala", "Unicare Hospital", "Arogyam Hospital", "IFL City", "Kharabwadi", "Sara City", "Mahalunge"],
        "Route 15 (KALEWADI)": ["16 No Bus Stop", "Kalewadi Phata", "Rahatani Phata", "Laxman Nagar Thergaon", "Old Jakat Naka", "Gawde Petrol Pump Link Road", "Chafekar Chowk", "Chaitanya Hall", "Giriraj Housing Society", "Hanuman Sweet", "Mhalsakant Chowk", "TJSB Bank", "Axis Bank", "Dattawadi", "Bajaj Chakan"],
        "Route 16 (SHAHU NAGAR)": ["Ajmera", "Amruteshwar Colony", "Shahu Garden", "Aryan Residency", "Ganpati Mandir", "Bajaj School", "Rameshwar Mandir", "Sambhaji Nagar Chowk", "Rajdeep Society", "Kasturi Market", "Shivarkar Chowk", "Polite Harmony", "Ganesh International School", "Shine City", "Bajaj Chakan"],
        "Route 17 (CHAKAN CHOWK)": ["Balaji Nagar", "Chakan Chowk", "Ghanvat Plaza", "Arogyam Hospital", "Nanekarwadi", "Kharabwadi", "Sara City", "Mahalunge", "Hotel Karwa"],
        "Route 18 (MOHAN NAGAR)": ["Mohan Nagar", "Morewasti", "Chikhali Rustic Paradise", "Talwade Chowk", "Mahindra Gate"],
        "Route 19 (ANKUSH CHOWK)": ["Ankush Chowk", "Triveni Nagar", "Tower Line", "Huma Bakery", "Ganesh Nagar", "Jotiba Nagar", "Talwade Chowk"],
        "Route 20 (NARAYANGAON)": ["Raut Hospital Narayangaon", "Narayangaon ST Stand", "Manchar Market Yard", "Manchar College Road", "Maxcare Hospital Manchar", "Paragaon Phata", "Market Yard", "Ambethan Chowk"],
        "Route 21 (NIGDI)": ["Big India", "Dattawadi", "LIC Corner", "Hatti Chowk"]
    }

}


@app.route("/get-shift-routes", methods=["GET"])
def get_shift_routes():
    shift = request.args.get("shift", "").strip()
    routes = SHIFT_ROUTES.get(shift, {})
    return jsonify({"status": "success", "shift": shift, "routes": routes})


# ==========================
# LIVE BUS LOCATION (shift + route based)
# ==========================
# Admin pushes their current lat/lng (from their phone) for a specific
# shift+route. Anyone viewing that route's tracking link (viewroots.html)
# polls this every couple of seconds to move the bus marker on the map.
#
# NOTE: this is stored in memory (not the database) since it's live,
# constantly-changing data — there's no need to keep old locations
# around. It will reset if the server restarts, which is expected.

LIVE_LOCATIONS = {}


def _location_key(shift, route):
    return (str(shift).strip(), str(route).strip())


@app.route("/update-live-location", methods=["POST"])
@login_required("admin")
def update_live_location():
    try:
        data = request.get_json()
        shift = data.get("shift", "")
        route = data.get("route", "")
        lat = data.get("lat")
        lng = data.get("lng")

        if not shift or not route or lat is None or lng is None:
            return jsonify({"status": "error", "message": "shift, route, lat and lng are all required"})

        LIVE_LOCATIONS[_location_key(shift, route)] = {
            "lat": float(lat),
            "lng": float(lng),
            "updatedAt": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        }

        return jsonify({"status": "success", "message": "Location updated"})

    except Exception as e:
        return jsonify({"status": "error", "message": str(e)})


@app.route("/live-location/<shift>/<route>", methods=["GET"])
def get_live_location(shift, route):
    entry = LIVE_LOCATIONS.get(_location_key(shift, route))

    if not entry:
        return jsonify({"success": False, "message": "No live location yet for this route"})

    return jsonify({
        "success": True,
        "lat": entry["lat"],
        "lng": entry["lng"],
        "updatedAt": entry["updatedAt"]
    })


# ==========================
# LIVE TRACKING LINKS (40 ROUTES)
# ==========================
# Admin pastes a live-tracking link (e.g. a GPS/fleet-tracking share link)
# against each route number. Employees see the link for their own route.

ROUTE_TRACKING_FILE = "route_tracking.xlsx"
TOTAL_ROUTES = 40


def create_route_tracking_file():
    default_df = pd.DataFrame({
        "routeNumber": [f"Route {i}" for i in range(1, TOTAL_ROUTES + 1)],
        "trackingLink": [""] * TOTAL_ROUTES,
        "updatedDate": [""] * TOTAL_ROUTES
    })
    init_table_if_needed(
        "route_tracking",
        ROUTE_TRACKING_FILE,
        list(default_df.columns),
        default_df=default_df
    )


def read_route_tracking():
    df = read_table("route_tracking", ROUTE_TRACKING_FILE)
    df.columns = df.columns.str.strip()
    return df


def save_route_tracking(df):
    save_table(df, "route_tracking", ROUTE_TRACKING_FILE)


create_route_tracking_file()


@app.route("/admin/route-tracking", methods=["GET"])
@login_required("admin")
def admin_get_route_tracking():
    try:
        df = read_route_tracking().fillna("")
        return jsonify(df.to_dict(orient="records"))
    except Exception as e:
        return jsonify({"status": "error", "message": str(e)})


@app.route("/admin/update-route-tracking", methods=["POST"])
@login_required("admin")
def admin_update_route_tracking():
    try:
        data = request.get_json()
        route_number = str(data.get("routeNumber", "")).strip()
        tracking_link = str(data.get("trackingLink", "")).strip()

        if route_number == "":
            return jsonify({"status": "error", "message": "Route Number Required"})

        df = read_route_tracking()
        match = df["routeNumber"].astype(str).str.strip() == route_number

        if not match.any():
            return jsonify({"status": "error", "message": "Route Not Found"})

        df.loc[match, "trackingLink"] = tracking_link
        df.loc[match, "updatedDate"] = datetime.now().strftime("%Y-%m-%d %H:%M")

        save_route_tracking(df)

        return jsonify({"status": "success", "message": "Tracking Link Updated"})

    except Exception as e:
        return jsonify({"status": "error", "message": str(e)})


@app.route("/my-route-tracking", methods=["GET"])
@login_required("employee")
def my_route_tracking():
    try:
        employee_code = session.get("employeeCode")

        emp_df = load_df()
        emp_match = emp_df["employeeCode"].astype(str).str.strip() == str(employee_code).strip()

        if not emp_match.any():
            return jsonify({"status": "error", "message": "Employee Not Found"})

        route_number = str(emp_df.loc[emp_match, "routeNumber"].iloc[0]).strip()

        if route_number == "" or route_number.lower() == "nan":
            return jsonify({"status": "error", "message": "No Route Assigned Yet"})

        track_df = read_route_tracking()
        track_match = track_df["routeNumber"].astype(str).str.strip() == route_number

        if not track_match.any():
            return jsonify({"status": "error", "message": "Route Not Found"})

        link = str(track_df.loc[track_match, "trackingLink"].iloc[0]).strip()

        if link == "" or link.lower() == "nan":
            return jsonify({
                "status": "error",
                "message": "Live tracking link not added yet for your route"
            })

        return jsonify({
            "status": "success",
            "routeNumber": route_number,
            "trackingLink": link
        })

    except Exception as e:
        return jsonify({"status": "error", "message": str(e)})


# ==========================
# REPORTS / CHARTS DATA
# ==========================

@app.route("/admin/report-data", methods=["GET"])
@login_required("admin")
def admin_report_data():
    try:
        # ---- Employees by Department ----
        emp_df = load_df()
        dept_counts = (
            emp_df["department"].fillna("Unassigned").replace("", "Unassigned")
            .value_counts().to_dict()
            if "department" in emp_df.columns and not emp_df.empty else {}
        )

        # ---- Employees by Route ----
        route_counts = (
            emp_df["routeNumber"].fillna("Unassigned").replace("", "Unassigned")
            .value_counts().to_dict()
            if "routeNumber" in emp_df.columns and not emp_df.empty else {}
        )

        # ---- Pass Status (Permanent + Temporary combined) ----
        pass_status = {"Pending": 0, "Approved": 0, "Rejected": 0}

        perm_requests = read_pass_requests()
        if not perm_requests.empty and "status" in perm_requests.columns:
            for status, count in perm_requests["status"].value_counts().items():
                if status in pass_status:
                    pass_status[status] += int(count)

        temp_requests = read_temp_passes()
        if not temp_requests.empty and "status" in temp_requests.columns:
            for status, count in temp_requests["status"].value_counts().items():
                if status in pass_status:
                    pass_status[status] += int(count)

        # ---- Live Tracking Coverage ----
        track_df = read_route_tracking()
        with_link = 0
        without_link = TOTAL_ROUTES

        if not track_df.empty and "trackingLink" in track_df.columns:
            with_link = int((track_df["trackingLink"].astype(str).str.strip() != "").sum())
            without_link = len(track_df) - with_link

        return jsonify({
            "status": "success",
            "employeesByDepartment": dept_counts,
            "employeesByRoute": route_counts,
            "passStatus": pass_status,
            "trackingCoverage": {
                "withLink": with_link,
                "withoutLink": without_link
            }
        })

    except Exception as e:
        return jsonify({"status": "error", "message": str(e)})


# ==========================
# OT MANPOWER APPROVAL
# ==========================
# Public page — no login required. Saves each OT request and returns
# email details so the browser can open a pre-filled mailto: link.

import json as _json

OT_FILE = "ot_requests.xlsx"

OT_COLS = [
    "otId", "submittedBy", "shift", "department", "otherDepartment",
    "approvalDate", "approvalTime", "isSuddenRequirement",
    "canteenData", "transportData", "createdAt"
]

# The department list used on the OT Manpower Approval form.
OT_DEPARTMENTS = [
    "Production", "Quality", "Maintenance", "HR", "Administration",
    "Alluminium Shop", "BU HR -Operations", "Central Maintenance",
    "Assembly Planned Maint", "Chassis Sub Assembly Center", "Civil",
    "Dispensary", "Civil & Utility", "Engine Assembly Line-KTM",
    "Engine Assembly", "Export", "Exoprt Assembly CKD", "Export Open",
    "Facility Engineering", "Flying Start GT-OP", "HRD", "Machining",
    "Maintenance", "Manufacturing Check", "Manufacturing Engine",
    "ME (E&T)", "ME(Vehicle)", "Paint Shop", "Personnel", "PPC",
    "Production", "Production Planning", "Vehicel Assembly-Pulsar",
    "Quality", "Quality Assurance", "Reliability Sub Vehicle",
    "Reliability Supply Vehicle", "Safety", "Security", "Steel Shop",
    "Steel Shop (C-10)", "Time Office", "Tool Room", "TPM",
    "Utilities & Services", "Vehicle Assembly Electric", "Vehicle Dispatch ",
    "Vehicle Assembly", "Works Admin (C-01)"
]

# Who should receive the approval email. Set OT_APPROVAL_EMAIL as an
# environment variable to change this without touching the code.
OT_APPROVAL_EMAIL = os.environ.get("OT_APPROVAL_EMAIL", "transport.admin@bajaj.com")


def create_ot_file():
    init_table_if_needed("ot_requests", OT_FILE, OT_COLS)


create_ot_file()


@app.route("/get-departments", methods=["GET"])
def get_departments():
    return jsonify({"status": "success", "departments": OT_DEPARTMENTS})


@app.route("/submit-ot", methods=["POST"])
@limiter.limit("20 per hour")
def submit_ot():

    try:
        shift = request.form.get("shift", "").strip()
        department = request.form.get("department", "").strip()
        other_department = request.form.get("other_department", "").strip()
        approval_date = request.form.get("approval_date", "").strip()
        approval_time = request.form.get("approval_time", "").strip()
        canteen_data_raw = request.form.get("canteen_data", "[]")
        transport_data_raw = request.form.get("transport_data", "[]")
        user_name = request.form.get("user_name", "").strip() or "Unknown"
        is_sudden = request.form.get("is_sudden_requirement", "0")

        if department == "":
            return jsonify({"status": "error", "message": "Department is required"})

        try:
            canteen_data = _json.loads(canteen_data_raw)
        except Exception:
            canteen_data = []

        try:
            transport_data = _json.loads(transport_data_raw)
        except Exception:
            transport_data = []

        final_department = other_department if department == "Other" and other_department else department

        ot_id = "OT" + str(int(datetime.now().timestamp()))

        df = read_table("ot_requests", OT_FILE)

        new_row = {
            "otId": ot_id,
            "submittedBy": user_name,
            "shift": shift,
            "department": final_department,
            "otherDepartment": other_department,
            "approvalDate": approval_date,
            "approvalTime": approval_time,
            "isSuddenRequirement": is_sudden,
            "canteenData": _json.dumps(canteen_data),
            "transportData": _json.dumps(transport_data),
            "createdAt": datetime.now().strftime("%Y-%m-%d %H:%M")
        }

        df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
        save_table(df, "ot_requests", OT_FILE)

        # ---- Build a readable email for the approver ----
        subject = f"OT Manpower Request - {final_department} - {shift} - {approval_date}"
        if is_sudden == "1":
            subject = "[URGENT] " + subject

        lines = []
        lines.append(f"Department: {final_department}")
        lines.append(f"Shift: {shift}")
        lines.append(f"Date: {approval_date}   Time: {approval_time}")
        lines.append(f"Submitted By: {user_name}")
        if is_sudden == "1":
            lines.append("")
            lines.append("*** SUDDEN / ADDITIONAL OT REQUEST ***")
        lines.append("")
        lines.append("---- Canteen Facility Requirement ----")
        if canteen_data:
            for item in canteen_data:
                lines.append(
                    f"{item.get('category','')}: 2hr={item.get('hours_2',0)}, 3hr={item.get('hours_3',0)}"
                )
        else:
            lines.append("None requested")
        lines.append("")
        lines.append("---- Transportation Requirement ----")
        if transport_data:
            for item in transport_data:
                lines.append(
                    f"{item.get('route_name','')} ({item.get('bus_stops','')}): "
                    f"2hr={item.get('hours_2',0)}, 3hr={item.get('hours_3',0)}"
                )
        else:
            lines.append("None requested")

        body = "\n".join(lines)

        return jsonify({
            "status": "success",
            "message": "OT request submitted successfully",
            "to": OT_APPROVAL_EMAIL,
            "email_subject": subject,
            "email_body": body
        })

    except Exception as e:
        return jsonify({"status": "error", "message": str(e)})


# ==========================
# RUN SERVER
# ==========================

if __name__ == "__main__":
    # debug mode should NEVER be on when the site is public (it leaks
    # internal code/data on errors). Set FLASK_DEBUG=True locally if needed.
    app.run(debug=os.environ.get("FLASK_DEBUG", "False") == "True")

