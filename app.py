from flask import Flask, request, jsonify, send_from_directory
from flask_cors import CORS
import pandas as pd
import os

app = Flask(__name__)
CORS(app)

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

        # Password Check
        db_password = str(row["password"]).strip()

        if db_password != password:
            return jsonify({
                "status": "error",
                "message": "Wrong Password"
            })

        # Login Success
        return jsonify({

            "status": "success",
            "message": "Login Successful",

            "passNumber": str(row.get("passNumber", "")),
            "employeeCode": str(row.get("employeeCode", "")),
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

    if email == "admin@bajaj.com" and password == "admin123":
        return jsonify({
            "status": "success",
            "message": "Login Successful"
        })

    return jsonify({
        "status": "error",
        "message": "Wrong Credentials"
    })





# ==========================
# ADMIN DASHBOARD STATS
# ==========================

@app.route("/dashboard-stats")
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
def employees():

    try:

        df = load_df()

        return df.to_html(index=False)

    except Exception as e:

        return f"<h3>Error : {str(e)}</h3>"
    

# ==========================
# ASSIGN ROUTE
# ==========================

@app.route("/assign-route", methods=["POST"])
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
def bus_pass():

    try:

        data = request.get_json()

        employeeCode = data.get("employeeCode")

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
def check_pass(employeeCode):

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
def my_pass(employeeCode):

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
# GET ALL PASS REQUESTS
# ==========================

@app.route("/admin/pass-requests", methods=["GET"])
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
def employees_json():

    try:

        # employees.xlsx मधील सर्व employees load करा
        df = load_df()

        # रिकामे values "" करा
        df = df.fillna("")

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
# RUN SERVER
# ==========================

if __name__ == "__main__":
    app.run(debug=True)

