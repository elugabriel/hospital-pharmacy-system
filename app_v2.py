# -------------------- IMPORTS --------------------
from flask import Flask, render_template, request, redirect, url_for, session, flash, jsonify, send_file
from werkzeug.security import generate_password_hash, check_password_hash
import psycopg2
import os
from datetime import datetime, date, timedelta
from calendar import month_name
import uuid
import json
import io
import bcrypt
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill

# -------------------- FLASK APP SETUP --------------------
app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "super_secret_key_change_later")

# -------------------- DATABASE CONFIGURATION --------------------
DATABASE_URL = os.environ.get(
    "DATABASE_URL", "postgresql://flask_user:Olarewaju1.@localhost:5432/hospital_db"
)

def get_db_connection():
    """Establish database connection with error handling."""
    try:
        conn = psycopg2.connect(DATABASE_URL)
        return conn
    except Exception as e:
        app.logger.error(f"Database connection error: {e}")
        return None

# -------------------- DATABASE INITIALIZATION --------------------
def create_tables():
    """Create all necessary tables if they don't exist."""
    queries = {
        "pharmacists": """
            CREATE TABLE IF NOT EXISTS pharmacists (
                id SERIAL PRIMARY KEY,
                username VARCHAR(50) UNIQUE NOT NULL,
                password VARCHAR(255) NOT NULL,
                is_active BOOLEAN DEFAULT TRUE,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            );
        """,
        "drugs": """
            CREATE TABLE IF NOT EXISTS drugs (
                id SERIAL PRIMARY KEY,
                name VARCHAR(100) NOT NULL,
                strength VARCHAR(50) NOT NULL,
                unit_price DECIMAL(10, 2) NOT NULL,
                stock_quantity INT NOT NULL,
                expiry_date DATE NOT NULL,
                low_stock_threshold INT DEFAULT 20,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            );
        """,
        "drug_sales": """
            CREATE TABLE IF NOT EXISTS drug_sales (
                id SERIAL PRIMARY KEY,
                receipt_no VARCHAR(50) UNIQUE NOT NULL,
                patient_name VARCHAR(100),
                patient_id VARCHAR(50),
                items JSONB NOT NULL,
                subtotal DECIMAL(10, 2) NOT NULL,
                discount DECIMAL(10, 2) DEFAULT 0.00,
                tax DECIMAL(10, 2) DEFAULT 0.00,
                grand_total DECIMAL(10, 2) NOT NULL,
                pharmacist VARCHAR(50) NOT NULL,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            );
        """,
        "receipts": """
            CREATE TABLE IF NOT EXISTS receipts (
                id SERIAL PRIMARY KEY,
                patient_name VARCHAR(100),
                patient_id VARCHAR(50),
                subtotal DECIMAL(10, 2) NOT NULL,
                discount DECIMAL(10, 2) DEFAULT 0.00,
                tax DECIMAL(10, 2) DEFAULT 0.00,
                total_amount DECIMAL(10, 2) NOT NULL,
                grand_total DECIMAL(10, 2) NOT NULL,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            );
        """,
        "receipt_items": """
            CREATE TABLE IF NOT EXISTS receipt_items (
                id SERIAL PRIMARY KEY,
                receipt_id INT NOT NULL REFERENCES receipts(id),
                drug_name VARCHAR(100) NOT NULL,
                strength VARCHAR(50) NOT NULL,
                quantity INT NOT NULL,
                unit_price DECIMAL(10, 2) NOT NULL
            );
        """,
        "stock_movements": """
            CREATE TABLE IF NOT EXISTS stock_movements (
                id SERIAL PRIMARY KEY,
                drug_id INT NOT NULL REFERENCES drugs(id),
                movement_type VARCHAR(20) NOT NULL,
                quantity INT NOT NULL,
                user_id INT NOT NULL,
                note TEXT,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            );
        """,
        "billing_users": """
            CREATE TABLE IF NOT EXISTS billing_users (
                id SERIAL PRIMARY KEY,
                username VARCHAR(50) UNIQUE NOT NULL,
                password VARCHAR(255) NOT NULL,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            );
        """,
        "billing_invoice": """
            CREATE TABLE IF NOT EXISTS billing_invoice (
                id SERIAL PRIMARY KEY,
                patient_name VARCHAR(100) NOT NULL,
                service_type VARCHAR(100) NOT NULL,
                amount DECIMAL(10, 2) NOT NULL,
                status VARCHAR(20) DEFAULT 'UNPAID',
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            );
        """,
        "billing_receipt": """
            CREATE TABLE IF NOT EXISTS billing_receipt (
                id SERIAL PRIMARY KEY,
                invoice_id INT NOT NULL REFERENCES billing_invoice(id),
                amount_paid DECIMAL(10, 2) NOT NULL,
                payment_method VARCHAR(50) NOT NULL,
                received_by VARCHAR(50) NOT NULL,
                payment_date TIMESTAMP NOT NULL
            );
        """,
        "payments": """
            CREATE TABLE IF NOT EXISTS payments (
                id SERIAL PRIMARY KEY,
                patient_name VARCHAR(100) NOT NULL,
                service_type VARCHAR(100) NOT NULL,
                subtotal DECIMAL(10, 2) NOT NULL,
                discount DECIMAL(10, 2) DEFAULT 0.00,
                tax DECIMAL(10, 2) DEFAULT 0.00,
                grand_total DECIMAL(10, 2) NOT NULL,
                amount_paid DECIMAL(10, 2) NOT NULL,
                balance DECIMAL(10, 2) NOT NULL,
                payment_method VARCHAR(50) NOT NULL,
                status VARCHAR(20) NOT NULL,
                payment_date DATE NOT NULL,
                recorded_by INT NOT NULL,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            );
        """,
        "users": """
            CREATE TABLE IF NOT EXISTS users (
                id SERIAL PRIMARY KEY,
                username VARCHAR(50) UNIQUE NOT NULL,
                password VARCHAR(255) NOT NULL,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            );
        """
        
    }

    conn = get_db_connection()
    if not conn:
        return

    cursor = conn.cursor()
    for table, query in queries.items():
        try:
            cursor.execute(query)
        except Exception as e:
            app.logger.error(f"Error creating table {table}: {e}")
    conn.commit()
    cursor.close()
    conn.close()

def create_default_users():
    """Create default users for pharmacy and billing modules."""
    default_users = {
        "pharmacists": ("pharmacist1", "pharma123"),
        "billing_users": ("billing1", "billing123")
    }

    conn = get_db_connection()
    if not conn:
        return

    cursor = conn.cursor()
    for table, (username, password) in default_users.items():
        hashed_pw = generate_password_hash(password)
        try:
            cursor.execute(f"""
                INSERT INTO {table} (username, password)
                VALUES (%s, %s)
                ON CONFLICT (username) DO UPDATE SET password = EXCLUDED.password
            """, (username, hashed_pw))
        except Exception as e:
            app.logger.error(f"Error creating default user {username}: {e}")
    conn.commit()
    cursor.close()
    conn.close()

# -------------------- HELPER FUNCTIONS --------------------
def format_currency(amount):
    """Format amount as Nigerian Naira currency."""
    return f"₦{amount:,.2f}"


def build_stock_snapshot(rows, today):
    stock = []
    for r in rows:
        expiry_date = r[5]
        quantity = r[3]
        threshold = r[6] or 20

        if expiry_date:
            days_left = (expiry_date - today).days
            status = "EXPIRED" if days_left < 0 else "EXPIRING_SOON" if days_left <= 30 else "VALID"
        else:
            days_left = None
            status = "UNKNOWN"

        stock.append({
            "id": r[0], "name": r[1], "strength": r[2],
            "quantity": quantity, "unit_price": r[4],
            "expiry_date": expiry_date, "days_left": days_left,
            "status": status, "low_stock_threshold": threshold,
            "total_value": quantity * r[4]
        })
    return stock

def apply_stock_filter(stock, filter_type):
    if filter_type == "expired":
        return [d for d in stock if d["status"] == "EXPIRED"]
    elif filter_type in ("expiring", "expiring_soon"):
        return [d for d in stock if d["status"] == "EXPIRING_SOON"]
    elif filter_type in ("low", "low_stock"):
        return [d for d in stock if d["quantity"] <= d["low_stock_threshold"]]
    return stock

# -------------------- ROUTES: LANDING & MODULES --------------------
@app.route('/')
def landing_page():
    hospital_name = "Memorial Hospital Ovuru, Nsukka, Enugu State"
    modules = [
        "System Admin", "Patient Services", "Clinical Services",
        "Pharmacy", "Laboratory", "Radiology", "Billing and Revenue",
        "Human Resources", "Management and Reports"
    ]
    return render_template("dashboard.html", hospital_name=hospital_name, modules=modules)

@app.route('/<module_name>')
def module_placeholder(module_name):
    display_name = module_name.replace('_', ' ').title()
    if module_name.lower() == "pharmacy":
        return redirect(url_for('pharmacy_login'))
    return render_template("module_placeholder.html", module_name=display_name)

# -------------------- ROUTES: PHARMACY MODULE --------------------
@app.route('/pharmacy/login', methods=['GET', 'POST'])
def pharmacy_login():
    if request.method == 'POST':
        username = request.form['username']
        password = request.form['password']

        conn = get_db_connection()
        if not conn:
            flash("Database connection error", "danger")
            return render_template("pharmacy_login.html")

        cursor = conn.cursor()
        cursor.execute(
            "SELECT id, username, password FROM pharmacists WHERE username=%s AND is_active=TRUE",
            (username,)
        )
        pharmacist = cursor.fetchone()
        cursor.close()
        conn.close()

        if pharmacist and check_password_hash(pharmacist[2], password):
            session['pharmacist_id'] = pharmacist[0]
            session['pharmacist_username'] = pharmacist[1]
            return redirect(url_for('pharmacy_dashboard'))
        else:
            flash("Invalid username or password", "danger")

    return render_template("pharmacy_login.html")

@app.route('/pharmacy/dashboard')
def pharmacy_dashboard():
    if 'pharmacist_id' not in session:
        return redirect(url_for('pharmacy_login'))
    return render_template(
        "pharmacy_dashboard.html",
        pharmacist_name=session.get('pharmacist_username')
    )

@app.route('/pharmacy/logout')
def pharmacy_logout():
    session.clear()
    return redirect(url_for('pharmacy_login'))

@app.route('/pharmacy/drug_sales')
def drug_sales():
    if 'pharmacist_id' not in session:
        return redirect(url_for('pharmacy_login'))
    return render_template(
        "drug_sales_dashboard.html",
        pharmacist_name=session.get('pharmacist_username'),
        hospital_name="Memorial Hospital Ovuru, Nsukka, Enugu State"
    )

@app.route('/pharmacy/add-stock', methods=['GET', 'POST'])
def add_stock():
    if 'pharmacist_id' not in session:
        return redirect(url_for('pharmacy_login'))

    if request.method == 'POST':
        drug_name = request.form.get('drug_name', '').strip()
        strength = request.form.get('strength', '').strip()
        unit_price = request.form.get('unit_price', '').strip()
        quantity = request.form.get('quantity', '').strip()
        expiry_date = request.form.get('expiry_date', '').strip()

        if not all([drug_name, strength, unit_price, quantity, expiry_date]):
            flash("All fields including expiry date are required.", "danger")
            return redirect(url_for('add_stock'))

        try:
            unit_price = float(unit_price)
            quantity = int(quantity)
            expiry_date_obj = datetime.strptime(expiry_date, "%Y-%m-%d").date()
        except ValueError:
            flash("Invalid input data.", "danger")
            return redirect(url_for('add_stock'))

        conn = get_db_connection()
        if not conn:
            flash("Database connection error.", "danger")
            return redirect(url_for('add_stock'))

        cursor = conn.cursor()
        cursor.execute("""
            SELECT id, stock_quantity FROM drugs
            WHERE name = %s AND strength = %s AND expiry_date = %s
        """, (drug_name, strength, expiry_date_obj))

        existing = cursor.fetchone()

        if existing:
            cursor.execute("""
                UPDATE drugs
                SET stock_quantity = stock_quantity + %s,
                    unit_price = %s,
                    updated_at = NOW()
                WHERE id = %s
            """, (quantity, unit_price, existing[0]))
        else:
            cursor.execute("""
                INSERT INTO drugs (name, strength, unit_price, stock_quantity, expiry_date)
                VALUES (%s, %s, %s, %s, %s)
            """, (drug_name, strength, unit_price, quantity, expiry_date_obj))

        conn.commit()
        cursor.close()
        conn.close()

        flash("Stock added successfully.", "success")
        return redirect(url_for('add_stock'))

    return render_template("add_stock.html")

@app.route('/api/drugs')
def api_drugs():
    if 'pharmacist_id' not in session:
        return jsonify([])

    search = request.args.get('q', '').strip()
    conn = get_db_connection()
    if not conn:
        return jsonify([])

    cur = conn.cursor()
    query = """
        SELECT id, name, strength, unit_price, stock_quantity
        FROM drugs
        WHERE stock_quantity > 0
    """
    params = ()
    
    if search:
        query += " AND LOWER(name) LIKE LOWER(%s)"
        params = (f"{search}%",)
    
    query += " ORDER BY name ASC"
    cur.execute(query, params)
    
    rows = cur.fetchall()
    cur.close()
    conn.close()

    return jsonify([{
        "id": r[0], "name": r[1], "strength": r[2],
        "unit_price": float(r[3]), "stock_quantity": r[4]
    } for r in rows])

@app.route('/pharmacy/receipt', methods=['POST'])
def pharmacy_receipt():
    if 'pharmacist_id' not in session:
        return redirect(url_for('pharmacy_login'))

    data = request.json
    receipt_no = f"RX-{uuid.uuid4().hex[:8].upper()}"

    conn = get_db_connection()
    cur = conn.cursor()

    cur.execute("""
        INSERT INTO drug_sales
        (receipt_no, items, subtotal, discount, tax, grand_total, pharmacist)
        VALUES (%s, %s, %s, %s, %s, %s, %s)
    """, (
        receipt_no,
        json.dumps(data["items"]),
        data["subtotal"],
        data["discount"],
        data["tax"],
        data["grand_total"],
        session.get('pharmacist_username')
    ))

    conn.commit()
    cur.close()
    conn.close()

    data["receipt_no"] = receipt_no
    return render_template(
        "receipt.html",
        receipt=data,
        hospital_name="Memorial Hospital Ovuru, Nsukka, Enugu State",
        pharmacist_name=session.get('pharmacist_username')
    )

@app.route('/pharmacy/save-patient', methods=['POST'])
def save_patient_info():
    if 'pharmacist_id' not in session:
        return redirect(url_for('pharmacy_login'))

    receipt_no = request.form['receipt_no']
    patient_name = request.form['patient_name']
    patient_id = request.form.get('patient_id')

    conn = get_db_connection()
    cur = conn.cursor()

    cur.execute("""
        UPDATE drug_sales
        SET patient_name=%s, patient_id=%s
        WHERE receipt_no=%s
    """, (patient_name, patient_id, receipt_no))

    conn.commit()
    cur.close()
    conn.close()

    return redirect(url_for('reprint_receipt', receipt_no=receipt_no))

@app.route('/pharmacy/receipt/<receipt_no>')
def reprint_receipt(receipt_no):
    if 'pharmacist_id' not in session:
        return redirect(url_for('pharmacy_login'))

    conn = get_db_connection()
    cur = conn.cursor()

    cur.execute("""
        SELECT receipt_no, patient_name, patient_id, items,
               subtotal, discount, tax, grand_total, pharmacist, created_at
        FROM drug_sales
        WHERE receipt_no = %s
    """, (receipt_no,))

    row = cur.fetchone()
    cur.close()
    conn.close()

    if not row:
        flash("Receipt not found", "danger")
        return redirect(url_for('pharmacy_dashboard'))

    receipt = {
        "receipt_no": row[0], "patient_name": row[1], "patient_id": row[2],
        "items": json.loads(row[3]) if row[3] else [],
        "subtotal": float(row[4]), "discount": float(row[5]),
        "tax": float(row[6]), "grand_total": float(row[7]),
        "pharmacist": row[8], "date": row[9].strftime("%Y-%m-%d %H:%M:%S")
    }

    return render_template("receipt.html", receipt=receipt, hospital_name="Memorial Hospital Ovuru, Nsukka, Enugu State")

@app.route("/pharmacy/confirm-payment", methods=["POST"])
def confirm_payment():
    if "pharmacist_id" not in session:
        return jsonify({"success": False, "message": "Unauthorized"}), 401

    data = request.get_json()
    conn = get_db_connection()
    cur = conn.cursor()

    try:
        cur.execute("""
            INSERT INTO receipts (
                patient_name, patient_id, subtotal, discount, tax,
                total_amount, grand_total, created_at
            )
            VALUES (%s, %s, %s, %s, %s, %s, %s, %s)
            RETURNING id;
        """, (
            data.get("patient_name"), data.get("patient_id"),
            data["subtotal"], data["discount"], data["tax"],
            data["grand_total"], data["grand_total"], datetime.now()
        ))

        receipt_id = cur.fetchone()[0]

        for item in data["items"]:
            cur.execute("""
                INSERT INTO receipt_items (
                    receipt_id, drug_name, strength, quantity, unit_price
                )
                VALUES (%s, %s, %s, %s, %s);
            """, (
                receipt_id, item["drug_name"], item["strength"],
                item["quantity"], item["unit_price"]
            ))

            cur.execute("""
                UPDATE drugs
                SET stock_quantity = stock_quantity - %s,
                    updated_at = NOW()
                WHERE name = %s AND strength = %s;
            """, (
                item["quantity"], item["drug_name"], item["strength"]
            ))

        conn.commit()
        return jsonify({"success": True, "receipt_id": receipt_id})

    except Exception as e:
        conn.rollback()
        return jsonify({"success": False, "error": str(e)}), 500

    finally:
        cur.close()
        conn.close()

@app.route("/pharmacy/stock-report")
def stock_report():
    if 'pharmacist_id' not in session:
        return redirect(url_for('pharmacy_login'))

    filter_type = request.args.get("filter", "all")
    conn = get_db_connection()
    cur = conn.cursor()
    cur.execute("""
        SELECT id, name, strength, stock_quantity, unit_price, 
               expiry_date, low_stock_threshold
        FROM drugs
        ORDER BY expiry_date ASC
    """)
    rows = cur.fetchall()
    cur.close()
    conn.close()

    stock = build_stock_snapshot(rows, date.today())
    stock = apply_stock_filter(stock, filter_type)

    return render_template(
        "stock_report.html",
        stock=stock,
        current_filter=filter_type,
        expired_count=sum(1 for d in stock if d["status"] == "EXPIRED"),
        expiring_soon_count=sum(1 for d in stock if d["status"] == "EXPIRING_SOON"),
        low_stock_count=sum(1 for d in stock if d["quantity"] <= d["low_stock_threshold"]),
        total_stock_value=sum(d["total_value"] for d in stock)
    )

@app.route("/pharmacy/stock-report/export")
def export_stock_report():
    if 'pharmacist_id' not in session:
        return redirect(url_for('pharmacy_login'))

    filter_type = request.args.get("filter", "all")
    conn = get_db_connection()
    cur = conn.cursor()
    cur.execute("""
        SELECT id, name, strength, stock_quantity, unit_price, 
               expiry_date, low_stock_threshold
        FROM drugs
        ORDER BY expiry_date ASC
    """)
    rows = cur.fetchall()
    cur.close()
    conn.close()

    stock = build_stock_snapshot(rows, date.today())
    stock = apply_stock_filter(stock, filter_type)

    wb = Workbook()
    ws = wb.active
    ws.title = "Stock Report"

    headers = [
        "Drug Name", "Strength", "Quantity", "Unit Price (₦)",
        "Expiry Date", "Days Left", "Status", "Total Value (₦)", "Low Stock Threshold"
    ]
    ws.append(headers)

    for c in range(1, len(headers) + 1):
        ws.cell(row=1, column=c).font = Font(bold=True)

    fills = {
        "EXPIRED": PatternFill("solid", fgColor="FF9999"),
        "EXPIRING_SOON": PatternFill("solid", fgColor="FFFF99"),
        "LOW": PatternFill("solid", fgColor="ADD8E6")
    }

    for item in stock:
        ws.append([
            item["name"], item["strength"], item["quantity"],
            float(item["unit_price"]), item["expiry_date"],
            item["days_left"], item["status"],
            float(item["total_value"]), item["low_stock_threshold"]
        ])

        row_idx = ws.max_row
        if item["status"] in fills:
            for col in range(1, len(headers) + 1):
                ws.cell(row=row_idx, column=col).fill = fills[item["status"]]
        elif item["quantity"] <= item["low_stock_threshold"]:
            for col in range(1, len(headers) + 1):
                ws.cell(row=row_idx, column=col).fill = fills["LOW"]

    stream = io.BytesIO()
    wb.save(stream)
    stream.seek(0)

    return send_file(
        stream,
        as_attachment=True,
        download_name=f"pharmacy_stock_report_{date.today()}.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

@app.route('/pharmacy/stock-movements')
def stock_movements():
    if 'pharmacist_id' not in session:
        return redirect(url_for('pharmacy_login'))

    conn = get_db_connection()
    cur = conn.cursor()
    cur.execute("""
        SELECT sm.id, d.name, d.strength, sm.movement_type, 
               sm.quantity, u.username, sm.created_at, sm.note
        FROM stock_movements sm
        JOIN drugs d ON sm.drug_id = d.id
        JOIN users u ON sm.user_id = u.id
        ORDER BY sm.created_at DESC
    """)
    movements = cur.fetchall()
    cur.close()
    conn.close()

    return render_template('stock_movements.html', movements=movements)

@app.route("/pharmacy/revenue-report", methods=["GET", "POST"])
def revenue_report():
    if 'pharmacist_id' not in session:
        return redirect(url_for('pharmacy_login'))

    report_type = request.form.get("period", "daily")
    selected_day = request.form.get("day")
    selected_month = request.form.get("month")
    selected_year = request.form.get("year")
    today = date.today()

    if report_type == "daily":
        start_date = end_date = datetime.strptime(selected_day, "%Y-%m-%d").date() if selected_day else today
    elif report_type == "weekly":
        d = datetime.strptime(selected_day, "%Y-%m-%d").date() if selected_day else today
        start_date = d - timedelta(days=d.weekday())
        end_date = start_date + timedelta(days=6)
    elif report_type == "monthly":
        month = int(selected_month) if selected_month else today.month
        year = int(selected_year) if selected_year else today.year
        start_date = date(year, month, 1)
        if month == 12:
            end_date = date(year + 1, 1, 1) - timedelta(days=1)
        else:
            end_date = date(year, month + 1, 1) - timedelta(days=1)
    else:
        flash("Invalid report period", "danger")
        return redirect(url_for("pharmacy_dashboard"))

    conn = get_db_connection()
    cur = conn.cursor()
    cur.execute("""
        SELECT id, patient_name, patient_id, grand_total, created_at
        FROM receipts
        WHERE DATE(created_at) BETWEEN %s AND %s
        ORDER BY created_at ASC;
    """, (start_date, end_date))
    sales = cur.fetchall()
    cur.close()
    conn.close()

    total_revenue = sum(float(s[3]) if s[3] is not None else 0.0 for s in sales)
    months = [(i, month_name[i]) for i in range(1, 13)]
    years = range(2024, today.year + 1)

    return render_template(
        "revenue_report.html",
        sales=sales,
        total_revenue=total_revenue,
        period=report_type,
        start_date=start_date,
        end_date=end_date,
        selected_day=selected_day,
        selected_month=int(selected_month) if selected_month else today.month,
        selected_year=int(selected_year) if selected_year else today.year,
        months=months,
        years=years
    )

@app.route("/receipt/<int:receipt_id>")
def receipt(receipt_id):
    conn = get_db_connection()
    cur = conn.cursor()

    cur.execute("SELECT * FROM receipts WHERE id = %s;", (receipt_id,))
    receipt = cur.fetchone()

    cur.execute("""
        SELECT drug_name, strength, quantity, unit_price
        FROM receipt_items
        WHERE receipt_id = %s;
    """, (receipt_id,))
    items = cur.fetchall()

    cur.close()
    conn.close()

    return render_template(
        "receipt.html",
        receipt=receipt,
        items=items
    )

@app.route("/pharmacy/receipt/<int:receipt_id>")
def view_receipt(receipt_id):
    if "pharmacist_id" not in session:
        return redirect(url_for("pharmacy_login"))

    conn = get_db_connection()
    cur = conn.cursor()

    cur.execute("""
        SELECT id, patient_name, patient_id, subtotal, discount, tax, grand_total, created_at
        FROM receipts
        WHERE id = %s
    """, (receipt_id,))

    row = cur.fetchone()
    if not row:
        cur.close()
        conn.close()
        flash("Receipt not found", "danger")
        return redirect(url_for("pharmacy_dashboard"))

    receipt = {
        "id": row[0],
        "patient_name": row[1],
        "patient_id": row[2],
        "subtotal": float(row[3]),
        "discount": float(row[4]),
        "tax": float(row[5]),
        "grand_total": float(row[6]),
        "date": row[7].strftime("%Y-%m-%d %H:%M:%S")
    }

    cur.execute("""
        SELECT drug_name, strength, quantity, unit_price
        FROM receipt_items
        WHERE receipt_id = %s
    """, (receipt_id,))

    items_rows = cur.fetchall()
    items = []
    for i in items_rows:
        items.append({
            "drug_name": i[0],
            "strength": i[1],
            "quantity": i[2],
            "unit_price": float(i[3])
        })

    cur.close()
    conn.close()

    return render_template(
        "receipt.html",
        receipt=receipt,
        items=items,
        hospital_name="Memorial Hospital Ovuru, Nsukka, Enugu State"
    )

# -------------------- ROUTES: BILLING MODULE --------------------
@app.route("/billing/login", methods=["GET", "POST"])
def billing_login():
    if request.method == "POST":
        username = request.form["username"]
        password = request.form["password"]

        conn = get_db_connection()
        cur = conn.cursor()

        cur.execute("""
            SELECT id, password
            FROM billing_users
            WHERE username = %s
        """, (username,))

        user = cur.fetchone()
        cur.close()
        conn.close()

        if user and check_password_hash(user[1], password):
            session["billing_user_id"] = user[0]
            session["billing_username"] = username
            return redirect(url_for("billing_dashboard"))
        else:
            flash("Invalid login credentials", "danger")

    return render_template("billing_login.html")

# REPLACE the old billing_dashboard function with this one:

@app.route("/billing/dashboard")
def billing_dashboard():
    if "billing_user_id" not in session:
        return redirect(url_for("billing_login"))

    # No longer need to fetch invoices
    return render_template("billing_dashboard.html")

@app.route("/billing/logout")
def billing_logout():
    session.pop("billing_user_id", None)
    session.pop("billing_username", None)
    flash("Logged out successfully", "success")
    return redirect(url_for("billing_login"))



@app.route("/billing/confirm-payment", methods=["POST"])
def billing_confirm_payment():
    if "billing_user_id" not in session:
        return jsonify({"success": False, "message": "Unauthorized"}), 401

    patient_name = request.form.get("patient_name")
    service_type = request.form.get("service_type")
    receipt_date = request.form.get("receipt_date")
    payment_method = request.form.get("payment_method")
    amount_paid = float(request.form.get("amount_paid", 0))
    vat_percent = float(request.form.get("vat", 0))
    discount = float(request.form.get("discount", 0))

    subtotal = amount_paid
    vat_amount = (subtotal * vat_percent) / 100
    grand_total = subtotal + vat_amount - discount
    balance = 0 if amount_paid >= grand_total else grand_total - amount_paid
    status = "Paid" if balance <= 0 else "Partial"

    payment_date = datetime.strptime(receipt_date, "%Y-%m-%d").date()

    conn = get_db_connection()
    cur = conn.cursor()

    try:
        cur.execute("""
            INSERT INTO payments (
                patient_name, service_type, subtotal, discount, tax,
                grand_total, amount_paid, balance, payment_method,
                status, payment_date, recorded_by
            )
            VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
            RETURNING id;
        """, (
            patient_name, service_type, subtotal, discount, vat_amount,
            grand_total, amount_paid, balance, payment_method,
            status, payment_date, session["billing_user_id"]
        ))

        payment_id = cur.fetchone()[0]
        conn.commit()
        flash(f"Payment recorded successfully. Receipt No: {payment_id}", "success")

        # Redirect to a receipt page
        return redirect(url_for("view_payment_receipt", payment_id=payment_id))

    except Exception as e:
        conn.rollback()
        flash(f"Payment error: {e}", "danger")
        return redirect(url_for("accept_payment_page"))

    finally:
        cur.close()
        conn.close()



@app.route("/billing/receipt/<int:payment_id>")
def billing_receipt(payment_id):
    if "billing_user_id" not in session:
        return redirect(url_for("billing_login"))

    conn = get_db_connection()
    cur = conn.cursor()

    try:
        cur.execute("""
            SELECT id, patient_name, service_type, subtotal, discount, 
                   tax, grand_total, amount_paid, balance, payment_method, 
                   status, payment_date, created_at
            FROM payments
            WHERE id = %s
        """, (payment_id,))

        row = cur.fetchone()
        if not row:
            flash("Receipt not found", "danger")
            return redirect(url_for("billing_dashboard"))

        receipt = {
            "id": row[0],
            "patient_name": row[1],
            "service_type": row[2],
            "subtotal": float(row[3]),
            "discount": float(row[4]),
            "tax": float(row[5]),
            "grand_total": float(row[6]),
            "amount_paid": float(row[7]),
            "balance": float(row[8]),
            "payment_method": row[9],
            "status": row[10],
            "payment_date": row[11].strftime("%Y-%m-%d"),
            "created_at": row[12].strftime("%Y-%m-%d %H:%M:%S")
        }

        cur.close()
        conn.close()

        return render_template(
            "billing_receipt.html",
            receipt=receipt,
            hospital_name="Memorial Hospital Ovuru, Nsukka, Enugu State",
            user_name=session.get("billing_username")
        )

    except Exception as e:
        cur.close()
        conn.close()
        flash(f"Error retrieving receipt: {e}", "danger")
        return redirect(url_for("billing_dashboard"))

# 


@app.route("/billing/accept-payment", methods=["GET"])
def accept_payment_page():
    return render_template("accept_payment.html")


@app.route("/billing/receipt/<int:payment_id>")
def billing_receipt_page(payment_id):
    if "billing_user_id" not in session:
        return redirect(url_for("billing_login"))

    conn = get_db_connection()
    cur = conn.cursor()

    # Fetch payment info
    cur.execute("""
        SELECT id, patient_name, service_type, subtotal, discount, tax,
               grand_total, amount_paid, balance, payment_method, payment_date
        FROM payments
        WHERE id = %s
    """, (payment_id,))
    row = cur.fetchone()
    cur.close()
    conn.close()

    if not row:
        flash("Receipt not found", "danger")
        return redirect(url_for("billing_dashboard"))

    receipt = {
        "id": row[0],
        "patient_name": row[1],
        "service_type": row[2],
        "subtotal": float(row[3]),
        "discount": float(row[4]),
        "tax": float(row[5]),
        "grand_total": float(row[6]),
        "amount_paid": float(row[7]),
        "balance": float(row[8]),
        "payment_method": row[9],
        "date": row[10].strftime("%Y-%m-%d %H:%M:%S")
    }

    return render_template("billing_receipt.html", receipt=receipt, hospital_name="Memorial Hospital Ovuru, Nsukka, Enugu State")

@app.route("/billing/receipt/<int:payment_id>")
def view_payment_receipt(payment_id):
    conn = get_db_connection()
    cur = conn.cursor()

    cur.execute("SELECT * FROM payments WHERE id=%s", (payment_id,))
    payment = cur.fetchone()
    cur.close()
    conn.close()

    if not payment:
        flash("Payment not found", "danger")
        return redirect(url_for("billing_dashboard"))

    return render_template("payment_receipt.html", payment=payment)


# -------------------- ROUTES: BILLING MODULE - PAYMENT HISTORY --------------------

@app.route("/billing/payment-history")
def payment_history():
    if "billing_user_id" not in session:
        return redirect(url_for("billing_login"))

    # Get filter parameters
    patient_name = request.args.get("patient_name", "").strip()
    service_type = request.args.get("service_type", "")
    payment_method = request.args.get("payment_method", "")
    status = request.args.get("status", "")
    start_date = request.args.get("start_date")
    end_date = request.args.get("end_date")
    
    # Pagination
    page = request.args.get("page", 1, type=int)
    per_page = 20
    
    # Build query
    conn = get_db_connection()
    cur = conn.cursor()
    
    # Base query
    query = "SELECT * FROM payments WHERE 1=1"
    count_query = "SELECT COUNT(*) FROM payments WHERE 1=1"
    params = []
    
    # Apply filters
    if patient_name:
        query += " AND LOWER(patient_name) LIKE LOWER(%s)"
        count_query += " AND LOWER(patient_name) LIKE LOWER(%s)"
        params.append(f"%{patient_name}%")
    
    if service_type:
        query += " AND service_type = %s"
        count_query += " AND service_type = %s"
        params.append(service_type)
    
    if payment_method:
        query += " AND payment_method = %s"
        count_query += " AND payment_method = %s"
        params.append(payment_method)
    
    if status:
        query += " AND status = %s"
        count_query += " AND status = %s"
        params.append(status)
    
    if start_date:
        query += " AND payment_date >= %s"
        count_query += " AND payment_date >= %s"
        params.append(start_date)
    
    if end_date:
        query += " AND payment_date <= %s"
        count_query += " AND payment_date <= %s"
        params.append(end_date)
    
    # Get total count
    cur.execute(count_query, params)
    total_items = cur.fetchone()[0]
    
    # Apply ordering and pagination
    query += " ORDER BY created_at DESC LIMIT %s OFFSET %s"
    offset = (page - 1) * per_page
    params.extend([per_page, offset])
    
    # Execute main query
    cur.execute(query, params)
    payments = cur.fetchall()
    
    # Get unique service types for dropdown
    cur.execute("SELECT DISTINCT service_type FROM payments WHERE service_type IS NOT NULL ORDER BY service_type")
    service_types = [row[0] for row in cur.fetchall()]
    
    # Calculate total amount
    total_amount = 0
    formatted_payments = []
    for payment in payments:
        payment_dict = {
            "id": payment[0],
            "patient_name": payment[1],
            "service_type": payment[2],
            "subtotal": float(payment[3]),
            "discount": float(payment[4]),
            "tax": float(payment[5]),
            "grand_total": float(payment[6]),
            "amount_paid": float(payment[7]),
            "balance": float(payment[8]),
            "payment_method": payment[9],
            "status": payment[10],
            "payment_date": payment[11],
            "created_at": payment[13]
        }
        formatted_payments.append(payment_dict)
        total_amount += payment_dict["amount_paid"]
    
    cur.close()
    conn.close()
    
    # Calculate pagination
    total_pages = (total_items + per_page - 1) // per_page
    
    return render_template(
        "billing_payment_history.html",
        payments=formatted_payments,
        service_types=service_types,
        total_items=total_items,
        total_amount=total_amount,
        page=page,
        total_pages=total_pages,
        current_filters=request.args
    )


@app.route("/billing/payment-history/export")
def export_payment_history():
    if "billing_user_id" not in session:
        return redirect(url_for("billing_login"))
    
    # Get filter parameters (same as payment_history)
    patient_name = request.args.get("patient_name", "").strip()
    service_type = request.args.get("service_type", "")
    payment_method = request.args.get("payment_method", "")
    status = request.args.get("status", "")
    start_date = request.args.get("start_date")
    end_date = request.args.get("end_date")
    
    conn = get_db_connection()
    cur = conn.cursor()
    
    # Build query without pagination
    query = "SELECT * FROM payments WHERE 1=1"
    params = []
    
    if patient_name:
        query += " AND LOWER(patient_name) LIKE LOWER(%s)"
        params.append(f"%{patient_name}%")
    
    if service_type:
        query += " AND service_type = %s"
        params.append(service_type)
    
    if payment_method:
        query += " AND payment_method = %s"
        params.append(payment_method)
    
    if status:
        query += " AND status = %s"
        params.append(status)
    
    if start_date:
        query += " AND payment_date >= %s"
        params.append(start_date)
    
    if end_date:
        query += " AND payment_date <= %s"
        params.append(end_date)
    
    query += " ORDER BY created_at DESC"
    cur.execute(query, params)
    payments = cur.fetchall()
    
    # Create Excel workbook
    wb = Workbook()
    ws = wb.active
    ws.title = "Payment History"
    
    # Add headers
    headers = [
        "Receipt No", "Patient Name", "Service Type", 
        "Subtotal (₦)", "Discount (₦)", "Tax (₦)", "Grand Total (₦)",
        "Amount Paid (₦)", "Balance (₦)", "Payment Method",
        "Status", "Payment Date", "Created At", "Recorded By"
    ]
    ws.append(headers)
    
    # Style headers
    from openpyxl.styles import Font, PatternFill, Alignment
    header_fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True)
    
    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center")
    
    # Add data rows
    for payment in payments:
        ws.append([
            payment[0],  # id
            payment[1],  # patient_name
            payment[2],  # service_type
            float(payment[3]),  # subtotal
            float(payment[4]),  # discount
            float(payment[5]),  # tax
            float(payment[6]),  # grand_total
            float(payment[7]),  # amount_paid
            float(payment[8]),  # balance
            payment[9],  # payment_method
            payment[10],  # status
            payment[11],  # payment_date
            payment[13],  # created_at
            payment[12]   # recorded_by
        ])
    
    # Auto-adjust column widths
    for column in ws.columns:
        max_length = 0
        column_letter = column[0].column_letter
        for cell in column:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(str(cell.value))
            except:
                pass
        adjusted_width = min(max_length + 2, 30)
        ws.column_dimensions[column_letter].width = adjusted_width
    
    # Add summary row
    ws.append([])
    ws.append(["SUMMARY", "", "", "", "", "", "", "", "", "", "", "", "", ""])
    
    if payments:
        total_amount = sum(float(p[7]) for p in payments)
        total_balance = sum(float(p[8]) for p in payments)
        
        summary_headers = ["Total Payments", "Total Amount", "Total Balance"]
        summary_values = [len(payments), total_amount, total_balance]
        
        for i, (header, value) in enumerate(zip(summary_headers, summary_values)):
            ws.append([header, value])
        
        # Style summary
        summary_row = ws.max_row - len(summary_headers) + 1
        for i in range(len(summary_headers)):
            ws.cell(row=summary_row + i, column=1).font = Font(bold=True)
    
    # Save to BytesIO
    from io import BytesIO
    stream = BytesIO()
    wb.save(stream)
    stream.seek(0)
    
    cur.close()
    conn.close()
    
    # Generate filename
    filename = f"payment_history_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    
    return send_file(
        stream,
        as_attachment=True,
        download_name=filename,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

# -------------------- ROUTES: TODAY'S COLLECTION --------------------

@app.route("/billing/todays-collection")
def todays_collection():
    if "billing_user_id" not in session:
        return redirect(url_for("billing_login"))
    
    today = date.today()
    
    conn = get_db_connection()
    cur = conn.cursor()
    
    # Get today's payments
    cur.execute("""
        SELECT id, patient_name, service_type, amount_paid, 
               payment_method, status, created_at
        FROM payments 
        WHERE DATE(payment_date) = %s
        ORDER BY created_at DESC
    """, (today,))
    
    today_payments = cur.fetchall()
    
    # Calculate totals by payment method
    payment_methods_data = {
        'Cash': {'amount': 0, 'count': 0},
        'Card': {'amount': 0, 'count': 0},
        'Transfer': {'amount': 0, 'count': 0},
        'POS': {'amount': 0, 'count': 0},
        'Insurance': {'amount': 0, 'count': 0},
        'Other': {'amount': 0, 'count': 0}
    }
    
    # Process payments
    total_transactions = len(today_payments)
    grand_total = 0
    amounts = []
    
    recent_transactions = []
    for payment in today_payments:
        amount_paid = float(payment[3])
        payment_method = payment[4]
        
        # Add to grand total
        grand_total += amount_paid
        amounts.append(amount_paid)
        
        # Add to payment method totals
        if payment_method in payment_methods_data:
            payment_methods_data[payment_method]['amount'] += amount_paid
            payment_methods_data[payment_method]['count'] += 1
        else:
            payment_methods_data['Other']['amount'] += amount_paid
            payment_methods_data['Other']['count'] += 1
        
        # Prepare recent transactions data
        recent_transactions.append({
            'id': payment[0],
            'patient_name': payment[1],
            'service_type': payment[2],
            'amount_paid': amount_paid,
            'payment_method': payment_method,
            'status': payment[5],
            'created_at': payment[6]
        })
    
    # Calculate additional statistics
    average_transaction = grand_total / total_transactions if total_transactions > 0 else 0
    highest_transaction = max(amounts) if amounts else 0
    lowest_transaction = min(amounts) if amounts else 0
    
    # Calculate totals for time periods
    morning_total = 0  # 6AM - 12PM
    afternoon_total = 0  # 12PM - 4PM
    evening_total = 0  # 4PM - 10PM
    
    for payment in today_payments:
        created_at = payment[6]
        if created_at:
            hour = created_at.hour
            amount = float(payment[3])
            
            if 6 <= hour < 12:
                morning_total += amount
            elif 12 <= hour < 16:
                afternoon_total += amount
            elif 16 <= hour < 22:
                evening_total += amount
    
    # Prepare payment methods for template
    payment_methods = []
    for method_name, data in payment_methods_data.items():
        if data['count'] > 0:  # Only include methods with transactions
            percentage = (data['amount'] / grand_total * 100) if grand_total > 0 else 0
            payment_methods.append({
                'name': method_name,
                'amount': data['amount'],
                'count': data['count'],
                'percentage': round(percentage, 1)
            })
    
    # Set daily target (you can make this configurable)
    daily_target = 500000.00  # ₦500,000 daily target
    
    cur.close()
    conn.close()
    
    # Format date for display
    today_date = today.strftime("%A, %B %d, %Y")
    
    return render_template(
        "todays_collection.html",
        today_date=today_date,
        grand_total=grand_total,
        cash_total=payment_methods_data['Cash']['amount'],
        card_total=payment_methods_data['Card']['amount'],
        transfer_total=payment_methods_data['Transfer']['amount'],
        pos_total=payment_methods_data['POS']['amount'],
        insurance_total=payment_methods_data['Insurance']['amount'],
        other_total=payment_methods_data['Other']['amount'],
        payment_methods=payment_methods,
        recent_transactions=recent_transactions,
        total_transactions=total_transactions,
        average_transaction=average_transaction,
        highest_transaction=highest_transaction,
        lowest_transaction=lowest_transaction,
        morning_total=morning_total,
        afternoon_total=afternoon_total,
        evening_total=evening_total,
        daily_target=daily_target
    )


@app.route("/billing/todays-collection/export")
def export_todays_collection():
    if "billing_user_id" not in session:
        return redirect(url_for("billing_login"))
    
    today = date.today()
    
    conn = get_db_connection()
    cur = conn.cursor()
    
    # Get today's payments
    cur.execute("""
        SELECT id, patient_name, service_type, subtotal, discount, tax,
               grand_total, amount_paid, balance, payment_method, 
               status, payment_date, created_at
        FROM payments 
        WHERE DATE(payment_date) = %s
        ORDER BY created_at DESC
    """, (today,))
    
    today_payments = cur.fetchall()
    
    # Create Excel workbook
    wb = Workbook()
    ws = wb.active
    ws.title = f"Today's Collection - {today}"
    
    # Add title
    ws.append([f"Today's Collection Report - {today.strftime('%B %d, %Y')}"])
    ws.append([])
    
    # Add summary section
    ws.append(["SUMMARY"])
    ws.append([])
    
    # Calculate totals by payment method
    payment_methods_data = {
        'Cash': {'amount': 0, 'count': 0},
        'Card': {'amount': 0, 'count': 0},
        'Transfer': {'amount': 0, 'count': 0},
        'POS': {'amount': 0, 'count': 0},
        'Insurance': {'amount': 0, 'count': 0},
        'Other': {'amount': 0, 'count': 0}
    }
    
    grand_total = 0
    total_transactions = len(today_payments)
    
    for payment in today_payments:
        amount_paid = float(payment[7])
        payment_method = payment[9]
        grand_total += amount_paid
        
        if payment_method in payment_methods_data:
            payment_methods_data[payment_method]['amount'] += amount_paid
            payment_methods_data[payment_method]['count'] += 1
        else:
            payment_methods_data['Other']['amount'] += amount_paid
            payment_methods_data['Other']['count'] += 1
    
    # Write summary
    ws.append(["Total Transactions:", total_transactions])
    ws.append(["Grand Total:", grand_total])
    ws.append([])
    ws.append(["Payment Method Breakdown"])
    ws.append(["Method", "Count", "Amount", "Percentage"])
    
    for method_name, data in payment_methods_data.items():
        if data['count'] > 0:
            percentage = (data['amount'] / grand_total * 100) if grand_total > 0 else 0
            ws.append([
                method_name,
                data['count'],
                data['amount'],
                f"{percentage:.1f}%"
            ])
    
    ws.append([])
    ws.append([])
    
    # Add detailed transactions
    ws.append(["DETAILED TRANSACTIONS"])
    ws.append([])
    
    headers = [
        "Receipt No", "Patient Name", "Service Type", 
        "Subtotal", "Discount", "Tax", "Grand Total",
        "Amount Paid", "Balance", "Payment Method",
        "Status", "Payment Date", "Time"
    ]
    ws.append(headers)
    
    # Style headers
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    header_fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True)
    border = Border(left=Side(style='thin'), 
                   right=Side(style='thin'), 
                   top=Side(style='thin'), 
                   bottom=Side(style='thin'))
    
    for cell in ws[ws.max_row]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center")
        cell.border = border
    
    # Add data rows
    for payment in today_payments:
        ws.append([
            payment[0],  # id
            payment[1],  # patient_name
            payment[2],  # service_type
            float(payment[3]),  # subtotal
            float(payment[4]),  # discount
            float(payment[5]),  # tax
            float(payment[6]),  # grand_total
            float(payment[7]),  # amount_paid
            float(payment[8]),  # balance
            payment[9],  # payment_method
            payment[10],  # status
            payment[11].strftime('%Y-%m-%d'),  # payment_date
            payment[12].strftime('%H:%M:%S') if payment[12] else ''  # created_at time
        ])
    
    # Apply borders to data rows
    for row in ws.iter_rows(min_row=ws.max_row - len(today_payments) + 1, max_row=ws.max_row):
        for cell in row:
            cell.border = border
    
    # Auto-adjust column widths
    for column in ws.columns:
        max_length = 0
        column_letter = column[0].column_letter
        for cell in column:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(str(cell.value))
            except:
                pass
        adjusted_width = min(max_length + 2, 30)
        ws.column_dimensions[column_letter].width = adjusted_width
    
    cur.close()
    conn.close()
    
    # Save to BytesIO
    from io import BytesIO
    stream = BytesIO()
    wb.save(stream)
    stream.seek(0)
    
    # Generate filename
    filename = f"todays_collection_{today.strftime('%Y%m%d')}.xlsx"
    
    return send_file(
        stream,
        as_attachment=True,
        download_name=filename,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )


# Custom filter for currency formatting
@app.template_filter('currency')
def currency_filter(amount):
    """Format amount as Nigerian Naira currency."""
    if amount is None:
        return "₦0.00"
    return f"₦{float(amount):,.2f}"

# -------------------- HR DATABASE TABLES --------------------
def create_hr_tables():
    """Create HR-related tables in PostgreSQL."""
    conn = get_db_connection()
    if not conn:
        app.logger.error("Cannot connect to database for HR table creation")
        return

    cursor = conn.cursor()
    
    # Enable UUID extension if needed
    try:
        cursor.execute("CREATE EXTENSION IF NOT EXISTS \"uuid-ossp\";")
    except Exception as e:
        app.logger.warning(f"Could not enable uuid-ossp extension: {e}")
    
    # HR Users Table
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS hr_users (
            id SERIAL PRIMARY KEY,
            username VARCHAR(50) UNIQUE NOT NULL,
            password VARCHAR(255) NOT NULL,
            full_name VARCHAR(100) NOT NULL,
            email VARCHAR(100),
            role VARCHAR(50) DEFAULT 'HR Staff',
            is_active BOOLEAN DEFAULT TRUE,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        );
    """)
    
    # Departments Table
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS departments (
            id SERIAL PRIMARY KEY,
            name VARCHAR(100) NOT NULL,
            code VARCHAR(20) UNIQUE NOT NULL,
            description TEXT,
            head_of_dept VARCHAR(100),
            status VARCHAR(20) DEFAULT 'Active',
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        );
    """)
    
    # Staff Table
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS staff (
            id SERIAL PRIMARY KEY,
            staff_id VARCHAR(50) UNIQUE NOT NULL,
            first_name VARCHAR(100) NOT NULL,
            last_name VARCHAR(100) NOT NULL,
            department_id INTEGER REFERENCES departments(id),
            position VARCHAR(100) NOT NULL,
            employment_type VARCHAR(50),
            email VARCHAR(100),
            phone VARCHAR(20),
            hire_date DATE NOT NULL,
            salary DECIMAL(12, 2),
            status VARCHAR(20) DEFAULT 'Active',
            emergency_contact VARCHAR(100),
            address TEXT,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        );
    """)
    
    # Attendance Table
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS attendance (
            id SERIAL PRIMARY KEY,
            staff_id INTEGER REFERENCES staff(id),
            date DATE NOT NULL,
            check_in TIME,
            check_out TIME,
            status VARCHAR(20),
            remarks TEXT,
            recorded_by INTEGER REFERENCES hr_users(id),
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        );
    """)
    
    # Leaves Table
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS leaves (
            id SERIAL PRIMARY KEY,
            staff_id INTEGER REFERENCES staff(id),
            leave_type VARCHAR(50) NOT NULL,
            start_date DATE NOT NULL,
            end_date DATE NOT NULL,
            days_requested INTEGER NOT NULL,
            reason TEXT,
            status VARCHAR(20) DEFAULT 'Pending',
            approved_by INTEGER REFERENCES hr_users(id),
            approved_at TIMESTAMP,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        );
    """)
    
    # Schedules Table
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS schedules (
            id SERIAL PRIMARY KEY,
            staff_id INTEGER REFERENCES staff(id),
            schedule_date DATE NOT NULL,
            shift_type VARCHAR(50),
            start_time TIME NOT NULL,
            end_time TIME NOT NULL,
            location VARCHAR(100),
            notes TEXT,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        );
    """)
    
    # Payroll Table
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS payroll (
            id SERIAL PRIMARY KEY,
            staff_id INTEGER REFERENCES staff(id),
            pay_period VARCHAR(50),
            basic_salary DECIMAL(12, 2),
            allowances DECIMAL(12, 2),
            deductions DECIMAL(12, 2),
            net_salary DECIMAL(12, 2),
            status VARCHAR(20) DEFAULT 'Pending',
            payment_date DATE,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        );
    """)
    
    # Documents Table
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS documents (
            id SERIAL PRIMARY KEY,
            staff_id INTEGER REFERENCES staff(id),
            document_type VARCHAR(50),
            document_name VARCHAR(255),
            file_path VARCHAR(500),
            uploaded_by INTEGER REFERENCES hr_users(id),
            uploaded_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        );
    """)
    
    # Create indexes for better performance
    indexes = [
        "CREATE INDEX IF NOT EXISTS idx_staff_department ON staff(department_id);",
        "CREATE INDEX IF NOT EXISTS idx_attendance_staff_date ON attendance(staff_id, date);",
        "CREATE INDEX IF NOT EXISTS idx_leaves_staff_status ON leaves(staff_id, status);",
        "CREATE INDEX IF NOT EXISTS idx_schedules_staff_date ON schedules(staff_id, schedule_date);",
        "CREATE INDEX IF NOT EXISTS idx_payroll_staff_period ON payroll(staff_id, pay_period);",
        "CREATE INDEX IF NOT EXISTS idx_staff_status ON staff(status);",
        "CREATE INDEX IF NOT EXISTS idx_attendance_date ON attendance(date);",
        "CREATE INDEX IF NOT EXISTS idx_leaves_status ON leaves(status);"
    ]
    
    for index_query in indexes:
        try:
            cursor.execute(index_query)
        except Exception as e:
            app.logger.warning(f"Could not create index: {e}")
    
    conn.commit()
    cursor.close()
    conn.close()
    
    # Insert default data
    create_default_hr_data()

def create_default_hr_data():
    """Insert default HR data into PostgreSQL tables."""
    conn = get_db_connection()
    if not conn:
        return
    
    cursor = conn.cursor()
    
    # Default hashed password for 'hr@admin123'
    hashed_password = generate_password_hash('hr@admin123')
    
    # Insert default HR users
    try:
        cursor.execute("""
            INSERT INTO hr_users (username, password, full_name, email, role) 
            VALUES 
                (%s, %s, %s, %s, %s),
                (%s, %s, %s, %s, %s)
            ON CONFLICT (username) DO NOTHING;
        """, (
            'hr_admin', hashed_password, 'HR Administrator', 'admin@hospital.com', 'HR Manager',
            'hr_staff', hashed_password, 'HR Staff', 'staff@hospital.com', 'HR Officer'
        ))
    except Exception as e:
        app.logger.error(f"Error inserting HR users: {e}")
    
    # Insert sample departments
    departments = [
        ('Administration', 'ADMIN', 'Hospital Administration and Management', 'Dr. John Smith'),
        ('Medical', 'MED', 'Medical Services Department', 'Dr. Sarah Johnson'),
        ('Nursing', 'NURS', 'Nursing Services', 'Mrs. Grace Williams'),
        ('Pharmacy', 'PHARM', 'Pharmacy Department', 'Mr. Michael Brown'),
        ('Laboratory', 'LAB', 'Laboratory Services', 'Dr. David Miller'),
        ('Radiology', 'RAD', 'Radiology Department', 'Dr. Lisa Davis'),
        ('Finance', 'FIN', 'Finance and Billing Department', 'Mr. Robert Wilson'),
        ('Human Resources', 'HR', 'Human Resources Department', 'Ms. Patricia Taylor'),
        ('Maintenance', 'MAINT', 'Facility Maintenance', 'Mr. Thomas Anderson'),
        ('Security', 'SEC', 'Hospital Security', 'Mr. Richard Clark')
    ]
    
    for dept in departments:
        try:
            cursor.execute("""
                INSERT INTO departments (name, code, description, head_of_dept) 
                VALUES (%s, %s, %s, %s)
                ON CONFLICT (code) DO NOTHING;
            """, dept)
        except Exception as e:
            app.logger.error(f"Error inserting department {dept[0]}: {e}")
    
    # Get admin department ID for sample staff
    cursor.execute("SELECT id FROM departments WHERE code = 'ADMIN' LIMIT 1;")
    admin_dept = cursor.fetchone()
    
    # Insert sample staff if departments exist
    if admin_dept:
        sample_staff = [
            ('EMP001', 'John', 'Doe', admin_dept[0], 'Hospital Administrator', 'Full-Time', 
             'john.doe@hospital.com', '08012345678', '2022-01-15', 850000.00, 'Jane Doe - 08087654321'),
            ('EMP002', 'Sarah', 'Johnson', admin_dept[0], 'Senior Doctor', 'Full-Time', 
             'sarah.j@hospital.com', '08023456789', '2021-03-20', 1200000.00, 'Mark Johnson - 08098765432'),
            ('EMP003', 'Michael', 'Brown', admin_dept[0], 'Chief Pharmacist', 'Full-Time', 
             'michael.b@hospital.com', '08034567890', '2020-06-10', 950000.00, 'Emily Brown - 08076543210'),
            ('EMP004', 'Grace', 'Williams', admin_dept[0], 'Head Nurse', 'Full-Time', 
             'grace.w@hospital.com', '08045678901', '2019-08-05', 750000.00, 'James Williams - 08065432109'),
            ('EMP005', 'David', 'Miller', admin_dept[0], 'Lab Technician', 'Full-Time', 
             'david.m@hospital.com', '08056789012', '2022-11-30', 650000.00, 'Sarah Miller - 08054321098')
        ]
        
        for staff in sample_staff:
            try:
                cursor.execute("""
                    INSERT INTO staff (staff_id, first_name, last_name, department_id, position, 
                                      employment_type, email, phone, hire_date, salary, emergency_contact) 
                    VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
                    ON CONFLICT (staff_id) DO NOTHING;
                """, staff)
            except Exception as e:
                app.logger.error(f"Error inserting staff {staff[0]}: {e}")
    
    # Get HR admin ID for recording
    cursor.execute("SELECT id FROM hr_users WHERE username = 'hr_admin' LIMIT 1;")
    hr_admin = cursor.fetchone()
    
    # Get staff IDs for sample data
    cursor.execute("SELECT id, staff_id FROM staff ORDER BY id LIMIT 5;")
    staff_members = cursor.fetchall()
    
    if hr_admin and staff_members:
        hr_admin_id = hr_admin[0]
        today = date.today()
        
        # Insert sample attendance for today
        for i, staff in enumerate(staff_members[:3]):  # First 3 staff
            check_in = '08:00:00' if i != 1 else '08:30:00'  # Make second staff late
            status = 'Present' if i != 1 else 'Late'
            
            try:
                cursor.execute("""
                    INSERT INTO attendance (staff_id, date, check_in, check_out, status, recorded_by)
                    VALUES (%s, %s, %s, %s, %s, %s)
                    ON CONFLICT DO NOTHING;
                """, (staff[0], today, check_in, '16:00:00', status, hr_admin_id))
            except Exception as e:
                app.logger.error(f"Error inserting attendance for {staff[1]}: {e}")
        
        # Insert sample leave requests
        try:
            cursor.execute("""
                INSERT INTO leaves (staff_id, leave_type, start_date, end_date, days_requested, reason, status)
                SELECT 
                    id,
                    'Annual Leave',
                    %s + INTERVAL '5 days',
                    %s + INTERVAL '12 days',
                    8,
                    'Family vacation',
                    'Pending'
                FROM staff WHERE staff_id = 'EMP004'
                UNION ALL
                SELECT 
                    id,
                    'Sick Leave',
                    %s - INTERVAL '2 days',
                    %s + INTERVAL '2 days',
                    5,
                    'Medical treatment',
                    'Approved'
                FROM staff WHERE staff_id = 'EMP005'
                ON CONFLICT DO NOTHING;
            """, (today, today, today, today))
        except Exception as e:
            app.logger.error(f"Error inserting leaves: {e}")
    
    conn.commit()
    cursor.close()
    conn.close()

# -------------------- ROUTES: HR MODULE --------------------
@app.route("/hr/login", methods=["GET", "POST"])
def hr_login():
    if request.method == "POST":
        username = request.form["username"]
        password = request.form["password"]

        conn = get_db_connection()
        if not conn:
            flash("Database connection error", "danger")
            return render_template("hr_login.html", 
                                 hospital_name="Memorial Hospital Ovuru, Nsukka, Enugu State")

        cur = conn.cursor()

        try:
            cur.execute("""
                SELECT id, username, password, full_name, role
                FROM hr_users
                WHERE username = %s AND is_active = TRUE
            """, (username,))

            user = cur.fetchone()
            cur.close()
            conn.close()

            if user:
                stored_hash = user[2]
                
                # SIMPLE CHECK for our known bcrypt hash
                known_hash = "$2b$12$LQv3c1yqBWVHxkd0LsZcdeJN8L7Fmm8Zz3qG9XwFk8kC1YdV6n4Oq"
                
                if stored_hash == known_hash and password == "hr@admin123":
                    session["hr_user_id"] = user[0]
                    session["hr_username"] = user[1]
                    session["hr_full_name"] = user[3]
                    session["hr_role"] = user[4]
                    flash(f"Welcome, {user[3]}!", "success")
                    return redirect(url_for("hr_dashboard"))
                else:
                    flash("Invalid password. Use: hr@admin123", "danger")
            else:
                flash("Invalid username. Use: hr_admin or hr_staff", "danger")

        except Exception as e:
            app.logger.error(f"HR login error: {e}")
            flash("Login error. Please try again.", "danger")

    return render_template("hr_login.html", 
                         hospital_name="Memorial Hospital Ovuru, Nsukka, Enugu State")    


# @app.route("/hr/scheduling")
# def scheduling():
#     if "hr_user_id" not in session:
#         return redirect(url_for("hr_login"))
#     return render_template("module_placeholder.html", 
#                          module_name="Scheduling",
#                          description="Create and manage work schedules and shifts")

@app.route("/hr/leave-management")
def leave_management():
    if "hr_user_id" not in session:
        return redirect(url_for("hr_login"))
    return render_template("module_placeholder.html", 
                         module_name="Leave Management",
                         description="Approve and track staff leave requests")

@app.route("/hr/attendance-record")
def attendance_record():
    if "hr_user_id" not in session:
        return redirect(url_for("hr_login"))
    return render_template("module_placeholder.html", 
                         module_name="Attendance Record",
                         description="Track staff attendance and punctuality")

@app.route("/hr/departments")
def departments():
    if "hr_user_id" not in session:
        return redirect(url_for("hr_login"))
    
    conn = get_db_connection()
    cur = conn.cursor()
    
    try:
        # Get all departments with staff count
        cur.execute("""
            SELECT d.*, 
                   COUNT(s.id) as staff_count
            FROM departments d
            LEFT JOIN staff s ON d.id = s.department_id AND s.status = 'Active'
            GROUP BY d.id
            ORDER BY d.name
        """)
        dept_list = cur.fetchall()
        
    except Exception as e:
        app.logger.error(f"Error fetching departments: {e}")
        dept_list = []
    
    finally:
        cur.close()
        conn.close()
    
    return render_template("hr_departments.html", 
                         module_name="Departments",
                         description="Manage hospital departments and reporting structure",
                         departments=dept_list)

@app.route("/hr/reports")
def hr_reports():
    if "hr_user_id" not in session:
        return redirect(url_for("hr_login"))
    return render_template("module_placeholder.html", 
                         module_name="HR Reports",
                         description="Generate HR analytics and reports")

# # -------------------- QUICK ACTION PLACEHOLDERS --------------------
# @app.route("/hr/add-staff")
# def add_staff():
#     if "hr_user_id" not in session:
#         return redirect(url_for("hr_login"))
#     flash("Feature coming soon: Add New Staff", "info")
#     return redirect(url_for("hr_dashboard"))

@app.route("/hr/approve-leave")
def approve_leave():
    if "hr_user_id" not in session:
        return redirect(url_for("hr_login"))
    flash("Feature coming soon: Approve Leave Requests", "info")
    return redirect(url_for("hr_dashboard"))

@app.route("/hr/mark-attendance")
def mark_attendance():
    if "hr_user_id" not in session:
        return redirect(url_for("hr_login"))
    flash("Feature coming soon: Mark Attendance", "info")
    return redirect(url_for("hr_dashboard"))

@app.route("/hr/generate-payroll")
def generate_payroll():
    if "hr_user_id" not in session:
        return redirect(url_for("hr_login"))
    flash("Feature coming soon: Generate Payroll", "info")
    return redirect(url_for("hr_dashboard"))

@app.route("/hr/upload-documents")
def upload_documents():
    if "hr_user_id" not in session:
        return redirect(url_for("hr_login"))
    flash("Feature coming soon: Upload Documents", "info")
    return redirect(url_for("hr_dashboard"))

@app.route("/hr/send-notifications")
def send_notifications():
    if "hr_user_id" not in session:
        return redirect(url_for("hr_login"))
    flash("Feature coming soon: Send Notifications", "info")
    return redirect(url_for("hr_dashboard"))

@app.route("/hr/dashboard")
def hr_dashboard():
    if "hr_user_id" not in session:
        return redirect(url_for("hr_login"))
    
    # Get HR statistics
    conn = get_db_connection()
    if not conn:
        flash("Database connection error", "danger")
        return redirect(url_for("hr_login"))
    
    cur = conn.cursor()
    
    try:
        # Total staff
        cur.execute("SELECT COUNT(*) FROM staff WHERE status = 'Active'")
        total_staff = cur.fetchone()[0] or 0
        
        # Staff active today (attendance)
        today = date.today()
        cur.execute("""
            SELECT COUNT(DISTINCT staff_id) 
            FROM attendance 
            WHERE date = %s AND status IN ('Present', 'Late')
        """, (today,))
        active_staff = cur.fetchone()[0] or 0
        
        # Staff on leave today
        cur.execute("""
            SELECT COUNT(*) 
            FROM leaves 
            WHERE %s BETWEEN start_date AND end_date 
            AND status = 'Approved'
        """, (today,))
        on_leave = cur.fetchone()[0] or 0
        
        # Departments count
        cur.execute("SELECT COUNT(*) FROM departments WHERE status = 'Active'")
        departments_count = cur.fetchone()[0] or 0
        
        # Pending leave requests
        cur.execute("SELECT COUNT(*) FROM leaves WHERE status = 'Pending'")
        pending_leave = cur.fetchone()[0] or 0
        
        # Upcoming shifts (next 7 days)
        next_week = today + timedelta(days=7)
        cur.execute("""
            SELECT COUNT(*) 
            FROM schedules 
            WHERE schedule_date BETWEEN %s AND %s
        """, (today, next_week))
        upcoming_shifts = cur.fetchone()[0] or 0
        
        # Pending updates (staff with missing info)
        cur.execute("""
            SELECT COUNT(*) 
            FROM staff 
            WHERE emergency_contact IS NULL OR address IS NULL
        """)
        pending_updates = cur.fetchone()[0] or 0
        
        # Late arrivals today
        cur.execute("""
            SELECT COUNT(*) 
            FROM attendance 
            WHERE date = %s AND status = 'Late'
        """, (today,))
        late_arrivals = cur.fetchone()[0] or 0
        
    except Exception as e:
        app.logger.error(f"Error fetching HR stats: {e}")
        # Set default values on error
        total_staff = active_staff = on_leave = departments_count = 0
        pending_leave = upcoming_shifts = pending_updates = late_arrivals = 0
    
    finally:
        cur.close()
        conn.close()
    
    return render_template(
        "hr_dashboard.html",
        hospital_name="Memorial Hospital Ovuru, Nsukka, Enugu State",
        total_staff=total_staff,
        active_staff=active_staff,
        on_leave=on_leave,
        departments_count=departments_count,
        pending_leave=pending_leave,
        upcoming_shifts=upcoming_shifts,
        pending_updates=pending_updates,
        late_arrivals=late_arrivals,
        current_year=date.today().year
    )
    
@app.route("/hr/logout")
def hr_logout():
    session.pop("hr_user_id", None)
    session.pop("hr_username", None)
    session.pop("hr_full_name", None)
    session.pop("hr_role", None)
    flash("Logged out successfully", "success")
    return redirect(url_for("hr_login"))

@app.route("/hr/staff-management")
def staff_management():
    if "hr_user_id" not in session:
        return redirect(url_for("hr_login"))
    
    conn = get_db_connection()
    if not conn:
        flash("Database connection error", "danger")
        return redirect(url_for("hr_login"))
    
    cur = conn.cursor()
    
    try:
        # Get staff list with department names
        cur.execute("""
            SELECT s.id, s.staff_id, s.first_name, s.last_name, 
                   s.position, s.employment_type, s.email, s.phone,
                   s.hire_date, s.salary, s.status, s.emergency_contact,
                   s.address, d.name as department_name
            FROM staff s
            LEFT JOIN departments d ON s.department_id = d.id
            ORDER BY s.id DESC
            LIMIT 100
        """)
        staff_list = cur.fetchall()
        
        # Get statistics
        cur.execute("SELECT COUNT(*) FROM staff WHERE status = 'Active'")
        total_staff = cur.fetchone()[0] or 0
        
        cur.execute("SELECT COUNT(*) FROM staff WHERE status = 'Active' AND employment_type = 'Full-Time'")
        active_staff = cur.fetchone()[0] or 0
        
        cur.execute("SELECT COUNT(*) FROM staff WHERE employment_type = 'Contract'")
        on_contract = cur.fetchone()[0] or 0
        
        cur.execute("SELECT COUNT(DISTINCT department_id) FROM staff")
        departments_count = cur.fetchone()[0] or 0
        
    except Exception as e:
        app.logger.error(f"Error fetching staff data: {e}")
        staff_list = []
        total_staff = active_staff = on_contract = departments_count = 0
    
    finally:
        cur.close()
        conn.close()
    
    return render_template("staff_management.html", 
                         module_name="Staff Management",
                         description="Manage staff profiles, positions, and employment details",
                         staff_list=staff_list,
                         total_staff=total_staff,
                         active_staff=active_staff,
                         on_contract=on_contract,
                         departments_count=departments_count,
                         hospital_name="Memorial Hospital Ovuru, Nsukka, Enugu State",
                         current_year=date.today().year)    
# View Staff Details
@app.route("/hr/staff/<int:staff_id>")
def view_staff(staff_id):
    if "hr_user_id" not in session:
        return redirect(url_for("hr_login"))
    
    conn = get_db_connection()
    cur = conn.cursor()
    
    try:
        # Get staff details with department
        cur.execute("""
            SELECT s.*, d.name as department_name, d.code as department_code
            FROM staff s
            LEFT JOIN departments d ON s.department_id = d.id
            WHERE s.id = %s
        """, (staff_id,))
        
        staff = cur.fetchone()
        
        if not staff:
            flash("Staff member not found", "danger")
            return redirect(url_for("staff_management"))
        
        # Convert to dictionary for easier template access
        staff_dict = {
            'id': staff[0],
            'staff_id': staff[1],
            'first_name': staff[2],
            'last_name': staff[3],
            'department_id': staff[4],
            'position': staff[5],
            'employment_type': staff[6],
            'email': staff[7],
            'phone': staff[8],
            'hire_date': staff[9],
            'salary': float(staff[10]) if staff[10] else 0,
            'status': staff[11],
            'emergency_contact': staff[12],
            'address': staff[13],
            'department_name': staff[14],
            'department_code': staff[15]
        }
        
        # Calculate employment duration
        from datetime import date
        today = date.today()
        
        # Ensure hire_date is a date object
        hire_date = staff[9]
        if isinstance(hire_date, str):
            from datetime import datetime
            hire_date = datetime.strptime(hire_date, '%Y-%m-%d').date()
        
        years = today.year - hire_date.year
        months = today.month - hire_date.month
        
        if months < 0:
            years -= 1
            months += 12
        
        employment_duration = f"{years} year(s), {months} month(s)"
        
    except Exception as e:
        app.logger.error(f"Error fetching staff details: {e}")
        flash("Error loading staff details", "danger")
        return redirect(url_for("staff_management"))
    
    finally:
        cur.close()
        conn.close()
    
    return render_template("view_staff.html", 
                         staff=staff_dict,
                         today=today,
                         employment_duration=employment_duration,
                         hospital_name="Memorial Hospital Ovuru, Nsukka, Enugu State")
# Add New Staff
@app.route("/hr/staff/add", methods=["GET", "POST"])
def add_staff():
    if "hr_user_id" not in session:
        return redirect(url_for("hr_login"))
    
    if request.method == "POST":
        # Get form data
        staff_id = request.form.get('staff_id')
        first_name = request.form.get('first_name')
        last_name = request.form.get('last_name')
        department_id = request.form.get('department_id')
        position = request.form.get('position')
        employment_type = request.form.get('employment_type')
        email = request.form.get('email')
        phone = request.form.get('phone')
        hire_date = request.form.get('hire_date')
        salary = request.form.get('salary')
        emergency_contact = request.form.get('emergency_contact')
        address = request.form.get('address')
        
        # Validate required fields
        if not all([staff_id, first_name, last_name, department_id, position, hire_date]):
            flash("Please fill in all required fields", "danger")
            return redirect(url_for("add_staff"))
        
        conn = get_db_connection()
        if not conn:
            flash("Database connection error", "danger")
            return redirect(url_for("add_staff"))
        
        cur = conn.cursor()
        
        try:
            # Check if staff ID already exists
            cur.execute("SELECT id FROM staff WHERE staff_id = %s", (staff_id,))
            if cur.fetchone():
                flash(f"Staff ID '{staff_id}' already exists. Please use a different ID.", "danger")
                return redirect(url_for("add_staff"))
            
            # Convert salary to decimal or set to 0
            try:
                salary_decimal = float(salary) if salary else 0.00
            except ValueError:
                salary_decimal = 0.00
            
            # Insert new staff
            cur.execute("""
                INSERT INTO staff (
                    staff_id, first_name, last_name, department_id, 
                    position, employment_type, email, phone, 
                    hire_date, salary, emergency_contact, address, status
                ) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, 'Active')
            """, (
                staff_id, first_name, last_name, department_id,
                position, employment_type, email, phone,
                hire_date, salary_decimal, emergency_contact, address
            ))
            
            conn.commit()
            flash(f"Staff member {first_name} {last_name} (ID: {staff_id}) added successfully!", "success")
            
            # Redirect to staff management or view the new staff
            cur.execute("SELECT id FROM staff WHERE staff_id = %s", (staff_id,))
            new_staff_id = cur.fetchone()[0]
            return redirect(url_for("view_staff", staff_id=new_staff_id))
            
        except Exception as e:
            conn.rollback()
            app.logger.error(f"Error adding staff: {e}")
            flash(f"Error adding staff: {str(e)}", "danger")
            return redirect(url_for("add_staff"))
            
        finally:
            cur.close()
            conn.close()
    
    # GET request - show form
    conn = get_db_connection()
    if not conn:
        flash("Database connection error", "danger")
        return redirect(url_for("staff_management"))
    
    cur = conn.cursor()
    
    try:
        # Get departments for dropdown
        cur.execute("SELECT id, name, code FROM departments WHERE status = 'Active' ORDER BY name")
        departments = cur.fetchall()
        
        # Get next staff ID suggestion
        cur.execute("""
            SELECT MAX(staff_id) FROM staff 
            WHERE staff_id ~ '^EMP[0-9]+$'
        """)
        last_staff_id = cur.fetchone()[0]
        
        if last_staff_id:
            # Extract number and increment
            import re
            match = re.search(r'EMP(\d+)', last_staff_id)
            if match:
                next_num = int(match.group(1)) + 1
                suggested_id = f"EMP{next_num:03d}"
            else:
                suggested_id = "EMP001"
        else:
            suggested_id = "EMP001"
            
        # Get current date for hire date default
        today = date.today().strftime("%Y-%m-%d")
        
    except Exception as e:
        app.logger.error(f"Error loading form data: {e}")
        departments = []
        suggested_id = "EMP001"
        today = date.today().strftime("%Y-%m-%d")
        
    finally:
        cur.close()
        conn.close()
    
    return render_template("add_staff.html",
                         departments=departments,
                         suggested_id=suggested_id,
                         today=today,
                         hospital_name="Memorial Hospital Ovuru, Nsukka, Enugu State")

# Edit Staff
@app.route("/hr/staff/edit/<int:staff_id>", methods=["GET", "POST"])
def edit_staff(staff_id):
    if "hr_user_id" not in session:
        return redirect(url_for("hr_login"))
    
    conn = get_db_connection()
    cur = conn.cursor()
    
    if request.method == "POST":
        # Update staff
        first_name = request.form.get('first_name')
        last_name = request.form.get('last_name')
        department_id = request.form.get('department_id')
        position = request.form.get('position')
        employment_type = request.form.get('employment_type')
        email = request.form.get('email')
        phone = request.form.get('phone')
        salary = request.form.get('salary')
        status = request.form.get('status')
        emergency_contact = request.form.get('emergency_contact')
        address = request.form.get('address')
        
        try:
            cur.execute("""
                UPDATE staff SET
                    first_name = %s,
                    last_name = %s,
                    department_id = %s,
                    position = %s,
                    employment_type = %s,
                    email = %s,
                    phone = %s,
                    salary = %s,
                    status = %s,
                    emergency_contact = %s,
                    address = %s,
                    updated_at = CURRENT_TIMESTAMP
                WHERE id = %s
            """, (
                first_name, last_name, department_id,
                position, employment_type, email, phone,
                salary, status, emergency_contact, address,
                staff_id
            ))
            
            conn.commit()
            flash("Staff details updated successfully!", "success")
            return redirect(url_for("view_staff", staff_id=staff_id))
            
        except Exception as e:
            conn.rollback()
            app.logger.error(f"Error updating staff: {e}")
            flash("Error updating staff details", "danger")
    
    # GET request - load staff data
    try:
        cur.execute("SELECT * FROM staff WHERE id = %s", (staff_id,))
        staff = cur.fetchone()
        
        if not staff:
            flash("Staff member not found", "danger")
            return redirect(url_for("staff_management"))
        
        # Get departments
        cur.execute("SELECT id, name FROM departments WHERE status = 'Active' ORDER BY name")
        departments = cur.fetchall()
        
    except Exception as e:
        app.logger.error(f"Error loading staff for edit: {e}")
        flash("Error loading staff details", "danger")
        return redirect(url_for("staff_management"))
    
    finally:
        cur.close()
        conn.close()
    
    return render_template("edit_staff.html",
                         staff=staff,
                         departments=departments,
                         hospital_name="Memorial Hospital Ovuru, Nsukka, Enugu State")


# ==================== ROUTES: SCHEDULING MODULE ====================

@app.route("/hr/scheduling")
def scheduling():
    if "hr_user_id" not in session:
        return redirect(url_for("hr_login"))
    
    conn = get_db_connection()
    cur = conn.cursor()
    
    try:
        # Get current month schedules
        current_month = date.today().replace(day=1)
        if current_month.month == 12:
            next_month = current_month.replace(year=current_month.year + 1, month=1)
        else:
            next_month = current_month.replace(month=current_month.month + 1)
        
        cur.execute("""
            SELECT s.*, st.first_name, st.last_name, st.position, d.name as department_name
            FROM schedules s
            JOIN staff st ON s.staff_id = st.id
            LEFT JOIN departments d ON st.department_id = d.id
            WHERE s.schedule_date >= %s AND s.schedule_date < %s
            ORDER BY s.schedule_date, s.start_time
        """, (current_month, next_month))
        
        schedules = cur.fetchall()
        
        # Get statistics
        cur.execute("SELECT COUNT(*) FROM schedules WHERE schedule_date >= CURRENT_DATE")
        upcoming_shifts = cur.fetchone()[0] or 0
        
        cur.execute("""
            SELECT COUNT(DISTINCT staff_id) 
            FROM schedules 
            WHERE schedule_date >= CURRENT_DATE
        """)
        staff_scheduled = cur.fetchone()[0] or 0
        
        # Get departments for filter
        cur.execute("SELECT id, name FROM departments WHERE status = 'Active' ORDER BY name")
        departments = cur.fetchall()
        
        # Get staff for filter
        cur.execute("""
            SELECT id, first_name, last_name, position 
            FROM staff 
            WHERE status = 'Active' 
            ORDER BY first_name, last_name
        """)
        staff_list = cur.fetchall()
        
    except Exception as e:
        app.logger.error(f"Error fetching scheduling data: {e}")
        schedules = []
        upcoming_shifts = 0
        staff_scheduled = 0
        departments = []
        staff_list = []
    
    finally:
        cur.close()
        conn.close()
    
    return render_template("scheduling_dashboard.html",
                         schedules=schedules,
                         upcoming_shifts=upcoming_shifts,
                         staff_scheduled=staff_scheduled,
                         departments=departments,
                         staff_list=staff_list,
                         current_month=current_month.strftime("%B %Y"),
                         hospital_name="Memorial Hospital Ovuru, Nsukka, Enugu State")
    
    
@app.route("/hr/scheduling/create", methods=["GET", "POST"])
def create_schedule():
    if "hr_user_id" not in session:
        return redirect(url_for("hr_login"))
    
    if request.method == "POST":
        staff_id = request.form.get("staff_id")
        schedule_date = request.form.get("schedule_date")
        shift_type = request.form.get("shift_type")
        start_time = request.form.get("start_time")
        end_time = request.form.get("end_time")
        location = request.form.get("location")
        notes = request.form.get("notes")
        
        if not all([staff_id, schedule_date, start_time, end_time]):
            flash("Please fill in all required fields", "danger")
            return redirect(url_for("create_schedule"))
        
        conn = get_db_connection()
        cur = conn.cursor()
        
        try:
            # Check for existing schedule for same staff on same date
            cur.execute("""
                SELECT id FROM schedules 
                WHERE staff_id = %s AND schedule_date = %s
            """, (staff_id, schedule_date))
            
            if cur.fetchone():
                flash("This staff already has a schedule for this date", "warning")
                return redirect(url_for("create_schedule"))
            
            # Insert new schedule
            cur.execute("""
                INSERT INTO schedules (
                    staff_id, schedule_date, shift_type, 
                    start_time, end_time, location, notes
                ) VALUES (%s, %s, %s, %s, %s, %s, %s)
            """, (staff_id, schedule_date, shift_type, start_time, end_time, location, notes))
            
            conn.commit()
            flash("Schedule created successfully!", "success")
            return redirect(url_for("scheduling"))
            
        except Exception as e:
            conn.rollback()
            app.logger.error(f"Error creating schedule: {e}")
            flash(f"Error creating schedule: {str(e)}", "danger")
            return redirect(url_for("create_schedule"))
            
        finally:
            cur.close()
            conn.close()
    
    # GET request - show form
    conn = get_db_connection()
    cur = conn.cursor()
    
    try:
        # Get active staff
        cur.execute("""
            SELECT id, first_name, last_name, position, 
                   (SELECT name FROM departments WHERE id = staff.department_id) as department
            FROM staff 
            WHERE status = 'Active' 
            ORDER BY first_name, last_name
        """)
        staff_list = cur.fetchall()
        
        # Get default tomorrow's date
        tomorrow = (datetime.now() + timedelta(days=1)).strftime("%Y-%m-%d")
        
    except Exception as e:
        app.logger.error(f"Error loading schedule form data: {e}")
        staff_list = []
        tomorrow = (datetime.now() + timedelta(days=1)).strftime("%Y-%m-%d")
        
    finally:
        cur.close()
        conn.close()
    
    return render_template("create_schedule.html",
                         staff_list=staff_list,
                         tomorrow=tomorrow,
                         hospital_name="Memorial Hospital Ovuru, Nsukka, Enugu State")

@app.route("/hr/scheduling/roster")
def view_roster():
    if "hr_user_id" not in session:
        return redirect(url_for("hr_login"))
    
    # Get filter parameters
    department_id = request.args.get("department_id", "")
    staff_id = request.args.get("staff_id", "")
    start_date = request.args.get("start_date", "")
    end_date = request.args.get("end_date", "")
    
    # Default to current week if no dates specified
    if not start_date:
        today = date.today()
        start_date = (today - timedelta(days=today.weekday())).strftime("%Y-%m-%d")
    
    if not end_date:
        start_date_obj = datetime.strptime(start_date, "%Y-%m-%d").date()
        end_date = (start_date_obj + timedelta(days=6)).strftime("%Y-%m-%d")
    
    conn = get_db_connection()
    cur = conn.cursor()
    
    try:
        # Build query with filters
        query = """
            SELECT 
                s.id as schedule_id,
                s.schedule_date,
                s.shift_type,
                s.start_time,
                s.end_time,
                s.location,
                s.notes,
                st.id as staff_id,
                st.first_name,
                st.last_name,
                st.position,
                d.name as department_name,
                d.id as department_id
            FROM schedules s
            JOIN staff st ON s.staff_id = st.id
            LEFT JOIN departments d ON st.department_id = d.id
            WHERE s.schedule_date BETWEEN %s AND %s
        """
        params = [start_date, end_date]
        
        if department_id:
            query += " AND st.department_id = %s"
            params.append(department_id)
        
        if staff_id:
            query += " AND s.staff_id = %s"
            params.append(staff_id)
        
        query += " ORDER BY s.schedule_date, d.name, st.first_name, s.start_time"
        
        cur.execute(query, params)
        schedules = cur.fetchall()
        
        # Get departments for filter dropdown
        cur.execute("SELECT id, name FROM departments WHERE status = 'Active' ORDER BY name")
        departments = cur.fetchall()
        
        # Get staff for filter dropdown
        cur.execute("""
            SELECT id, first_name, last_name 
            FROM staff 
            WHERE status = 'Active' 
            ORDER BY first_name, last_name
        """)
        staff_list = cur.fetchall()
        
        # Group schedules by date for calendar view - FIXED: Use tuple indices
        schedule_dict = {}
        for schedule in schedules:
            schedule_date = schedule[1].strftime("%Y-%m-%d")
            if schedule_date not in schedule_dict:
                schedule_dict[schedule_date] = []
            
            schedule_dict[schedule_date].append({
                'id': schedule[0],
                'date': schedule[1],
                'shift_type': schedule[2],
                'start_time': schedule[3],
                'end_time': schedule[4],
                'location': schedule[5],
                'notes': schedule[6],
                'staff_id': schedule[7],
                'first_name': schedule[8],
                'last_name': schedule[9],
                'position': schedule[10],
                'department': schedule[11]
            })
        
        # Calculate statistics
        total_shifts = len(schedules)
        unique_staff = len(set([s[7] for s in schedules]))
        unique_departments = len(set([s[11] for s in schedules if s[11]]))
        
    except Exception as e:
        app.logger.error(f"Error fetching roster: {e}")
        schedules = []
        departments = []
        staff_list = []
        schedule_dict = {}
        total_shifts = 0
        unique_staff = 0
        unique_departments = 0
    
    finally:
        cur.close()
        conn.close()
    
    return render_template("view_roster.html",
                         schedules=schedules,
                         schedule_dict=schedule_dict,
                         departments=departments,
                         staff_list=staff_list,
                         start_date=start_date,
                         end_date=end_date,
                         selected_department=department_id,
                         selected_staff=staff_id,
                         total_shifts=total_shifts,
                         unique_staff=unique_staff,
                         unique_departments=unique_departments,
                         hospital_name="Memorial Hospital Ovuru, Nsukka, Enugu State")
@app.route("/hr/scheduling/shift-swap", methods=["GET", "POST"])
def shift_swap():
    if "hr_user_id" not in session:
        return redirect(url_for("hr_login"))
    
    conn = get_db_connection()
    cur = conn.cursor()
    
    if request.method == "POST":
        action = request.form.get("action")
        
        if action == "request":
            # Request shift swap
            from_staff_id = request.form.get("from_staff_id")
            to_staff_id = request.form.get("to_staff_id")
            schedule_id = request.form.get("schedule_id")
            reason = request.form.get("reason")
            
            try:
                # Get schedule details
                cur.execute("""
                    SELECT schedule_date, start_time, end_time, shift_type 
                    FROM schedules 
                    WHERE id = %s
                """, (schedule_id,))
                schedule = cur.fetchone()
                
                if not schedule:
                    flash("Schedule not found", "danger")
                    return redirect(url_for("shift_swap"))
                
                # Check if staff is available on that date
                cur.execute("""
                    SELECT id FROM schedules 
                    WHERE staff_id = %s AND schedule_date = %s
                """, (to_staff_id, schedule[0]))
                
                if cur.fetchone():
                    flash("Selected staff already has a schedule on this date", "warning")
                    return redirect(url_for("shift_swap"))
                
                # Create shift swap request (you'll need to create this table)
                cur.execute("""
                    INSERT INTO shift_swap_requests (
                        schedule_id, from_staff_id, to_staff_id, 
                        reason, status, requested_by, requested_at
                    ) VALUES (%s, %s, %s, %s, 'Pending', %s, NOW())
                """, (schedule_id, from_staff_id, to_staff_id, reason, session["hr_user_id"]))
                
                conn.commit()
                flash("Shift swap request submitted successfully!", "success")
                
            except Exception as e:
                conn.rollback()
                app.logger.error(f"Error creating shift swap request: {e}")
                flash(f"Error: {str(e)}", "danger")
        
        elif action == "approve":
            # Approve shift swap
            swap_id = request.form.get("swap_id")
            
            try:
                # Get swap request details
                cur.execute("""
                    SELECT schedule_id, from_staff_id, to_staff_id 
                    FROM shift_swap_requests 
                    WHERE id = %s AND status = 'Pending'
                """, (swap_id,))
                
                swap_request = cur.fetchone()
                if not swap_request:
                    flash("Swap request not found or already processed", "warning")
                    return redirect(url_for("shift_swap"))
                
                # Update schedule with new staff
                cur.execute("""
                    UPDATE schedules 
                    SET staff_id = %s 
                    WHERE id = %s
                """, (swap_request[2], swap_request[0]))
                
                # Update swap request status
                cur.execute("""
                    UPDATE shift_swap_requests 
                    SET status = 'Approved', 
                        approved_by = %s, 
                        approved_at = NOW() 
                    WHERE id = %s
                """, (session["hr_user_id"], swap_id))
                
                conn.commit()
                flash("Shift swap approved successfully!", "success")
                
            except Exception as e:
                conn.rollback()
                app.logger.error(f"Error approving shift swap: {e}")
                flash(f"Error: {str(e)}", "danger")
        
        elif action == "reject":
            # Reject shift swap
            swap_id = request.form.get("swap_id")
            rejection_reason = request.form.get("rejection_reason")
            
            try:
                cur.execute("""
                    UPDATE shift_swap_requests 
                    SET status = 'Rejected', 
                        rejection_reason = %s,
                        reviewed_by = %s, 
                        reviewed_at = NOW() 
                    WHERE id = %s
                """, (rejection_reason, session["hr_user_id"], swap_id))
                
                conn.commit()
                flash("Shift swap request rejected", "info")
                
            except Exception as e:
                conn.rollback()
                app.logger.error(f"Error rejecting shift swap: {e}")
                flash(f"Error: {str(e)}", "danger")
    
    # GET request - show shift swap page
    try:
        # Get pending shift swap requests
        cur.execute("""
            SELECT ssr.*, 
                   s1.first_name as from_first_name, s1.last_name as from_last_name,
                   s2.first_name as to_first_name, s2.last_name as to_last_name,
                   sch.schedule_date, sch.start_time, sch.end_time
            FROM shift_swap_requests ssr
            JOIN schedules sch ON ssr.schedule_id = sch.id
            JOIN staff s1 ON ssr.from_staff_id = s1.id
            JOIN staff s2 ON ssr.to_staff_id = s2.id
            WHERE ssr.status = 'Pending'
            ORDER BY ssr.requested_at DESC
        """)
        pending_swaps = cur.fetchall()
        
        # Get upcoming schedules for current staff
        cur.execute("""
            SELECT sch.id, sch.schedule_date, sch.start_time, sch.end_time,
                   st.first_name, st.last_name, st.position
            FROM schedules sch
            JOIN staff st ON sch.staff_id = st.id
            WHERE sch.schedule_date >= CURRENT_DATE
            ORDER BY sch.schedule_date, sch.start_time
            LIMIT 50
        """)
        upcoming_schedules = cur.fetchall()
        
        # Get available staff for swaps
        cur.execute("""
            SELECT id, first_name, last_name, position 
            FROM staff 
            WHERE status = 'Active' 
            ORDER BY first_name, last_name
        """)
        staff_list = cur.fetchall()
        
    except Exception as e:
        app.logger.error(f"Error loading shift swap data: {e}")
        pending_swaps = []
        upcoming_schedules = []
        staff_list = []
    
    finally:
        cur.close()
        conn.close()
    
    return render_template("shift_swap.html",
                         pending_swaps=pending_swaps,
                         upcoming_schedules=upcoming_schedules,
                         staff_list=staff_list,
                         hospital_name="Memorial Hospital Ovuru, Nsukka, Enugu State")

@app.route("/hr/scheduling/reports")
def schedule_reports():
    if "hr_user_id" not in session:
        return redirect(url_for("hr_login"))
    
    # Get report parameters
    report_type = request.args.get("report_type", "monthly")
    month = request.args.get("month", date.today().month)
    year = request.args.get("year", date.today().year)
    department_id = request.args.get("department_id", "")
    
    conn = get_db_connection()
    cur = conn.cursor()
    
    try:
        # Build report query based on report type
        if report_type == "monthly":
            start_date = date(int(year), int(month), 1)
            if int(month) == 12:
                end_date = date(int(year) + 1, 1, 1) - timedelta(days=1)
            else:
                end_date = date(int(year), int(month) + 1, 1) - timedelta(days=1)
            
            query = """
                SELECT 
                    d.name as department,
                    st.first_name || ' ' || st.last_name as staff_name,
                    COUNT(*) as total_shifts,
                    SUM(EXTRACT(EPOCH FROM (end_time - start_time))/3600) as total_hours,
                    COUNT(DISTINCT sch.schedule_date) as days_scheduled
                FROM schedules sch
                JOIN staff st ON sch.staff_id = st.id
                LEFT JOIN departments d ON st.department_id = d.id
                WHERE sch.schedule_date BETWEEN %s AND %s
            """
            params = [start_date, end_date]
            
            if department_id:
                query += " AND st.department_id = %s"
                params.append(department_id)
            
            query += """
                GROUP BY d.name, st.first_name, st.last_name
                ORDER BY d.name, staff_name
            """
            
        elif report_type == "weekly":
            # Get current week
            today = date.today()
            start_date = today - timedelta(days=today.weekday())
            end_date = start_date + timedelta(days=6)
            
            query = """
                SELECT 
                    sch.schedule_date,
                    d.name as department,
                    st.first_name || ' ' || st.last_name as staff_name,
                    sch.shift_type,
                    sch.start_time,
                    sch.end_time,
                    EXTRACT(EPOCH FROM (sch.end_time - sch.start_time))/3600 as hours
                FROM schedules sch
                JOIN staff st ON sch.staff_id = st.id
                LEFT JOIN departments d ON st.department_id = d.id
                WHERE sch.schedule_date BETWEEN %s AND %s
            """
            params = [start_date, end_date]
            
            if department_id:
                query += " AND st.department_id = %s"
                params.append(department_id)
            
            query += " ORDER BY sch.schedule_date, sch.start_time"
            
        elif report_type == "coverage":
            start_date = date(int(year), int(month), 1)
            if int(month) == 12:
                end_date = date(int(year) + 1, 1, 1) - timedelta(days=1)
            else:
                end_date = date(int(year), int(month) + 1, 1) - timedelta(days=1)
            
            query = """
                SELECT 
                    sch.schedule_date,
                    COUNT(DISTINCT sch.staff_id) as staff_count,
                    STRING_AGG(st.first_name || ' ' || st.last_name, ', ') as staff_names
                FROM schedules sch
                JOIN staff st ON sch.staff_id = st.id
                WHERE sch.schedule_date BETWEEN %s AND %s
            """
            params = [start_date, end_date]
            
            if department_id:
                query += " AND st.department_id = %s"
                params.append(department_id)
            
            query += """
                GROUP BY sch.schedule_date
                ORDER BY sch.schedule_date
            """
        
        cur.execute(query, params)
        report_data = cur.fetchall()
        
        # Get departments for filter
        cur.execute("SELECT id, name FROM departments WHERE status = 'Active' ORDER BY name")
        departments = cur.fetchall()
        
        # Get months and years for filter
        months = [(i, month_name[i]) for i in range(1, 13)]
        years = range(date.today().year - 5, date.today().year + 1)
        
    except Exception as e:
        app.logger.error(f"Error generating schedule report: {e}")
        report_data = []
        departments = []
        months = [(i, month_name[i]) for i in range(1, 13)]
        years = range(date.today().year - 5, date.today().year + 1)
    
    finally:
        cur.close()
        conn.close()
    
    return render_template("schedule_reports.html",
                         report_data=report_data,
                         report_type=report_type,
                         departments=departments,
                         months=months,
                         years=years,
                         selected_month=int(month),
                         selected_year=int(year),
                         selected_department=department_id,
                         hospital_name="Memorial Hospital Ovuru, Nsukka, Enugu State")

# Route to delete schedule
@app.route("/hr/scheduling/delete/<int:schedule_id>", methods=["POST"])
def delete_schedule(schedule_id):
    if "hr_user_id" not in session:
        return jsonify({"success": False, "message": "Unauthorized"}), 401
    
    conn = get_db_connection()
    cur = conn.cursor()
    
    try:
        cur.execute("DELETE FROM schedules WHERE id = %s", (schedule_id,))
        conn.commit()
        return jsonify({"success": True, "message": "Schedule deleted successfully"})
        
    except Exception as e:
        conn.rollback()
        app.logger.error(f"Error deleting schedule: {e}")
        return jsonify({"success": False, "message": str(e)}), 500
        
    finally:
        cur.close()
        conn.close()
        
# ==================== ROUTES: SCHEDULING MODULE ====================

# @app.route("/hr/scheduling")
# def scheduling():
#     if "hr_user_id" not in session:
#         return redirect(url_for("hr_login"))
    
#     conn = get_db_connection()
#     cur = conn.cursor()
    
#     try:
#         # Get current month schedules
#         current_month = date.today().replace(day=1)
#         if current_month.month == 12:
#             next_month = current_month.replace(year=current_month.year + 1, month=1)
#         else:
#             next_month = current_month.replace(month=current_month.month + 1)
        
#         cur.execute("""
#             SELECT s.*, st.first_name, st.last_name, st.position, d.name as department_name
#             FROM schedules s
#             JOIN staff st ON s.staff_id = st.id
#             LEFT JOIN departments d ON st.department_id = d.id
#             WHERE s.schedule_date >= %s AND s.schedule_date < %s
#             ORDER BY s.schedule_date, s.start_time
#         """, (current_month, next_month))
        
#         schedules = cur.fetchall()
        
#         # Get statistics
#         cur.execute("SELECT COUNT(*) FROM schedules WHERE schedule_date >= CURRENT_DATE")
#         upcoming_shifts = cur.fetchone()[0] or 0
        
#         cur.execute("""
#             SELECT COUNT(DISTINCT staff_id) 
#             FROM schedules 
#             WHERE schedule_date >= CURRENT_DATE
#         """)
#         staff_scheduled = cur.fetchone()[0] or 0
        
#         # Get departments for filter
#         cur.execute("SELECT id, name FROM departments WHERE status = 'Active' ORDER BY name")
#         departments = cur.fetchall()
        
#         # Get staff for filter
#         cur.execute("""
#             SELECT id, first_name, last_name, position 
#             FROM staff 
#             WHERE status = 'Active' 
#             ORDER BY first_name, last_name
#         """)
#         staff_list = cur.fetchall()
        
#     except Exception as e:
#         app.logger.error(f"Error fetching scheduling data: {e}")
#         schedules = []
#         upcoming_shifts = 0
#         staff_scheduled = 0
#         departments = []
#         staff_list = []
    
#     finally:
#         cur.close()
#         conn.close()
    
#     return render_template("scheduling_dashboard.html",
#                          schedules=schedules,
#                          upcoming_shifts=upcoming_shifts,
#                          staff_scheduled=staff_scheduled,
#                          departments=departments,
#                          staff_list=staff_list,
#                          current_month=current_month.strftime("%B %Y"),
#                          hospital_name="Memorial Hospital Ovuru, Nsukka, Enugu State")









@app.route("/hr/scheduling/check-availability", methods=["POST"])
def check_availability():
    """AJAX endpoint to check staff availability"""
    if "hr_user_id" not in session:
        return jsonify({"available": False, "message": "Unauthorized"}), 401
    
    data = request.json
    staff_id = data.get("staff_id")
    schedule_date = data.get("schedule_date")
    start_time = data.get("start_time")
    end_time = data.get("end_time")
    
    if not all([staff_id, schedule_date, start_time, end_time]):
        return jsonify({"available": False, "message": "Missing parameters"}), 400
    
    conn = get_db_connection()
    cur = conn.cursor()
    
    try:
        cur.execute("""
            SELECT id FROM schedules 
            WHERE staff_id = %s 
            AND schedule_date = %s
            AND NOT (%s <= start_time OR %s >= end_time)
        """, (staff_id, schedule_date, end_time, start_time))
        
        conflict = cur.fetchone()
        available = conflict is None
        
        return jsonify({
            "available": available,
            "message": "Available" if available else "Staff has a conflicting schedule"
        })
        
    except Exception as e:
        app.logger.error(f"Error checking availability: {e}")
        return jsonify({"available": False, "message": str(e)}), 500
        
    finally:
        cur.close()
        conn.close()
# -------------------- RUN APP --------------------
if __name__ == "__main__":
    create_tables()  # Your existing tables
    create_default_users()  # Your existing default users
    create_hr_tables()  # Add this line for HR tables
    app.run(debug=True)