# -------------------- IMPORTS --------------------
from flask import Flask, render_template, request, redirect, url_for, session, flash
from werkzeug.security import generate_password_hash, check_password_hash
import psycopg2
import os
from datetime import datetime, date


# -------------------- FLASK APP SETUP --------------------
app = Flask(__name__)
app.secret_key = "super_secret_key_change_later"  # change for production

# -------------------- DATABASE SETUP --------------------
# PostgreSQL connection string
DATABASE_URL = "postgresql://flask_user:Olarewaju1.@localhost:5432/hospital_db"

def get_db_connection():
    try:
        conn = psycopg2.connect(DATABASE_URL)
        return conn
    except Exception as e:
        print(f"Database connection error: {e}")
        return None

# -------------------- DATABASE INITIALIZATION --------------------
def create_pharmacists_table():
    """Create pharmacists table if it does not exist"""
    conn = get_db_connection()
    if not conn:
        print("Cannot create table, no DB connection")
        return

    cursor = conn.cursor()
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS pharmacists (
            id SERIAL PRIMARY KEY,
            username VARCHAR(50) UNIQUE NOT NULL,
            password VARCHAR(255) NOT NULL,
            is_active BOOLEAN DEFAULT TRUE,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    """)
    conn.commit()
    cursor.close()
    conn.close()

def create_default_pharmacist():
    """Insert or update a default pharmacist account"""
    conn = get_db_connection()
    if not conn:
        print("Cannot insert default pharmacist, no DB connection")
        return

    cursor = conn.cursor()
    hashed_password = generate_password_hash("pharma123")  # default password

    # Insert or update password if pharmacist already exists
    cursor.execute("""
        INSERT INTO pharmacists (username, password)
        VALUES (%s, %s)
        ON CONFLICT (username)
        DO UPDATE SET password = EXCLUDED.password
    """, ("pharmacist1", hashed_password))

    conn.commit()
    cursor.close()
    conn.close()

# -------------------- LANDING PAGE / DASHBOARD --------------------
@app.route('/')
def landing_page():
    hospital_name = "Memorial Hospital Ovuru, Nsukka, Enugu State"
    modules = [
        "System Admin",
        "Patient Services",
        "Clinical Services",
        "Pharmacy",
        "Laboratory",
        "Radiology",
        "Billing and Revenue",
        "Human Resources",
        "Management and Reports"
    ]
    return render_template("dashboard.html", hospital_name=hospital_name, modules=modules)

# -------------------- MODULE ROUTES PLACEHOLDERS --------------------
@app.route('/<module_name>')
def module_placeholder(module_name):
    display_name = module_name.replace('_', ' ').title()
    # Redirect pharmacy module to login page
    if module_name.lower() == "pharmacy":
        return redirect(url_for('pharmacy_login'))
    return render_template("module_placeholder.html", module_name=display_name)

# -------------------- PHARMACY LOGIN --------------------
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

# -------------------- PHARMACY DASHBOARD --------------------
@app.route('/pharmacy/dashboard')
def pharmacy_dashboard():
    if 'pharmacist_id' not in session:
        return redirect(url_for('pharmacy_login'))

    return render_template(
        "pharmacy_dashboard.html",
        pharmacist_name=session.get('pharmacist_username')
    )

# -------------------- PHARMACY LOGOUT --------------------
@app.route('/pharmacy/logout')
def pharmacy_logout():
    session.clear()
    return redirect(url_for('pharmacy_login'))


# # -------------------- DRUG SALES DASHBOARD --------------------# Pharmacy Dashboard

@app.route('/pharmacy/drug_sales')
def drug_sales():
    if 'pharmacist_id' not in session:
        return redirect(url_for('pharmacy_login'))

    return render_template(
        "drug_sales_dashboard.html",
        pharmacist_name=session.get('pharmacist_username'),
        hospital_name="Memorial Hospital Ovuru, Nsukka, Enugu State"
    )


# -------------------- ADD or RESTOCK --------------------

@app.route('/pharmacy/add-stock', methods=['GET', 'POST'])
def add_stock():
    if 'pharmacist_id' not in session:
        return redirect(url_for('pharmacy_login'))

    if request.method == 'POST':
        drug_name = request.form.get('drug_name', '').strip()
        strength = request.form.get('strength', '').strip()
        unit_price = request.form.get('unit_price', '').strip()
        quantity = request.form.get('quantity', '').strip()
        expiry_date = request.form.get('expiry_date', '').strip()  # New field

        # Validation
        if not all([drug_name, strength, unit_price, quantity, expiry_date]):
            flash("All fields including expiry date are required.", "danger")
            return redirect(url_for('add_stock'))

        try:
            unit_price = float(unit_price)
            quantity = int(quantity)
        except ValueError:
            flash("Invalid price or quantity.", "danger")
            return redirect(url_for('add_stock'))

        # Validate expiry date format
        try:
            expiry_date_obj = datetime.strptime(expiry_date, "%Y-%m-%d").date()
        except ValueError:
            flash("Invalid expiry date format.", "danger")
            return redirect(url_for('add_stock'))

        conn = get_db_connection()
        if not conn:
            flash("Database connection error.", "danger")
            return redirect(url_for('add_stock'))

        cursor = conn.cursor()

        # Check if the drug + strength + expiry already exists
        cursor.execute("""
            SELECT id, stock_quantity
            FROM drugs
            WHERE name = %s AND strength = %s AND expiry_date = %s
        """, (drug_name, strength, expiry_date_obj))

        existing = cursor.fetchone()

        if existing:
            # Increase stock if the same batch exists
            cursor.execute("""
                UPDATE drugs
                SET stock_quantity = stock_quantity + %s,
                    unit_price = %s,
                    updated_at = NOW()
                WHERE id = %s
            """, (quantity, unit_price, existing[0]))
        else:
            # Insert new drug batch
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



from flask import jsonify

@app.route('/api/drugs')
def api_drugs():
    if 'pharmacist_id' not in session:
        return jsonify([])

    search = request.args.get('q', '').strip()

    conn = get_db_connection()
    cur = conn.cursor()

    if search:
        cur.execute("""
            SELECT id, name, strength, unit_price, stock_quantity
            FROM drugs
            WHERE stock_quantity > 0
              AND LOWER(name) LIKE LOWER(%s)
            ORDER BY name ASC
        """, (f"{search}%",))
    else:
        cur.execute("""
            SELECT id, name, strength, unit_price, stock_quantity
            FROM drugs
            WHERE stock_quantity > 0
            ORDER BY name ASC
        """)

    rows = cur.fetchall()
    cur.close()
    conn.close()

    return jsonify([
        {
            "id": r[0],
            "name": r[1],
            "strength": r[2],
            "unit_price": float(r[3]),
            "stock_quantity": r[4]
        }
        for r in rows
    ])

import uuid
from flask import jsonify

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

    sale = cur.fetchone()
    cur.close()
    conn.close()

    if not sale:
        flash("Receipt not found", "danger")
        return redirect(url_for('pharmacy_dashboard'))

    return render_template(
        "receipt.html",
        receipt={
            "receipt_no": sale[0],
            "patient_name": sale[1],
            "patient_id": sale[2],
            "items": sale[3],
            "subtotal": sale[4],
            "discount": sale[5],
            "tax": sale[6],
            "grand_total": sale[7],
            "date": sale[9].strftime("%Y-%m-%d %H:%M")
        },
        pharmacist_name=sale[8],
        hospital_name="Memorial Hospital Ovuru, Nsukka, Enugu State"
    )


@app.route("/receipt_search", methods=["GET", "POST"])
def receipt_search():
    receipts = []

    if request.method == "POST":
        query = request.form.get("query")

        conn = get_db_connection()
        cur = conn.cursor()

        cur.execute("""
            SELECT id, patient_name, patient_id, created_at, grand_total
            FROM receipts
            WHERE patient_name ILIKE %s
               OR patient_id ILIKE %s
            ORDER BY created_at DESC;
        """, (f"%{query}%", f"%{query}%"))

        receipts = cur.fetchall()
        cur.close()
        conn.close()

    return render_template("receipt_search.html", receipts=receipts)


@app.route("/pharmacy/confirm-payment", methods=["POST"])
def confirm_payment():
    data = request.get_json()

    conn = get_db_connection()
    cur = conn.cursor()

    cur.execute("""
        INSERT INTO receipts (
            patient_name,
            patient_id,
            subtotal,
            discount,
            tax,
            total_amount,
            grand_total
        )
        VALUES (%s, %s, %s, %s, %s, %s, %s)
        RETURNING id;
    """, (
        data["patient_name"],
        data["patient_id"],
        data["subtotal"],
        data["discount"],
        data["tax"],
        data["grand_total"],
        data["grand_total"]
    ))

    receipt_id = cur.fetchone()[0]

    for item in data["items"]:
        cur.execute("""
            INSERT INTO receipt_items (
                receipt_id,
                drug_name,
                strength,
                quantity,
                unit_price
            )
            VALUES (%s, %s, %s, %s, %s);
        """, (
            receipt_id,
            item["drug_name"],
            item["strength"],
            item["quantity"],
            item["unit_price"]
        ))

        cur.execute("""
            UPDATE drugs
            SET stock_quantity = stock_quantity - %s
            WHERE name = %s AND strength = %s;
        """, (
            item["quantity"],
            item["drug_name"],
            item["strength"]
        ))

    conn.commit()
    cur.close()
    conn.close()

    return jsonify({"success": True, "receipt_id": receipt_id})



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

    # Fetch receipt header
    cur.execute("""
        SELECT id, patient_name, patient_id, subtotal, discount, tax, grand_total, created_at
        FROM receipts
        WHERE id = %s
    """, (receipt_id,))

    receipt = cur.fetchone()

    if not receipt:
        cur.close()
        conn.close()
        flash("Receipt not found", "danger")
        return redirect(url_for("pharmacy_dashboard"))

    # Fetch receipt items
    cur.execute("""
        SELECT drug_name, strength, quantity, unit_price
        FROM receipt_items
        WHERE receipt_id = %s
    """, (receipt_id,))

    items = cur.fetchall()

    cur.close()
    conn.close()

    return render_template(
        "receipt.html",
        receipt=receipt,
        items=items,
        hospital_name="Memorial Hospital Ovuru, Nsukka, Enugu State"
    )

def build_stock_snapshot(rows, today):
    stock = []

    for r in rows:
        expiry_date = r[5]
        quantity = r[3]
        threshold = r[6] if r[6] else 20

        if expiry_date:
            days_left = (expiry_date - today).days
            if days_left < 0:
                status = "EXPIRED"
            elif days_left <= 30:
                status = "EXPIRING_SOON"
            else:
                status = "VALID"
        else:
            days_left = None
            status = "UNKNOWN"

        stock.append({
            "id": r[0],
            "name": r[1],
            "strength": r[2],
            "quantity": quantity,
            "unit_price": r[4],
            "expiry_date": expiry_date,
            "days_left": days_left,
            "status": status,
            "low_stock_threshold": threshold,
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



@app.route("/pharmacy/stock-report")
def stock_report():
    if 'pharmacist_id' not in session:
        return redirect(url_for('pharmacy_login'))

    filter_type = request.args.get("filter", "all")

    conn = get_db_connection()
    cur = conn.cursor()
    cur.execute("""
        SELECT
            id,
            name,
            strength,
            stock_quantity,
            unit_price,
            expiry_date,
            low_stock_threshold
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




from openpyxl import Workbook
from openpyxl.styles import Font
from flask import send_file
from datetime import date
import io

from flask import request
from openpyxl.styles import PatternFill

@app.route("/pharmacy/stock-report/export")
def export_stock_report():
    if 'pharmacist_id' not in session:
        return redirect(url_for('pharmacy_login'))

    filter_type = request.args.get("filter", "all")

    conn = get_db_connection()
    cur = conn.cursor()
    cur.execute("""
        SELECT
            id,
            name,
            strength,
            stock_quantity,
            unit_price,
            expiry_date,
            low_stock_threshold
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
            item["name"],
            item["strength"],
            item["quantity"],
            float(item["unit_price"]),
            item["expiry_date"],
            item["days_left"],
            item["status"],
            float(item["total_value"]),
            item["low_stock_threshold"]
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




def log_stock_movement(drug_id, movement_type, quantity, user_id):
    conn = get_db_connection()
    cur = conn.cursor()
    cur.execute("""
        INSERT INTO stock_movements (drug_id, movement_type, quantity, user_id)
        VALUES (%s, %s, %s, %s)
    """, (drug_id, movement_type, quantity, user_id))
    conn.commit()
    cur.close()
    conn.close()

@app.route('/pharmacy/stock-movements')
def stock_movements():
    if 'pharmacist_id' not in session:
        return redirect(url_for('pharmacy_login'))

    conn = get_db_connection()
    cur = conn.cursor()
    cur.execute("""
        SELECT sm.id, d.name, d.strength, sm.movement_type, sm.quantity, u.username, sm.created_at, sm.note
        FROM stock_movements sm
        JOIN drugs d ON sm.drug_id = d.id
        JOIN users u ON sm.user_id = u.id
        ORDER BY sm.created_at DESC
    """)
    movements = cur.fetchall()
    cur.close()
    conn.close()

    return render_template('stock_movements.html', movements=movements)


@app.route("/pharmacy/stock-report/monthly-valuation")
def monthly_valuation():
    if 'pharmacist_id' not in session:
        return redirect(url_for('pharmacy_login'))

    conn = get_db_connection()
    cur = conn.cursor()

    cur.execute("""
        SELECT DATE_TRUNC('month', created_at) AS month, 
               SUM(quantity * unit_price) AS total_value
        FROM stock_movements sm
        JOIN drugs d ON sm.drug_id = d.id
        GROUP BY month
        ORDER BY month DESC;
    """)
    monthly_data = cur.fetchall()
    cur.close()
    conn.close()

    return render_template('monthly_valuation.html', monthly_data=monthly_data)


# -------------------- RUN APP --------------------
if __name__ == "__main__":
    create_pharmacists_table()
    create_default_pharmacist()
    app.run(debug=True)
