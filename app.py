import streamlit as st
import sqlite3
import pandas as pd
from datetime import datetime
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
import io

# ==========================================
# 1. DATABASE SETUP & INITIALIZATION
# ==========================================
DB_FILE = "inventory.db"
ADMIN_PIN = "1234"  # Default Admin PIN

def get_connection():
    return sqlite3.connect(DB_FILE, check_same_thread=False)

def init_db():
    conn = get_connection()
    c = conn.cursor()
    
    # Products Table
    c.execute('''
        CREATE TABLE IF NOT EXISTS products (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT UNIQUE NOT NULL,
            category TEXT NOT NULL,
            cost_price REAL NOT NULL DEFAULT 0.0,
            price REAL NOT NULL,
            stock INTEGER NOT NULL
        )
    ''')
    
    # Auto-migrate older database versions to include cost_price
    try:
        c.execute("ALTER TABLE products ADD COLUMN cost_price REAL NOT NULL DEFAULT 0.0")
    except sqlite3.OperationalError:
        pass  # Column already exists
    
    # Sales Table
    c.execute('''
        CREATE TABLE IF NOT EXISTS sales (
            receipt_id TEXT NOT NULL,
            timestamp TEXT NOT NULL,
            cashier TEXT NOT NULL,
            customer_name TEXT NOT NULL DEFAULT 'Walk-in',
            payment_type TEXT NOT NULL DEFAULT 'Paid',
            product_name TEXT NOT NULL,
            quantity INTEGER NOT NULL,
            cost_price REAL NOT NULL DEFAULT 0.0,
            unit_price REAL NOT NULL,
            total_price REAL NOT NULL,
            profit REAL NOT NULL DEFAULT 0.0
        )
    ''')
    
    # Auto-migrate older database versions for new sales fields
    for col_def in [
        ("customer_name", "TEXT NOT NULL DEFAULT 'Walk-in'"),
        ("payment_type", "TEXT NOT NULL DEFAULT 'Paid'"),
        ("cost_price", "REAL NOT NULL DEFAULT 0.0"),
        ("profit", "REAL NOT NULL DEFAULT 0.0")
    ]:
        try:
            c.execute(f"ALTER TABLE sales ADD COLUMN {col_def[0]} {col_def[1]}")
        except sqlite3.OperationalError:
            pass

    # Generic sample products setup
    c.execute("SELECT COUNT(*) FROM products")
    if c.fetchone()[0] == 0:
        default_items = [
            ("Standard Product A", "General", 10.00, 25.00, 50),
            ("Standard Product B", "General", 15.00, 35.00, 40),
            ("Basic Service Package", "Services", 20.00, 60.00, 100),
            ("Premium Item", "General", 50.00, 120.00, 20)
        ]
        c.executemany("INSERT INTO products (name, category, cost_price, price, stock) VALUES (?, ?, ?, ?, ?)", default_items)
    
    conn.commit()
    conn.close()

init_db()

# ==========================================
# 2. PAGE CONFIGURATION & STYLING
# ==========================================
st.set_page_config(page_title="Point of Sale System", layout="wide", page_icon="🛒")

st.markdown("""
    <style>
        .main { padding: 1rem; }
        .stButton>button { width: 100%; border-radius: 5px; height: 2.7em; }
    </style>
""", unsafe_allow_html=True)

# ==========================================
# 3. SESSION STATE INITIALIZATION
# ==========================================
if "cart" not in st.session_state:
    st.session_state.cart = []

if "logged_in" not in st.session_state:
    st.session_state.logged_in = False

if "cashier_name" not in st.session_state:
    st.session_state.cashier_name = ""

if "is_admin" not in st.session_state:
    st.session_state.is_admin = False

if "checkout_step" not in st.session_state:
    st.session_state.checkout_step = "catalog"

# ==========================================
# 4. HELPER FUNCTIONS
# ==========================================
def generate_docx_receipt(receipt_id, cashier, customer, pay_type, cart_items, total_amount):
    doc = Document()
    
    for section in doc.sections:
        section.top_margin = Inches(0.2)
        section.bottom_margin = Inches(0.2)
        section.left_margin = Inches(0.2)
        section.right_margin = Inches(0.2)

    title = doc.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = title.add_run("SALES RECEIPT\n")
    run.bold = True
    run.font.size = Pt(13)
    
    sub = title.add_run("Official Transaction Record\n")
    sub.font.size = Pt(9)
    
    doc.add_paragraph("-" * 35)
    
    info = doc.add_paragraph()
    info.paragraph_format.space_after = Pt(2)
    info.add_run(f"Receipt ID: {receipt_id}\n")
    info.add_run(f"Date: {datetime.now().strftime('%Y-%m-%d %H:%M')}\n")
    info.add_run(f"Cashier: {cashier}\n")
    info.add_run(f"Customer: {customer}\n")
    info.add_run(f"Payment Status: {pay_type.upper()}\n")
    
    doc.add_paragraph("-" * 35)
    
    table = doc.add_table(rows=1, cols=4)
    hdr_cells = table.rows[0].cells
    hdr_cells[0].text = 'Item'
    hdr_cells[1].text = 'Qty'
    hdr_cells[2].text = 'Price'
    hdr_cells[3].text = 'Total'
    
    for item in cart_items:
        row_cells = table.add_row().cells
        row_cells[0].text = str(item['name'])
        row_cells[1].text = str(item['qty'])
        row_cells[2].text = f"${item['price']:.2f}"
        row_cells[3].text = f"${item['total']:.2f}"

    doc.add_paragraph("-" * 35)
    
    tot_p = doc.add_paragraph()
    tot_p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    r_tot = tot_p.add_run(f"TOTAL: ${total_amount:.2f}")
    r_tot.bold = True
    r_tot.font.size = Pt(12)
    
    footer = doc.add_paragraph()
    footer.alignment = WD_ALIGN_PARAGRAPH.CENTER
    footer.add_run("\nThank you for your business!")
    
    bio = io.BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio

# ==========================================
# 5. AUTHENTICATION & SIDEBAR
# ==========================================
st.title("🛒 Point of Sale System")

with st.sidebar:
    st.header("Cashier Login")
    if not st.session_state.logged_in:
        user_input = st.text_input("Enter Cashier Name")
        if st.button("Login"):
            if user_input.strip() != "":
                st.session_state.logged_in = True
                st.session_state.cashier_name = user_input.strip()
                st.rerun()
            else:
                st.error("Please enter a name.")
    else:
        st.success(f"Cashier: **{st.session_state.cashier_name}**")
        
        st.markdown("---")
        if not st.session_state.is_admin:
            pin = st.text_input("Unlock Admin Rights", type="password", key="admin_pin_input")
            if pin == ADMIN_PIN:
                st.session_state.is_admin = True
                st.success("Admin Mode Active")
                st.rerun()
            elif pin != "":
                st.error("Incorrect PIN")
        else:
            st.info("🔓 Admin Access Enabled")
            if st.button("Lock Admin"):
                st.session_state.is_admin = False
                st.rerun()

        st.markdown("---")
        if st.button("Logout"):
            st.session_state.logged_in = False
            st.session_state.cashier_name = ""
            st.session_state.is_admin = False
            st.session_state.cart = []
            st.session_state.checkout_step = "catalog"
            st.rerun()

if not st.session_state.logged_in:
    st.warning("Please log in from the sidebar to start transactions.")
    st.stop()

# ==========================================
# 6. APPLICATION NAVIGATION
# ==========================================
if st.session_state.is_admin:
    tabs = st.tabs(["🛒 Cashier Terminal", "📦 Inventory Control", "📊 Business Reports"])
    tab_pos = tabs[0]
    tab_inv = tabs[1]
    tab_reports = tabs[2]
else:
    tab_pos = st.container()
    tab_inv = None
    tab_reports = None

# ------------------------------------------
# TAB 1: CASHIER TERMINAL (MULTI-STEP FLOW)
# ------------------------------------------
with tab_pos:
    if st.session_state.checkout_step == "catalog":
        st.subheader("Step 1: Select Products & Services")
        
        conn = get_connection()
        df_products = pd.read_sql_query("SELECT * FROM products WHERE stock > 0", conn)
        conn.close()

        search_term = st.text_input("Search catalog...", key="pos_search")
        if not df_products.empty:
            filtered_df = df_products[df_products['name'].str.contains(search_term, case=False)]
            
            for idx, row in filtered_df.iterrows():
                with st.container():
                    c1, c2, c3, c4 = st.columns([3, 1, 1, 1])
                    c1.write(f"**{row['name']}**")
                    c2.write(f"${row['price']:.2f}")
                    c3.write(f"Stock: {row['stock']}")
                    if c4.button("Add to Order", key=f"add_{row['id']}"):
                        existing_item = next((item for item in st.session_state.cart if item['id'] == row['id']), None)
                        
                        # Safely retrieve cost_price or default to 0.0
                        item_cost = float(row['cost_price']) if 'cost_price' in row and pd.notna(row['cost_price']) else 0.0
                        
                        if existing_item:
                            if existing_item['qty'] < row['stock']:
                                existing_item['qty'] += 1
                                existing_item['total'] = existing_item['qty'] * existing_item['price']
                            else:
                                st.error("Cannot add more than available stock.")
                        else:
                            st.session_state.cart.append({
                                'id': row['id'],
                                'name': row['name'],
                                'cost': item_cost,
                                'price': float(row['price']),
                                'qty': 1,
                                'total': float(row['price'])
                            })
                        st.rerun()
        else:
            st.info("No available inventory.")

        st.markdown("---")
        cart_count = sum(item['qty'] for item in st.session_state.cart)
        
        if cart_count > 0:
            if st.button(f"➡️ Proceed to Current Order Summary ({cart_count} items)", type="primary"):
                st.session_state.checkout_step = "summary"
                st.rerun()

    elif st.session_state.checkout_step == "summary":
        st.subheader("Step 2: Current Order Summary & Checkout")
        
        if st.button("⬅️ Back to Product Catalog"):
            st.session_state.checkout_step = "catalog"
            st.rerun()

        st.markdown("---")
        
        if len(st.session_state.cart) == 0:
            st.warning("Your cart is empty.")
            if st.button("Return to Catalog"):
                st.session_state.checkout_step = "catalog"
                st.rerun()
        else:
            col_summary, col_payment = st.columns([1.5, 1])
            
            with col_summary:
                st.markdown("### Order Items")
                for index, item in enumerate(st.session_state.cart):
                    col_item_name, col_item_qty, col_item_price, col_item_del = st.columns([3, 1, 1, 0.5])
                    col_item_name.write(item['name'])
                    col_item_qty.write(f"x{item['qty']}")
                    col_item_price.write(f"${item['total']:.2f}")
                    if col_item_del.button("❌", key=f"del_{index}"):
                        st.session_state.cart.pop(index)
                        if len(st.session_state.cart) == 0:
                            st.session_state.checkout_step = "catalog"
                        st.rerun()
                
                grand_total = sum(item['total'] for item in st.session_state.cart)
                st.markdown(f"## **Grand Total: ${grand_total:.2f}**")

            with col_payment:
                st.markdown("### Customer & Payment Details")
                customer_name = st.text_input("Customer Name", value="Walk-in", key="cust_name_field")
                payment_type = st.radio("Payment Status", ["Paid", "Pay Later / Tab"], key="pay_type_field")
                
                if st.button("Complete Transaction & Generate Receipt", type="primary"):
                    receipt_id = f"REC-{int(datetime.now().timestamp())}"
                    timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                    
                    conn = get_connection()
                    cursor = conn.cursor()
                    
                    for item in st.session_state.cart:
                        item_profit = (item['price'] - item['cost']) * item['qty']
                        cursor.execute("UPDATE products SET stock = stock - ? WHERE id = ?", (item['qty'], item['id']))
                        cursor.execute("""
                            INSERT INTO sales (receipt_id, timestamp, cashier, customer_name, payment_type, product_name, quantity, cost_price, unit_price, total_price, profit)
                            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                        """, (receipt_id, timestamp, st.session_state.cashier_name, customer_name, payment_type, item['name'], item['qty'], item['cost'], item['price'], item['total'], item_profit))
                    
                    conn.commit()
                    conn.close()
                    
                    doc_file = generate_docx_receipt(receipt_id, st.session_state.cashier_name, customer_name, payment_type, st.session_state.cart, grand_total)
                    
                    st.success(f"Sale Recorded! Receipt ID: {receipt_id}")
                    st.download_button(
                        label="📄 Download / Print Receipt",
                        data=doc_file,
                        file_name=f"{receipt_id}.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    )
                    
                    st.session_state.cart = []
                    st.session_state.checkout_step = "catalog"

# ------------------------------------------
# TAB 2 & 3: ADMIN INVENTORY & REPORTS
# ------------------------------------------
if st.session_state.is_admin:
    with tab_inv:
        st.subheader("Inventory Management")
        
        conn = get_connection()
        df_all_products = pd.read_sql_query("SELECT * FROM products", conn)
        conn.close()

        st.dataframe(df_all_products, use_container_width=True)

        st.markdown("---")
        st.subheader("Add or Update Product")
        
        with st.form("inventory_form", clear_on_submit=True):
            prod_name = st.text_input("Product / Service Name")
            prod_cat = st.selectbox("Category", ["General", "Services", "Supplies", "Other"])
            prod_cost = st.number_input("Cost Price", min_value=0.0, step=1.0)
            prod_price = st.number_input("Selling Price", min_value=0.0, step=1.0)
            prod_stock = st.number_input("Stock Quantity", min_value=0, step=1)
            
            submitted = st.form_submit_button("Save Item")
            if submitted:
                if prod_name.strip() != "":
                    conn = get_connection()
                    c = conn.cursor()
                    c.execute("""
                        INSERT INTO products (name, category, cost_price, price, stock)
                        VALUES (?, ?, ?, ?, ?)
                        ON CONFLICT(name) DO UPDATE SET
                            category=excluded.category,
                            cost_price=excluded.cost_price,
                            price=excluded.price,
                            stock=products.stock + excluded.stock
                    """, (prod_name.strip(), prod_cat, prod_cost, prod_price, prod_stock))
                    conn.commit()
                    conn.close()
                    st.success(f"Saved: {prod_name}")
                    st.rerun()
                else:
                    st.error("Item name required.")

    with tab_reports:
        st.subheader("Business Reports")
        
        conn = get_connection()
        df_sales = pd.read_sql_query("SELECT * FROM sales ORDER BY timestamp DESC", conn)
        conn.close()
        
        if not df_sales.empty:
            col_m1, col_m2, col_m3, col_m4 = st.columns(4)
            total_rev = df_sales['total_price'].sum() if 'total_price' in df_sales else 0.0
            total_profit = df_sales['profit'].sum() if 'profit' in df_sales else 0.0
            
            if 'payment_type' in df_sales:
                unpaid_total = df_sales[df_sales['payment_type'] == 'Pay Later / Tab']['total_price'].sum()
            else:
                unpaid_total = 0.0
            
            col_m1.metric("Gross Revenue", f"${total_rev:.2f}")
            col_m2.metric("Net Profit", f"${total_profit:.2f}")
            col_m3.metric("Unpaid Customer Debt", f"${unpaid_total:.2f}")
            col_m4.metric("Total Sales Count", len(df_sales))
            
            st.markdown("---")
            st.subheader("Outstanding Customer Tabs")
            if 'payment_type' in df_sales:
                debtors = df_sales[df_sales['payment_type'] == 'Pay Later / Tab']
                if not debtors.empty:
                    st.dataframe(debtors[['receipt_id', 'timestamp', 'customer_name', 'product_name', 'total_price']], use_container_width=True)
                else:
                    st.info("No outstanding unpaid tabs.")
            else:
                st.info("No outstanding unpaid tabs.")
                
            st.markdown("---")
            st.subheader("All Sales History")
            st.dataframe(df_sales, use_container_width=True)
            
            csv = df_sales.to_csv(index=False).encode('utf-8')
            st.download_button(
                label="📥 Export Sales to CSV",
                data=csv,
                file_name=f"sales_report_{datetime.now().strftime('%Y_%m_%d')}.csv",
                mime="text/csv"
            )
        else:
            st.info("No transaction history.")
