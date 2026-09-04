```python
import io
import sqlite3
import uuid
from datetime import datetime, date

import pandas as pd
import streamlit as st
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Inches, Pt


# ============================================================
# APPLICATION CONFIGURATION
# ============================================================

DB_FILE = "inventory.db"
ADMIN_PIN = "1234"

APP_TITLE = "Point of Sale System"
CURRENCY_SYMBOL = "$"


# ============================================================
# PAGE CONFIGURATION
# ============================================================

st.set_page_config(
    page_title=APP_TITLE,
    page_icon="🛒",
    layout="wide",
    initial_sidebar_state="expanded",
)


st.markdown(
    """
    <style>
        .main {
            padding: 1rem;
        }

        .stButton > button {
            width: 100%;
            min-height: 2.6rem;
            border-radius: 6px;
        }

        div[data-testid="stMetric"] {
            border: 1px solid rgba(128, 128, 128, 0.25);
            border-radius: 8px;
            padding: 10px;
        }

        .product-card {
            border: 1px solid rgba(128, 128, 128, 0.25);
            border-radius: 8px;
            padding: 12px;
            margin-bottom: 8px;
        }

        .total-box {
            border: 2px solid rgba(128, 128, 128, 0.35);
            border-radius: 8px;
            padding: 15px;
            margin-top: 10px;
        }
    </style>
    """,
    unsafe_allow_html=True,
)


# ============================================================
# DATABASE
# ============================================================

def get_connection():
    """
    Create a short-lived SQLite connection.

    Connections are intentionally not stored in session state.
    This reduces the chance of SQLite thread/locking problems
    during Streamlit reruns.
    """
    conn = sqlite3.connect(
        DB_FILE,
        timeout=15,
        check_same_thread=False,
    )
    conn.row_factory = sqlite3.Row
    return conn


def table_columns(cursor, table_name):
    """Return existing column names for a SQLite table."""
    cursor.execute(f"PRAGMA table_info({table_name})")
    return {row["name"] for row in cursor.fetchall()}


def add_missing_columns(cursor, table_name, required_columns):
    """
    Safely add columns that are missing from an existing database.
    """
    existing = table_columns(cursor, table_name)

    for column_name, column_definition in required_columns.items():
        if column_name not in existing:
            cursor.execute(
                f"ALTER TABLE {table_name} "
                f"ADD COLUMN {column_name} {column_definition}"
            )


def init_db():
    """
    Initialize the database and automatically migrate older databases.

    PRAGMA table_info() is deliberately used before ALTER TABLE so
    deployments can reuse an existing inventory.db safely.
    """
    conn = get_connection()

    try:
        cursor = conn.cursor()

        # ----------------------------------------------------
        # Products
        # ----------------------------------------------------

        cursor.execute(
            """
            CREATE TABLE IF NOT EXISTS products (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT UNIQUE NOT NULL,
                category TEXT NOT NULL DEFAULT 'General',
                cost_price REAL NOT NULL DEFAULT 0.0,
                price REAL NOT NULL DEFAULT 0.0,
                stock INTEGER NOT NULL DEFAULT 0
            )
            """
        )

        add_missing_columns(
            cursor,
            "products",
            {
                "name": "TEXT",
                "category": "TEXT NOT NULL DEFAULT 'General'",
                "cost_price": "REAL NOT NULL DEFAULT 0.0",
                "price": "REAL NOT NULL DEFAULT 0.0",
                "stock": "INTEGER NOT NULL DEFAULT 0",
            },
        )

        # ----------------------------------------------------
        # Sales
        # ----------------------------------------------------

        cursor.execute(
            """
            CREATE TABLE IF NOT EXISTS sales (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                receipt_id TEXT NOT NULL,
                timestamp TEXT NOT NULL,
                cashier TEXT NOT NULL,
                customer_name TEXT NOT NULL DEFAULT 'Walk-in',
                payment_type TEXT NOT NULL DEFAULT 'Paid',
                product_name TEXT NOT NULL,
                quantity INTEGER NOT NULL DEFAULT 1,
                cost_price REAL NOT NULL DEFAULT 0.0,
                unit_price REAL NOT NULL DEFAULT 0.0,
                total_price REAL NOT NULL DEFAULT 0.0,
                profit REAL NOT NULL DEFAULT 0.0
            )
            """
        )

        add_missing_columns(
            cursor,
            "sales",
            {
                "receipt_id": "TEXT",
                "timestamp": "TEXT",
                "cashier": "TEXT NOT NULL DEFAULT 'Unknown'",
                "customer_name": "TEXT NOT NULL DEFAULT 'Walk-in'",
                "payment_type": "TEXT NOT NULL DEFAULT 'Paid'",
                "product_name": "TEXT NOT NULL DEFAULT 'Unknown'",
                "quantity": "INTEGER NOT NULL DEFAULT 1",
                "cost_price": "REAL NOT NULL DEFAULT 0.0",
                "unit_price": "REAL NOT NULL DEFAULT 0.0",
                "total_price": "REAL NOT NULL DEFAULT 0.0",
                "profit": "REAL NOT NULL DEFAULT 0.0",
            },
        )

        # ----------------------------------------------------
        # Indexes
        # ----------------------------------------------------

        cursor.execute(
            """
            CREATE INDEX IF NOT EXISTS idx_sales_receipt
            ON sales(receipt_id)
            """
        )

        cursor.execute(
            """
            CREATE INDEX IF NOT EXISTS idx_sales_timestamp
            ON sales(timestamp)
            """
        )

        cursor.execute(
            """
            CREATE INDEX IF NOT EXISTS idx_products_name
            ON products(name)
            """
        )

        # ----------------------------------------------------
        # Generic sample inventory only when database is empty
        # ----------------------------------------------------

        cursor.execute("SELECT COUNT(*) AS count FROM products")
        product_count = cursor.fetchone()["count"]

        if product_count == 0:
            default_items = [
                (
                    "Standard Product A",
                    "General",
                    10.00,
                    25.00,
                    50,
                ),
                (
                    "Standard Product B",
                    "General",
                    15.00,
                    35.00,
                    40,
                ),
                (
                    "Basic Service Package",
                    "Services",
                    20.00,
                    60.00,
                    100,
                ),
                (
                    "Premium Item",
                    "General",
                    50.00,
                    120.00,
                    20,
                ),
            ]

            cursor.executemany(
                """
                INSERT INTO products
                (name, category, cost_price, price, stock)
                VALUES (?, ?, ?, ?, ?)
                """,
                default_items,
            )

        conn.commit()

    except Exception:
        conn.rollback()
        raise

    finally:
        conn.close()


# Initialize database before the application UI.
init_db()


# ============================================================
# SESSION STATE
# ============================================================

DEFAULT_SESSION_STATE = {
    "cart": [],
    "logged_in": False,
    "cashier_name": "",
    "is_admin": False,
    "checkout_step": "catalog",
    "last_receipt_id": "",
    "last_receipt_file": None,
    "admin_pin_error": "",
}


for key, default_value in DEFAULT_SESSION_STATE.items():
    if key not in st.session_state:
        st.session_state[key] = default_value


# ============================================================
# GENERAL HELPERS
# ============================================================

def money(value):
    """Safely format a monetary value."""
    try:
        return f"{CURRENCY_SYMBOL}{float(value):,.2f}"
    except (TypeError, ValueError):
        return f"{CURRENCY_SYMBOL}0.00"


def safe_float(value, default=0.0):
    try:
        if pd.isna(value):
            return default
        return float(value)
    except (TypeError, ValueError):
        return default


def safe_int(value, default=0):
    try:
        if pd.isna(value):
            return default
        return int(value)
    except (TypeError, ValueError):
        return default


def get_cart_quantity(product_id):
    """Return the quantity currently in the cart for a product."""
    return sum(
        safe_int(item.get("qty"), 0)
        for item in st.session_state.cart
        if item.get("id") == product_id
    )


def calculate_cart_total():
    return sum(
        safe_float(item.get("total"), 0.0)
        for item in st.session_state.cart
    )


def normalize_cart():
    """
    Ensure older/session cart records contain all fields required
    by the current application.
    """
    normalized = []

    for item in st.session_state.cart:
        try:
            price = safe_float(item.get("price"), 0.0)
            qty = max(1, safe_int(item.get("qty"), 1))
            cost = safe_float(item.get("cost"), 0.0)

            normalized.append(
                {
                    "id": item.get("id"),
                    "name": str(item.get("name", "Unknown")),
                    "cost": cost,
                    "price": price,
                    "qty": qty,
                    "total": price * qty,
                }
            )
        except Exception:
            continue

    st.session_state.cart = normalized


normalize_cart()


def generate_receipt_id():
    """
    Generate a unique receipt ID.
    """
    timestamp = datetime.now().strftime("%Y%m%d%H%M%S")
    unique_part = uuid.uuid4().hex[:6].upper()
    return f"REC-{timestamp}-{unique_part}"


# ============================================================
# DATABASE READ FUNCTIONS
# ============================================================

def load_products(available_only=False):
    conn = get_connection()

    try:
        if available_only:
            query = """
                SELECT
                    id,
                    name,
                    category,
                    cost_price,
                    price,
                    stock
                FROM products
                WHERE stock > 0
                ORDER BY name COLLATE NOCASE
            """
        else:
            query = """
                SELECT
                    id,
                    name,
                    category,
                    cost_price,
                    price,
                    stock
                FROM products
                ORDER BY name COLLATE NOCASE
            """

        return pd.read_sql_query(query, conn)

    except Exception:
        return pd.DataFrame()

    finally:
        conn.close()


def load_sales():
    conn = get_connection()

    try:
        return pd.read_sql_query(
            """
            SELECT
                *
            FROM sales
            ORDER BY timestamp DESC, id DESC
            """,
            conn,
        )

    except Exception:
        return pd.DataFrame()

    finally:
        conn.close()


def get_product(product_id):
    conn = get_connection()

    try:
        cursor = conn.cursor()

        cursor.execute(
            """
            SELECT
                id,
                name,
                category,
                cost_price,
                price,
                stock
            FROM products
            WHERE id = ?
            """,
            (product_id,),
        )

        return cursor.fetchone()

    finally:
        conn.close()


# ============================================================
# CART OPERATIONS
# ============================================================

def add_product_to_cart(product):
    """
    Add one unit of a product to the cart while checking
    current inventory.
    """
    product_id = safe_int(product["id"], 0)

    current_qty = get_cart_quantity(product_id)
    available_stock = safe_int(product["stock"], 0)

    if current_qty >= available_stock:
        st.error("Cannot add more than the available stock.")
        return

    existing_item = next(
        (
            item
            for item in st.session_state.cart
            if item.get("id") == product_id
        ),
        None,
    )

    cost_price = safe_float(product["cost_price"], 0.0)
    selling_price = safe_float(product["price"], 0.0)

    if existing_item:
        existing_item["qty"] += 1
        existing_item["total"] = (
            existing_item["qty"] * existing_item["price"]
        )
    else:
        st.session_state.cart.append(
            {
                "id": product_id,
                "name": str(product["name"]),
                "cost": cost_price,
                "price": selling_price,
                "qty": 1,
                "total": selling_price,
            }
        )


def remove_cart_item(index):
    if 0 <= index < len(st.session_state.cart):
        st.session_state.cart.pop(index)

    if not st.session_state.cart:
        st.session_state.checkout_step = "catalog"


def clear_cart():
    st.session_state.cart = []
    st.session_state.checkout_step = "catalog"


# ============================================================
# RECEIPT GENERATION
# ============================================================

def generate_docx_receipt(
    receipt_id,
    cashier,
    customer,
    payment_type,
    cart_items,
    total_amount,
    transaction_time=None,
):
    """
    Generate a generic DOCX sales receipt.
    """
    doc = Document()

    for section in doc.sections:
        section.top_margin = Inches(0.25)
        section.bottom_margin = Inches(0.25)
        section.left_margin = Inches(0.3)
        section.right_margin = Inches(0.3)

    # --------------------------------------------------------
    # Header
    # --------------------------------------------------------

    title = doc.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER

    run = title.add_run("SALES RECEIPT\n")
    run.bold = True
    run.font.size = Pt(14)

    subtitle = title.add_run("Point of Sale Transaction Record")
    subtitle.font.size = Pt(9)

    separator = doc.add_paragraph()
    separator.alignment = WD_ALIGN_PARAGRAPH.CENTER
    separator.add_run("-" * 42)

    # --------------------------------------------------------
    # Transaction information
    # --------------------------------------------------------

    transaction_time = transaction_time or datetime.now().strftime(
        "%Y-%m-%d %H:%M:%S"
    )

    info = doc.add_paragraph()
    info.paragraph_format.space_after = Pt(2)

    info.add_run(f"Receipt ID: {receipt_id}\n")
    info.add_run(f"Date: {transaction_time}\n")
    info.add_run(f"Cashier: {cashier}\n")
    info.add_run(f"Customer: {customer}\n")
    info.add_run(f"Payment Status: {payment_type}\n")

    separator = doc.add_paragraph()
    separator.alignment = WD_ALIGN_PARAGRAPH.CENTER
    separator.add_run("-" * 42)

    # --------------------------------------------------------
    # Items
    # --------------------------------------------------------

    table = doc.add_table(rows=1, cols=4)

    headers = ["Item", "Qty", "Price", "Total"]

    for cell, header in zip(table.rows[0].cells, headers):
        cell.text = header

    for item in cart_items:
        row = table.add_row().cells

        row[0].text = str(item.get("name", "Unknown"))
        row[1].text = str(safe_int(item.get("qty"), 0))
        row[2].text = money(item.get("price", 0.0))
        row[3].text = money(item.get("total", 0.0))

    separator = doc.add_paragraph()
    separator.alignment = WD_ALIGN_PARAGRAPH.CENTER
    separator.add_run("-" * 42)

    # --------------------------------------------------------
    # Total
    # --------------------------------------------------------

    total_paragraph = doc.add_paragraph()
    total_paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT

    total_run = total_paragraph.add_run(
        f"TOTAL: {money(total_amount)}"
    )
    total_run.bold = True
    total_run.font.size = Pt(13)

    # --------------------------------------------------------
    # Footer
    # --------------------------------------------------------

    footer = doc.add_paragraph()
    footer.alignment = WD_ALIGN_PARAGRAPH.CENTER

    footer_run = footer.add_run(
        "\nThank you for your business!"
    )
    footer_run.font.size = Pt(9)

    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)

    return buffer


# ============================================================
# COMPLETE TRANSACTION
# ============================================================

def complete_transaction(customer_name, payment_type):
    """
    Complete a sale using a single SQLite transaction.

    Inventory is rechecked immediately before updating stock.
    This protects against selling more stock than is currently
    available if the database was changed after the cart was created.
    """
    normalize_cart()

    if not st.session_state.cart:
        return False, "The cart is empty.", None, None

    customer_name = customer_name.strip() or "Walk-in"
    payment_type = payment_type.strip() or "Paid"

    receipt_id = generate_receipt_id()
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    conn = get_connection()

    try:
        cursor = conn.cursor()

        # ----------------------------------------------------
        # Validate every cart item against current inventory
        # ----------------------------------------------------

        validated_items = []

        for item in st.session_state.cart:
            product_id = safe_int(item.get("id"), 0)
            requested_qty = safe_int(item.get("qty"), 0)

            if product_id <= 0 or requested_qty <= 0:
                raise ValueError(
                    "One or more cart items are invalid."
                )

            cursor.execute(
                """
                SELECT
                    id,
                    name,
                    cost_price,
                    price,
                    stock
                FROM products
                WHERE id = ?
                """,
                (product_id,),
            )

            product = cursor.fetchone()

            if product is None:
                raise ValueError(
                    f"Product '{item.get('name', 'Unknown')}' "
                    "is no longer available."
                )

            current_stock = safe_int(product["stock"], 0)

            if requested_qty > current_stock:
                raise ValueError(
                    f"Insufficient stock for '{product['name']}'. "
                    f"Available: {current_stock}, "
                    f"requested: {requested_qty}."
                )

            current_cost = safe_float(product["cost_price"], 0.0)
            current_price = safe_float(product["price"], 0.0)

            total_price = current_price * requested_qty
            profit = (
                current_price - current_cost
            ) * requested_qty

            validated_items.append(
                {
                    "id": product_id,
                    "name": str(product["name"]),
                    "cost": current_cost,
                    "price": current_price,
                    "qty": requested_qty,
                    "total": total_price,
                    "profit": profit,
                }
            )

        # ----------------------------------------------------
        # Record sale + decrease inventory
        # ----------------------------------------------------

        for item in validated_items:
            cursor.execute(
                """
                UPDATE products
                SET stock = stock - ?
                WHERE id = ?
                  AND stock >= ?
                """,
                (
                    item["qty"],
                    item["id"],
                    item["qty"],
                ),
            )

            if cursor.rowcount != 1:
                raise ValueError(
                    f"Unable to update inventory for "
                    f"'{item['name']}'."
                )

            cursor.execute(
                """
                INSERT INTO sales (
                    receipt_id,
                    timestamp,
                    cashier,
                    customer_name,
                    payment_type,
                    product_name,
                    quantity,
                    cost_price,
                    unit_price,
                    total_price,
                    profit
                )
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    receipt_id,
                    timestamp,
                    st.session_state.cashier_name,
                    customer_name,
                    payment_type,
                    item["name"],
                    item["qty"],
                    item["cost"],
                    item["price"],
                    item["total"],
                    item["profit"],
                ),
            )

        conn.commit()

        grand_total = sum(
            item["total"]
            for item in validated_items
        )

        receipt_file = generate_docx_receipt(
            receipt_id=receipt_id,
            cashier=st.session_state.cashier_name,
            customer=customer_name,
            payment_type=payment_type,
            cart_items=validated_items,
            total_amount=grand_total,
            transaction_time=timestamp,
        )

        return (
            True,
            receipt_id,
            receipt_file,
            validated_items,
        )

    except Exception as exc:
        conn.rollback()

        return (
            False,
            str(exc),
            None,
            None,
        )

    finally:
        conn.close()


# ============================================================
# SIDEBAR / AUTHENTICATION
# ============================================================

st.title("🛒 Point of Sale System")

with st.sidebar:
    st.header("Cashier Login")

    if not st.session_state.logged_in:

        cashier_input = st.text_input(
            "Cashier Name",
            placeholder="Enter cashier name",
            key="cashier_login_input",
        )

        if st.button(
            "Login",
            type="primary",
            key="cashier_login_button",
        ):
            cashier_name = cashier_input.strip()

            if cashier_name:
                st.session_state.logged_in = True
                st.session_state.cashier_name = cashier_name
                st.session_state.is_admin = False
                st.session_state.checkout_step = "catalog"
                st.session_state.admin_pin_error = ""
                st.rerun()
            else:
                st.error("Please enter a cashier name.")

    else:
        st.success(
            f"Cashier: **{st.session_state.cashier_name}**"
        )

        st.markdown("---")

        # ----------------------------------------------------
        # Admin access
        # ----------------------------------------------------

        if not st.session_state.is_admin:

            st.subheader("Admin Access")

            admin_pin = st.text_input(
                "Admin PIN",
                type="password",
                key="admin_pin_input",
            )

            if st.button(
                "Unlock Admin",
                key="unlock_admin_button",
            ):
                if admin_pin == ADMIN_PIN:
                    st.session_state.is_admin = True
                    st.session_state.admin_pin_error = ""
                    st.rerun()
                else:
                    st.session_state.admin_pin_error = (
                        "Incorrect Admin PIN."
                    )

            if st.session_state.admin_pin_error:
                st.error(
                    st.session_state.admin_pin_error
                )

        else:
            st.success("🔓 Admin Mode Active")

            if st.button(
                "Lock Admin",
                key="lock_admin_button",
            ):
                st.session_state.is_admin = False
                st.session_state.admin_pin_error = ""
                st.rerun()

        st.markdown("---")

        if st.button(
            "Logout",
            key="logout_button",
        ):
            st.session_state.logged_in = False
            st.session_state.cashier_name = ""
            st.session_state.is_admin = False
            st.session_state.cart = []
            st.session_state.checkout_step = "catalog"
            st.session_state.last_receipt_id = ""
            st.session_state.last_receipt_file = None
            st.session_state.admin_pin_error = ""
            st.rerun()


# ============================================================
# LOGIN GATE
# ============================================================

if not st.session_state.logged_in:
    st.info(
        "Please log in from the sidebar to access the "
        "Point of Sale System."
    )
    st.stop()


# ============================================================
# NAVIGATION
# ============================================================

if st.session_state.is_admin:
    tab_pos, tab_inventory, tab_reports = st.tabs(
        [
            "🛒 Cashier Terminal",
            "📦 Inventory Control",
            "📊 Business Reports",
        ]
    )
else:
    tab_pos = st.container()
    tab_inventory = None
    tab_reports = None


# ============================================================
# CASHIER TERMINAL
# ============================================================

with tab_pos:

    # ========================================================
    # STEP 1 — PRODUCT CATALOG
    # ========================================================

    if st.session_state.checkout_step == "catalog":

        st.subheader(
            "Step 1: Product Selection"
        )

        cart_count = sum(
            safe_int(item.get("qty"), 0)
            for item in st.session_state.cart
        )

        top_left, top_right = st.columns([3, 1])

        with top_left:
            search_term = st.text_input(
                "Search products or services",
                placeholder="Search by name...",
                key="catalog_search",
            )

        with top_right:
            st.metric(
                "Items in Order",
                cart_count,
            )

        st.markdown("---")

        df_products = load_products(
            available_only=True
        )

        if df_products.empty:
            st.info(
                "No products or services are currently "
                "available in inventory."
            )

        else:

            # ------------------------------------------------
            # Defensive column handling
            # ------------------------------------------------

            if "name" not in df_products:
                df_products["name"] = "Unknown"

            if "category" not in df_products:
                df_products["category"] = "General"

            if "price" not in df_products:
                df_products["price"] = 0.0

            if "stock" not in df_products:
                df_products["stock"] = 0

            filtered_df = df_products.copy()

            search_term = search_term.strip()

            if search_term:
                filtered_df = filtered_df[
                    filtered_df["name"]
                    .astype(str)
                    .str.contains(
                        search_term,
                        case=False,
                        na=False,
                    )
                ]

            if filtered_df.empty:
                st.info(
                    "No products match your search."
                )

            else:

                for _, row in filtered_df.iterrows():

                    product_name = str(
                        row.get("name", "Unknown")
                    )

                    category = str(
                        row.get("category", "General")
                    )

                    price = safe_float(
                        row.get("price", 0.0)
                    )

                    stock = safe_int(
                        row.get("stock", 0)
                    )

                    current_cart_qty = get_cart_quantity(
                        safe_int(row.get("id"), 0)
                    )

                    c1, c2, c3, c4 = st.columns(
                        [3, 1.2, 1.2, 1.3]
                    )

                    with c1:
                        st.markdown(
                            f"**{product_name}**"
                        )
                        st.caption(category)

                    with c2:
                        st.write(
                            f"**{money(price)}**"
                        )

                    with c3:
                        st.write(
                            f"Stock: {stock}"
                        )

                        if current_cart_qty:
                            st.caption(
                                f"In order: "
                                f"{current_cart_qty}"
                            )

                    with c4:
                        if st.button(
                            "Add",
                            key=f"add_product_{row['id']}",
                        ):
                            add_product_to_cart(row)
                            st.rerun()

        st.markdown("---")

        # ----------------------------------------------------
        # Cart actions
        # ----------------------------------------------------

        cart_count = sum(
            safe_int(item.get("qty"), 0)
            for item in st.session_state.cart
        )

        if cart_count > 0:

            c1, c2 = st.columns(2)

            with c1:
                if st.button(
                    f"➡️ Continue to Order "
                    f"({cart_count} items)",
                    type="primary",
                    key="continue_to_summary",
                ):
                    st.session_state.checkout_step = "summary"
                    st.rerun()

            with c2:
                if st.button(
                    "🗑️ Clear Order",
                    key="clear_catalog_cart",
                ):
                    clear_cart()
                    st.rerun()

    # ========================================================
    # STEP 2 — ORDER SUMMARY / CHECKOUT
    # ========================================================

    elif st.session_state.checkout_step == "summary":

        st.subheader(
            "Step 2: Order Summary & Checkout"
        )

        if st.button(
            "⬅️ Back to Product Selection",
            key="back_to_catalog",
        ):
            st.session_state.checkout_step = "catalog"
            st.rerun()

        st.markdown("---")

        normalize_cart()

        if not st.session_state.cart:

            st.warning(
                "Your order is empty."
            )

            if st.button(
                "Return to Product Selection",
                key="return_empty_cart",
            ):
                st.session_state.checkout_step = "catalog"
                st.rerun()

        else:

            summary_col, payment_col = st.columns(
                [1.6, 1]
            )

            # ------------------------------------------------
            # Order summary
            # ------------------------------------------------

            with summary_col:

                st.markdown("### Order Items")

                items_to_remove = None

                for index, item in enumerate(
                    st.session_state.cart
                ):

                    c1, c2, c3, c4 = st.columns(
                        [3, 0.8, 1.2, 0.6]
                    )

                    with c1:
                        st.write(
                            f"**{item['name']}**"
                        )
                        st.caption(
                            f"{money(item['price'])} "
                            f"each"
                        )

                    with c2:
                        st.write(
                            f"x{item['qty']}"
                        )

                    with c3:
                        st.write(
                            money(item["total"])
                        )

                    with c4:
                        if st.button(
                            "❌",
                            key=f"remove_cart_{index}",
                        ):
                            items_to_remove = index

                if items_to_remove is not None:
                    remove_cart_item(items_to_remove)
                    st.rerun()

                grand_total = calculate_cart_total()

                st.markdown(
                    f"""
                    <div class="total-box">
                        <h2 style="text-align:right;">
                            Total: {money(grand_total)}
                        </h2>
                    </div>
                    """,
                    unsafe_allow_html=True,
                )

                st.markdown("")

                if st.button(
                    "🗑️ Clear Entire Order",
                    key="clear_summary_cart",
                ):
                    clear_cart()
                    st.rerun()

            # ------------------------------------------------
            # Payment information
            # ------------------------------------------------

            with payment_col:

                st.markdown(
                    "### Customer & Payment"
                )

                customer_name = st.text_input(
                    "Customer Name",
                    value="Walk-in",
                    key="customer_name_checkout",
                )

                payment_type = st.radio(
                    "Payment Status",
                    [
                        "Paid",
                        "Pay Later / Tab",
                    ],
                    key="payment_type_checkout",
                )

                st.markdown("")

                if st.button(
                    "✅ Complete Transaction",
                    type="primary",
                    key="complete_transaction",
                ):

                    with st.spinner(
                        "Processing transaction..."
                    ):
                        success, result, receipt_file, items = (
                            complete_transaction(
                                customer_name,
                                payment_type,
                            )
                        )

                    if success:

                        receipt_id = result

                        st.session_state.last_receipt_id = (
                            receipt_id
                        )

                        st.session_state.last_receipt_file = (
                            receipt_file.getvalue()
                        )

                        st.session_state.cart = []
                        st.session_state.checkout_step = (
                            "catalog"
                        )

                        st.success(
                            f"Transaction completed successfully. "
                            f"Receipt: **{receipt_id}**"
                        )

                        st.download_button(
                            label="📄 Download / Print Receipt",
                            data=receipt_file.getvalue(),
                            file_name=(
                                f"{receipt_id}.docx"
                            ),
                            mime=(
                                "application/vnd.openxmlformats-"
                                "officedocument.wordprocessingml.document"
                            ),
                            key=f"receipt_download_{receipt_id}",
                        )

                        st.info(
                            "The order has been cleared and "
                            "inventory has been updated."
                        )

                    else:
                        st.error(
                            f"Transaction could not be completed: "
                            f"{result}"
                        )


# ============================================================
# INVENTORY CONTROL
# ============================================================

if st.session_state.is_admin and tab_inventory is not None:

    with tab_inventory:

        st.subheader(
            "📦 Inventory Management"
        )

        df_products = load_products(
            available_only=False
        )

        if df_products.empty:
            st.info(
                "No inventory records found."
            )
        else:

            # Defensive display columns
            display_df = df_products.copy()

            expected_columns = [
                "id",
                "name",
                "category",
                "cost_price",
                "price",
                "stock",
            ]

            for column in expected_columns:
                if column not in display_df.columns:
                    if column in (
                        "cost_price",
                        "price",
                    ):
                        display_df[column] = 0.0
                    elif column == "stock":
                        display_df[column] = 0
                    elif column == "category":
                        display_df[column] = "General"
                    else:
                        display_df[column] = ""

            display_df["Margin"] = (
                display_df["price"]
                - display_df["cost_price"]
            )

            display_df["Stock Value"] = (
                display_df["cost_price"]
                * display_df["stock"]
            )

            st.dataframe(
                display_df[
                    [
                        "id",
                        "name",
                        "category",
                        "cost_price",
                        "price",
                        "stock",
                        "Margin",
                        "Stock Value",
                    ]
                ],
                use_container_width=True,
                hide_index=True,
            )

        st.markdown("---")

        # ----------------------------------------------------
        # Inventory summary
        # ----------------------------------------------------

        if not df_products.empty:

            total_items = len(df_products)

            total_units = (
                pd.to_numeric(
                    df_products["stock"]
                    if "stock" in df_products
                    else pd.Series(dtype=float),
                    errors="coerce",
                )
                .fillna(0)
                .sum()
            )

            total_inventory_value = (
                (
                    pd.to_numeric(
                        df_products["cost_price"]
                        if "cost_price" in df_products
                        else pd.Series(dtype=float),
                        errors="coerce",
                    ).fillna(0)
                    *
                    pd.to_numeric(
                        df_products["stock"]
                        if "stock" in df_products
                        else pd.Series(dtype=float),
                        errors="coerce",
                    ).fillna(0)
                )
                .sum()
            )

            low_stock_count = (
                (
                    pd.to_numeric(
                        df_products["stock"]
                        if "stock" in df_products
                        else pd.Series(dtype=float),
                        errors="coerce",
                    ).fillna(0)
                    <= 5
                )
                .sum()
            )

            m1, m2, m3, m4 = st.columns(4)

            m1.metric(
                "Products",
                total_items,
            )

            m2.metric(
                "Total Units",
                int(total_units),
            )

            m3.metric(
                "Inventory Cost Value",
                money(total_inventory_value),
            )

            m4.metric(
                "Low Stock Items",
                int(low_stock_count),
            )

        st.markdown("---")

        # ----------------------------------------------------
        # Add / update product
        # ----------------------------------------------------

        st.subheader(
            "Add or Update Product"
        )

        with st.form(
            "inventory_form",
            clear_on_submit=True,
        ):

            prod_name = st.text_input(
                "Product / Service Name",
                placeholder="Enter item name",
            )

            prod_cat = st.selectbox(
                "Category",
                [
                    "General",
                    "Services",
                    "Supplies",
                    "Other",
                ],
            )

            col1, col2, col3 = st.columns(3)

            with col1:
                prod_cost = st.number_input(
                    "Cost Price",
                    min_value=0.0,
                    value=0.0,
                    step=0.01,
                )

            with col2:
                prod_price = st.number_input(
                    "Selling Price",
                    min_value=0.0,
                    value=0.0,
                    step=0.01,
                )

            with col3:
                prod_stock = st.number_input(
                    "Stock to Add",
                    min_value=0,
                    value=0,
                    step=1,
                )

            submitted = st.form_submit_button(
                "Save Item",
                type="primary",
            )

            if submitted:

                cleaned_name = prod_name.strip()

                if not cleaned_name:
                    st.error(
                        "Product / service name is required."
                    )

                elif prod_price < 0:
                    st.error(
                        "Selling price cannot be negative."
                    )

                elif prod_cost < 0:
                    st.error(
                        "Cost price cannot be negative."
                    )

                else:

                    conn = get_connection()

                    try:
                        cursor = conn.cursor()

                        # Existing product?
                        cursor.execute(
                            """
                            SELECT id
                            FROM products
                            WHERE name = ?
                            COLLATE NOCASE
                            """,
                            (cleaned_name,),
                        )

                        existing = cursor.fetchone()

                        if existing:

                            cursor.execute(
                                """
                                UPDATE products
                                SET
                                    category = ?,
                                    cost_price = ?,
                                    price = ?,
                                    stock = stock + ?
                                WHERE id = ?
                                """,
                                (
                                    prod_cat,
                                    prod_cost,
                                    prod_price,
                                    prod_stock,
                                    existing["id"],
                                ),
                            )

                            message = (
                                f"Updated: {cleaned_name}"
                            )

                        else:

                            cursor.execute(
                                """
                                INSERT INTO products
                                (
                                    name,
                                    category,
                                    cost_price,
                                    price,
                                    stock
                                )
                                VALUES (?, ?, ?, ?, ?)
                                """,
                                (
                                    cleaned_name,
                                    prod_cat,
                                    prod_cost,
                                    prod_price,
                                    prod_stock,
                                ),
                            )

                            message = (
                                f"Added: {cleaned_name}"
                            )

                        conn.commit()
                        st.success(message)

                    except sqlite3.IntegrityError:
                        conn.rollback()
                        st.error(
                            "Unable to save the item because "
                            "another record uses the same name."
                        )

                    except Exception as exc:
                        conn.rollback()
                        st.error(
                            f"Unable to save item: {exc}"
                        )

                    finally:
                        conn.close()

                    st.rerun()

        # ----------------------------------------------------
        # Product editing / deletion
        # ----------------------------------------------------

        st.markdown("---")

        st.subheader(
            "Manage Existing Item"
        )

        if not df_products.empty:

            product_options = {
                f"{row['name']} "
                f"(ID {row['id']})": row["id"]
                for _, row in df_products.iterrows()
            }

            selected_label = st.selectbox(
                "Select item",
                list(product_options.keys()),
                key="selected_inventory_product",
            )

            selected_id = product_options[
                selected_label
            ]

            selected_product = get_product(
                selected_id
            )

            if selected_product:

                with st.form(
                    "edit_inventory_form"
                ):

                    edit_name = st.text_input(
                        "Name",
                        value=str(
                            selected_product["name"]
                        ),
                    )

                    edit_category = st.text_input(
                        "Category",
                        value=str(
                            selected_product["category"]
                            or "General"
                        ),
                    )

                    e1, e2, e3 = st.columns(3)

                    with e1:
                        edit_cost = st.number_input(
                            "Cost Price",
                            min_value=0.0,
                            value=safe_float(
                                selected_product[
                                    "cost_price"
                                ]
                            ),
                            step=0.01,
                        )

                    with e2:
                        edit_price = st.number_input(
                            "Selling Price",
                            min_value=0.0,
                            value=safe_float(
                                selected_product[
                                    "price"
                                ]
                            ),
                            step=0.01,
                        )

                    with e3:
                        edit_stock = st.number_input(
                            "Stock Quantity",
                            min_value=0,
                            value=max(
                                0,
                                safe_int(
                                    selected_product[
                                        "stock"
                                    ]
                                ),
                            ),
                            step=1,
                        )

                    update_button = st.form_submit_button(
                        "Update Item"
                    )

                    if update_button:

                        cleaned_edit_name = (
                            edit_name.strip()
                        )

                        if not cleaned_edit_name:
                            st.error(
                                "Name is required."
                            )

                        else:

                            conn = get_connection()

                            try:
                                cursor = conn.cursor()

                                cursor.execute(
                                    """
                                    UPDATE products
                                    SET
                                        name = ?,
                                        category = ?,
                                        cost_price = ?,
                                        price = ?,
                                        stock = ?
                                    WHERE id = ?
                                    """,
                                    (
                                        cleaned_edit_name,
                                        edit_category.strip()
                                        or "General",
                                        edit_cost,
                                        edit_price,
                                        edit_stock,
                                        selected_id,
                                    ),
                                )

                                conn.commit()

                                st.success(
                                    "Item updated successfully."
                                )

                            except sqlite3.IntegrityError:
                                conn.rollback()
                                st.error(
                                    "Another item already "
                                    "uses that name."
                                )

                            except Exception as exc:
                                conn.rollback()
                                st.error(
                                    f"Update failed: {exc}"
                                )

                            finally:
                                conn.close()

                            st.rerun()

                st.markdown("")

                confirm_delete = st.checkbox(
                    "Enable deletion for this item",
                    key=f"confirm_delete_{selected_id}",
                )

                if confirm_delete:

                    if st.button(
                        "Delete Selected Item",
                        type="secondary",
                        key=f"delete_product_{selected_id}",
                    ):

                        conn = get_connection()

                        try:
                            cursor = conn.cursor()

                            cursor.execute(
                                """
                                DELETE FROM products
                                WHERE id = ?
                                """,
                                (selected_id,),
                            )

                            conn.commit()

                            # Remove deleted product from cart
                            st.session_state.cart = [
                                item
                                for item
                                in st.session_state.cart
                                if item.get("id")
                                != selected_id
                            ]

                            st.success(
                                "Item deleted successfully."
                            )

                        except Exception as exc:
                            conn.rollback()
                            st.error(
                                f"Unable to delete item: {exc}"
                            )

                        finally:
                            conn.close()

                        st.rerun()


# ============================================================
# BUSINESS REPORTS
# ============================================================

if st.session_state.is_admin and tab_reports is not None:

    with tab_reports:

        st.subheader(
            "📊 Business Reports"
        )

        df_sales = load_sales()

        if df_sales.empty:

            st.info(
                "No transaction history is available yet."
            )

        else:

            # ------------------------------------------------
            # Defensive sales columns
            # ------------------------------------------------

            sales_defaults = {
                "receipt_id": "",
                "timestamp": "",
                "cashier": "Unknown",
                "customer_name": "Walk-in",
                "payment_type": "Paid",
                "product_name": "Unknown",
                "quantity": 0,
                "cost_price": 0.0,
                "unit_price": 0.0,
                "total_price": 0.0,
                "profit": 0.0,
            }

            for column, default in sales_defaults.items():
                if column not in df_sales.columns:
                    df_sales[column] = default

            df_sales["total_price"] = pd.to_numeric(
                df_sales["total_price"],
                errors="coerce",
            ).fillna(0.0)

            df_sales["profit"] = pd.to_numeric(
                df_sales["profit"],
                errors="coerce",
            ).fillna(0.0)

            df_sales["quantity"] = pd.to_numeric(
                df_sales["quantity"],
                errors="coerce",
            ).fillna(0)

            # ------------------------------------------------
            # Date range filter
            # ------------------------------------------------

            st.markdown(
                "### Report Filters"
            )

            try:
                df_sales["_date"] = pd.to_datetime(
                    df_sales["timestamp"],
                    errors="coerce",
                ).dt.date
            except Exception:
                df_sales["_date"] = pd.NaT

            valid_dates = df_sales["_date"].dropna()

            if not valid_dates.empty:

                min_date = valid_dates.min()
                max_date = valid_dates.max()

                f1, f2 = st.columns(2)

                with f1:
                    start_date = st.date_input(
                        "From",
                        value=min_date,
                        min_value=min_date,
                        max_value=max_date,
                        key="report_start_date",
                    )

                with f2:
                    end_date = st.date_input(
                        "To",
                        value=max_date,
                        min_value=min_date,
                        max_value=max_date,
                        key="report_end_date",
                    )

                if start_date > end_date:
                    st.error(
                        "The start date cannot be after "
                        "the end date."
                    )

                    filtered_sales = df_sales.iloc[0:0]

                else:
                    filtered_sales = df_sales[
                        (
                            df_sales["_date"]
                            >= start_date
                        )
                        &
                        (
                            df_sales["_date"]
                            <= end_date
                        )
                    ].copy()

            else:
                filtered_sales = df_sales.copy()

            st.markdown("---")

            # ------------------------------------------------
            # Metrics
            # ------------------------------------------------

            total_revenue = (
                filtered_sales["total_price"].sum()
                if "total_price" in filtered_sales
                else 0.0
            )

            total_profit = (
                filtered_sales["profit"].sum()
                if "profit" in filtered_sales
                else 0.0
            )

            unpaid_total = 0.0

            if "payment_type" in filtered_sales:

                unpaid_total = (
                    filtered_sales.loc[
                        filtered_sales["payment_type"]
                        == "Pay Later / Tab",
                        "total_price",
                    ].sum()
                )

            transaction_count = (
                filtered_sales["receipt_id"]
                .nunique()
                if "receipt_id" in filtered_sales
                else 0
            )

            units_sold = (
                filtered_sales["quantity"].sum()
                if "quantity" in filtered_sales
                else 0
            )

            m1, m2, m3, m4, m5 = st.columns(5)

            m1.metric(
                "Gross Revenue",
                money(total_revenue),
            )

            m2.metric(
                "Profit",
                money(total_profit),
            )

            m3.metric(
                "Outstanding Tabs",
                money(unpaid_total),
            )

            m4.metric(
                "Transactions",
                int(transaction_count),
            )

            m5.metric(
                "Units Sold",
                int(units_sold),
            )

            # ------------------------------------------------
            # Sales by product
            # ------------------------------------------------

            st.markdown("---")

            st.subheader(
                "Sales by Product"
            )

            if not filtered_sales.empty:

                product_summary = (
                    filtered_sales
                    .groupby(
                        "product_name",
                        as_index=False,
                    )
                    .agg(
                        Quantity=("quantity", "sum"),
                        Revenue=(
                            "total_price",
                            "sum",
                        ),
                        Profit=(
                            "profit",
                            "sum",
                        ),
                    )
                    .sort_values(
                        "Revenue",
                        ascending=False,
                    )
                )

                st.dataframe(
                    product_summary,
                    use_container_width=True,
                    hide_index=True,
                )

            else:
                st.info(
                    "No sales match the selected date range."
                )

            # ------------------------------------------------
            # Outstanding tabs
            # ------------------------------------------------

            st.markdown("---")

            st.subheader(
                "Outstanding Customer Tabs"
            )

            if "payment_type" in filtered_sales:

                debtors = filtered_sales[
                    filtered_sales["payment_type"]
                    == "Pay Later / Tab"
                ].copy()

                if not debtors.empty:

                    debtor_columns = [
                        "receipt_id",
                        "timestamp",
                        "customer_name",
                        "product_name",
                        "quantity",
                        "total_price",
                    ]

                    debtor_columns = [
                        col
                        for col in debtor_columns
                        if col in debtors.columns
                    ]

                    st.dataframe(
                        debtors[debtor_columns],
                        use_container_width=True,
                        hide_index=True,
                    )

                else:
                    st.info(
                        "No outstanding unpaid tabs."
                    )

            else:
                st.info(
                    "Payment status information is "
                    "not available."
                )

            # ------------------------------------------------
            # Cashier performance
            # ------------------------------------------------

            st.markdown("---")

            st.subheader(
                "Sales by Cashier"
            )

            if not filtered_sales.empty:

                cashier_summary = (
                    filtered_sales
                    .groupby(
                        "cashier",
                        as_index=False,
                    )
                    .agg(
                        Transactions=(
                            "receipt_id",
                            "nunique",
                        ),
                        Units=(
                            "quantity",
                            "sum",
                        ),
                        Revenue=(
                            "total_price",
                            "sum",
                        ),
                        Profit=(
                            "profit",
                            "sum",
                        ),
                    )
                    .sort_values(
                        "Revenue",
                        ascending=False,
                    )
                )

                st.dataframe(
                    cashier_summary,
                    use_container_width=True,
                    hide_index=True,
                )

            # ------------------------------------------------
            # Payment breakdown
            # ------------------------------------------------

            st.markdown("---")

            st.subheader(
                "Payment Status Breakdown"
            )

            if not filtered_sales.empty:

                payment_summary = (
                    filtered_sales
                    .groupby(
                        "payment_type",
                        as_index=False,
                    )
                    .agg(
                        Transactions=(
                            "receipt_id",
                            "nunique",
                        ),
                        Amount=(
                            "total_price",
                            "sum",
                        ),
                    )
                    .sort_values(
                        "Amount",
                        ascending=False,
                    )
                )

                st.dataframe(
                    payment_summary,
                    use_container_width=True,
                    hide_index=True,
                )

            # ------------------------------------------------
            # Full sales history
            # ------------------------------------------------

            st.markdown("---")

            st.subheader(
                "Sales History"
            )

            history_display = filtered_sales.copy()

            if "_date" in history_display.columns:
                history_display = history_display.drop(
                    columns=["_date"]
                )

            if "id" in history_display.columns:
                # Keep the database ID available, but put
                # receipt information first in the display.
                ordered_columns = [
                    "receipt_id",
                    "timestamp",
                    "cashier",
                    "customer_name",
                    "payment_type",
                    "product_name",
                    "quantity",
                    "cost_price",
                    "unit_price",
                    "total_price",
                    "profit",
                ]

                ordered_columns = [
                    col
                    for col in ordered_columns
                    if col in history_display.columns
                ]

                history_display = history_display[
                    ordered_columns
                ]

            st.dataframe(
                history_display,
                use_container_width=True,
                hide_index=True,
            )

            # ------------------------------------------------
            # CSV export
            # ------------------------------------------------

            csv_data = history_display.to_csv(
                index=False
            ).encode("utf-8")

            st.download_button(
                label="📥 Export Report to CSV",
                data=csv_data,
                file_name=(
                    "sales_report_"
                    f"{datetime.now().strftime('%Y_%m_%d')}.csv"
                ),
                mime="text/csv",
                key="export_sales_csv",
            )

            # ------------------------------------------------
            # Refresh
            # ------------------------------------------------

            st.markdown("---")

            if st.button(
                "🔄 Refresh Reports",
                key="refresh_reports",
            ):
                st.rerun()
```
