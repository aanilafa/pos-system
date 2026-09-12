import streamlit as st
import sqlite3
import pandas as pd
from datetime import datetime, date
import io
import random
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT

st.set_page_config(page_title="POS & Inventory System", layout="wide", page_icon="🧾")

# ---------------------------------------------------------
# STYLING
# ---------------------------------------------------------
st.markdown("""
    <style>
    .block-container {
        padding-top: 1.5rem;
        padding-bottom: 2rem;
    }
    div[data-testid="stHorizontalBlock"] button {
        border-radius: 6px !important;
        font-weight: 600 !important;
        height: 42px !important;
    }
    .product-btn > button {
        height: 85px !important;
        white-space: pre-wrap !important;
        border-radius: 8px !important;
        border: 1px solid #e0e0e0 !important;
        box-shadow: 0 2px 4px rgba(0, 0, 0, 0.04) !important;
        transition: all 0.2s ease;
    }
    .product-btn > button:hover {
        border-color: #0066cc !important;
        box-shadow: 0 4px 8px rgba(0, 0, 0, 0.08) !important;
    }
    .product-btn > button p {
        font-size: 15px !important;
        font-weight: 600 !important;
        margin: 0 !important;
    }
    .cart-summary-box {
        background-color: #f8f9fa;
        border: 1px solid #e9ecef;
        border-radius: 8px;
        padding: 16px;
        margin-top: 12px;
        margin-bottom: 16px;
    }
    .cart-line {
        border-bottom: 1px solid #eee;
        padding: 8px 0;
    }
    .stock-pill {
        display: inline-block;
        padding: 2px 8px;
        border-radius: 10px;
        font-size: 12px;
        font-weight: 600;
    }
    </style>
""", unsafe_allow_html=True)

CATEGORIES = ["Patch", "Tube", "Tires", "Car Wash", "Others"]
PAYMENT_METHODS = ["Cash", "Card", "Bank Transfer"]


def get_connection():
    conn = sqlite3.connect("inventory.db")
    conn.execute("PRAGMA foreign_keys = ON")
    return conn


def init_db():
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS products (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT UNIQUE,
            category TEXT,
            cost_price REAL DEFAULT 0,
            price REAL,
            stock INTEGER
        )
    """)
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS cashiers (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT UNIQUE
        )
    """)
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS sales (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            receipt_id TEXT,
            product_name TEXT,
            quantity INTEGER,
            unit_price REAL DEFAULT 0,
            discount_amount REAL DEFAULT 0,
            total_price REAL,
            payment_method TEXT,
            cashier TEXT,
            timestamp DATETIME DEFAULT CURRENT_TIMESTAMP
        )
    """)

    # Auto-migrations for columns added after the tables already existed.
    cursor.execute("PRAGMA table_info(sales)")
    sales_columns = [column[1] for column in cursor.fetchall()]
    if "receipt_id" not in sales_columns:
        cursor.execute("ALTER TABLE sales ADD COLUMN receipt_id TEXT")
    if "discount_amount" not in sales_columns:
        cursor.execute("ALTER TABLE sales ADD COLUMN discount_amount REAL DEFAULT 0")
    if "unit_price" not in sales_columns:
        cursor.execute("ALTER TABLE sales ADD COLUMN unit_price REAL DEFAULT 0")

    cursor.execute("PRAGMA table_info(products)")
    product_columns = [column[1] for column in cursor.fetchall()]
    if "cost_price" not in product_columns:
        cursor.execute("ALTER TABLE products ADD COLUMN cost_price REAL DEFAULT 0")

    conn.commit()
    conn.close()


init_db()


# ---------------------------------------------------------
# RECEIPT GENERATION
# ---------------------------------------------------------
def generate_receipt_docx(receipt_id, cashier_name, pay_method, cart_items, subtotal_amount, discount_amount, total_amount, discount_label=""):
    doc = Document()

    for section in doc.sections:
        section.top_margin = Inches(0.5)
        section.bottom_margin = Inches(0.5)
        section.left_margin = Inches(0.5)
        section.right_margin = Inches(0.5)

    title_p = doc.add_paragraph()
    title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_title = title_p.add_run("SALES RECEIPT\n")
    run_title.bold = True
    run_title.font.size = Pt(18)
    run_sub = title_p.add_run("Point of Sale System\n")
    run_sub.font.size = Pt(11)
    run_sub.font.color.rgb = RGBColor(120, 120, 120)

    doc.add_paragraph("-" * 100)

    meta_p = doc.add_paragraph()
    meta_p.add_run("Receipt Ref: ").bold = True
    meta_p.add_run(f"{receipt_id}\n")
    meta_p.add_run("Date & Time: ").bold = True
    meta_p.add_run(f"{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
    meta_p.add_run("Cashier: ").bold = True
    meta_p.add_run(f"{cashier_name}\n")
    meta_p.add_run("Payment Method: ").bold = True
    meta_p.add_run(f"{pay_method}\n")

    table = doc.add_table(rows=1, cols=4)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    hdr_cells = table.rows[0].cells
    headers = ['Item Description', 'Unit Price', 'Qty', 'Subtotal']
    for i, header_text in enumerate(headers):
        hdr_cells[i].text = header_text
        hdr_cells[i].paragraphs[0].runs[0].font.bold = True

    for item in cart_items:
        row_cells = table.add_row().cells
        subtotal = item['Quantity'] * item['Unit Price ($)']
        row_cells[0].text = str(item['Product Name'])
        row_cells[1].text = f"${item['Unit Price ($)']:.2f}"
        row_cells[2].text = str(item['Quantity'])
        row_cells[3].text = f"${subtotal:.2f}"

    doc.add_paragraph("-" * 100)

    subtotal_p = doc.add_paragraph()
    subtotal_p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    subtotal_p.add_run(f"Subtotal: ${subtotal_amount:.2f}")

    if discount_amount > 0:
        discount_p = doc.add_paragraph()
        discount_p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        label = f"Discount ({discount_label}): " if discount_label else "Discount: "
        run_discount = discount_p.add_run(f"{label}-${discount_amount:.2f}")
        run_discount.font.color.rgb = RGBColor(180, 0, 0)

    total_p = doc.add_paragraph()
    total_p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    run_total = total_p.add_run(f"TOTAL AMOUNT: ${total_amount:.2f}")
    run_total.bold = True
    run_total.font.size = Pt(15)

    footer_p = doc.add_paragraph()
    footer_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_foot = footer_p.add_run("\nThank you for your purchase!")
    run_foot.font.italic = True
    run_foot.font.size = Pt(10)

    target_stream = io.BytesIO()
    doc.save(target_stream)
    return target_stream.getvalue()


# ---------------------------------------------------------
# SESSION STATE
# ---------------------------------------------------------
if "cart" not in st.session_state:
    st.session_state.cart = []
if "selected_category" not in st.session_state:
    st.session_state.selected_category = "All"
if "active_cashier" not in st.session_state:
    st.session_state.active_cashier = ""
if "last_receipt" not in st.session_state:
    st.session_state.last_receipt = None
if "confirm_delete_product" not in st.session_state:
    st.session_state.confirm_delete_product = None

# ---------------------------------------------------------
# SIDEBAR
# ---------------------------------------------------------
st.sidebar.markdown("### 🧾 POS System")
st.sidebar.caption("Manage sales, stock, and reports")
role = st.sidebar.radio(
    "Navigation",
    ["🛒 Cashier Terminal", "📦 Stock Inventory", "📊 Admin Dashboard"],
    label_visibility="collapsed",
)

# ---------------------------------------------------------
# VIEW 1: CASHIER TERMINAL
# ---------------------------------------------------------
if role == "🛒 Cashier Terminal":
    conn = get_connection()
    df_products = pd.read_sql_query("SELECT * FROM products", conn)
    df_cashiers = pd.read_sql_query("SELECT name FROM cashiers", conn)
    conn.close()

    cashier_list = df_cashiers['name'].tolist() if not df_cashiers.empty else []

    top_col1, top_col2 = st.columns([3, 1])
    with top_col1:
        st.title("🛒 Cashier Terminal")
    with top_col2:
        if cashier_list:
            default_idx = (
                cashier_list.index(st.session_state.active_cashier)
                if st.session_state.active_cashier in cashier_list
                else 0
            )
            st.session_state.active_cashier = st.selectbox("Active Cashier", cashier_list, index=default_idx)
        else:
            st.session_state.active_cashier = st.text_input("Cashier Name", value="Default Cashier")

    st.divider()

    if st.session_state.last_receipt:
        st.success(f"✅ Transaction completed. Reference: #{st.session_state.last_receipt['id']}")
        rc_col1, rc_col2 = st.columns([2, 1])
        with rc_col1:
            st.download_button(
                label=f"📄 Download Word Receipt ({st.session_state.last_receipt['id']})",
                data=st.session_state.last_receipt['docx_data'],
                file_name=f"Receipt_{st.session_state.last_receipt['id']}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True,
            )
        with rc_col2:
            if st.button("Dismiss", use_container_width=True):
                st.session_state.last_receipt = None
                st.rerun()

    if df_products.empty:
        st.info("No items available yet. Add inventory under **Stock Inventory**.")
    else:
        col_products, col_cart = st.columns([3, 2], gap="large")

        # --- LEFT PANEL: CATALOG & SEARCH ---
        with col_products:
            st.markdown("##### Product Catalog")
            categories = ["All"] + CATEGORIES
            cat_cols = st.columns(len(categories))
            for idx, cat in enumerate(categories):
                with cat_cols[idx]:
                    btn_type = "primary" if st.session_state.selected_category == cat else "secondary"
                    if st.button(cat, key=f"cat_btn_{cat}", type=btn_type, use_container_width=True):
                        st.session_state.selected_category = cat
                        st.rerun()

            if st.session_state.selected_category != "All":
                filtered_df = df_products[df_products['category'] == st.session_state.selected_category]
            else:
                filtered_df = df_products

            search_query = st.text_input("🔍 Filter Catalog", placeholder="Search by item name...")
            if search_query:
                # regex=False avoids crashes on names containing (, ), +, etc.
                filtered_df = filtered_df[
                    filtered_df['name'].str.contains(search_query, case=False, regex=False)
                ]

            st.markdown("---")

            if filtered_df.empty:
                st.caption("No products match this filter.")
            else:
                grid_cols = st.columns(2)
                # Use a sequential position for column alternation instead of the
                # original (possibly non-sequential) DataFrame index.
                for position, (_, row) in enumerate(filtered_df.iterrows()):
                    col_idx = position % 2
                    with grid_cols[col_idx]:
                        stock_qty = int(row['stock'])
                        if stock_qty <= 0:
                            stock_status = "Out of Stock"
                        elif stock_qty <= 3:
                            stock_status = f"Low Stock ({stock_qty})"
                        else:
                            stock_status = f"Stock: {stock_qty}"

                        btn_label = f"{row['name']}\n${row['price']:.2f} | {stock_status}"

                        st.markdown('<div class="product-btn">', unsafe_allow_html=True)
                        disabled = stock_qty <= 0
                        if st.button(
                            btn_label,
                            key=f"prod_btn_{row['id']}",
                            use_container_width=True,
                            disabled=disabled,
                        ):
                            existing = next(
                                (item for item in st.session_state.cart if item['id'] == row['id']), None
                            )
                            if existing:
                                if existing['Quantity'] < stock_qty:
                                    existing['Quantity'] += 1
                                    st.toast(f"Added another {row['name']}", icon="🛒")
                                else:
                                    st.warning("Quantity limit reached based on available inventory.")
                            else:
                                st.session_state.cart.append({
                                    "id": row['id'],
                                    "Product Name": row['name'],
                                    "Unit Price ($)": row['price'],
                                    "Quantity": 1,
                                    "max_stock": stock_qty,
                                })
                                st.toast(f"{row['name']} added to cart", icon="🛒")
                            st.rerun()
                        st.markdown('</div>', unsafe_allow_html=True)

        # --- RIGHT PANEL: ORDER CHECKOUT TERMINAL ---
        with col_cart:
            st.markdown("##### Current Order Summary")

            if not st.session_state.cart:
                st.caption("No items added to the cart yet — tap a product to get started.")
            else:
                st.markdown('<div class="cart-summary-box">', unsafe_allow_html=True)

                # Simple, safe +/- steppers per line item instead of a freeform
                # data editor (which could desync from st.session_state.cart
                # whenever rows were added/removed inside the editor itself).
                items_to_remove = []
                for item in st.session_state.cart:
                    line_col1, line_col2, line_col3, line_col4 = st.columns([3, 1.3, 1, 0.6])
                    with line_col1:
                        st.markdown(f"**{item['Product Name']}**")
                        st.caption(f"${item['Unit Price ($)']:.2f} each")
                    with line_col2:
                        minus_col, qty_col, plus_col = st.columns([1, 1, 1])
                        with minus_col:
                            if st.button("−", key=f"minus_{item['id']}"):
                                item['Quantity'] = max(1, item['Quantity'] - 1)
                                st.rerun()
                        with qty_col:
                            st.markdown(f"<div style='text-align:center;padding-top:6px'>{item['Quantity']}</div>", unsafe_allow_html=True)
                        with plus_col:
                            if st.button("+", key=f"plus_{item['id']}"):
                                if item['Quantity'] < item['max_stock']:
                                    item['Quantity'] += 1
                                else:
                                    st.warning(f"Only {item['max_stock']} in stock.")
                                st.rerun()
                    with line_col3:
                        st.markdown(
                            f"<div style='text-align:right;padding-top:6px'>${item['Quantity'] * item['Unit Price ($)']:.2f}</div>",
                            unsafe_allow_html=True,
                        )
                    with line_col4:
                        if st.button("🗑️", key=f"rem_{item['id']}"):
                            items_to_remove.append(item['id'])

                if items_to_remove:
                    st.session_state.cart = [
                        item for item in st.session_state.cart if item['id'] not in items_to_remove
                    ]
                    st.rerun()

                st.markdown('</div>', unsafe_allow_html=True)

                subtotal = sum(item["Quantity"] * item["Unit Price ($)"] for item in st.session_state.cart)

                # --- Discount (kept outside the form so the total updates live) ---
                st.markdown("###### Discount")
                disc_col1, disc_col2 = st.columns([1.3, 1])
                with disc_col1:
                    discount_type = st.selectbox(
                        "Discount type",
                        ["None", "Percentage (%)", "Fixed Amount ($)"],
                        key="discount_type",
                        label_visibility="collapsed",
                    )
                with disc_col2:
                    discount_value = 0.0
                    if discount_type != "None":
                        discount_value = st.number_input(
                            "Discount value",
                            min_value=0.0,
                            step=1.0,
                            format="%.2f",
                            key="discount_value",
                            label_visibility="collapsed",
                        )

                if discount_type == "Percentage (%)":
                    discount_value = min(discount_value, 100.0)
                    discount_amount = subtotal * (discount_value / 100.0)
                    discount_label = f"{discount_value:.0f}%"
                elif discount_type == "Fixed Amount ($)":
                    discount_amount = min(discount_value, subtotal)
                    discount_label = "fixed"
                else:
                    discount_amount = 0.0
                    discount_label = ""

                grand_total = max(0.0, subtotal - discount_amount)

                st.markdown(f"Subtotal: ${subtotal:.2f}")
                if discount_amount > 0:
                    st.markdown(f"Discount: −${discount_amount:.2f}")
                st.markdown(f"### Total: **${grand_total:.2f}**")

                with st.form("multi_checkout_form"):
                    pay_method = st.selectbox("Payment Method", PAYMENT_METHODS)
                    submit_sale = st.form_submit_button("✅ Complete Transaction", type="primary", use_container_width=True)

                    if submit_sale:
                        if not st.session_state.cart:
                            st.error("Cart contains no items.")
                        else:
                            # Re-verify current stock right before committing the
                            # sale, in case it changed since items were added.
                            conn = get_connection()
                            cursor = conn.cursor()
                            stock_problem = None
                            for item in st.session_state.cart:
                                cursor.execute("SELECT stock FROM products WHERE id = ?", (item['id'],))
                                row = cursor.fetchone()
                                current_stock = row[0] if row else 0
                                if item['Quantity'] > current_stock:
                                    stock_problem = f"{item['Product Name']} only has {current_stock} left in stock."
                                    break

                            if stock_problem:
                                conn.close()
                                st.error(f"⚠️ {stock_problem} Please adjust the quantity.")
                            else:
                                receipt_id = f"REF-{random.randint(100000, 999999)}"
                                # Spread the discount across line items proportionally
                                # to their share of the subtotal, so per-line and
                                # reporting totals still add up to the discounted total.
                                for item in st.session_state.cart:
                                    item_subtotal = item['Quantity'] * item['Unit Price ($)']
                                    item_discount_share = (
                                        discount_amount * (item_subtotal / subtotal) if subtotal > 0 else 0.0
                                    )
                                    item_total = item_subtotal - item_discount_share
                                    cursor.execute(
                                        """INSERT INTO sales (receipt_id, product_name, quantity, unit_price, discount_amount, total_price, payment_method, cashier)
                                           VALUES (?, ?, ?, ?, ?, ?, ?, ?)""",
                                        (
                                            receipt_id,
                                            item['Product Name'],
                                            item['Quantity'],
                                            item['Unit Price ($)'],
                                            item_discount_share,
                                            item_total,
                                            pay_method,
                                            st.session_state.active_cashier,
                                        )
                                    )
                                    cursor.execute("UPDATE products SET stock = stock - ? WHERE id = ?", (item['Quantity'], item['id']))

                                conn.commit()
                                conn.close()

                                docx_bytes = generate_receipt_docx(
                                    receipt_id, st.session_state.active_cashier, pay_method, st.session_state.cart,
                                    subtotal, discount_amount, grand_total, discount_label,
                                )
                                st.session_state.last_receipt = {"id": receipt_id, "docx_data": docx_bytes}
                                st.session_state.cart = []
                                st.session_state.discount_type = "None"
                                st.rerun()

                if st.button("🚫 Cancel Entire Order", type="secondary", use_container_width=True):
                    st.session_state.cart = []
                    st.rerun()

# ---------------------------------------------------------
# VIEW 2: STOCK INVENTORY MANAGEMENT
# ---------------------------------------------------------
elif role == "📦 Stock Inventory":
    st.title("📦 Stock Inventory Management")

    conn = get_connection()
    df_products = pd.read_sql_query("SELECT * FROM products", conn)
    conn.close()

    st.markdown("##### Current Stock Levels")
    if not df_products.empty:
        df_display = df_products.copy()
        df_display["Status"] = df_display["stock"].apply(
            lambda x: "Low Stock" if 0 < x <= 3 else ("Out of Stock" if x <= 0 else "In Stock")
        )
        df_display = df_display.rename(columns={"cost_price": "Unit Cost", "price": "Selling Price"})
        df_display["Margin"] = df_display["Selling Price"] - df_display["Unit Cost"]
        st.dataframe(
            df_display,
            use_container_width=True,
            hide_index=True,
            column_config={
                "Unit Cost": st.column_config.NumberColumn(format="$%.2f"),
                "Selling Price": st.column_config.NumberColumn(format="$%.2f"),
                "Margin": st.column_config.NumberColumn(format="$%.2f"),
            },
        )
    else:
        st.info("No items found in stock database yet.")

    st.divider()

    inv_tab1, inv_tab2, inv_tab3 = st.tabs(["➕ Add Product", "✏️ Edit Product", "🗑️ Delete Product"])

    with inv_tab1:
        with st.form("product_form"):
            p_name = st.text_input("Product Name")
            p_cat = st.selectbox("Category", CATEGORIES)
            cost_col, price_col = st.columns(2)
            with cost_col:
                p_cost = st.number_input("Unit Cost ($)", min_value=0.0, format="%.2f", help="What you pay to acquire/stock this item.")
            with price_col:
                p_price = st.number_input("Selling Price ($)", min_value=0.0, format="%.2f", help="What the customer pays.")
            if p_price < p_cost:
                st.caption("⚠️ Selling price is below unit cost — this item would sell at a loss.")
            p_stock = st.number_input("Stock Quantity", min_value=0, step=1)
            st.caption("If this name already exists, the quantity entered here will be **added** to existing stock (restock), not replace it.")

            save = st.form_submit_button("Save Item", type="primary")

            if save:
                if not p_name.strip():
                    st.error("Product name cannot be empty.")
                else:
                    conn = get_connection()
                    cursor = conn.cursor()
                    cursor.execute("SELECT id FROM products WHERE name = ?", (p_name.strip(),))
                    existing = cursor.fetchone()

                    if existing:
                        cursor.execute(
                            "UPDATE products SET category = ?, cost_price = ?, price = ?, stock = stock + ? WHERE name = ?",
                            (p_cat, p_cost, p_price, p_stock, p_name.strip()),
                        )
                        conn.commit()
                        conn.close()
                        st.success(f"'{p_name}' already existed — restocked by {p_stock} units.")
                    else:
                        cursor.execute(
                            "INSERT INTO products (name, category, cost_price, price, stock) VALUES (?, ?, ?, ?, ?)",
                            (p_name.strip(), p_cat, p_cost, p_price, p_stock),
                        )
                        conn.commit()
                        conn.close()
                        st.success(f"Inventory record for '{p_name}' created.")
                    st.rerun()

    with inv_tab2:
        if df_products.empty:
            st.caption("No products to edit yet.")
        else:
            with st.form("edit_product_form"):
                selected_prod = st.selectbox("Select Product", df_products["name"].tolist())
                current_item = df_products[df_products["name"] == selected_prod].iloc[0]

                edit_name = st.text_input("Product Name", value=current_item["name"])
                edit_cat = st.selectbox(
                    "Category", CATEGORIES, index=CATEGORIES.index(current_item["category"])
                )
                edit_cost_col, edit_price_col = st.columns(2)
                with edit_cost_col:
                    edit_cost = st.number_input(
                        "Unit Cost ($)", min_value=0.0, value=float(current_item.get("cost_price", 0.0) or 0.0), format="%.2f"
                    )
                with edit_price_col:
                    edit_price = st.number_input(
                        "Selling Price ($)", min_value=0.0, value=float(current_item["price"]), format="%.2f"
                    )
                if edit_price < edit_cost:
                    st.caption("⚠️ Selling price is below unit cost — this item would sell at a loss.")
                edit_stock = st.number_input(
                    "Exact Stock Count", min_value=0, value=int(current_item["stock"]), step=1
                )
                st.caption("This sets the **exact** stock count (overwrite), unlike Add Product which restocks additively.")

                update_submit = st.form_submit_button("Update Product", type="primary")

                if update_submit:
                    if not edit_name.strip():
                        st.error("Product name cannot be empty.")
                    else:
                        conn = get_connection()
                        cursor = conn.cursor()
                        cursor.execute(
                            "UPDATE products SET name = ?, category = ?, cost_price = ?, price = ?, stock = ? WHERE id = ?",
                            (edit_name.strip(), edit_cat, edit_cost, edit_price, edit_stock, int(current_item["id"])),
                        )
                        conn.commit()
                        conn.close()
                        st.success(f"Updated product details for '{edit_name}'.")
                        st.rerun()

    with inv_tab3:
        if df_products.empty:
            st.caption("No products to delete yet.")
        else:
            delete_prod_name = st.selectbox("Select Product to Remove", df_products["name"].tolist(), key="del_select")

            if st.session_state.confirm_delete_product != delete_prod_name:
                if st.button("🗑️ Delete Product", type="secondary"):
                    st.session_state.confirm_delete_product = delete_prod_name
                    st.rerun()
            else:
                st.warning(f"Are you sure you want to permanently delete **{delete_prod_name}**? This cannot be undone.")
                conf_col1, conf_col2 = st.columns(2)
                with conf_col1:
                    if st.button("Yes, delete it", type="primary", use_container_width=True):
                        conn = get_connection()
                        cursor = conn.cursor()
                        cursor.execute("DELETE FROM products WHERE name = ?", (delete_prod_name,))
                        conn.commit()
                        conn.close()
                        st.session_state.confirm_delete_product = None
                        st.success(f"Removed '{delete_prod_name}' from inventory.")
                        st.rerun()
                with conf_col2:
                    if st.button("Cancel", use_container_width=True):
                        st.session_state.confirm_delete_product = None
                        st.rerun()

# ---------------------------------------------------------
# VIEW 3: ADMIN DASHBOARD & REPORTS
# ---------------------------------------------------------
elif role == "📊 Admin Dashboard":
    st.title("📊 Admin Dashboard & Analytics")

    admin_tab1, admin_tab2 = st.tabs(["📈 Sales Reporting", "👤 Cashier Management"])

    with admin_tab1:
        conn = get_connection()
        df_sales = pd.read_sql_query("SELECT * FROM sales ORDER BY timestamp DESC", conn)
        df_cashiers = pd.read_sql_query("SELECT name FROM cashiers", conn)
        conn.close()

        f_col1, f_col2, f_col3 = st.columns(3)
        with f_col1:
            selected_date = st.date_input("Filter Date", date.today())
        with f_col2:
            cashier_options = ["All"] + (df_cashiers["name"].tolist() if not df_cashiers.empty else [])
            selected_cashier = st.selectbox("Filter Cashier", cashier_options)
        with f_col3:
            selected_payment = st.selectbox("Filter Payment Method", ["All"] + PAYMENT_METHODS)

        if not df_sales.empty:
            df_sales['date_only'] = pd.to_datetime(df_sales['timestamp']).dt.date
            filtered_df = df_sales[df_sales['date_only'] == selected_date].copy()

            if selected_cashier != "All":
                filtered_df = filtered_df[filtered_df['cashier'] == selected_cashier]
            if selected_payment != "All":
                filtered_df = filtered_df[filtered_df['payment_method'] == selected_payment]

            filtered_df.drop(columns=['date_only'], inplace=True, errors='ignore')
        else:
            filtered_df = pd.DataFrame()

        col1, col2, col3, col4 = st.columns(4)
        daily_revenue = filtered_df['total_price'].sum() if not filtered_df.empty else 0.0
        daily_items = filtered_df['quantity'].sum() if not filtered_df.empty else 0
        daily_discounts = (
            filtered_df['discount_amount'].sum()
            if not filtered_df.empty and 'discount_amount' in filtered_df.columns
            else 0.0
        )
        top_product = (
            filtered_df.groupby('product_name')['quantity'].sum().idxmax()
            if not filtered_df.empty else "N/A"
        )

        col1.metric("Revenue", f"${daily_revenue:.2f}")
        col2.metric("Items Sold", int(daily_items))
        col3.metric("Discounts Given", f"${daily_discounts:.2f}")
        col4.metric("Top Moving Product", top_product)

        st.markdown(f"##### Sales Log — {selected_date}")
        if filtered_df.empty:
            st.caption("No sales recorded for this filter combination.")
        else:
            st.dataframe(filtered_df, use_container_width=True, hide_index=True)

            excel_buffer = io.BytesIO()
            with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
                sheet_name = f'Sales_{selected_date}'[:31]  # Excel sheet name limit
                filtered_df.to_excel(writer, index=False, sheet_name=sheet_name)

                worksheet = writer.sheets[sheet_name]
                for col in worksheet.columns:
                    max_len = max(len(str(cell.value or '')) for cell in col)
                    col_letter = col[0].column_letter
                    worksheet.column_dimensions[col_letter].width = max(max_len + 3, 12)

            excel_data = excel_buffer.getvalue()

            st.download_button(
                label=f"📊 Download Excel Sales Report ({selected_date})",
                data=excel_data,
                file_name=f"Sales_Report_{selected_date}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

    with admin_tab2:
        st.markdown("##### Add Cashier Profile")
        with st.form("add_cashier_form"):
            new_cashier_name = st.text_input("Full Name")
            add_cashier_btn = st.form_submit_button("Register Cashier", type="primary")

            if add_cashier_btn:
                if not new_cashier_name.strip():
                    st.error("Name cannot be empty.")
                else:
                    try:
                        conn = get_connection()
                        cursor = conn.cursor()
                        cursor.execute("INSERT INTO cashiers (name) VALUES (?)", (new_cashier_name.strip(),))
                        conn.commit()
                        conn.close()
                        st.success(f"Cashier '{new_cashier_name}' registered.")
                        st.rerun()
                    except sqlite3.IntegrityError:
                        st.error("Cashier already exists.")

        st.markdown("##### Registered Cashiers")
        conn = get_connection()
        df_cashiers = pd.read_sql_query("SELECT * FROM cashiers", conn)
        conn.close()

        if not df_cashiers.empty:
            st.dataframe(df_cashiers, use_container_width=True, hide_index=True)
        else:
            st.caption("No registered cashiers yet.")
