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

st.set_page_config(page_title="POS & Inventory System", layout="wide")

# Standard Modern Interface Styling
st.markdown("""
    <style>
    /* Global Container Adjustments */
    .block-container {
        padding-top: 1.5rem;
        padding-bottom: 2rem;
    }
    
    /* Category Selector Styling */
    div[data-testid="stHorizontalBlock"] button {
        border-radius: 6px !important;
        font-weight: 600 !important;
        height: 42px !important;
    }
    
    /* Product Card Button Styling */
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
    
    /* Summary Container Cards */
    .cart-summary-box {
        background-color: #f8f9fa;
        border: 1px solid #e9ecef;
        border-radius: 8px;
        padding: 16px;
        margin-top: 12px;
        margin-bottom: 16px;
    }
    </style>
""", unsafe_allow_html=True)

def get_connection():
    return sqlite3.connect("inventory.db")

def init_db():
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS products (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT UNIQUE,
            category TEXT,
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
            total_price REAL,
            payment_method TEXT,
            cashier TEXT,
            timestamp DATETIME DEFAULT CURRENT_TIMESTAMP
        )
    """)
    
    # Auto-migration check for receipt_id
    cursor.execute("PRAGMA table_info(sales)")
    columns = [column[1] for column in cursor.fetchall()]
    if "receipt_id" not in columns:
        cursor.execute("ALTER TABLE sales ADD COLUMN receipt_id TEXT")
        
    conn.commit()
    conn.close()

init_db()

# Helper: Generate Standard Printable MS Word Receipt
def generate_receipt_docx(receipt_id, cashier_name, pay_method, cart_items, total_amount):
    doc = Document()
    
    sections = doc.sections
    for section in sections:
        section.top_margin = Inches(0.5)
        section.bottom_margin = Inches(0.5)
        section.left_margin = Inches(0.5)
        section.right_margin = Inches(0.5)

    # Standard Header
    title_p = doc.add_paragraph()
    title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_title = title_p.add_run("SALES RECEIPT\n")
    run_title.bold = True
    run_title.font.size = Pt(18)
    run_sub = title_p.add_run("Point of Sale System\n")
    run_sub.font.size = Pt(11)
    run_sub.font.color.rgb = RGBColor(120, 120, 120)

    doc.add_paragraph("--------------------------------------------------------------------------------------------------")

    # Transaction Metadata
    meta_p = doc.add_paragraph()
    meta_p.add_run(f"Receipt Ref: ").bold = True
    meta_p.add_run(f"{receipt_id}\n")
    meta_p.add_run(f"Date & Time: ").bold = True
    meta_p.add_run(f"{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
    meta_p.add_run(f"Cashier: ").bold = True
    meta_p.add_run(f"{cashier_name}\n")
    meta_p.add_run(f"Payment Method: ").bold = True
    meta_p.add_run(f"{pay_method}\n")

    # Table of Purchased Items
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

    doc.add_paragraph("--------------------------------------------------------------------------------------------------")

    # Total Section
    total_p = doc.add_paragraph()
    total_p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    run_total = total_p.add_run(f"TOTAL AMOUNT: ${total_amount:.2f}")
    run_total.bold = True
    run_total.font.size = Pt(15)

    # Footer
    footer_p = doc.add_paragraph()
    footer_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_foot = footer_p.add_run("\nThank you for your purchase!")
    run_foot.font.italic = True
    run_foot.font.size = Pt(10)

    target_stream = io.BytesIO()
    doc.save(target_stream)
    return target_stream.getvalue()

# Initialize State Variables
if "cart" not in st.session_state:
    st.session_state.cart = []

if "selected_category" not in st.session_state:
    st.session_state.selected_category = "All"

if "active_cashier" not in st.session_state:
    st.session_state.active_cashier = ""

if "last_receipt" not in st.session_state:
    st.session_state.last_receipt = None

# Sidebar Navigation
st.sidebar.markdown("### POS System")
role = st.sidebar.radio("Navigation View", ["Cashier Terminal", "Stock Inventory", "Admin Dashboard"])

# ---------------------------------------------------------
# VIEW 1: PROFESSIONAL CASHIER TERMINAL
# ---------------------------------------------------------
if role == "Cashier Terminal":
    conn = get_connection()
    df_products = pd.read_sql_query("SELECT * FROM products", conn)
    df_cashiers = pd.read_sql_query("SELECT name FROM cashiers", conn)
    conn.close()
    
    cashier_list = df_cashiers['name'].tolist() if not df_cashiers.empty else []
    
    # Top Action Bar
    top_col1, top_col2 = st.columns([3, 1])
    with top_col1:
        st.title("Cashier Terminal")
    with top_col2:
        if cashier_list:
            st.session_state.active_cashier = st.selectbox(
                "Active Cashier", 
                cashier_list, 
                index=0 if not st.session_state.active_cashier else cashier_list.index(st.session_state.active_cashier) if st.session_state.active_cashier in cashier_list else 0
            )
        else:
            st.session_state.active_cashier = st.text_input("Cashier Name", value="Default Cashier")

    st.divider()

    # Optional Receipt Print Download Prompt
    if st.session_state.last_receipt:
        st.success(f"Transaction completed successfully. Reference: #{st.session_state.last_receipt['id']}")
        rc_col1, rc_col2 = st.columns([2, 1])
        with rc_col1:
            st.download_button(
                label=f"📄 Download MS Word Receipt ({st.session_state.last_receipt['id']})",
                data=st.session_state.last_receipt['docx_data'],
                file_name=f"Receipt_{st.session_state.last_receipt['id']}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True
            )
        with rc_col2:
            if st.button("Dismiss", use_container_width=True):
                st.session_state.last_receipt = None
                st.rerun()

    if df_products.empty:
        st.info("No items available. Add inventory under the Stock Inventory tab.")
    else:
        col_products, col_cart = st.columns([3, 2], gap="large")
        
        # --- LEFT PANEL: CATALOG & SEARCH ---
        with col_products:
            st.markdown("##### Product Catalog")
            categories = ["All", "Patch", "Tube", "Tires", "Car Wash", "Others"]
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

            search_query = st.text_input("Filter Catalog", placeholder="Search by item name...")
            if search_query:
                filtered_df = filtered_df[filtered_df['name'].str.contains(search_query, case=False)]

            st.markdown("---")
            
            grid_cols = st.columns(2)
            for idx, row in filtered_df.iterrows():
                col_idx = idx % 2
                with grid_cols[col_idx]:
                    stock_qty = row['stock']
                    if stock_qty <= 0:
                        stock_status = "Out of Stock"
                    elif stock_qty <= 3:
                        stock_status = f"Low Stock ({stock_qty})"
                    else:
                        stock_status = f"Stock: {stock_qty}"
                        
                    btn_label = f"{row['name']}\n${row['price']:.2f} | {stock_status}"
                    
                    st.markdown('<div class="product-btn">', unsafe_allow_html=True)
                    if st.button(btn_label, key=f"prod_btn_{row['id']}", use_container_width=True):
                        if stock_qty <= 0:
                            st.error(f"{row['name']} is currently out of stock.")
                        else:
                            existing = next((item for item in st.session_state.cart if item['id'] == row['id']), None)
                            if existing:
                                if existing['Quantity'] < stock_qty:
                                    existing['Quantity'] += 1
                                else:
                                    st.warning("Quantity limit reached based on available inventory.")
                            else:
                                st.session_state.cart.append({
                                    "id": row['id'],
                                    "Product Name": row['name'],
                                    "Unit Price ($)": row['price'],
                                    "Quantity": 1,
                                    "max_stock": stock_qty
                                })
                            st.rerun()
                    st.markdown('</div>', unsafe_allow_html=True)

        # --- RIGHT PANEL: ORDER CHECKOUT TERMINAL ---
        with col_cart:
            st.markdown("##### Current Order Summary")
            
            if not st.session_state.cart:
                st.caption("No items added to the cart.")
            else:
                cart_df = pd.DataFrame(st.session_state.cart)
                
                # Render Data Editor for Quantities
                edited_cart = st.data_editor(
                    cart_df[["Product Name", "Unit Price ($)", "Quantity"]],
                    num_rows="dynamic",
                    use_container_width=True,
                    column_config={
                        "Product Name": st.column_config.TextColumn(disabled=True),
                        "Unit Price ($)": st.column_config.NumberColumn(format="$%.2f", disabled=True),
                        "Quantity": st.column_config.NumberColumn(min_value=1, step=1),
                    },
                    key="cart_table_editor"
                )
                
                # Fill NaN values safely and sync quantities back to session state
                edited_cart["Quantity"] = edited_cart["Quantity"].fillna(1).astype(int)
                edited_cart["Quantity"] = edited_cart["Quantity"].apply(lambda x: max(1, x))

                for idx, row in edited_cart.iterrows():
                    st.session_state.cart[idx]["Quantity"] = int(row["Quantity"])

                st.markdown("---")
                st.markdown("###### Item Removal")

                # Multi-select check to remove single or multiple items simultaneously
                selected_for_removal = []
                rem_cols = st.columns(2)
                for idx, item in enumerate(st.session_state.cart):
                    col_idx = idx % 2
                    with rem_cols[col_idx]:
                        if st.checkbox(f"Remove {item['Product Name']}", key=f"rem_chk_{item['id']}"):
                            selected_for_removal.append(item['id'])

                if selected_for_removal:
                    if st.button("🗑️ Remove Selected Items", type="secondary", use_container_width=True):
                        st.session_state.cart = [
                            item for item in st.session_state.cart if item['id'] not in selected_for_removal
                        ]
                        st.rerun()

                # Recalculate Total directly from updated session state cart
                grand_total = sum(item["Quantity"] * item["Unit Price ($)"] for item in st.session_state.cart)
                
                st.markdown("---")
                st.markdown(f"### Total: **${grand_total:.2f}**")
                
                # Checkout Controls
                with st.form("multi_checkout_form"):
                    pay_method = st.selectbox("Payment Method", ["Cash", "Card", "Bank Transfer"])
                    submit_sale = st.form_submit_button("Complete Transaction", type="primary", use_container_width=True)
                    
                    if submit_sale:
                        if not st.session_state.cart:
                            st.error("Cart contains no items.")
                        else:
                            receipt_id = f"REF-{random.randint(100000, 999999)}"
                            conn = get_connection()
                            cursor = conn.cursor()
                            
                            for item in st.session_state.cart:
                                item_total = item['Quantity'] * item['Unit Price ($)']
                                cursor.execute(
                                    """INSERT INTO sales (receipt_id, product_name, quantity, total_price, payment_method, cashier) 
                                       VALUES (?, ?, ?, ?, ?, ?)""",
                                    (receipt_id, item['Product Name'], item['Quantity'], item_total, pay_method, st.session_state.active_cashier)
                                )
                                cursor.execute("UPDATE products SET stock = stock - ? WHERE id = ?", (item['Quantity'], item['id']))
                            
                            conn.commit()
                            conn.close()
                            
                            docx_bytes = generate_receipt_docx(
                                receipt_id, st.session_state.active_cashier, pay_method, st.session_state.cart, grand_total
                            )
                            st.session_state.last_receipt = {
                                "id": receipt_id,
                                "docx_data": docx_bytes
                            }
                            
                            st.session_state.cart = []
                            st.rerun()

                # Cancel Entire Order Button
                if st.button("🚫 Cancel Entire Order", type="secondary", use_container_width=True):
                    st.session_state.cart = []
                    st.rerun()

# ---------------------------------------------------------
# VIEW 2: STOCK INVENTORY MANAGEMENT
# ---------------------------------------------------------
elif role == "Stock Inventory":
    st.title("Stock Inventory Management")
    
    conn = get_connection()
    df_products = pd.read_sql_query("SELECT * FROM products", conn)
    conn.close()
    
    st.markdown("##### Current Stock Levels")
    if not df_products.empty:
        df_display = df_products.copy()
        df_display["Status"] = df_display["stock"].apply(
            lambda x: "Low Stock" if 0 < x <= 3 else ("Out of Stock" if x <= 0 else "In Stock")
        )
        st.dataframe(df_display, use_container_width=True)
    else:
        st.info("No items found in stock database.")
    
    st.divider()
    
    inv_tab1, inv_tab2, inv_tab3 = st.tabs(["Add Product", "Edit Product", "Delete Product"])
    
    with inv_tab1:
        with st.form("product_form"):
            p_name = st.text_input("Product Name")
            p_cat = st.selectbox("Category", ["Patch", "Tube", "Tires", "Car Wash", "Others"])
            p_price = st.number_input("Unit Price ($)", min_value=0.0, format="%.2f")
            p_stock = st.number_input("Stock Quantity", min_value=0, step=1)
            
            save = st.form_submit_button("Save Item", type="primary")
            
            if save and p_name:
                conn = get_connection()
                cursor = conn.cursor()
                cursor.execute("SELECT id FROM products WHERE name = ?", (p_name,))
                existing = cursor.fetchone()
                
                if existing:
                    cursor.execute("UPDATE products SET category = ?, price = ?, stock = stock + ? WHERE name = ?", 
                                   (p_cat, p_price, p_stock, p_name))
                else:
                    cursor.execute("INSERT INTO products (name, category, price, stock) VALUES (?, ?, ?, ?)", 
                                   (p_name, p_cat, p_price, p_stock))
                
                conn.commit()
                conn.close()
                st.success(f"Inventory record for '{p_name}' saved.")
                st.rerun()

    with inv_tab2:
        if not df_products.empty:
            with st.form("edit_product_form"):
                selected_prod = st.selectbox("Select Product", df_products["name"].tolist())
                current_item = df_products[df_products["name"] == selected_prod].iloc[0]
                
                edit_name = st.text_input("Product Name", value=current_item["name"])
                edit_cat = st.selectbox("Category", ["Patch", "Tube", "Tires", "Car Wash", "Others"], 
                                        index=["Patch", "Tube", "Tires", "Car Wash", "Others"].index(current_item["category"]))
                edit_price = st.number_input("Unit Price ($)", min_value=0.0, value=float(current_item["price"]), format="%.2f")
                edit_stock = st.number_input("Exact Stock Count", min_value=0, value=int(current_item["stock"]), step=1)
                
                update_submit = st.form_submit_button("Update Product", type="primary")
                
                if update_submit:
                    conn = get_connection()
                    cursor = conn.cursor()
                    cursor.execute("UPDATE products SET name = ?, category = ?, price = ?, stock = ? WHERE id = ?", 
                                   (edit_name, edit_cat, edit_price, edit_stock, current_item["id"]))
                    conn.commit()
                    conn.close()
                    st.success(f"Updated product details for '{edit_name}'.")
                    st.rerun()

    with inv_tab3:
        if not df_products.empty:
            with st.form("delete_product_form"):
                delete_prod_name = st.selectbox("Select Product to Remove", df_products["name"].tolist())
                delete_submit = st.form_submit_button("Delete Product", type="secondary")
                
                if delete_submit:
                    conn = get_connection()
                    cursor = conn.cursor()
                    cursor.execute("DELETE FROM products WHERE name = ?", (delete_prod_name,))
                    conn.commit()
                    conn.close()
                    st.success(f"Removed '{delete_prod_name}' from inventory database.")
                    st.rerun()

# ---------------------------------------------------------
# VIEW 3: ADMIN DASHBOARD & REPORTS
# ---------------------------------------------------------
elif role == "Admin Dashboard":
    st.title("Admin Dashboard & Analytics")
    
    admin_tab1, admin_tab2 = st.tabs(["Sales Reporting", "Cashier Management"])
    
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
            selected_payment = st.selectbox("Filter Payment Method", ["All", "Cash", "Card", "Bank Transfer"])

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
        
        col1, col2, col3 = st.columns(3)
        daily_revenue = filtered_df['total_price'].sum() if not filtered_df.empty else 0.0
        daily_items = filtered_df['quantity'].sum() if not filtered_df.empty else 0
        top_product = filtered_df.groupby('product_name')['quantity'].sum().idxmax() if not filtered_df.empty else "N/A"
        
        col1.metric("Revenue", f"${daily_revenue:.2f}")
        col2.metric("Items Sold", int(daily_items))
        col3.metric("Top Moving Product", top_product)
        
        st.markdown(f"##### Sales Log — {selected_date}")
        st.dataframe(filtered_df, use_container_width=True)
        
        if not filtered_df.empty:
            excel_buffer = io.BytesIO()
            with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
                filtered_df.to_excel(writer, index=False, sheet_name=f'Sales_{selected_date}')
                
                workbook = writer.book
                worksheet = writer.sheets[f'Sales_{selected_date}']
                for col in worksheet.columns:
                    max_len = max(len(str(cell.value or '')) for cell in col)
                    col_letter = col[0].column_letter
                    worksheet.column_dimensions[col_letter].width = max(max_len + 3, 12)
                    
            excel_data = excel_buffer.getvalue()
            
            st.download_button(
                label=f"📊 Download Excel Sales Report ({selected_date})",
                data=excel_data,
                file_name=f"Sales_Report_{selected_date}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    with admin_tab2:
        st.markdown("##### Add Cashier Profile")
        with st.form("add_cashier_form"):
            new_cashier_name = st.text_input("Full Name")
            add_cashier_btn = st.form_submit_button("Register Cashier", type="primary")
            
            if add_cashier_btn and new_cashier_name:
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
            st.dataframe(df_cashiers, use_container_width=True)
        else:
            st.caption("No registered cashiers.")
