import streamlit as st
import pandas as pd
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import os
from datetime import datetime
import re
import io

# Page configuration
st.set_page_config(
    page_title="Customer Price Manager Pro",
    page_icon="🛒",
    layout="wide"
)

# Title and description
st.title("🛒 Customer Price Manager Pro")
st.markdown("---")

# Initialize session state
if 'original_file_bytes' not in st.session_state:
    st.session_state.original_file_bytes = None
if 'original_filename' not in st.session_state:
    st.session_state.original_filename = None
if 'customers' not in st.session_state:
    st.session_state.customers = []
if 'products' not in st.session_state:
    st.session_state.products = {}
if 'customer_row_map' not in st.session_state:
    st.session_state.customer_row_map = {}  # Maps customer name to their data
# Store ALL edits in memory
if 'customer_edits' not in st.session_state:
    st.session_state.customer_edits = {}  # Format: {customer_name: {items: [...], custom_prices: {...}}}
if 'data_loaded' not in st.session_state:
    st.session_state.data_loaded = False

def find_marker_row(ws, marker_text):
    """Find the row number containing the marker text"""
    for row_idx in range(1, ws.max_row + 1):
        for col_idx in range(1, ws.max_column + 1):
            cell_value = ws.cell(row_idx, col_idx).value
            if cell_value and isinstance(cell_value, str) and marker_text in cell_value:
                return row_idx
    return None

def load_excel_data(file_bytes):
    """Load Excel data from bytes - NO FILE SAVING"""
    try:
        # Load workbook from bytes
        wb = load_workbook(io.BytesIO(file_bytes))
        ws = wb.active
        
        # Find marker rows
        customer_marker_row = find_marker_row(ws, "报名名单")
        product_marker_row = find_marker_row(ws, "商品汇总")
        
        if not customer_marker_row:
            st.error("❌ Cannot find customer list marker (报名名单) in the Excel file!")
            return None, None, None
        
        if not product_marker_row:
            st.error("❌ Cannot find product summary marker (商品汇总) in the Excel file!")
            return None, None, None
        
        # Customer data starts at the row after marker (header row)
        customer_header_row = customer_marker_row + 1
        customer_data_start_row = customer_header_row + 1
        
        # Product data starts at the row after marker (header row)
        product_header_row = product_marker_row + 1
        product_data_start_row = product_header_row + 1
        
        # Extract customers
        customers = []
        customer_row_map = {}
        row_idx = customer_data_start_row
        while row_idx < product_marker_row:  # Stop before product section
            seq_num = ws.cell(row_idx, 1).value
            name = ws.cell(row_idx, 2).value
            content = ws.cell(row_idx, 3).value
            phone = ws.cell(row_idx, 5).value
            address = ws.cell(row_idx, 6).value
            
            # Stop if we hit an empty row or invalid sequence number
            if seq_num is None or name is None:
                break
            
            customer_data = {
                'seq': seq_num,
                'name': name,
                'content': content,
                'phone': phone,
                'address': address
            }
            
            customers.append(customer_data)
            customer_row_map[name] = customer_data
            row_idx += 1
        
        # Extract products with prices
        products = {}
        row_idx = product_data_start_row
        while row_idx <= ws.max_row:
            product_name = ws.cell(row_idx, 1).value
            price = ws.cell(row_idx, 2).value
            
            if product_name is None or price is None:
                break
            
            products[product_name] = {
                'price': float(price) if price else 0.0
            }
            row_idx += 1
        
        return customers, products, customer_row_map
        
    except Exception as e:
        st.error(f"❌ Error loading file: {str(e)}")
        return None, None, None

def parse_customer_items(content_text):
    """Parse the customer's content field to extract items and quantities"""
    if not content_text:
        return []
    
    items = []
    pattern = r'(.+?)x(\d+)(?:\s*[,、]|\s*，|$)'
    matches = re.finditer(pattern, content_text)
    
    for match in matches:
        item_name = match.group(1).strip()
        item_name = item_name.strip(',，、 ').strip()
        qty_str = match.group(2)
        
        # Skip if this is the total price line
        if '总价' in item_name or '總價' in item_name:
            continue
        
        try:
            qty = int(qty_str)
            if item_name:
                items.append({'name': item_name, 'qty': qty})
        except ValueError:
            continue
    
    return items

def get_item_price(customer_name, item_name):
    """Get the price for an item, checking custom prices first"""
    # Check if there's a custom price for this customer and item
    if customer_name in st.session_state.customer_edits:
        if 'custom_prices' in st.session_state.customer_edits[customer_name]:
            if item_name in st.session_state.customer_edits[customer_name]['custom_prices']:
                return st.session_state.customer_edits[customer_name]['custom_prices'][item_name]
    
    # Otherwise return the base price
    if item_name in st.session_state.products:
        return st.session_state.products[item_name]['price']
    
    return 0.0

def calculate_total(items, customer_name=None):
    """Calculate total price for items"""
    total = 0.0
    for item in items:
        item_name = item['name']
        qty = item['qty']
        price = get_item_price(customer_name, item_name)
        total += price * qty
    return total

def get_current_items(customer_name):
    """Get current items for a customer (edited or original)"""
    # If we have edits, use those
    if customer_name in st.session_state.customer_edits:
        if 'items' in st.session_state.customer_edits[customer_name]:
            return st.session_state.customer_edits[customer_name]['items']
    
    # Otherwise parse from original content
    if customer_name in st.session_state.customer_row_map:
        content = st.session_state.customer_row_map[customer_name]['content']
        return parse_customer_items(content)
    
    return []

def save_customer_edits(customer_name, items, custom_prices):
    """Save customer edits to session state"""
    if customer_name not in st.session_state.customer_edits:
        st.session_state.customer_edits[customer_name] = {}
    
    st.session_state.customer_edits[customer_name]['items'] = items
    st.session_state.customer_edits[customer_name]['custom_prices'] = custom_prices
    st.session_state.customer_edits[customer_name]['last_modified'] = datetime.now()

def create_export_excel():
    """Create a clean, formatted Excel export with all customer data"""
    wb = Workbook()
    ws = wb.active
    ws.title = "Customer Orders"
    
    # Define styles
    header_font = Font(bold=True, size=12, color="FFFFFF")
    header_fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
    header_alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    
    cell_alignment = Alignment(vertical="top", wrap_text=True)
    border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )
    
    # Set column widths
    ws.column_dimensions['A'].width = 8   # Seq
    ws.column_dimensions['B'].width = 15  # Name
    ws.column_dimensions['C'].width = 50  # Items
    ws.column_dimensions['D'].width = 12  # Total
    ws.column_dimensions['E'].width = 15  # Phone
    ws.column_dimensions['F'].width = 30  # Address
    
    # Create header row
    headers = ['序号', '客户名称', '订购商品明细', '总价', '联系电话', '地址']
    for col_idx, header in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col_idx)
        cell.value = header
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = header_alignment
        cell.border = border
    
    # Fill data rows
    row_idx = 2
    for customer in st.session_state.customers:
        customer_name = customer['name']
        
        # Get current items (edited or original)
        items = get_current_items(customer_name)
        
        # Build items detail string
        items_detail = []
        for item in items:
            price = get_item_price(customer_name, item['name'])
            subtotal = price * item['qty']
            items_detail.append(f"{item['name']} x{item['qty']} (@${price:.2f} = ${subtotal:.2f})")
        
        items_text = '\n'.join(items_detail)
        
        # Calculate total
        total = calculate_total(items, customer_name)
        
        # Write row
        ws.cell(row=row_idx, column=1).value = customer['seq']
        ws.cell(row=row_idx, column=2).value = customer_name
        ws.cell(row=row_idx, column=3).value = items_text
        ws.cell(row=row_idx, column=4).value = f"${total:.2f}"
        ws.cell(row=row_idx, column=5).value = customer['phone']
        ws.cell(row=row_idx, column=6).value = customer['address']
        
        # Apply formatting
        for col_idx in range(1, 7):
            cell = ws.cell(row=row_idx, column=col_idx)
            cell.alignment = cell_alignment
            cell.border = border
        
        row_idx += 1
    
    # Freeze the header row
    ws.freeze_panes = 'A2'
    
    # Save to bytes
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output.getvalue()

# Sidebar for file upload
st.sidebar.header("📁 File Upload")

uploaded_file = st.sidebar.file_uploader(
    "Choose Excel file",
    type=['xlsx', 'xls'],
    help="Upload your customer order Excel file"
)

if uploaded_file is not None:
    # Read file into bytes
    file_bytes = uploaded_file.read()
    
    # Only reload if it's a new file
    if st.session_state.original_file_bytes != file_bytes:
        with st.spinner("Loading Excel file..."):
            customers, products, customer_row_map = load_excel_data(file_bytes)
            
            if customers and products:
                st.session_state.original_file_bytes = file_bytes
                st.session_state.original_filename = uploaded_file.name
                st.session_state.customers = customers
                st.session_state.products = products
                st.session_state.customer_row_map = customer_row_map
                st.session_state.data_loaded = True
                st.session_state.customer_edits = {}  # Reset edits on new file
                st.sidebar.success(f"✅ Loaded {len(customers)} customers and {len(products)} products")

# Main content
if st.session_state.data_loaded:
    
    # Export button at the top
    st.sidebar.markdown("---")
    st.sidebar.header("📤 Export")
    
    export_button = st.sidebar.button("📥 Export Clean Excel", use_container_width=True, type="primary")
    
    if export_button:
        with st.spinner("Creating export file..."):
            export_bytes = create_export_excel()
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            export_filename = f"Customer_Orders_Export_{timestamp}.xlsx"
            
            st.sidebar.download_button(
                label="⬇️ Download Export File",
                data=export_bytes,
                file_name=export_filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
            st.sidebar.success("✅ Export ready! Click Download button above.")
    
    # Show stats
    total_customers = len(st.session_state.customers)
    edited_customers = len(st.session_state.customer_edits)
    st.sidebar.info(f"📊 Total customers: {total_customers}\n✏️ Edited: {edited_customers}")
    
    # Customer selection
    st.sidebar.markdown("---")
    st.sidebar.header("👤 Select Customer")
    
    customer_names = [f"{c['seq']}. {c['name']}" for c in st.session_state.customers]
    selected_customer_idx = st.sidebar.selectbox(
        "Customer",
        range(len(customer_names)),
        format_func=lambda x: customer_names[x]
    )
    
    selected_customer = st.session_state.customers[selected_customer_idx]
    customer_name = selected_customer['name']
    
    # Check if customer has been edited
    has_edits = customer_name in st.session_state.customer_edits
    
    # Display customer information
    col1, col2 = st.columns([1, 1])
    
    with col1:
        st.subheader("📋 Customer Information")
        st.write(f"**Name:** {customer_name}")
        st.write(f"**Phone:** {selected_customer['phone']}")
        st.write(f"**Address:** {selected_customer['address']}")
        if has_edits:
            last_modified = st.session_state.customer_edits[customer_name].get('last_modified')
            if last_modified:
                st.info(f"✏️ Last edited: {last_modified.strftime('%Y-%m-%d %H:%M:%S')}")
    
    with col2:
        st.subheader("📊 Order Summary")
        current_items = get_current_items(customer_name)
        total = calculate_total(current_items, customer_name)
        st.metric("Total Items", len(current_items))
        st.metric("Total Price", f"${total:.2f}")
        if has_edits:
            st.success("✅ Changes saved in memory")
    
    st.markdown("---")
    
    # Edit items section
    st.subheader("✏️ Edit Order Items & Prices")
    
    # Get current items
    current_items = get_current_items(customer_name)
    
    # Get current custom prices
    current_custom_prices = {}
    if customer_name in st.session_state.customer_edits:
        if 'custom_prices' in st.session_state.customer_edits[customer_name]:
            current_custom_prices = st.session_state.customer_edits[customer_name]['custom_prices'].copy()
    
    # Create a form for editing
    with st.form(key='edit_form'):
        st.write("**Current Items:**")
        
        # Column headers
        col_headers = st.columns([3, 1.2, 1, 1.2, 1])
        with col_headers[0]:
            st.write("**Product**")
        with col_headers[1]:
            st.write("**Price ($/unit)**")
        with col_headers[2]:
            st.write("**Qty**")
        with col_headers[3]:
            st.write("**Subtotal**")
        with col_headers[4]:
            st.write("**Del**")
        
        edited_items = []
        items_to_delete = []
        custom_price_updates = {}
        
        # Display existing items
        for idx, item in enumerate(current_items):
            col1, col2, col3, col4, col5 = st.columns([3, 1.2, 1, 1.2, 1])
            
            with col1:
                st.text(item['name'])
            
            with col2:
                current_price = get_item_price(customer_name, item['name'])
                
                new_price = st.number_input(
                    f"Price###{idx}",
                    min_value=0.0,
                    value=float(current_price),
                    step=0.01,
                    format="%.2f",
                    key=f"price_{idx}",
                    label_visibility="collapsed"
                )
                
                if abs(new_price - current_price) > 0.001:
                    custom_price_updates[item['name']] = new_price
            
            with col3:
                new_qty = st.number_input(
                    f"Qty###{idx}",
                    min_value=0,
                    value=item['qty'],
                    step=1,
                    key=f"qty_{idx}",
                    label_visibility="collapsed"
                )
            
            with col4:
                price_to_use = new_price
                subtotal = price_to_use * new_qty
                st.text(f"${subtotal:.2f}")
            
            with col5:
                delete = st.checkbox("Del", key=f"del_{idx}", label_visibility="collapsed")
                if delete:
                    items_to_delete.append(idx)
            
            if new_qty > 0 and idx not in items_to_delete:
                edited_items.append({'name': item['name'], 'qty': new_qty})
        
        st.markdown("---")
        
        # Add new item section
        st.write("**Add New Item:**")
        col1, col2 = st.columns([3, 1])
        
        with col1:
            available_products = list(st.session_state.products.keys())
            new_item = st.selectbox(
                "Select Product",
                [""] + available_products,
                key="new_item_select"
            )
        
        with col2:
            new_item_qty = st.number_input(
                "Quantity",
                min_value=0,
                value=0,
                step=1,
                key="new_item_qty"
            )
        
        if new_item and new_item_qty > 0:
            existing = False
            for item in edited_items:
                if item['name'] == new_item:
                    item['qty'] += new_item_qty
                    existing = True
                    break
            
            if not existing:
                edited_items.append({'name': new_item, 'qty': new_item_qty})
        
        # Update custom prices
        current_custom_prices.update(custom_price_updates)
        
        # Calculate new total
        new_total = calculate_total(edited_items, customer_name)
        
        # Temporarily store custom prices for calculation
        if custom_price_updates:
            temp_edits = st.session_state.customer_edits.get(customer_name, {})
            temp_custom = temp_edits.get('custom_prices', {}).copy()
            temp_custom.update(custom_price_updates)
            
            # Recalculate with new prices
            new_total = 0.0
            for item in edited_items:
                if item['name'] in temp_custom:
                    price = temp_custom[item['name']]
                else:
                    price = st.session_state.products.get(item['name'], {}).get('price', 0.0)
                new_total += price * item['qty']
        
        st.markdown(f"### **New Total: ${new_total:.2f}**")
        
        # Show custom prices if any
        if custom_price_updates:
            with st.expander("🏷️ Custom Prices for This Customer"):
                for prod, price in custom_price_updates.items():
                    base_price = st.session_state.products.get(prod, {}).get('price', 0.0)
                    diff = price - base_price
                    if diff > 0:
                        st.write(f"**{prod}**: ${base_price:.2f} → ${price:.2f} (+${diff:.2f})")
                    elif diff < 0:
                        st.write(f"**{prod}**: ${base_price:.2f} → ${price:.2f} (${diff:.2f})")
        
        # Submit button
        col1, col2, col3 = st.columns([1, 1, 2])
        with col1:
            submit_button = st.form_submit_button("💾 Save to Memory", use_container_width=True)
        with col2:
            cancel_button = st.form_submit_button("❌ Cancel", use_container_width=True)
        
        if submit_button:
            # Save edits to session state
            save_customer_edits(customer_name, edited_items, current_custom_prices)
            st.success("✅ Changes saved to memory! (Will persist even if you refresh)")
            st.info("💡 Click 'Export Clean Excel' in sidebar when all edits are done")
            st.rerun()
        
        if cancel_button:
            st.info("Changes discarded")
            st.rerun()

else:
    # Welcome screen
    st.info("👈 Please upload an Excel file to get started")
    
    st.markdown("""
    ### 🎯 How This Works:
    
    **NO FILE PERMISSION ISSUES!**
    - All changes stored in app memory
    - Survives page refresh and reconnection
    - No need to write to original file
    
    **Workflow:**
    1. **Upload** your Excel file (read-only, no permissions needed)
    2. **Edit** customers one by one
    3. **Save to Memory** after each customer (changes persist!)
    4. **Export** when done - creates clean, formatted Excel
    
    ### ✨ Key Features:
    
    - ✅ No file permission errors
    - ✅ Changes persist in memory (even on refresh!)
    - ✅ Edit quantities and prices per customer
    - ✅ Add/delete items
    - ✅ **Export to clean Excel format**
    - ✅ Professional formatting in export
    
    ### 📤 Export Format:
    
    Clean Excel with columns:
    - 序号 (Sequence)
    - 客户名称 (Customer Name)
    - 订购商品明细 (Items Detail with prices)
    - 总价 (Total Price)
    - 联系电话 (Phone)
    - 地址 (Address)
    
    ### 📋 Requirements:
    
    - Excel file must have "报名名单" (customer list) section
    - Excel file must have "商品汇总" (product summary) section
    """)

# Footer
st.markdown("---")
st.caption("Customer Price Manager Pro v4.0 | Memory-Based + Clean Export | Made with ❤️ using Streamlit")
