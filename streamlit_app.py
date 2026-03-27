#!/usr/bin/env python3
"""
E-Statement Bank Converter
Konversi e-statement bank (PDF) ke Excel/CSV

Cara deploy ke Streamlit Cloud:
1. Buat akun di https://streamlit.io/
2. Klik "New app"
3. Upload file ini (app.py) atau paste kodenya
4. Klik "Deploy"

Requirements akan otomatis terinstall dari comments di bawah:
"""

# Requirements (Streamlit Cloud akan otomatis install):
# streamlit
# pdfplumber
# pandas
# openpyxl
# xlsxwriter

import streamlit as st
import pdfplumber
import pandas as pd
import re
import io
from datetime import datetime
from typing import Dict, List, Optional, Tuple, Any

# Page config
st.set_page_config(
    page_title="E-Statement Converter",
    page_icon="🏦",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS
st.markdown("""
<style>
    .main-header {
        font-size: 2.5rem;
        font-weight: bold;
        color: #1F4E79;
        text-align: center;
        margin-bottom: 1rem;
    }
    .sub-header {
        font-size: 1.2rem;
        color: #666;
        text-align: center;
        margin-bottom: 2rem;
    }
    .stat-card {
        background: linear-gradient(135deg, #1F4E79 0%, #2E75B6 100%);
        color: white;
        padding: 1rem;
        border-radius: 10px;
        text-align: center;
    }
    .stat-value {
        font-size: 1.5rem;
        font-weight: bold;
    }
    .stat-label {
        font-size: 0.9rem;
        opacity: 0.9;
    }
    div[data-testid="stDataFrame"] {
        overflow-x: auto;
    }
    .stDataFrame table {
        min-width: 100%;
    }
</style>
""", unsafe_allow_html=True)


def clean_balance(balance_str: Any) -> float:
    """Membersihkan string amount menjadi float"""
    if not balance_str:
        return 0.0
    # Remove dots as thousand separators and convert comma to dot for decimal
    cleaned = str(balance_str).replace('.', '').replace(',', '.')
    # Remove any non-numeric characters except dot and minus
    cleaned = re.sub(r'[^\d.\-]', '', cleaned)
    try:
        return float(cleaned)
    except ValueError:
        return 0.0


def extract_account_info(text: str) -> Dict[str, str]:
    """Ekstrak informasi akun dari header PDF"""
    info = {
        'account_name': '',
        'account_number': '',
        'account_type': '',
        'period': '',
        'currency': 'IDR',
        'ledger_balance': 0.0
    }
    
    lines = text.split('\n')
    
    # Extract account name - usually after "ACCOUNT STATEMENT"
    for i, line in enumerate(lines):
        if 'ACCOUNT STATEMENT' in line and i + 1 < len(lines):
            next_line = lines[i + 1]
            name_match = re.match(r'^([A-Z][A-Z\s]+?)\s+Account No', next_line)
            if name_match:
                info['account_name'] = name_match.group(1).strip()
            break
    
    # Extract account number
    acc_match = re.search(r'Account No\.?\s*:\s*([\d]+(?:\s*/\s*[A-Z]+)?)', text)
    if acc_match:
        info['account_number'] = acc_match.group(1).strip()
    
    # Extract account type
    type_match = re.search(r'Account Type\s*:\s*(\w+)', text)
    if type_match:
        info['account_type'] = type_match.group(1)
    
    # Extract period
    period_match = re.search(r'Period\s*:\s*([\d]+-[A-Za-z]+-[\d]+\s*-\s*[\d]+-[A-Za-z]+-[\d]+)', text)
    if period_match:
        info['period'] = period_match.group(1)
    
    # Extract ledger balance
    ledger_match = re.search(r'Ledger Balance:\s*([\d.,]+)', text)
    if ledger_match:
        info['ledger_balance'] = clean_balance(ledger_match.group(1))
    
    return info


def parse_row(row: List, prev_balance: Optional[float], num_cols: int) -> Tuple[Optional[Dict], Optional[float]]:
    """Parse satu baris tabel menjadi transaksi"""

    # Skip header rows
    first_cell = str(row[0]) if row[0] else ''
    if 'Posting Date' in first_cell or 'Ledger Balance' in first_cell or first_cell == '' or first_cell == 'None':
        return None, prev_balance

    # Check for valid date
    if not re.search(r'\d{2}/\d{2}/\d{4}', first_cell):
        return None, prev_balance

    posting_date = str(row[0]).strip()

    # Handle different column structures
    # BNI PDF: Page 1 has different structure than other pages
    if num_cols >= 10:
        # Page 1 structure: Posting Date | Description | Effective Date | Branch | Journal | Description2 | Amount | D/K | ? | Balance
        effective_date = str(row[2]).strip() if len(row) > 2 and row[2] else posting_date
        branch = str(row[3]).replace('\n', ' ').strip() if len(row) > 3 and row[3] else ''
        journal = str(row[4]).strip() if len(row) > 4 and row[4] else ''
        # Combine description from multiple columns if available
        desc1 = str(row[1]).replace('\n', ' ').strip() if len(row) > 1 and row[1] else ''
        desc2 = str(row[5]).replace('\n', ' ').strip() if len(row) > 5 and row[5] else ''
        description = f"{desc1} {desc2}".strip() if desc1 or desc2 else ''
        # D/K can be in column 7 or 6 depending on exact structure
        db_cr_raw = str(row[7]).strip() if len(row) > 7 else ''
        if db_cr_raw.upper() not in ['D', 'K']:
            db_cr_raw = str(row[6]).strip() if len(row) > 6 else ''
        db_cr = db_cr_raw.upper() if db_cr_raw.upper() in ['D', 'K'] else ''
        balance_raw = row[9] if len(row) > 9 else None
    elif num_cols >= 8:
        # Page 2+ structure: 8 columns
        effective_date = str(row[1]).strip() if len(row) > 1 and row[1] else posting_date
        branch = str(row[2]).replace('\n', ' ').strip() if len(row) > 2 and row[2] else ''
        journal = str(row[3]).strip() if len(row) > 3 and row[3] else ''
        description = str(row[4]).replace('\n', ' ').strip() if len(row) > 4 and row[4] else ''
        # D/K is typically in column 6
        db_cr_raw = str(row[6]).strip() if len(row) > 6 else ''
        db_cr = db_cr_raw.upper() if db_cr_raw.upper() in ['D', 'K'] else ''
        balance_raw = row[7] if len(row) > 7 else None
    else:
        # Try to extract from any structure
        effective_date = posting_date
        branch = ''
        journal = ''
        description = ''
        db_cr = ''
        balance_raw = None
        # Search for D/K and balance in the row
        for i, cell in enumerate(row):
            cell_str = str(cell).strip().upper() if cell else ''
            if cell_str in ['D', 'K']:
                db_cr = cell_str
                # Balance is typically the next cell after D/K
                if i + 1 < len(row):
                    balance_raw = row[i + 1]

    # Clean up text
    branch = re.sub(r'\s+', ' ', branch)
    description = re.sub(r'\s+', ' ', description)

    # Parse balance
    current_balance = clean_balance(balance_raw)

    # First transaction establishes prev_balance
    if prev_balance is None:
        if current_balance > 0:
            return None, current_balance
        return None, prev_balance

    # Skip if balance is same as previous (no actual transaction)
    if current_balance == prev_balance:
        return None, prev_balance

    # Calculate amount from balance difference
    if current_balance < prev_balance:
        amount = prev_balance - current_balance
        db_cr = 'D'  # Debit - money out
    else:
        amount = current_balance - prev_balance
        db_cr = 'K'  # Credit - money in

    # Skip if no valid amount
    if amount == 0:
        return None, prev_balance

    trans = {
        'posting_date': posting_date,
        'effective_date': effective_date,
        'branch': branch,
        'journal': journal,
        'description': description,
        'amount': abs(amount),
        'db_cr': db_cr,
        'balance': current_balance
    }

    return trans, current_balance


def extract_transactions(pdf_file, debug_mode=False) -> Tuple[Dict[str, str], List[Dict], List[str]]:
    """Ekstrak semua transaksi dari PDF"""
    all_transactions = []
    account_info = {}
    prev_balance = None
    debug_logs = []

    with pdfplumber.open(pdf_file) as pdf:
        # Get account info from first page
        first_text = pdf.pages[0].extract_text() or ""
        account_info = extract_account_info(first_text)

        # Get initial ledger balance
        ledger_match = re.search(r'Ledger Balance:\s*([\d.,]+)', first_text)
        if ledger_match:
            prev_balance = clean_balance(ledger_match.group(1))
            debug_logs.append(f"Initial Ledger Balance: {prev_balance}")
        else:
            # Try to find any balance at the beginning
            balance_patterns = [
                r'Balance:\s*([\d.,]+)',
                r'Saldo:\s*([\d.,]+)',
                r'([\d.,]+)\s*Cr\.?$'
            ]
            for pattern in balance_patterns:
                match = re.search(pattern, first_text, re.MULTILINE)
                if match:
                    prev_balance = clean_balance(match.group(1))
                    debug_logs.append(f"Found initial balance via pattern: {prev_balance}")
                    break

        # Process all pages
        for page_num, page in enumerate(pdf.pages):
            # Try extracting tables with default settings
            tables = page.extract_tables()

            # If no tables found, try with different settings
            if not tables or all(len(t) == 0 for t in tables):
                tables = page.extract_tables(table_settings={
                    "vertical_strategy": "text",
                    "horizontal_strategy": "text"
                })

            debug_logs.append(f"\n=== Page {page_num + 1} - Found {len(tables)} tables ===")

            for table_idx, table in enumerate(tables):
                if not table or len(table) < 1:
                    continue

                num_cols = len(table[0]) if table[0] else 0
                debug_logs.append(f"Table {table_idx + 1}: {len(table)} rows x {num_cols} cols")

                if debug_mode and len(table) > 0:
                    # Log first few rows
                    for i, row in enumerate(table[:5]):
                        debug_logs.append(f"  Row {i}: {row}")

                for row in table:
                    trans, prev_balance = parse_row(row, prev_balance, num_cols)
                    if trans:
                        all_transactions.append(trans)

    return account_info, all_transactions, debug_logs


def format_currency(amount: float) -> str:
    """Format number to Indonesian Rupiah"""
    return f"Rp {amount:,.2f}"


def main():
    # Header
    st.markdown('<h1 class="main-header">🏦 E-Statement Bank Converter</h1>', unsafe_allow_html=True)
    st.markdown('<p class="sub-header">Konversi e-statement bank (PDF) ke Excel/CSV dengan mudah</p>', unsafe_allow_html=True)
    
    # Sidebar
    st.sidebar.header("⚙️ Pengaturan")

    # Debug mode toggle
    debug_mode = st.sidebar.checkbox("🔧 Debug Mode", value=False,
                                      help="Tampilkan informasi debug untuk troubleshooting")

    # Supported banks info
    st.sidebar.markdown("### Bank yang Didukung")
    st.sidebar.markdown("""
    - ✅ BNI
    - ✅ BCA
    - ✅ Mandiri
    - ✅ BRI
    - ✅ Dan bank lainnya
    """)
    
    # Main content
    col1, col2 = st.columns([2, 1])
    
    with col1:
        st.header("📄 Upload E-Statement")
        
        uploaded_file = st.file_uploader(
            "Pilih file PDF e-statement bank",
            type=['pdf'],
            help="Upload file PDF e-statement dari bank Anda"
        )
    
    with col2:
        st.header("📊 Info")
        st.info("""
        **Cara Penggunaan:**
        1. Upload file PDF e-statement
        2. Klik tombol "Proses"
        3. Lihat hasil dan download
        
        **Fitur:**
        - ✅ Ekstrak transaksi otomatis
        - ✅ Export ke Excel
        - ✅ Export ke CSV
        - ✅ Ringkasan transaksi
        """)
    
    # Process file
    if uploaded_file is not None:
        st.markdown("---")
        
        col_btn1, col_btn2, col_btn3 = st.columns([1, 1, 1])
        
        with col_btn2:
            process_button = st.button("🔄 Proses E-Statement", type="primary", use_container_width=True)
        
        if process_button:
            with st.spinner("Membaca dan mengekstrak data dari PDF..."):
                try:
                    # Extract data
                    account_info, transactions, debug_logs = extract_transactions(uploaded_file, debug_mode)

                    # Store debug logs in session state
                    st.session_state['debug_logs'] = debug_logs

                    if not transactions:
                        st.warning("⚠️ Tidak ada transaksi yang dapat diekstrak dari file ini.")
                        if debug_mode:
                            st.markdown("### 🔧 Debug Logs")
                            with st.expander("Lihat Detail Ekstraksi", expanded=True):
                                for log in debug_logs:
                                    st.text(log)
                    else:
                        # Store in session state
                        st.session_state['account_info'] = account_info
                        st.session_state['transactions'] = transactions
                        st.session_state['processed'] = True

                        st.success(f"✅ Berhasil mengekstrak {len(transactions)} transaksi!")

                        if debug_mode:
                            st.markdown("### 🔧 Debug Logs")
                            with st.expander("Lihat Detail Ekstraksi", expanded=False):
                                for log in debug_logs:
                                    st.text(log)

                except Exception as e:
                    st.error(f"❌ Terjadi kesalahan: {str(e)}")
                    import traceback
                    st.code(traceback.format_exc())
    
    # Display results if processed
    if st.session_state.get('processed', False):
        account_info = st.session_state.get('account_info', {})
        transactions = st.session_state.get('transactions', [])
        
        if transactions:
            # Create DataFrame
            df = pd.DataFrame(transactions)
            
            # Calculate summary
            debit_total = df[df['db_cr'] == 'D']['amount'].sum()
            credit_total = df[df['db_cr'] == 'K']['amount'].sum()
            net_flow = credit_total - debit_total
            
            # Account Info Section
            st.markdown("---")
            st.header("📋 Informasi Rekening")
            
            col_a, col_b, col_c = st.columns(3)
            with col_a:
                st.markdown(f"""
                <div class="stat-card">
                    <div class="stat-label">Nama Rekening</div>
                    <div class="stat-value" style="font-size: 1rem;">{account_info.get('account_name', 'N/A')}</div>
                </div>
                """, unsafe_allow_html=True)
            
            with col_b:
                st.markdown(f"""
                <div class="stat-card">
                    <div class="stat-label">Nomor Rekening</div>
                    <div class="stat-value" style="font-size: 1rem;">{account_info.get('account_number', 'N/A')}</div>
                </div>
                """, unsafe_allow_html=True)
            
            with col_c:
                st.markdown(f"""
                <div class="stat-card">
                    <div class="stat-label">Periode</div>
                    <div class="stat-value" style="font-size: 1rem;">{account_info.get('period', 'N/A')}</div>
                </div>
                """, unsafe_allow_html=True)
            
            # Summary Section
            st.markdown("---")
            st.header("📊 Ringkasan Transaksi")
            
            col_d, col_e, col_f, col_g = st.columns(4)
            
            with col_d:
                st.metric("Total Transaksi", f"{len(transactions):,}")
            
            with col_e:
                st.metric("Total Debit", format_currency(debit_total))
            
            with col_f:
                st.metric("Total Kredit", format_currency(credit_total))
            
            with col_g:
                st.metric("Net Flow", format_currency(net_flow), 
                         delta=f"{'+' if net_flow >= 0 else ''}{format_currency(net_flow)}")
            
            # Tabs
            st.markdown("---")
            tab1, tab2 = st.tabs(["📋 Transaksi", "📈 Analisis"])
            
            with tab1:
                # Filters
                col_filter1, col_filter2 = st.columns([1, 1])
                
                with col_filter1:
                    filter_type = st.selectbox("Filter Transaksi", ["Semua", "Debit Saja", "Kredit Saja"])
                
                with col_filter2:
                    search_term = st.text_input("Cari Deskripsi", placeholder="Ketik kata kunci...")
                
                # Apply filters
                filtered_df = df.copy()
                if filter_type == "Debit Saja":
                    filtered_df = filtered_df[filtered_df['db_cr'] == 'D']
                elif filter_type == "Kredit Saja":
                    filtered_df = filtered_df[filtered_df['db_cr'] == 'K']
                
                if search_term:
                    filtered_df = filtered_df[filtered_df['description'].str.contains(search_term, case=False, na=False)]
                
                # Display data
                st.dataframe(
                    filtered_df,
                    use_container_width=True,
                    hide_index=True,
                    column_config={
                        'posting_date': st.column_config.TextColumn('Posting Date'),
                        'effective_date': st.column_config.TextColumn('Effective Date'),
                        'branch': st.column_config.TextColumn('Branch'),
                        'journal': st.column_config.TextColumn('Journal'),
                        'description': st.column_config.TextColumn('Description', width='large'),
                        'amount': st.column_config.NumberColumn('Amount', format='%,.2f'),
                        'db_cr': st.column_config.TextColumn('DB/CR'),
                        'balance': st.column_config.NumberColumn('Balance', format='%,.2f'),
                    }
                )
                
                # Export Section
                st.markdown("---")
                st.header("📥 Export Data")
                
                col_exp1, col_exp2 = st.columns(2)
                
                with col_exp1:
                    # Excel export
                    excel_buffer = io.BytesIO()
                    with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
                        # Summary sheet
                        summary_data = pd.DataFrame({
                            'Keterangan': ['Nama Rekening', 'Nomor Rekening', 'Tipe Rekening', 'Periode', 
                                          'Total Transaksi', 'Total Debit', 'Total Kredit', 'Net Flow'],
                            'Nilai': [
                                account_info.get('account_name', 'N/A'),
                                account_info.get('account_number', 'N/A'),
                                account_info.get('account_type', 'N/A'),
                                account_info.get('period', 'N/A'),
                                len(transactions),
                                debit_total,
                                credit_total,
                                net_flow
                            ]
                        })
                        summary_data.to_excel(writer, sheet_name='Ringkasan', index=False)
                        
                        # Transactions sheet
                        df.to_excel(writer, sheet_name='Transaksi', index=False)
                    
                    excel_buffer.seek(0)
                    
                    st.download_button(
                        label="📥 Download Excel (.xlsx)",
                        data=excel_buffer,
                        file_name=f"estatement_{account_info.get('account_number', 'export')}_{datetime.now().strftime('%Y%m%d')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )
                
                with col_exp2:
                    # CSV export
                    csv_buffer = io.StringIO()
                    df.to_csv(csv_buffer, index=False)
                    csv_buffer.seek(0)
                    
                    st.download_button(
                        label="📥 Download CSV",
                        data=csv_buffer.getvalue(),
                        file_name=f"estatement_{account_info.get('account_number', 'export')}_{datetime.now().strftime('%Y%m%d')}.csv",
                        mime="text/csv",
                        use_container_width=True
                    )
            
            with tab2:
                # Top transactions
                st.subheader("Top 10 Transaksi Terbesar")
                top_trans = df.nlargest(10, 'amount')[['posting_date', 'description', 'amount', 'db_cr']]
                st.dataframe(top_trans, use_container_width=True, hide_index=True)
                
                # Statistics
                st.subheader("Statistik")
                
                stats_col1, stats_col2 = st.columns(2)
                
                with stats_col1:
                    st.metric("Rata-rata Transaksi", format_currency(df['amount'].mean()))
                    st.metric("Transaksi Terbesar", format_currency(df['amount'].max()))
                
                with stats_col2:
                    st.metric("Transaksi Terkecil", format_currency(df['amount'].min()))
                    st.metric("Jumlah Hari", f"{df['posting_date'].nunique()} hari")
            
            # Reset Button
            st.markdown("---")
            if st.button("🔄 Upload File Baru", use_container_width=True):
                st.session_state['processed'] = False
                st.session_state['transactions'] = []
                st.session_state['account_info'] = {}
                st.rerun()
    
    # Footer
    st.markdown("---")
    st.markdown("""
    <div style="text-align: center; color: #666; padding: 1rem;">
        <p>🏦 E-Statement Converter | Dibuat dengan ❤️ menggunakan Streamlit</p>
        <p style="font-size: 0.8rem;">Mendukung berbagai format e-statement bank Indonesia</p>
    </div>
    """, unsafe_allow_html=True)


if __name__ == "__main__":
    main()
