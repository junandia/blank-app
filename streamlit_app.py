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


def clean_amount(amount_str: Any) -> float:
    """Membersihkan string amount menjadi float - menangani format yang rusak"""
    if not amount_str:
        return 0.0

    amount_str = str(amount_str)

    # Handle corrupted format like "33,,550000,,000000..0000"
    # This happens when PDF has overlapping text
    # Try to extract the correct number pattern

    # Remove any non-numeric except dots, commas, minus
    cleaned = re.sub(r'[^\d.,\-]', '', amount_str)

    # If there are multiple decimal separators, fix it
    # Pattern like "33,,550000,,000000..0000" means overlapping digits
    # We need to extract unique digits in order

    # Try to find the most likely correct format
    # Split by double separators
    parts = re.split(r'[.,]{2,}', cleaned)

    if len(parts) > 1:
        # Reconstruct: take alternating or deduplicated parts
        # "33,,550000,,000000..0000" -> might be "3,550,000.00"
        # Extract digits only and reconstruct
        all_digits = re.sub(r'[^\d]', '', cleaned)

        # If we have many digits, it's likely a corrupted format
        # Try to parse as Indonesian format: 3.550.000,00
        if len(all_digits) > 10:
            # Take unique consecutive digits
            unique_digits = []
            prev = None
            for d in all_digits:
                if d != prev:
                    unique_digits.append(d)
                    prev = d
            all_digits = ''.join(unique_digits)

        # Try to format as number with 2 decimal places
        if len(all_digits) > 2:
            # Last 2 digits are cents
            try:
                integer_part = all_digits[:-2]
                decimal_part = all_digits[-2:]
                return float(f"{integer_part}.{decimal_part}")
            except:
                pass

    # Standard cleaning for normal format
    # Indonesian format: 1.234.567,89 -> 1234567.89
    cleaned = cleaned.replace('.', '').replace(',', '.')

    try:
        return float(cleaned)
    except ValueError:
        return 0.0


def extract_balance_from_row(row: list, num_cols: int) -> Tuple[Optional[float], int]:
    """Extract balance from row and return with column index"""
    balance = None
    balance_col = -1

    # Find the last numeric value in the row (that's usually the balance)
    for i in range(len(row) - 1, -1, -1):
        val = row[i]
        if val and str(val).strip():
            cleaned = clean_amount(val)
            if cleaned > 0:
                balance = cleaned
                balance_col = i
                break

    return balance, balance_col


def extract_transactions(pdf_file, debug_mode=False) -> Tuple[Dict[str, str], List[Dict], List[str]]:
    """Ekstrak semua transaksi dari PDF"""
    all_transactions = []
    account_info = {
        'account_name': '',
        'account_number': '',
        'account_type': '',
        'period': '',
        'currency': 'IDR',
        'ledger_balance': 0.0
    }
    debug_logs = []
    prev_balance = None

    with pdfplumber.open(pdf_file) as pdf:
        # Get account info from first page text
        first_text = pdf.pages[0].extract_text() or ""

        # Extract account name
        name_match = re.search(r'^([A-Z][A-Z\s]+?)\s+Account No', first_text, re.MULTILINE)
        if name_match:
            account_info['account_name'] = name_match.group(1).strip()

        # Extract account number
        acc_match = re.search(r'Account No\.?\s*:\s*([\d]+(?:\s*/\s*[A-Z]+)?)', first_text)
        if acc_match:
            account_info['account_number'] = acc_match.group(1).strip()

        # Extract period
        period_match = re.search(r'Period\s*:\s*([\d]+-[A-Za-z]+-[\d]+\s*-\s*[\d]+-[A-Za-z]+-[\d]+)', first_text)
        if period_match:
            account_info['period'] = period_match.group(1)

        debug_logs.append(f"Account Info: {account_info}")

        # Process all pages
        for page_num, page in enumerate(pdf.pages):
            tables = page.extract_tables()
            debug_logs.append(f"\n=== Page {page_num + 1} - Found {len(tables)} tables ===")

            for table_idx, table in enumerate(tables):
                if not table or len(table) < 1:
                    continue

                num_cols = len(table[0]) if table[0] else 0
                debug_logs.append(f"Table {table_idx + 1}: {len(table)} rows x {num_cols} cols")

                if debug_mode and len(table) > 0:
                    for i, row in enumerate(table[:5]):
                        debug_logs.append(f"  Row {i}: {row}")

                for row_idx, row in enumerate(table):
                    # Check for Ledger Balance in first row (Page 1)
                    first_cell = str(row[0]) if row[0] else ''

                    if 'Ledger Balance' in first_cell:
                        # Extract balance from this row
                        for cell in row:
                            if cell and str(cell).strip():
                                val = clean_amount(cell)
                                if val > 0:
                                    prev_balance = val
                                    account_info['ledger_balance'] = val
                                    debug_logs.append(f"Found Ledger Balance: {prev_balance}")
                                    break
                        continue

                    # Skip header rows
                    if 'Posting Date' in first_cell or first_cell == '' or first_cell == 'None':
                        continue

                    # Check for valid date
                    if not re.search(r'\d{2}/\d{2}/\d{4}', first_cell):
                        continue

                    # Extract posting date
                    posting_date = str(row[0]).strip()

                    # Handle different column structures
                    if num_cols >= 10:
                        # Page 1 structure: 10 columns
                        # Row: [Posting Date, ?, Effective Date, Branch, Journal, Desc, Amount, D/K, ?, Balance]
                        effective_date = str(row[2]).strip() if len(row) > 2 and row[2] else posting_date
                        branch = str(row[3]).replace('\n', ' ').strip() if len(row) > 3 and row[3] else ''
                        journal = str(row[4]).strip() if len(row) > 4 and row[4] else ''

                        # Description could be in column 1 or 5
                        desc1 = str(row[1]).replace('\n', ' ').strip() if len(row) > 1 and row[1] else ''
                        desc2 = str(row[5]).replace('\n', ' ').strip() if len(row) > 5 and row[5] else ''
                        description = f"{desc1} {desc2}".strip() if desc1 or desc2 else ''

                        # D/K indicator
                        db_cr = ''
                        if len(row) > 7 and str(row[7]).strip().upper() in ['D', 'K']:
                            db_cr = str(row[7]).strip().upper()

                        # Balance is in the last column with numeric value
                        balance, balance_col = extract_balance_from_row(row, num_cols)

                    elif num_cols >= 8:
                        # Page 2+ structure: 8 columns
                        # Row: [Posting Date, Effective Date, Branch, Journal, Desc, Amount, D/K, Balance]
                        effective_date = str(row[1]).strip() if len(row) > 1 and row[1] else posting_date
                        branch = str(row[2]).replace('\n', ' ').strip() if len(row) > 2 and row[2] else ''
                        journal = str(row[3]).strip() if len(row) > 3 and row[3] else ''
                        description = str(row[4]).replace('\n', ' ').strip() if len(row) > 4 and row[4] else ''

                        # D/K indicator
                        db_cr = ''
                        if len(row) > 6 and str(row[6]).strip().upper() in ['D', 'K']:
                            db_cr = str(row[6]).strip().upper()

                        # Balance is in the last column
                        balance, balance_col = extract_balance_from_row(row, num_cols)
                    else:
                        continue

                    # Clean up text
                    branch = re.sub(r'\s+', ' ', branch)
                    description = re.sub(r'\s+', ' ', description)

                    # Skip if no valid balance
                    if balance is None or balance == 0:
                        continue

                    current_balance = balance

                    # First transaction sets prev_balance
                    if prev_balance is None:
                        prev_balance = current_balance
                        debug_logs.append(f"Setting initial prev_balance: {prev_balance}")
                        continue

                    # Skip if balance unchanged
                    if current_balance == prev_balance:
                        continue

                    # Calculate amount from balance difference
                    if current_balance < prev_balance:
                        amount = prev_balance - current_balance
                        db_cr = 'D'  # Debit - money out
                    else:
                        amount = current_balance - prev_balance
                        db_cr = 'K'  # Credit - money in

                    # Skip zero amounts
                    if amount == 0:
                        continue

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

                    all_transactions.append(trans)
                    prev_balance = current_balance

        debug_logs.append(f"\nTotal transactions extracted: {len(all_transactions)}")

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
