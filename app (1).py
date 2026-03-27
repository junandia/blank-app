#!/usr/bin/env python3
"""
E-Statement Bank Converter
Aplikasi Streamlit untuk mengkonversi e-statement bank (PDF) ke Excel/CSV
"""

import streamlit as st
import pdfplumber
import pandas as pd
import re
import io
import tempfile
from datetime import datetime
from typing import List, Dict, Optional, Tuple

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
        font-size: 2rem;
        font-weight: bold;
    }
    .stat-label {
        font-size: 0.9rem;
        opacity: 0.9;
    }
    .success-box {
        background-color: #d4edda;
        border: 1px solid #c3e6cb;
        border-radius: 5px;
        padding: 1rem;
        margin: 1rem 0;
    }
    .info-box {
        background-color: #e7f3ff;
        border: 1px solid #b6d4fe;
        border-radius: 5px;
        padding: 1rem;
        margin: 1rem 0;
    }
</style>
""", unsafe_allow_html=True)


def clean_amount(amount_str: str) -> float:
    """Membersihkan string amount menjadi float"""
    if not amount_str:
        return 0.0
    # Remove dots as thousand separators and convert comma to dot for decimal
    cleaned = str(amount_str).replace('.', '').replace(',', '.')
    # Remove any non-numeric characters except dot and minus
    cleaned = re.sub(r'[^\d.\-]', '', cleaned)
    try:
        return float(cleaned)
    except ValueError:
        return 0.0


def parse_date(date_str: str) -> Optional[str]:
    """Parse tanggal dari berbagai format"""
    if not date_str:
        return None
    
    # Try different date formats
    formats = [
        '%d/%m/%Y',
        '%d-%m-%Y',
        '%Y-%m-%d',
        '%d/%m/%y',
    ]
    
    for fmt in formats:
        try:
            parsed = datetime.strptime(date_str.strip(), fmt)
            return parsed.strftime('%Y-%m-%d')
        except ValueError:
            continue
    return date_str


def extract_account_info(text: str) -> Dict:
    """Ekstrak informasi akun dari header PDF"""
    info = {
        'account_name': '',
        'account_number': '',
        'account_type': '',
        'period': '',
        'currency': 'IDR'
    }
    
    # Extract account name
    name_match = re.search(r'^([A-Z][A-Z\s]+?)\s+Account No', text, re.MULTILINE)
    if name_match:
        info['account_name'] = name_match.group(1).strip()
    
    # Extract account number
    acc_match = re.search(r'Account No\.?\s*:\s*([\d\s/]+(?:\([A-Z]+\))?)', text)
    if acc_match:
        info['account_number'] = acc_match.group(1).strip()
    
    # Extract account type
    type_match = re.search(r'Account Type\s*:\s*(\w+)', text)
    if type_match:
        info['account_type'] = type_match.group(1)
    
    # Extract period
    period_match = re.search(r'Period\s*:\s*([\d\-]+\s*-\s*[\d\-]+)', text)
    if period_match:
        info['period'] = period_match.group(1)
    
    return info


def extract_transactions_from_page(page) -> List[Dict]:
    """Ekstrak transaksi dari satu halaman PDF"""
    transactions = []
    
    # Get text for context
    text = page.extract_text() or ""
    
    # Extract tables
    tables = page.extract_tables()
    
    for table in tables:
        for row in table:
            if not row or len(row) < 5:
                continue
            
            # Skip header rows
            row_text = ' '.join([str(cell) for cell in row if cell])
            if any(h in row_text.lower() for h in ['posting date', 'effective date', 'balance', 'branch', 'journal']):
                continue
            
            # Try to identify transaction rows
            # Format: [Posting Date, Effective Date, Branch, Journal, Description, Amount, DB/CR, Balance]
            try:
                # Find date pattern
                date_pattern = r'(\d{2}/\d{2}/\d{4})'
                dates = re.findall(date_pattern, row_text)
                
                if len(dates) >= 1:
                    posting_date = dates[0] if len(dates) > 0 else None
                    effective_date = dates[1] if len(dates) > 1 else posting_date
                    
                    # Find amount pattern
                    amount_pattern = r'[\d.,]+\d{2}'
                    amounts = re.findall(amount_pattern, row_text)
                    
                    # Find DB/CR
                    db_cr = 'D' if ' D ' in row_text or row_text.endswith(' D') else 'K' if ' K ' in row_text or row_text.endswith(' K') else None
                    
                    # Build transaction
                    trans = {
                        'posting_date': parse_date(posting_date) if posting_date else None,
                        'effective_date': parse_date(effective_date) if effective_date else None,
                        'description': row_text[:200] if row_text else '',
                        'amount': clean_amount(amounts[0]) if amounts else 0,
                        'db_cr': db_cr,
                        'balance': clean_amount(amounts[-1]) if len(amounts) > 1 else 0
                    }
                    
                    if trans['posting_date'] and trans['amount'] > 0:
                        transactions.append(trans)
            except Exception:
                continue
    
    return transactions


def extract_transactions_improved(pdf_path: str) -> Tuple[Dict, List[Dict]]:
    """
    Ekstrak transaksi dengan metode yang lebih robust
    """
    all_transactions = []
    account_info = {}
    
    with pdfplumber.open(pdf_path) as pdf:
        # First page - get account info
        first_page = pdf.pages[0]
        first_text = first_page.extract_text() or ""
        account_info = extract_account_info(first_text)
        
        # Current balance tracker
        running_balance = None
        
        for page_num, page in enumerate(pdf.pages):
            text = page.extract_text() or ""
            
            # Skip empty pages
            if not text.strip():
                continue
            
            # Parse transactions line by line
            lines = text.split('\n')
            
            for line in lines:
                # Pattern for transaction line with date
                date_match = re.match(r'(\d{2}/\d{2}/\d{4})\s+(\d{2}:\d{2}:\d{2})?\s*(\d{2}/\d{2}/\d{4})?\s*(\d{2}:\d{2}:\d{2})?', line)
                
                if date_match:
                    posting_date = date_match.group(1)
                    effective_date = date_match.group(3) if date_match.group(3) else posting_date
                    
                    # Find amount and balance
                    # Pattern: amount followed by D or K, then balance
                    amount_pattern = r'([\d.,]+)\s+([DK])\s+([\d.,]+)'
                    amount_match = re.search(amount_pattern, line)
                    
                    if amount_match:
                        amount = clean_amount(amount_match.group(1))
                        db_cr = amount_match.group(2)
                        balance = clean_amount(amount_match.group(3))
                        
                        # Extract description (between date and amount)
                        desc_start = date_match.end()
                        desc_end = amount_match.start()
                        description = line[desc_start:desc_end].strip()
                        
                        trans = {
                            'posting_date': parse_date(posting_date),
                            'effective_date': parse_date(effective_date),
                            'description': description,
                            'amount': amount,
                            'db_cr': db_cr,
                            'balance': balance,
                            'page': page_num + 1
                        }
                        
                        all_transactions.append(trans)
    
    return account_info, all_transactions


def extract_transactions_from_text(pdf_path: str) -> Tuple[Dict, List[Dict]]:
    """
    Ekstrak transaksi dengan parsing text yang lebih akurat
    """
    all_transactions = []
    account_info = {}
    
    with pdfplumber.open(pdf_path) as pdf:
        # First page - get account info
        first_page = pdf.pages[0]
        first_text = first_page.extract_text() or ""
        account_info = extract_account_info(first_text)
        
        # Extract ledger balance from first page
        ledger_match = re.search(r'Ledger Balance:\s*([\d.,]+)', first_text)
        if ledger_match:
            account_info['ledger_balance'] = clean_amount(ledger_match.group(1))
        
        full_text = ""
        
        # Combine all pages text
        for page_num, page in enumerate(pdf.pages):
            text = page.extract_text() or ""
            full_text += f"\n--- PAGE {page_num + 1} ---\n{text}\n"
        
        # Pattern for transaction lines
        # Format: DD/MM/YYYY HH.MM.SS DD/MM/YYYY HH.MM.SS BRANCH JOURNAL DESC... AMOUNT D/K BALANCE
        trans_pattern = r'''
            (\d{2}/\d{2}/\d{4})\s+      # Posting date
            (\d{2}[:.]\d{2}[:.]\d{2})?\s*  # Posting time (optional)
            (\d{2}/\d{2}/\d{4})?\s*     # Effective date (optional)
            (\d{2}[:.]\d{2}[:.]\d{2})?\s*  # Effective time (optional)
            (\d{6})?\s*                 # Branch code (optional)
            (.+?)\s+                    # Description
            ([\d.,]+)\s+                # Amount
            ([DK])\s+                   # D/C indicator
            ([\d.,]+)                   # Balance
        '''
        
        # Simpler pattern
        simple_pattern = r'(\d{2}/\d{2}/\d{4})[^\d]*?(\d{2}/\d{2}/\d{4})?[^\d]*?([\d.,]+)\s+([DK])\s+([\d.,]+)'
        
        lines = full_text.split('\n')
        
        for line in lines:
            line = line.strip()
            
            # Skip headers and empty lines
            if not line or 'Posting Date' in line or 'Account Statement' in line:
                continue
            
            # Try to match transaction pattern
            match = re.search(simple_pattern, line)
            
            if match:
                posting_date = match.group(1)
                effective_date = match.group(2) if match.group(2) else posting_date
                amount = clean_amount(match.group(3))
                db_cr = match.group(4)
                balance = clean_amount(match.group(5))
                
                # Get description
                desc_start = line.find(posting_date) + len(posting_date)
                desc_end = line.rfind(match.group(3))
                description = line[desc_start:desc_end].strip() if desc_end > desc_start else ""
                
                # Clean description
                description = re.sub(r'\s+', ' ', description)
                description = re.sub(r'^[\d\s]+', '', description).strip()
                
                trans = {
                    'posting_date': parse_date(posting_date),
                    'effective_date': parse_date(effective_date),
                    'description': description[:300] if description else '',
                    'amount': amount,
                    'db_cr': db_cr,
                    'balance': balance
                }
                
                all_transactions.append(trans)
    
    return account_info, all_transactions


def extract_with_table_method(pdf_path: str) -> Tuple[Dict, List[Dict]]:
    """
    Ekstrak menggunakan metode tabel pdfplumber
    """
    all_transactions = []
    account_info = {}
    
    with pdfplumber.open(pdf_path) as pdf:
        # Get account info from first page
        first_text = pdf.pages[0].extract_text() or ""
        account_info = extract_account_info(first_text)
        
        # Extract ledger balance
        ledger_match = re.search(r'Ledger Balance:\s*([\d.,]+)', first_text)
        if ledger_match:
            account_info['ledger_balance'] = clean_amount(ledger_match.group(1))
        
        for page_num, page in enumerate(pdf.pages):
            tables = page.extract_tables()
            
            for table in tables:
                if not table or len(table) < 2:
                    continue
                
                # Find header row
                header_idx = -1
                for i, row in enumerate(table):
                    row_text = ' '.join([str(c).lower() if c else '' for c in row])
                    if 'posting' in row_text or 'date' in row_text:
                        header_idx = i
                        break
                
                # Process data rows
                start_idx = header_idx + 1 if header_idx >= 0 else 0
                
                for row in table[start_idx:]:
                    if not row or all(not c for c in row):
                        continue
                    
                    # Try to extract transaction data
                    trans = parse_table_row(row)
                    if trans:
                        trans['page'] = page_num + 1
                        all_transactions.append(trans)
    
    return account_info, all_transactions


def parse_table_row(row: list) -> Optional[Dict]:
    """Parse satu baris tabel menjadi transaksi"""
    try:
        # Clean row
        cells = [str(c).strip() if c else '' for c in row]
        
        # Find dates in row
        dates = []
        for cell in cells:
            date_match = re.search(r'(\d{2}/\d{2}/\d{4})', cell)
            if date_match:
                dates.append(date_match.group(1))
        
        if len(dates) < 1:
            return None
        
        posting_date = dates[0] if len(dates) > 0 else None
        effective_date = dates[1] if len(dates) > 1 else posting_date
        
        # Find amount and balance
        amounts = []
        for cell in cells:
            # Amount pattern: numbers with dots and commas
            amount_match = re.findall(r'[\d]{1,3}(?:[.,][\d]{3})*[.,][\d]{2}', cell)
            amounts.extend(amount_match)
        
        # Find D/C indicator
        db_cr = None
        for cell in cells:
            if cell.strip() in ['D', 'K']:
                db_cr = cell.strip()
                break
        
        # Build description
        desc_parts = []
        for cell in cells:
            if cell and not re.match(r'^[\d/:\.\s]+$', cell) and cell not in ['D', 'K']:
                if not re.match(r'^[\d.,]+$', cell):
                    desc_parts.append(cell)
        
        description = ' '.join(desc_parts)[:300]
        
        # Determine amount and balance
        if len(amounts) >= 2:
            amount = clean_amount(amounts[-2])
            balance = clean_amount(amounts[-1])
        elif len(amounts) == 1:
            amount = clean_amount(amounts[0])
            balance = 0
        else:
            return None
        
        if amount > 0 and posting_date:
            return {
                'posting_date': parse_date(posting_date),
                'effective_date': parse_date(effective_date),
                'description': description,
                'amount': amount,
                'db_cr': db_cr or 'D',
                'balance': balance
            }
    except Exception:
        pass
    
    return None


def main():
    # Header
    st.markdown('<h1 class="main-header">🏦 E-Statement Bank Converter</h1>', unsafe_allow_html=True)
    st.markdown('<p class="sub-header">Konversi e-statement bank (PDF) ke Excel/CSV dengan mudah</p>', unsafe_allow_html=True)
    
    # Sidebar
    st.sidebar.header("⚙️ Pengaturan")
    
    extraction_method = st.sidebar.selectbox(
        "Metode Ekstraksi",
        ["Otomatis (Rekomendasi)", "Text Parsing", "Table Extraction"],
        help="Pilih metode ekstraksi yang sesuai dengan format PDF Anda"
    )
    
    # Main content
    col1, col2 = st.columns([2, 1])
    
    with col1:
        st.header("📄 Upload E-Statement")
        
        uploaded_file = st.file_uploader(
            "Pilih file PDF e-statement bank",
            type=['pdf'],
            help="Upload file PDF e-statement dari bank Anda"
        )
        
        # Sample file option
        use_sample = st.checkbox("Gunakan file contoh yang sudah diupload", value=False)
        
        if use_sample:
            sample_path = "/home/z/my-project/upload/Account_Statement_Download_Single_20240801070721973.pdf"
            try:
                with open(sample_path, 'rb') as f:
                    uploaded_file = io.BytesIO(f.read())
                    uploaded_file.name = "Account_Statement_Download_Single_20240801070721973.pdf"
            except FileNotFoundError:
                st.error("File contoh tidak ditemukan")
                uploaded_file = None
    
    with col2:
        st.header("📊 Info")
        st.info("""
        **Format yang didukung:**
        - Bank BNI E-Statement
        - Bank BCA E-Statement
        - Bank Mandiri E-Statement
        - Dan format bank lainnya
        
        **Fitur:**
        - ✅ Ekstrak transaksi otomatis
        - ✅ Export ke Excel
        - ✅ Export ke CSV
        - ✅ Ringkasan transaksi
        """)
    
    # Process file
    if uploaded_file is not None:
        st.markdown("---")
        st.header("🔄 Memproses File...")
        
        with st.spinner("Membaca dan mengekstrak data dari PDF..."):
            try:
                # Save uploaded file temporarily
                with tempfile.NamedTemporaryFile(delete=False, suffix='.pdf') as tmp_file:
                    tmp_file.write(uploaded_file.getvalue())
                    tmp_path = tmp_file.name
                
                # Extract data based on method
                method_map = {
                    "Otomatis (Rekomendasi)": extract_transactions_improved,
                    "Text Parsing": extract_transactions_from_text,
                    "Table Extraction": extract_with_table_method
                }
                
                extract_func = method_map[extraction_method]
                account_info, transactions = extract_func(tmp_path)
                
                # Clean up temp file
                import os
                os.unlink(tmp_path)
                
                if not transactions:
                    st.warning("⚠️ Tidak ada transaksi yang dapat diekstrak. Coba metode ekstraksi lain.")
                else:
                    # Success message
                    st.success(f"✅ Berhasil mengekstrak {len(transactions)} transaksi!")
                    
                    # Account Info Section
                    st.markdown("---")
                    st.header("📋 Informasi Rekening")
                    
                    col_a, col_b, col_c = st.columns(3)
                    with col_a:
                        st.markdown(f"""
                        <div class="stat-card">
                            <div class="stat-label">Nama Rekening</div>
                            <div class="stat-value" style="font-size: 1.2rem;">{account_info.get('account_name', 'N/A')}</div>
                        </div>
                        """, unsafe_allow_html=True)
                    
                    with col_b:
                        st.markdown(f"""
                        <div class="stat-card">
                            <div class="stat-label">Nomor Rekening</div>
                            <div class="stat-value" style="font-size: 1.2rem;">{account_info.get('account_number', 'N/A')}</div>
                        </div>
                        """, unsafe_allow_html=True)
                    
                    with col_c:
                        st.markdown(f"""
                        <div class="stat-card">
                            <div class="stat-label">Periode</div>
                            <div class="stat-value" style="font-size: 1.2rem;">{account_info.get('period', 'N/A')}</div>
                        </div>
                        """, unsafe_allow_html=True)
                    
                    # Create DataFrame
                    df = pd.DataFrame(transactions)
                    
                    # Calculate summary
                    debit_total = df[df['db_cr'] == 'D']['amount'].sum()
                    credit_total = df[df['db_cr'] == 'K']['amount'].sum()
                    
                    # Summary Section
                    st.markdown("---")
                    st.header("📊 Ringkasan Transaksi")
                    
                    col_d, col_e, col_f, col_g = st.columns(4)
                    
                    with col_d:
                        st.metric("Total Transaksi", f"{len(transactions):,}")
                    
                    with col_e:
                        st.metric("Total Debit", f"Rp {debit_total:,.2f}")
                    
                    with col_f:
                        st.metric("Total Kredit", f"Rp {credit_total:,.2f}")
                    
                    with col_g:
                        net_flow = credit_total - debit_total
                        st.metric("Net Flow", f"Rp {net_flow:,.2f}", 
                                 delta=f"{'+' if net_flow >= 0 else ''}{net_flow:,.2f}")
                    
                    # Data Preview
                    st.markdown("---")
                    st.header("👀 Preview Data")
                    
                    # Column config
                    column_config = {
                        'posting_date': st.column_config.DateColumn('Tanggal Posting', format='YYYY-MM-DD'),
                        'effective_date': st.column_config.DateColumn('Tanggal Efektif', format='YYYY-MM-DD'),
                        'description': st.column_config.TextColumn('Deskripsi', width='large'),
                        'amount': st.column_config.NumberColumn('Jumlah', format='%,.2f'),
                        'db_cr': st.column_config.TextColumn('D/K'),
                        'balance': st.column_config.NumberColumn('Saldo', format='%,.2f'),
                        'page': st.column_config.NumberColumn('Halaman')
                    }
                    
                    # Filter options
                    col_filter1, col_filter2 = st.columns(2)
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
                        column_config=column_config,
                        use_container_width=True,
                        hide_index=True,
                        height=400
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
                            summary_data = {
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
                                    credit_total - debit_total
                                ]
                            }
                            summary_df = pd.DataFrame(summary_data)
                            summary_df.to_excel(writer, sheet_name='Ringkasan', index=False)
                            
                            # Transactions sheet
                            df.to_excel(writer, sheet_name='Transaksi', index=False)
                        
                        excel_buffer.seek(0)
                        
                        st.download_button(
                            label="📥 Download Excel (.xlsx)",
                            data=excel_buffer,
                            file_name=f"estatement_{account_info.get('account_number', 'export')}_{datetime.now().strftime('%Y%m%d')}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
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
                            mime="text/csv"
                        )
                    
                    # Statistics
                    st.markdown("---")
                    st.header("📈 Statistik")
                    
                    # Daily transaction count
                    daily_counts = df.groupby('posting_date').size().reset_index(name='count')
                    daily_counts = daily_counts.sort_values('posting_date')
                    
                    st.subheader("Jumlah Transaksi per Hari")
                    st.bar_chart(daily_counts.set_index('posting_date'))
                    
                    # Top transactions
                    st.subheader("Top 10 Transaksi Terbesar")
                    top_trans = df.nlargest(10, 'amount')[['posting_date', 'description', 'amount', 'db_cr']]
                    st.dataframe(top_trans, use_container_width=True, hide_index=True)
            
            except Exception as e:
                st.error(f"❌ Terjadi kesalahan: {str(e)}")
                st.exception(e)
    
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
