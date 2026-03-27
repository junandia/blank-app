#!/usr/bin/env python3
"""
E-Statement Bank Converter
Konversi e-statement bank (PDF) ke Excel/CSV
"""

import streamlit as st
import pdfplumber
import pandas as pd
import re
import io
from datetime import datetime
from typing import Dict, List, Optional, Tuple, Any

st.set_page_config(
    page_title="E-Statement Converter",
    page_icon="🏦",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.markdown("""
<style>
    .main-header {
        font-size: 2.5rem;
        font-weight: bold;
        color: #1F4E79;
        text-align: center;
        margin-bottom: 1rem;
    }
    .stat-card {
        background: linear-gradient(135deg, #1F4E79 0%, #2E75B6 100%);
        color: white;
        padding: 1rem;
        border-radius: 10px;
        text-align: center;
    }
    .stat-value { font-size: 1.5rem; font-weight: bold; }
    .stat-label { font-size: 0.9rem; opacity: 0.9; }
</style>
""", unsafe_allow_html=True)

def clean_balance(balance_str: Any) -> float:
    if not balance_str:
        return 0.0
    cleaned = str(balance_str).replace('.', '').replace(',', '.')
    cleaned = re.sub(r'[^\d.\-]', '', cleaned)
    try:
        return float(cleaned)
    except ValueError:
        return 0.0

def extract_account_info(text: str) -> Dict[str, str]:
    info = {'account_name': '', 'account_number': '', 'account_type': '', 'period': '', 'currency': 'IDR', 'ledger_balance': 0.0}
    lines = text.split('\n')
    for i, line in enumerate(lines):
        if 'ACCOUNT STATEMENT' in line and i + 1 < len(lines):
            next_line = lines[i + 1]
            name_match = re.match(r'^([A-Z][A-Z\s]+?)\s+Account No', next_line)
            if name_match:
                info['account_name'] = name_match.group(1).strip()
            break
    acc_match = re.search(r'Account No\.?\s*:\s*([\d]+(?:\s*/\s*[A-Z]+)?)', text)
    if acc_match:
        info['account_number'] = acc_match.group(1).strip()
    type_match = re.search(r'Account Type\s*:\s*(\w+)', text)
    if type_match:
        info['account_type'] = type_match.group(1)
    period_match = re.search(r'Period\s*:\s*([\d]+-[A-Za-z]+-[\d]+\s*-\s*[\d]+-[A-Za-z]+-[\d]+)', text)
    if period_match:
        info['period'] = period_match.group(1)
    ledger_match = re.search(r'Ledger Balance:\s*([\d.,]+)', text)
    if ledger_match:
        info['ledger_balance'] = clean_balance(ledger_match.group(1))
    return info

def parse_row(row: List, prev_balance: Optional[float], num_cols: int) -> Tuple[Optional[Dict], Optional[float]]:
    first_cell = str(row[0]) if row[0] else ''
    if 'Posting Date' in first_cell or 'Ledger Balance' in first_cell or first_cell == '' or first_cell == 'None':
        return None, prev_balance
    if not re.search(r'\d{2}/\d{2}/\d{4}', first_cell):
        return None, prev_balance
    posting_date = str(row[0]).strip()
    if num_cols >= 10:
        effective_date = str(row[2]).strip() if len(row) > 2 and row[2] else posting_date
        branch = str(row[3]).replace('\n', ' ').strip() if len(row) > 3 and row[3] else ''
        journal = str(row[4]).strip() if len(row) > 4 and row[4] else ''
        description = str(row[5]).replace('\n', ' ').strip() if len(row) > 5 and row[5] else ''
        db_cr = str(row[7]).strip() if len(row) > 7 and row[7] in ['D', 'K'] else ''
        balance_raw = row[9] if len(row) > 9 else None
    else:
        effective_date = str(row[1]).strip() if len(row) > 1 and row[1] else posting_date
        branch = str(row[2]).replace('\n', ' ').strip() if len(row) > 2 and row[2] else ''
        journal = str(row[3]).strip() if len(row) > 3 and row[3] else ''
        description = str(row[4]).replace('\n', ' ').strip() if len(row) > 4 and row[4] else ''
        db_cr = str(row[6]).strip() if len(row) > 6 and row[6] in ['D', 'K'] else ''
        balance_raw = row[7] if len(row) > 7 else None
    branch = re.sub(r'\s+', ' ', branch)
    description = re.sub(r'\s+', ' ', description)
    current_balance = clean_balance(balance_raw)
    if not posting_date or not db_cr or current_balance == 0:
        return None, prev_balance
    if prev_balance is not None:
        if db_cr == 'D':
            amount = prev_balance - current_balance
        else:
            amount = current_balance - prev_balance
    else:
        prev_balance = current_balance
        return None, prev_balance
    trans = {'posting_date': posting_date, 'effective_date': effective_date, 'branch': branch, 
             'journal': journal, 'description': description, 'amount': abs(amount), 'db_cr': db_cr, 'balance': current_balance}
    return trans, current_balance

def extract_transactions(pdf_file) -> Tuple[Dict[str, str], List[Dict]]:
    all_transactions = []
    account_info = {}
    prev_balance = None
    with pdfplumber.open(pdf_file) as pdf:
        first_text = pdf.pages[0].extract_text() or ""
        account_info = extract_account_info(first_text)
        ledger_match = re.search(r'Ledger Balance:\s*([\d.,]+)', first_text)
        if ledger_match:
            prev_balance = clean_balance(ledger_match.group(1))
        for page_num, page in enumerate(pdf.pages):
            tables = page.extract_tables()
            for table in tables:
                if not table or len(table) < 2:
                    continue
                num_cols = len(table[0]) if table[0] else 0
                for row in table:
                    trans, prev_balance = parse_row(row, prev_balance, num_cols)
                    if trans:
                        all_transactions.append(trans)
    return account_info, all_transactions

def format_currency(amount: float) -> str:
    return f"Rp {amount:,.2f}"

def main():
    st.markdown('<h1 class="main-header">🏦 E-Statement Bank Converter</h1>', unsafe_allow_html=True)
    st.markdown('<p style="text-align:center;color:#666;">Konversi e-statement bank (PDF) ke Excel/CSV</p>', unsafe_allow_html=True)
    
    st.sidebar.header("⚙️ Pengaturan")
    st.sidebar.markdown("### Bank yang Didukung")
    st.sidebar.markdown("- ✅ BNI\n- ✅ BCA\n- ✅ Mandiri\n- ✅ BRI\n- ✅ Dan bank lainnya")
    
    uploaded_file = st.file_uploader("📄 Pilih file PDF e-statement bank", type=['pdf'])
    
    if uploaded_file is not None:
        if st.button("🔄 Proses E-Statement", type="primary"):
            with st.spinner("Membaca dan mengekstrak data dari PDF..."):
                try:
                    account_info, transactions = extract_transactions(uploaded_file)
                    if not transactions:
                        st.warning("⚠️ Tidak ada transaksi yang dapat diekstrak.")
                    else:
                        st.success(f"✅ Berhasil mengekstrak {len(transactions)} transaksi!")
                        df = pd.DataFrame(transactions)
                        debit_total = df[df['db_cr'] == 'D']['amount'].sum()
                        credit_total = df[df['db_cr'] == 'K']['amount'].sum()
                        net_flow = credit_total - debit_total
                        
                        st.header("📋 Informasi Rekening")
                        col_a, col_b, col_c = st.columns(3)
                        col_a.metric("Nama Rekening", account_info.get('account_name', 'N/A'))
                        col_b.metric("Nomor Rekening", account_info.get('account_number', 'N/A'))
                        col_c.metric("Periode", account_info.get('period', 'N/A'))
                        
                        st.header("📊 Ringkasan")
                        col_d, col_e, col_f, col_g = st.columns(4)
                        col_d.metric("Total Transaksi", f"{len(transactions):,}")
                        col_e.metric("Total Debit", format_currency(debit_total))
                        col_f.metric("Total Kredit", format_currency(credit_total))
                        col_g.metric("Net Flow", format_currency(net_flow))
                        
                        st.header("📋 Transaksi")
                        st.dataframe(df, use_container_width=True, hide_index=True)
                        
                        st.header("📥 Export Data")
                        col1, col2 = st.columns(2)
                        with col1:
                            excel_buffer = io.BytesIO()
                            with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
                                df.to_excel(writer, sheet_name='Transaksi', index=False)
                            excel_buffer.seek(0)
                            st.download_button("📥 Download Excel (.xlsx)", excel_buffer, 
                                file_name=f"estatement_{datetime.now().strftime('%Y%m%d')}.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                        with col2:
                            csv_buffer = io.StringIO()
                            df.to_csv(csv_buffer, index=False)
                            st.download_button("📥 Download CSV", csv_buffer.getvalue(),
                                file_name=f"estatement_{datetime.now().strftime('%Y%m%d')}.csv", mime="text/csv")
                except Exception as e:
                    st.error(f"❌ Terjadi kesalahan: {str(e)}")
    
    st.markdown("---")
    st.markdown("<p style='text-align:center;color:#666;'>🏦 E-Statement Converter | Powered by Streamlit</p>", unsafe_allow_html=True)

if __name__ == "__main__":
    main()
