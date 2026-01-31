"""
Card Reconciliation Tool - Web Version
Host on Streamlit Cloud for free: https://streamlit.io/cloud
"""

import streamlit as st
import pandas as pd
import re
from datetime import datetime
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Border, Side


# ============== PAGE CONFIG ==============
st.set_page_config(
    page_title="Card Reconciliation Tool",
    page_icon="💳",
    layout="wide"
)

# ============== STYLES ==============
st.markdown("""
<style>
    .main-header {
        font-size: 2.5rem;
        font-weight: bold;
        color: #1E88E5;
        text-align: center;
        margin-bottom: 1rem;
    }
    .sub-header {
        font-size: 1.2rem;
        color: #666;
        text-align: center;
        margin-bottom: 2rem;
    }
    .card-box {
        background-color: #f0f2f6;
        border-radius: 10px;
        padding: 1rem;
        margin: 0.5rem 0;
    }
    .success-box {
        background-color: #d4edda;
        border-radius: 10px;
        padding: 1rem;
        border-left: 5px solid #28a745;
    }
    .error-box {
        background-color: #f8d7da;
        border-radius: 10px;
        padding: 1rem;
        border-left: 5px solid #dc3545;
    }
</style>
""", unsafe_allow_html=True)


# ============== FUNCTIONS ==============
def extract_card_number(filename):
    """Extract card number from filename"""
    # Card Journal pattern: starts with card number like "178-D17-"
    card_match = re.match(r'^(\d+)-D17-', filename, re.IGNORECASE)
    if card_match:
        return card_match.group(1), 'card_journal'
    
    # P2P Statement pattern: contains "d17-XXX-"
    p2p_match = re.search(r'd17-(\d+)-', filename, re.IGNORECASE)
    if p2p_match:
        return p2p_match.group(1), 'p2p_statement'
    
    return None, None


def load_card_journal(file):
    """Load and parse the Card Journal CSV"""
    try:
        df = pd.read_csv(file, sep=';')
        if len(df.columns) == 1:
            file.seek(0)
            df = pd.read_csv(file, sep=',')
    except:
        file.seek(0)
        df = pd.read_csv(file, sep=',')
    
    df.columns = df.columns.str.strip()
    
    if 'DATE' in df.columns:
        df['DATE'] = df['DATE'].astype(str).str.strip()
        try:
            df['DATE'] = pd.to_datetime(df['DATE'], format='%d/%m/%Y')
        except:
            df['DATE'] = pd.to_datetime(df['DATE'], dayfirst=True)
    
    if 'MONTANT' in df.columns:
        df['MONTANT'] = df['MONTANT'].astype(str).str.replace(',', '.').astype(float)
    
    if 'LIBELLE' in df.columns:
        df['Auth'] = df['LIBELLE'].str.extract(r'(\d{6})$')
    
    return df


def load_p2p_statement(file):
    """Load and parse the P2P Statement CSV"""
    df = pd.read_csv(file)
    df.columns = df.columns.str.strip()
    
    if 'Date' in df.columns:
        df['Date'] = pd.to_datetime(df['Date'])
    
    # Convert Auth to string (handle NaN and float)
    if 'Auth' in df.columns:
        df['Auth'] = df['Auth'].fillna('').astype(str)
        df['Auth'] = df['Auth'].str.replace(r'\.0$', '', regex=True)
    
    return df


def get_in_transactions_card(df, from_date):
    """Get incoming transactions from Card Journal (Transfert du)"""
    filtered = df[df['DATE'] >= from_date].copy()
    if 'LIBELLE' in filtered.columns:
        incoming = filtered[filtered['LIBELLE'].str.contains('Transfert du', case=False, na=False)].copy()
        incoming = incoming[incoming['MONTANT'] > 0].copy()
        incoming['Auth'] = incoming['Auth'].astype(str)
        return incoming.reset_index(drop=True)
    return pd.DataFrame()


def get_out_transactions_card(df, from_date):
    """Get outgoing transactions from Card Journal (Transfert vers + Mandat)"""
    filtered = df[df['DATE'] >= from_date].copy()
    if 'LIBELLE' in filtered.columns:
        # Include both "Transfert vers" and "Mandat"
        outgoing = filtered[
            filtered['LIBELLE'].str.contains('Transfert vers', case=False, na=False) |
            filtered['LIBELLE'].str.contains('Mandat', case=False, na=False)
        ].copy()
        outgoing['MONTANT'] = outgoing['MONTANT'].abs()  # Make positive for comparison
        
        # Extract account number for "Transfert vers" (e.g., "Transfert vers 29526566 230314" -> "29526566")
        outgoing['ToAccount'] = outgoing['LIBELLE'].str.extract(r'Transfert vers (\d+)', expand=False)
        
        # For Mandat, we'll match by amount and date only
        outgoing['IsMandat'] = outgoing['LIBELLE'].str.contains('Mandat', case=False, na=False)
        
        # Create a match key: Date + Amount + ToAccount (for transfers) or Date + Amount (for mandats)
        outgoing['MatchKey'] = outgoing.apply(
            lambda row: f"{row['DATE'].strftime('%Y-%m-%d')}_{row['MONTANT']:.2f}_{row['ToAccount'] if pd.notna(row['ToAccount']) else 'MANDAT'}", 
            axis=1
        )
        
        return outgoing.reset_index(drop=True)
    return pd.DataFrame()


def get_in_transactions_p2p(df):
    """Get incoming transactions from P2P Statement (DEPOSIT)"""
    if 'Type' in df.columns:
        deposits = df[df['Type'] == 'DEPOSIT'].copy()
    else:
        deposits = df.copy()
    
    deposits = deposits[deposits['Auth'] != ''].copy()
    
    if 'Amount' in deposits.columns:
        deposits['Adjusted_Amount'] = deposits['Amount'].apply(
            lambda x: x * 0.99 if x > 40 else x
        )
    
    deposits['Auth'] = deposits['Auth'].astype(str)
    return deposits.reset_index(drop=True)


def get_out_transactions_p2p(df):
    """Get outgoing transactions from P2P Statement (WITHDRAWAL + CASHOUT)"""
    if 'Type' in df.columns:
        # Include both WITHDRAWAL and CASHOUT
        withdrawals = df[
            (df['Type'] == 'WITHDRAWAL') | (df['Type'] == 'CASHOUT')
        ].copy()
        
        if 'Amount' in withdrawals.columns:
            withdrawals['Amount'] = withdrawals['Amount'].abs()  # Make positive for comparison
            withdrawals['Adjusted_Amount'] = withdrawals['Amount']  # No fee adjustment for withdrawals
        
        # Get To Account for WITHDRAWAL, empty for CASHOUT
        if 'To Account' in withdrawals.columns:
            withdrawals['ToAccount'] = withdrawals['To Account'].apply(
                lambda x: str(int(x)) if pd.notna(x) else ''
            )
        else:
            withdrawals['ToAccount'] = ''
        
        # Create match key: Date + Amount + ToAccount (for WITHDRAWAL) or Date + Amount + MANDAT (for CASHOUT)
        withdrawals['MatchKey'] = withdrawals.apply(
            lambda row: f"{row['Date'].strftime('%Y-%m-%d')}_{row['Amount']:.2f}_{row['ToAccount'] if row['ToAccount'] != '' else 'MANDAT'}", 
            axis=1
        )
        
        return withdrawals.reset_index(drop=True)
    return pd.DataFrame()


def reconcile_transactions(card_df, p2p_df, direction='IN'):
    """Reconcile transactions between Card Journal and P2P Statement"""
    
    if card_df.empty and p2p_df.empty:
        return {
            'summary': {'card_count': 0, 'p2p_count': 0, 'matched': 0, 
                       'missing_in_p2p': 0, 'missing_in_card': 0},
            'missing_in_p2p': [], 'missing_in_card': [], 
            'matched_details': [], 'discrepancies': [],
            'totals': {'card': 0, 'p2p': 0, 'adjusted': 0, 'difference': 0},
            'missing_amount': 0
        }
    
    # For IN transactions, use Auth code; for OUT, use MatchKey
    if direction == 'IN':
        card_keys = set(card_df['Auth'].dropna()) if not card_df.empty else set()
        p2p_keys = set(p2p_df['Auth'].dropna()) if not p2p_df.empty else set()
        key_field_card = 'Auth'
        key_field_p2p = 'Auth'
    else:  # OUT
        card_keys = set(card_df['MatchKey'].dropna()) if not card_df.empty and 'MatchKey' in card_df.columns else set()
        p2p_keys = set(p2p_df['MatchKey'].dropna()) if not p2p_df.empty and 'MatchKey' in p2p_df.columns else set()
        key_field_card = 'MatchKey'
        key_field_p2p = 'MatchKey'
    
    missing_in_p2p = card_keys - p2p_keys
    missing_in_card = p2p_keys - card_keys
    matched = card_keys & p2p_keys
    
    results = {
        'summary': {
            'card_count': len(card_df),
            'p2p_count': len(p2p_df),
            'matched': len(matched),
            'missing_in_p2p': len(missing_in_p2p),
            'missing_in_card': len(missing_in_card)
        },
        'missing_in_p2p': [],
        'missing_in_card': [],
        'matched_details': [],
        'discrepancies': []
    }
    
    # Missing in P2P
    for key in missing_in_p2p:
        if not card_df.empty:
            rows = card_df[card_df[key_field_card] == key]
            if not rows.empty:
                row = rows.iloc[0]
                results['missing_in_p2p'].append({
                    'Auth': row.get('Auth', key) if direction == 'IN' else row.get('ToAccount', '-'),
                    'Date': row['DATE'].strftime('%d/%m/%Y'),
                    'Time': str(row.get('HEURE', '')).strip(),
                    'Amount': abs(row['MONTANT']),
                    'Description': row.get('LIBELLE', '')
                })
    
    # Missing in Card
    for key in missing_in_card:
        if not p2p_df.empty:
            rows = p2p_df[p2p_df[key_field_p2p] == key]
            if not rows.empty:
                row = rows.iloc[0]
                results['missing_in_card'].append({
                    'Auth': row.get('Auth', '-') if direction == 'IN' else row.get('ToAccount', '-'),
                    'Date': row['Date'].strftime('%d/%m/%Y'),
                    'Time': str(row.get('Time', '')),
                    'Amount': row['Amount'],
                    'Type': row.get('Type', '-')
                })
    
    # Matched details
    total_card = 0
    total_p2p = 0
    total_adjusted = 0
    
    for key in sorted(matched):
        card_rows = card_df[card_df[key_field_card] == key]
        p2p_rows = p2p_df[p2p_df[key_field_p2p] == key]
        
        if not card_rows.empty and not p2p_rows.empty:
            card_row = card_rows.iloc[0]
            p2p_row = p2p_rows.iloc[0]
            
            card_amt = abs(card_row['MONTANT'])
            p2p_amt = p2p_row['Amount']
            adj_amt = p2p_row['Adjusted_Amount']
            diff = round(card_amt - adj_amt, 2)
            
            total_card += card_amt
            total_p2p += p2p_amt
            total_adjusted += adj_amt
            
            status = "OK" if abs(diff) < 0.01 else "DIFF"
            fee_applied = "Yes" if p2p_amt > 40 and direction == 'IN' else "No"
            
            detail = {
                'Auth': card_row.get('Auth', '-') if direction == 'IN' else card_row.get('ToAccount', '-'),
                'Date': card_row['DATE'].strftime('%d/%m/%Y'),
                'P2P_Amount': p2p_amt,
                'Adjusted_Amount': round(adj_amt, 2),
                'Card_Amount': card_amt,
                'Difference': diff,
                'Status': status,
                'Fee_Applied': fee_applied
            }
            results['matched_details'].append(detail)
            
            if abs(diff) >= 0.01:
                results['discrepancies'].append(detail)
    
    results['totals'] = {
        'card': round(total_card, 2),
        'p2p': round(total_p2p, 2),
        'adjusted': round(total_adjusted, 2),
        'difference': round(total_card - total_adjusted, 2)
    }
    
    if results['missing_in_p2p']:
        results['missing_amount'] = sum(item['Amount'] for item in results['missing_in_p2p'])
    else:
        results['missing_amount'] = 0
    
    return results


def create_excel_report(all_results, from_date):
    """Create Excel report and return as bytes"""
    wb = Workbook()
    
    header_font = Font(bold=True, color="FFFFFF")
    header_fill_in = PatternFill("solid", fgColor="4472C4")  # Blue for IN
    header_fill_out = PatternFill("solid", fgColor="7030A0")  # Purple for OUT
    ok_fill = PatternFill("solid", fgColor="C6EFCE")
    error_fill = PatternFill("solid", fgColor="FFC7CE")
    warning_fill = PatternFill("solid", fgColor="FFEB9C")
    border = Border(
        left=Side(style='thin'), right=Side(style='thin'),
        top=Side(style='thin'), bottom=Side(style='thin')
    )
    
    # Summary sheet
    ws_summary = wb.active
    ws_summary.title = "Summary"
    
    ws_summary['A1'] = "RECONCILIATION SUMMARY - ALL CARDS"
    ws_summary['A1'].font = Font(bold=True, size=14)
    ws_summary.merge_cells('A1:I1')
    
    ws_summary['A3'] = f"From Date: {from_date.strftime('%d/%m/%Y')}"
    ws_summary['A4'] = "Rule: 1% fee removed from P2P IN amounts > 40"
    
    # IN Summary
    ws_summary['A6'] = "📥 IN TRANSACTIONS (Deposits)"
    ws_summary['A6'].font = Font(bold=True, size=12, color="4472C4")
    
    headers = ['Card', 'Card Journal', 'P2P Statement', 'Matched', 'Missing P2P', 'Missing Card', 'Missing Amount']
    for col, header in enumerate(headers, 1):
        cell = ws_summary.cell(row=7, column=col, value=header)
        cell.font = header_font
        cell.fill = header_fill_in
        cell.border = border
    
    row = 8
    total_missing_in = 0
    for card_num, results in sorted(all_results.items()):
        if 'in' in results:
            r = results['in']
            summary = r['summary']
            ws_summary.cell(row=row, column=1, value=card_num).border = border
            ws_summary.cell(row=row, column=2, value=summary['card_count']).border = border
            ws_summary.cell(row=row, column=3, value=summary['p2p_count']).border = border
            ws_summary.cell(row=row, column=4, value=summary['matched']).border = border
            
            cell_mp = ws_summary.cell(row=row, column=5, value=summary['missing_in_p2p'])
            cell_mp.border = border
            if summary['missing_in_p2p'] > 0:
                cell_mp.fill = error_fill
            
            cell_mc = ws_summary.cell(row=row, column=6, value=summary['missing_in_card'])
            cell_mc.border = border
            if summary['missing_in_card'] > 0:
                cell_mc.fill = error_fill
            
            cell_ma = ws_summary.cell(row=row, column=7, value=r['missing_amount'])
            cell_ma.border = border
            if r['missing_amount'] > 0:
                cell_ma.fill = error_fill
            
            total_missing_in += r['missing_amount']
            row += 1
    
    ws_summary.cell(row=row, column=1, value="TOTAL IN").font = Font(bold=True)
    ws_summary.cell(row=row, column=7, value=total_missing_in).font = Font(bold=True)
    if total_missing_in > 0:
        ws_summary.cell(row=row, column=7).fill = error_fill
    
    # OUT Summary
    row += 2
    ws_summary.cell(row=row, column=1, value="📤 OUT TRANSACTIONS (Withdrawals)")
    ws_summary.cell(row=row, column=1).font = Font(bold=True, size=12, color="7030A0")
    row += 1
    
    for col, header in enumerate(headers, 1):
        cell = ws_summary.cell(row=row, column=col, value=header)
        cell.font = header_font
        cell.fill = header_fill_out
        cell.border = border
    row += 1
    
    total_missing_out = 0
    for card_num, results in sorted(all_results.items()):
        if 'out' in results:
            r = results['out']
            summary = r['summary']
            ws_summary.cell(row=row, column=1, value=card_num).border = border
            ws_summary.cell(row=row, column=2, value=summary['card_count']).border = border
            ws_summary.cell(row=row, column=3, value=summary['p2p_count']).border = border
            ws_summary.cell(row=row, column=4, value=summary['matched']).border = border
            
            cell_mp = ws_summary.cell(row=row, column=5, value=summary['missing_in_p2p'])
            cell_mp.border = border
            if summary['missing_in_p2p'] > 0:
                cell_mp.fill = error_fill
            
            cell_mc = ws_summary.cell(row=row, column=6, value=summary['missing_in_card'])
            cell_mc.border = border
            if summary['missing_in_card'] > 0:
                cell_mc.fill = error_fill
            
            cell_ma = ws_summary.cell(row=row, column=7, value=r['missing_amount'])
            cell_ma.border = border
            if r['missing_amount'] > 0:
                cell_ma.fill = error_fill
            
            total_missing_out += r['missing_amount']
            row += 1
    
    ws_summary.cell(row=row, column=1, value="TOTAL OUT").font = Font(bold=True)
    ws_summary.cell(row=row, column=7, value=total_missing_out).font = Font(bold=True)
    if total_missing_out > 0:
        ws_summary.cell(row=row, column=7).fill = error_fill
    
    for col in range(1, 8):
        ws_summary.column_dimensions[chr(64 + col)].width = 18
    
    # Detail sheets per card
    for card_num, results in sorted(all_results.items()):
        for direction, label, header_fill in [('in', 'IN', header_fill_in), ('out', 'OUT', header_fill_out)]:
            if direction not in results:
                continue
            
            r = results[direction]
            if r['summary']['card_count'] == 0 and r['summary']['p2p_count'] == 0:
                continue
            
            ws = wb.create_sheet(f"Card {card_num} {label}")
            
            ws['A1'] = f"CARD {card_num} - {label} TRANSACTIONS"
            ws['A1'].font = Font(bold=True, size=12)
            ws.merge_cells('A1:H1')
            
            headers = ['Auth', 'Date', 'P2P Amt', 'Adjusted Amt', 'Card Amt', 'Diff', 'Status', 'Fee']
            for col, header in enumerate(headers, 1):
                cell = ws.cell(row=3, column=col, value=header)
                cell.font = header_font
                cell.fill = header_fill
                cell.border = border
            
            row = 4
            for item in r['matched_details']:
                ws.cell(row=row, column=1, value=item['Auth']).border = border
                ws.cell(row=row, column=2, value=item['Date']).border = border
                ws.cell(row=row, column=3, value=item['P2P_Amount']).border = border
                ws.cell(row=row, column=4, value=item['Adjusted_Amount']).border = border
                ws.cell(row=row, column=5, value=item['Card_Amount']).border = border
                ws.cell(row=row, column=6, value=item['Difference']).border = border
                
                status_cell = ws.cell(row=row, column=7, value=item['Status'])
                status_cell.border = border
                status_cell.fill = ok_fill if item['Status'] == 'OK' else warning_fill
                
                ws.cell(row=row, column=8, value=item['Fee_Applied']).border = border
                row += 1
            
            # Missing in P2P section
            if r['missing_in_p2p']:
                row += 2
                ws.cell(row=row, column=1, value="❌ MISSING IN P2P STATEMENT")
                ws.cell(row=row, column=1).font = Font(bold=True, color="FF0000", size=11)
                ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=5)
                row += 1
                
                missing_headers = ['Auth', 'Date', 'Time', 'Amount', 'Description']
                for col, header in enumerate(missing_headers, 1):
                    cell = ws.cell(row=row, column=col, value=header)
                    cell.font = Font(bold=True, color="FFFFFF")
                    cell.fill = PatternFill("solid", fgColor="C00000")
                    cell.border = border
                row += 1
                
                for item in r['missing_in_p2p']:
                    ws.cell(row=row, column=1, value=item['Auth']).border = border
                    ws.cell(row=row, column=2, value=item['Date']).border = border
                    ws.cell(row=row, column=3, value=item['Time']).border = border
                    ws.cell(row=row, column=4, value=item['Amount']).border = border
                    ws.cell(row=row, column=5, value=item['Description']).border = border
                    for c in range(1, 6):
                        ws.cell(row=row, column=c).fill = error_fill
                    row += 1
            
            # Missing in Card section
            if r['missing_in_card']:
                row += 2
                ws.cell(row=row, column=1, value="❌ MISSING IN CARD JOURNAL")
                ws.cell(row=row, column=1).font = Font(bold=True, color="FF0000", size=11)
                ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=4)
                row += 1
                
                missing_headers = ['Auth', 'Date', 'Time', 'Amount']
                for col, header in enumerate(missing_headers, 1):
                    cell = ws.cell(row=row, column=col, value=header)
                    cell.font = Font(bold=True, color="FFFFFF")
                    cell.fill = PatternFill("solid", fgColor="C00000")
                    cell.border = border
                row += 1
                
                for item in r['missing_in_card']:
                    ws.cell(row=row, column=1, value=item['Auth']).border = border
                    ws.cell(row=row, column=2, value=item['Date']).border = border
                    ws.cell(row=row, column=3, value=item['Time']).border = border
                    ws.cell(row=row, column=4, value=item['Amount']).border = border
                    for c in range(1, 5):
                        ws.cell(row=row, column=c).fill = error_fill
                    row += 1
            
            ws.column_dimensions['A'].width = 12
            ws.column_dimensions['B'].width = 12
            ws.column_dimensions['C'].width = 12
            ws.column_dimensions['D'].width = 14
            ws.column_dimensions['E'].width = 12
            ws.column_dimensions['F'].width = 10
            ws.column_dimensions['G'].width = 10
            ws.column_dimensions['H'].width = 8
    
    # Save to bytes
    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return output


# ============== MAIN APP ==============
st.markdown('<div class="main-header">💳 Card Reconciliation Tool</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-header">Compare Card Journal vs P2P Statement</div>', unsafe_allow_html=True)

st.info("""
**Rules:**
- 📥 **IN (Deposits):** Match by Auth code | 1% fee removed from P2P amounts > 40
- 📤 **OUT (Withdrawals/Cashouts):** Match by Date + Amount + Account | No fee adjustment
""")

# File upload section
st.markdown("---")
st.subheader("📁 Upload Files")

col1, col2 = st.columns(2)

with col1:
    st.markdown("**Card Journal**")
    st.caption("Files like: 178-D17-31-01-2026.csv")
    card_files = st.file_uploader(
        "Upload Card Journal CSV files",
        type=['csv'],
        accept_multiple_files=True,
        key="card_journals"
    )

with col2:
    st.markdown("**P2P Statement**")
    st.caption("Files like: statement-d17-178-xxxxx.csv")
    p2p_files = st.file_uploader(
        "Upload P2P Statement CSV files",
        type=['csv'],
        accept_multiple_files=True,
        key="p2p_statements"
    )

# Date input
st.markdown("---")
from_date = st.date_input(
    "📅 Starting Date",
    value=datetime(2026, 1, 27),
    format="DD/MM/YYYY"
)

# Process files
if card_files and p2p_files:
    # Organize files by card number
    card_journals = {}
    p2p_statements = {}
    
    for f in card_files:
        card_num, file_type = extract_card_number(f.name)
        if card_num and file_type == 'card_journal':
            card_journals[card_num] = f
    
    for f in p2p_files:
        card_num, file_type = extract_card_number(f.name)
        if card_num and file_type == 'p2p_statement':
            p2p_statements[card_num] = f
    
    # Show detected files
    st.markdown("---")
    st.subheader("📄 Detected Files")
    
    col1, col2 = st.columns(2)
    with col1:
        st.markdown("**Card Journals:**")
        for card, f in sorted(card_journals.items()):
            st.write(f"  Card {card}: {f.name}")
    
    with col2:
        st.markdown("**P2P Statements:**")
        for card, f in sorted(p2p_statements.items()):
            st.write(f"  Card {card}: {f.name}")
    
    # Find matched cards
    matched_cards = set(card_journals.keys()) & set(p2p_statements.keys())
    
    if matched_cards:
        st.success(f"✅ Matched cards: {', '.join(sorted(matched_cards))}")
        
        # Run reconciliation button
        if st.button("🚀 Run Reconciliation", type="primary", use_container_width=True):
            from_date_dt = datetime.combine(from_date, datetime.min.time())
            
            all_results = {}
            
            with st.spinner("Running reconciliation..."):
                for card_num in sorted(matched_cards):
                    try:
                        card_file = card_journals[card_num]
                        p2p_file = p2p_statements[card_num]
                        
                        card_file.seek(0)
                        p2p_file.seek(0)
                        
                        card_df = load_card_journal(card_file)
                        p2p_df = load_p2p_statement(p2p_file)
                        
                        # Get IN transactions
                        card_in = get_in_transactions_card(card_df, from_date_dt)
                        p2p_in = get_in_transactions_p2p(p2p_df)
                        results_in = reconcile_transactions(card_in, p2p_in, 'IN')
                        
                        # Get OUT transactions
                        card_out = get_out_transactions_card(card_df, from_date_dt)
                        p2p_out = get_out_transactions_p2p(p2p_df)
                        results_out = reconcile_transactions(card_out, p2p_out, 'OUT')
                        
                        all_results[card_num] = {
                            'in': results_in,
                            'out': results_out
                        }
                    except Exception as e:
                        st.error(f"Error processing card {card_num}: {e}")
            
            if all_results:
                # Show results
                st.markdown("---")
                st.subheader("📊 Results")
                
                # IN Summary
                st.markdown("### 📥 IN Transactions (Deposits)")
                in_data = []
                for card_num, results in sorted(all_results.items()):
                    r = results['in']
                    summary = r['summary']
                    in_data.append({
                        'Card': card_num,
                        'Card Journal': summary['card_count'],
                        'P2P Statement': summary['p2p_count'],
                        'Matched': summary['matched'],
                        'Missing (P2P)': summary['missing_in_p2p'],
                        'Missing (Card)': summary['missing_in_card'],
                        'Missing Amount': r['missing_amount']
                    })
                
                df_in = pd.DataFrame(in_data)
                st.dataframe(df_in, use_container_width=True, hide_index=True)
                
                total_missing_in = sum(r['in']['missing_amount'] for r in all_results.values())
                if total_missing_in > 0:
                    st.error(f"⚠️ Total Missing IN Amount: **{total_missing_in:.2f}**")
                else:
                    st.success("✅ All IN transactions matched!")
                
                # OUT Summary
                st.markdown("### 📤 OUT Transactions (Withdrawals)")
                out_data = []
                for card_num, results in sorted(all_results.items()):
                    r = results['out']
                    summary = r['summary']
                    out_data.append({
                        'Card': card_num,
                        'Card Journal': summary['card_count'],
                        'P2P Statement': summary['p2p_count'],
                        'Matched': summary['matched'],
                        'Missing (P2P)': summary['missing_in_p2p'],
                        'Missing (Card)': summary['missing_in_card'],
                        'Missing Amount': r['missing_amount']
                    })
                
                df_out = pd.DataFrame(out_data)
                st.dataframe(df_out, use_container_width=True, hide_index=True)
                
                total_missing_out = sum(r['out']['missing_amount'] for r in all_results.values())
                if total_missing_out > 0:
                    st.error(f"⚠️ Total Missing OUT Amount: **{total_missing_out:.2f}**")
                else:
                    st.success("✅ All OUT transactions matched!")
                
                # Details per card
                for card_num, results in sorted(all_results.items()):
                    with st.expander(f"📋 Card {card_num} Details"):
                        tab_in, tab_out = st.tabs(["📥 IN", "📤 OUT"])
                        
                        with tab_in:
                            r = results['in']
                            if r['missing_in_p2p']:
                                st.markdown("**❌ Missing in P2P Statement:**")
                                st.dataframe(pd.DataFrame(r['missing_in_p2p']), use_container_width=True, hide_index=True)
                            
                            if r['missing_in_card']:
                                st.markdown("**❌ Missing in Card Journal:**")
                                st.dataframe(pd.DataFrame(r['missing_in_card']), use_container_width=True, hide_index=True)
                            
                            if r['matched_details']:
                                st.markdown("**✅ Matched:**")
                                st.dataframe(pd.DataFrame(r['matched_details']), use_container_width=True, hide_index=True)
                        
                        with tab_out:
                            r = results['out']
                            if r['missing_in_p2p']:
                                st.markdown("**❌ Missing in P2P Statement:**")
                                st.dataframe(pd.DataFrame(r['missing_in_p2p']), use_container_width=True, hide_index=True)
                            
                            if r['missing_in_card']:
                                st.markdown("**❌ Missing in Card Journal:**")
                                st.dataframe(pd.DataFrame(r['missing_in_card']), use_container_width=True, hide_index=True)
                            
                            if r['matched_details']:
                                st.markdown("**✅ Matched:**")
                                st.dataframe(pd.DataFrame(r['matched_details']), use_container_width=True, hide_index=True)
                
                # Download Excel report
                st.markdown("---")
                excel_file = create_excel_report(all_results, from_date_dt)
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                
                st.download_button(
                    label="📥 Download Excel Report",
                    data=excel_file,
                    file_name=f"recon_report_{timestamp}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
    else:
        st.warning("⚠️ No matching card pairs found. Make sure you have both Card Journal and P2P Statement files for at least one card.")

else:
    st.markdown("---")
    st.markdown("### 👆 Upload your files to get started")
    st.markdown("""
    **File naming:**
    - Card Journal: `178-D17-xxxxx.csv` (card number at start)
    - P2P Statement: `statement-d17-178-xxxxx.csv` (card number after d17-)
    """)

# Footer
st.markdown("---")
st.caption("Card Reconciliation Tool v2.2 | IN: Auth match | OUT: Date+Amount+Account match")
