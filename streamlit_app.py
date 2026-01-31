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
    # Statement pattern: starts with card number like "178-D17-"
    stmt_match = re.match(r'^(\d+)-D17-', filename, re.IGNORECASE)
    if stmt_match:
        return stmt_match.group(1), 'statement'
    
    # Journal pattern: contains "d17-XXX-"
    journal_match = re.search(r'd17-(\d+)-', filename, re.IGNORECASE)
    if journal_match:
        return journal_match.group(1), 'journal'
    
    return None, None


def load_statement(file):
    """Load and parse the card statement CSV"""
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
        # Filter: Only "Transfert du" (incoming transfers)
        # Exclude: "Mandat", "Transfert vers" (outgoing)
        df = df[df['LIBELLE'].str.contains('Transfert du', case=False, na=False)].copy()
        df['Auth'] = df['LIBELLE'].str.extract(r'(\d{6})$')
    
    return df


def load_journal(file):
    """Load and parse the P2P journal/deposits CSV"""
    df = pd.read_csv(file)
    df.columns = df.columns.str.strip()
    
    if 'Date' in df.columns:
        df['Date'] = pd.to_datetime(df['Date'])
    
    # Filter only DEPOSIT transactions if Type column exists
    if 'Type' in df.columns:
        df = df[df['Type'] == 'DEPOSIT'].copy()
    
    # Convert Auth to string (handle NaN and float)
    if 'Auth' in df.columns:
        df['Auth'] = df['Auth'].fillna('').astype(str)
        df['Auth'] = df['Auth'].str.replace(r'\.0$', '', regex=True)
        df = df[df['Auth'] != ''].copy()
    
    # Apply 1% fee adjustment for amounts > 40
    if 'Amount' in df.columns:
        df['Adjusted_Amount'] = df['Amount'].apply(
            lambda x: x * 0.99 if x > 40 else x
        )
    
    return df


def reconcile(statement_df, journal_df, from_date):
    """Perform reconciliation between statement and journal"""
    statement_filtered = statement_df[statement_df['DATE'] >= from_date].copy()
    statement_incoming = statement_filtered[statement_filtered['MONTANT'] > 0].copy()
    statement_incoming = statement_incoming.reset_index(drop=True)
    statement_incoming['Auth'] = statement_incoming['Auth'].astype(str)
    
    journal_df = journal_df.copy()
    journal_df['Auth'] = journal_df['Auth'].astype(str)
    
    statement_auths = set(statement_incoming['Auth'].dropna())
    journal_auths = set(journal_df['Auth'].dropna())
    
    missing_in_journal = statement_auths - journal_auths
    missing_in_statement = journal_auths - statement_auths
    matched = statement_auths & journal_auths
    
    results = {
        'summary': {
            'statement_count': len(statement_incoming),
            'journal_count': len(journal_df),
            'matched': len(matched),
            'missing_in_journal': len(missing_in_journal),
            'missing_in_statement': len(missing_in_statement),
            'from_date': from_date.strftime('%d/%m/%Y')
        },
        'missing_in_journal': [],
        'missing_in_statement': [],
        'matched_details': [],
        'discrepancies': []
    }
    
    for auth in missing_in_journal:
        row = statement_incoming[statement_incoming['Auth'] == auth].iloc[0]
        results['missing_in_journal'].append({
            'Auth': auth,
            'Date': row['DATE'].strftime('%d/%m/%Y'),
            'Time': str(row.get('HEURE', '')).strip(),
            'Amount': row['MONTANT'],
            'Description': row.get('LIBELLE', '')
        })
    
    for auth in missing_in_statement:
        row = journal_df[journal_df['Auth'] == auth].iloc[0]
        results['missing_in_statement'].append({
            'Auth': auth,
            'Date': row['Date'].strftime('%d/%m/%Y'),
            'Time': str(row.get('Time', '')),
            'Amount': row['Amount']
        })
    
    total_statement = 0
    total_journal = 0
    total_adjusted = 0
    
    for auth in sorted(matched):
        stmt_row = statement_incoming[statement_incoming['Auth'] == auth].iloc[0]
        jour_row = journal_df[journal_df['Auth'] == auth].iloc[0]
        
        stmt_amt = stmt_row['MONTANT']
        jour_amt = jour_row['Amount']
        adj_amt = jour_row['Adjusted_Amount']
        diff = round(stmt_amt - adj_amt, 2)
        
        total_statement += stmt_amt
        total_journal += jour_amt
        total_adjusted += adj_amt
        
        status = "OK" if abs(diff) < 0.01 else "DIFF"
        fee_applied = "Yes" if jour_amt > 40 else "No"
        
        detail = {
            'Auth': auth,
            'Date': stmt_row['DATE'].strftime('%d/%m/%Y'),
            'Journal_Amount': jour_amt,
            'Adjusted_Amount': round(adj_amt, 2),
            'Statement_Amount': stmt_amt,
            'Difference': diff,
            'Status': status,
            'Fee_Applied': fee_applied
        }
        results['matched_details'].append(detail)
        
        if abs(diff) >= 0.01:
            results['discrepancies'].append(detail)
    
    results['totals'] = {
        'statement': round(total_statement, 2),
        'journal': round(total_journal, 2),
        'adjusted': round(total_adjusted, 2),
        'difference': round(total_statement - total_adjusted, 2)
    }
    
    if results['missing_in_journal']:
        results['missing_amount'] = sum(item['Amount'] for item in results['missing_in_journal'])
    else:
        results['missing_amount'] = 0
    
    return results


def create_excel_report(all_results, from_date):
    """Create Excel report and return as bytes"""
    wb = Workbook()
    
    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill("solid", fgColor="4472C4")
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
    ws_summary.merge_cells('A1:G1')
    
    ws_summary['A3'] = f"From Date: {from_date.strftime('%d/%m/%Y')}"
    ws_summary['A4'] = "Rule: 1% fee removed from P2P amounts > 40"
    
    headers = ['Card', 'Statement', 'Journal', 'Matched', 'Missing Journal', 'Missing Statement', 'Missing Amount']
    for col, header in enumerate(headers, 1):
        cell = ws_summary.cell(row=6, column=col, value=header)
        cell.font = header_font
        cell.fill = header_fill
        cell.border = border
    
    row = 7
    total_missing = 0
    for card_num, results in sorted(all_results.items()):
        summary = results['summary']
        ws_summary.cell(row=row, column=1, value=card_num).border = border
        ws_summary.cell(row=row, column=2, value=summary['statement_count']).border = border
        ws_summary.cell(row=row, column=3, value=summary['journal_count']).border = border
        ws_summary.cell(row=row, column=4, value=summary['matched']).border = border
        
        cell_mj = ws_summary.cell(row=row, column=5, value=summary['missing_in_journal'])
        cell_mj.border = border
        if summary['missing_in_journal'] > 0:
            cell_mj.fill = error_fill
        
        cell_ms = ws_summary.cell(row=row, column=6, value=summary['missing_in_statement'])
        cell_ms.border = border
        if summary['missing_in_statement'] > 0:
            cell_ms.fill = error_fill
        
        cell_ma = ws_summary.cell(row=row, column=7, value=results['missing_amount'])
        cell_ma.border = border
        if results['missing_amount'] > 0:
            cell_ma.fill = error_fill
        
        total_missing += results['missing_amount']
        row += 1
    
    ws_summary.cell(row=row, column=1, value="TOTAL").font = Font(bold=True)
    ws_summary.cell(row=row, column=7, value=total_missing).font = Font(bold=True)
    if total_missing > 0:
        ws_summary.cell(row=row, column=7).fill = error_fill
    
    for col in range(1, 8):
        ws_summary.column_dimensions[chr(64 + col)].width = 18
    
    # Detail sheets per card
    for card_num, results in sorted(all_results.items()):
        ws = wb.create_sheet(f"Card {card_num}")
        
        ws['A1'] = f"CARD {card_num} - RECONCILIATION DETAILS"
        ws['A1'].font = Font(bold=True, size=12)
        ws.merge_cells('A1:H1')
        
        headers = ['Auth', 'Date', 'Journal Amt', 'Adjusted Amt', 'Statement Amt', 'Diff', 'Status', 'Fee']
        for col, header in enumerate(headers, 1):
            cell = ws.cell(row=3, column=col, value=header)
            cell.font = header_font
            cell.fill = header_fill
            cell.border = border
        
        row = 4
        for item in results['matched_details']:
            ws.cell(row=row, column=1, value=item['Auth']).border = border
            ws.cell(row=row, column=2, value=item['Date']).border = border
            ws.cell(row=row, column=3, value=item['Journal_Amount']).border = border
            ws.cell(row=row, column=4, value=item['Adjusted_Amount']).border = border
            ws.cell(row=row, column=5, value=item['Statement_Amount']).border = border
            ws.cell(row=row, column=6, value=item['Difference']).border = border
            
            status_cell = ws.cell(row=row, column=7, value=item['Status'])
            status_cell.border = border
            status_cell.fill = ok_fill if item['Status'] == 'OK' else warning_fill
            
            ws.cell(row=row, column=8, value=item['Fee_Applied']).border = border
            row += 1
        
        # Missing in Journal section
        if results['missing_in_journal']:
            row += 2
            ws.cell(row=row, column=1, value="❌ MISSING IN JOURNAL (Not in Deposits)")
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
            
            for item in results['missing_in_journal']:
                ws.cell(row=row, column=1, value=item['Auth']).border = border
                ws.cell(row=row, column=2, value=item['Date']).border = border
                ws.cell(row=row, column=3, value=item['Time']).border = border
                ws.cell(row=row, column=4, value=item['Amount']).border = border
                ws.cell(row=row, column=5, value=item['Description']).border = border
                for c in range(1, 6):
                    ws.cell(row=row, column=c).fill = error_fill
                row += 1
        
        # Missing in Statement section
        if results['missing_in_statement']:
            row += 2
            ws.cell(row=row, column=1, value="❌ MISSING IN STATEMENT (Not in Card)")
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
            
            for item in results['missing_in_statement']:
                ws.cell(row=row, column=1, value=item['Auth']).border = border
                ws.cell(row=row, column=2, value=item['Date']).border = border
                ws.cell(row=row, column=3, value=item['Time']).border = border
                ws.cell(row=row, column=4, value=item['Amount']).border = border
                for c in range(1, 5):
                    ws.cell(row=row, column=c).fill = error_fill
                row += 1
        
        ws.column_dimensions['A'].width = 12
        ws.column_dimensions['B'].width = 12
        ws.column_dimensions['C'].width = 14
        ws.column_dimensions['D'].width = 14
        ws.column_dimensions['E'].width = 14
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
st.markdown('<div class="sub-header">Compare Card Statements vs P2P Journal/Deposits</div>', unsafe_allow_html=True)

st.info("**Rule:** 1% fee is removed from P2P amounts > 40 before comparing")

# File upload section
st.markdown("---")
st.subheader("📁 Upload Files")

col1, col2 = st.columns(2)

with col1:
    st.markdown("**Card Statements**")
    st.caption("Files like: 178-D17-31-01-2026.csv")
    statement_files = st.file_uploader(
        "Upload Statement CSV files",
        type=['csv'],
        accept_multiple_files=True,
        key="statements"
    )

with col2:
    st.markdown("**P2P Journals**")
    st.caption("Files like: statement-d17-178-xxxxx.csv")
    journal_files = st.file_uploader(
        "Upload Journal CSV files",
        type=['csv'],
        accept_multiple_files=True,
        key="journals"
    )

# Date input
st.markdown("---")
from_date = st.date_input(
    "📅 Starting Date",
    value=datetime(2026, 1, 27),
    format="DD/MM/YYYY"
)

# Process files
if statement_files and journal_files:
    # Organize files by card number
    statements = {}
    journals = {}
    
    for f in statement_files:
        card_num, file_type = extract_card_number(f.name)
        if card_num and file_type == 'statement':
            statements[card_num] = f
    
    for f in journal_files:
        card_num, file_type = extract_card_number(f.name)
        if card_num and file_type == 'journal':
            journals[card_num] = f
    
    # Show detected files
    st.markdown("---")
    st.subheader("📄 Detected Files")
    
    col1, col2 = st.columns(2)
    with col1:
        st.markdown("**Statements:**")
        for card, f in sorted(statements.items()):
            st.write(f"  Card {card}: {f.name}")
    
    with col2:
        st.markdown("**Journals:**")
        for card, f in sorted(journals.items()):
            st.write(f"  Card {card}: {f.name}")
    
    # Find matched cards
    matched_cards = set(statements.keys()) & set(journals.keys())
    
    if matched_cards:
        st.success(f"✅ Matched cards: {', '.join(sorted(matched_cards))}")
        
        # Run reconciliation button
        if st.button("🚀 Run Reconciliation", type="primary", use_container_width=True):
            from_date_dt = datetime.combine(from_date, datetime.min.time())
            
            all_results = {}
            
            with st.spinner("Running reconciliation..."):
                for card_num in sorted(matched_cards):
                    try:
                        stmt_file = statements[card_num]
                        jour_file = journals[card_num]
                        
                        stmt_file.seek(0)
                        jour_file.seek(0)
                        
                        statement_df = load_statement(stmt_file)
                        journal_df = load_journal(jour_file)
                        
                        results = reconcile(statement_df, journal_df, from_date_dt)
                        all_results[card_num] = results
                    except Exception as e:
                        st.error(f"Error processing card {card_num}: {e}")
            
            if all_results:
                # Show results
                st.markdown("---")
                st.subheader("📊 Results")
                
                # Summary table
                summary_data = []
                for card_num, results in sorted(all_results.items()):
                    summary = results['summary']
                    summary_data.append({
                        'Card': card_num,
                        'Statement': summary['statement_count'],
                        'Journal': summary['journal_count'],
                        'Matched': summary['matched'],
                        'Missing (Journal)': summary['missing_in_journal'],
                        'Missing (Statement)': summary['missing_in_statement'],
                        'Missing Amount': results['missing_amount']
                    })
                
                df_summary = pd.DataFrame(summary_data)
                st.dataframe(df_summary, use_container_width=True, hide_index=True)
                
                # Total missing
                total_missing = sum(r['missing_amount'] for r in all_results.values())
                if total_missing > 0:
                    st.error(f"⚠️ Total Missing Amount: **{total_missing:.2f}**")
                else:
                    st.success("✅ All transactions matched!")
                
                # Details per card
                for card_num, results in sorted(all_results.items()):
                    with st.expander(f"📋 Card {card_num} Details"):
                        if results['missing_in_journal']:
                            st.markdown("**❌ Missing in Journal:**")
                            df_missing = pd.DataFrame(results['missing_in_journal'])
                            st.dataframe(df_missing, use_container_width=True, hide_index=True)
                        
                        if results['missing_in_statement']:
                            st.markdown("**❌ Missing in Statement:**")
                            df_missing = pd.DataFrame(results['missing_in_statement'])
                            st.dataframe(df_missing, use_container_width=True, hide_index=True)
                        
                        if results['matched_details']:
                            st.markdown("**✅ Matched Transactions:**")
                            df_matched = pd.DataFrame(results['matched_details'])
                            st.dataframe(df_matched, use_container_width=True, hide_index=True)
                
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
        st.warning("⚠️ No matching card pairs found. Make sure you have both statement and journal files for at least one card.")

else:
    st.markdown("---")
    st.markdown("### 👆 Upload your files to get started")
    st.markdown("""
    **File naming:**
    - Statement: `178-D17-xxxxx.csv` (card number at start)
    - Journal: `statement-d17-178-xxxxx.csv` (card number after d17-)
    """)

# Footer
st.markdown("---")
st.caption("Card Reconciliation Tool v1.0 | Rule: 1% fee removed from P2P amounts > 40")
