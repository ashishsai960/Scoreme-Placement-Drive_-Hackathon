import re
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, Border, Side, Alignment
import pdfplumber

def extract_punjab_sind_transactions(pdf_path):
    all_text = ""
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text = page.extract_text(layout=True, x_tolerance=2, y_tolerance=2)
            if text:
                all_text += text + "\n"
    transactions = []
    lines = all_text.split('\n')
    i = 0
    n = len(lines)
    
    while i < n:
        line = lines[i].strip()
        if not line or "Statement of account" in line or "BANK NAME" in line:
            i += 1
            continue
        if re.match(r'^\d{2}-[A-Za-z]{3}-\d{4}', line):
            parts = re.split(r'\s{2,}', line)
            
            if len(parts) >= 4:
                date = parts[0]
                desc_parts = []
                amount = None
                balance = None
                
                for part in parts[1:]:
                    if re.match(r'^[\d,]+\.\d{2}(?:Dr|Cr)?$', part.replace(',', '')):
                        if amount is None:
                            amount = part
                        else:
                            balance = part
                    else:
                        desc_parts.append(part)
                
                description = ' '.join(desc_parts)
                
                if amount and balance:
                    amount_num = float(amount.replace(',', '').replace('Dr', '').replace('Cr', ''))
                    balance_num = float(balance.replace(',', '').replace('Dr', '').replace('Cr', ''))
                    trans_type = "Debit" if 'Dr' in amount else "Credit"
                    balance_type = "Dr" if 'Dr' in balance else "Cr"
                    
                    transactions.append({
                        'Date': date,
                        'Description': description,
                        'Amount': amount_num,
                        'Type': trans_type,
                        'Balance': balance_num,
                        'Balance Type': balance_type
                    })
            elif i+1 < n and not re.match(r'^\d{2}-[A-Za-z]{3}-\d{4}', lines[i+1].strip()):
                next_line = lines[i+1].strip()
                if next_line and not any(x in next_line for x in ['Page No:', 'REPORT PRINTED BY']):
                    combined_line = line + " " + next_line
                    parts = re.split(r'\s{2,}', combined_line)
                    
                    if len(parts) >= 4:
                        date = parts[0]
                        desc_parts = []
                        amount = None
                        balance = None
                        
                        for part in parts[1:]:
                            if re.match(r'^[\d,]+\.\d{2}(?:Dr|Cr)?$', part.replace(',', '')):
                                if amount is None:
                                    amount = part
                                else:
                                    balance = part
                            else:
                                desc_parts.append(part)
                        
                        description = ' '.join(desc_parts)
                        
                        if amount and balance:
                            amount_num = float(amount.replace(',', '').replace('Dr', '').replace('Cr', ''))
                            balance_num = float(balance.replace(',', '').replace('Dr', '').replace('Cr', ''))
                            
                            trans_type = "Debit" if 'Dr' in amount else "Credit"
                            balance_type = "Dr" if 'Dr' in balance else "Cr"
                            
                            transactions.append({
                                'Date': date,
                                'Description': description,
                                'Amount': amount_num,
                                'Type': trans_type,
                                'Balance': balance_num,
                                'Balance Type': balance_type
                            })
                            i += 1  
        
        i += 1
    
    return transactions

def save_to_excel(transactions, output_file):
    """Save transactions to formatted Excel file"""
    if not transactions:
        print("No transactions to save")
        return False
    

    df = pd.DataFrame(transactions)
    
    df['Date'] = pd.to_datetime(df['Date'], format='%d-%b-%Y')
    
    df['Formatted Balance'] = df.apply(
        lambda x: f"{x['Balance']:,.2f}{x['Balance Type']}", axis=1)
    
    df = df[['Date', 'Description', 'Amount', 'Type', 'Formatted Balance']]
    
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Transactions')
        
        workbook = writer.book
        worksheet = writer.sheets['Transactions']
        
        header_font = Font(bold=True)
        border = Border(left=Side(style='thin'), 
                      right=Side(style='thin'), 
                      top=Side(style='thin'), 
                      bottom=Side(style='thin'))
        
        for cell in worksheet[1]:
            cell.font = header_font
            cell.border = border
            cell.alignment = Alignment(horizontal='center')
        
        for row in worksheet.iter_rows(min_row=2, max_row=worksheet.max_row):
            for cell in row:
                cell.border = border
                cell.alignment = Alignment(horizontal='left')
        
        for col in ['C', 'E']:  
            for cell in worksheet[col][1:]:
                cell.number_format = '#,##0.00'
        
        for column in worksheet.columns:
            max_length = 0
            column_letter = column[0].column_letter
            for cell in column:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except:
                    pass
            adjusted_width = (max_length + 2)
            worksheet.column_dimensions[column_letter].width = adjusted_width
    
    return True

if __name__ == "__main__":
    input_pdf = "/Users/lakavathashishsainaik/Desktop/test3 (1).pdf"  
    output_excel = "punjab_sind_statement.xlsx"
    
    print(f"Extracting transactions from {input_pdf}...")
    transactions = extract_punjab_sind_transactions(input_pdf)
    
    if transactions:
        print(f"Found {len(transactions)} transactions")
        if save_to_excel(transactions, output_excel):
            print(f"Successfully saved to {output_excel}")
        else:
            print("Failed to save Excel file")
    else:
        print("No transactions found in the PDF")
