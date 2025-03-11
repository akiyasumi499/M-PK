import pandas as pd
from openpyxl import load_workbook

customer_data = pd.read_excel("customer_data.xlsx")

purchase_data = pd.read_excel("purchase_data.xlsx")

template_path = "invoice_template.xlsx"
wb = load_workbook(template_path)
sheet = wb["請求書"]

for index, customer in customer_data.iterrows():
    sheet["B2"] = customer['顧客名']  
    sheet["B3"] = customer['住所']   
    sheet["B4"] = customer['メールアドレス'] 

    
    customer_purchases = purchase_data[purchase_data['顧客名'] == customer['顧客名']]
    start_row = 12  
    for idx, (_, item) in enumerate(customer_purchases.iterrows(), start=start_row):
        sheet[f"A{idx}"] = item['商品名']
        sheet[f"B{idx}"] = item['数量']
        sheet[f"C{idx}"] = item['単価']
        sheet[f"D{idx}"] = item['数量'] * item['単価']

    
    subtotal = customer_purchases['数量'] * customer_purchases['単価']
    tax = subtotal.sum() * 0.1
    total = subtotal.sum() + tax

    sheet["D18"] = subtotal.sum()  
    sheet["D19"] = tax  
    sheet["D20"] = total  


    invoice_excel_path = f"Invoice_{customer['顧客名']}.xlsx"
    wb.save(invoice_excel_path)

    print(f"請求書が生成されました: {invoice_excel_path}")

