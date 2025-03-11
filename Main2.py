import pandas as pd
import os
from openpyxl import load_workbook

# PurchaseDataListに各顧客フォルダのパスをリスト形式で格納
DataList = [
    r"C:\Users\tomo9\OneDrive\デスクトップ\IT\参考コード\競技プログラミング\Python\AtCoder\松原さん送信\to\customer_folder1",
    r"C:\Users\tomo9\OneDrive\デスクトップ\IT\参考コード\競技プログラミング\Python\AtCoder\松原さん送信\to\customer_folder2",
    r"C:\Users\tomo9\OneDrive\デスクトップ\IT\参考コード\競技プログラミング\Python\AtCoder\松原さん送信\to\customer_folder3",
]

# 請求書テンプレートの絶対パスを指定
template_path = r"C:\Users\tomo9\OneDrive\デスクトップ\IT\参考コード\競技プログラミング\Python\AtCoder\松原さん送信\invoice_template.xlsx"

# テンプレートファイルを読み込み
if not os.path.exists(template_path):
    raise FileNotFoundError(f"テンプレートファイルが見つかりません: {template_path}")
template_wb = load_workbook(template_path)
template_sheet = template_wb["請求書"]  # シート名「請求書」を指定

# 各フォルダ内のファイルを処理
for folder_path in DataList:
    if os.path.exists(folder_path):  # フォルダが存在するか確認
        file_list = os.listdir(folder_path)  # フォルダ内のファイル一覧を取得
        print(f"フォルダ: {folder_path}, ファイル数: {len(file_list)}")

        for file_name in file_list:
            file_path = os.path.join(folder_path, file_name)  # ファイルのフルパス

            # .xlsxファイルのみ処理する
            if file_name.endswith(".xlsx"):
                print(f"処理中のファイル: {file_path}")

                # 顧客データと購入データを読み込み
                customer_data = pd.read_excel(file_path, sheet_name="顧客データ")
                purchase_data = pd.read_excel(file_path, sheet_name="購入データ")

                # 必要なセルにデータを転記
                template_sheet["E8"] = customer_data.loc[0, "顧客名"]  # 顧客名
                template_sheet["K5"] = customer_data.loc[0, "住所"]    # 住所
                template_sheet["M8"] = customer_data.loc[0, "メールアドレス"]  # メールアドレス

                # テンプレートを保存（各顧客ごとに異なるファイル名で保存）
                output_path = os.path.join(folder_path, f"invoice_{file_name}.xlsx")
                template_wb.save(output_path)
                print(f"保存しました: {output_path}")

    else:
        print(f"フォルダが見つかりません: {folder_path}")
