import pandas as pd

def main():
    # 讀取名為 '上市公司資料.csv' 的 CSV 檔案，並將其內容載入到 DataFrame (df) 中。
    df = pd.read_csv('上市公司資料.csv')

    # 移除 DataFrame (df) 中含有 NaN (缺失值) 的所有行，並將結果儲存到新的 DataFrame (df1) 中。
    # dropna() 會直接將有NaN的row移除
    df1 = df.dropna()

    # 使用 reindex 重新排列 DataFrame (df1) 的欄位順序，並只保留指定的欄位。
    # 新的 DataFrame (df2) 將只包含 '公司代號'、'出表日期'、'公司名稱'、'產業別'、'營業收入-當月營收' 和 '營業收入-上月營收' 這六個欄位。
    # 其餘欄位將被丟棄。
    df2 = df1.reindex(columns=['公司代號','出表日期','公司名稱','產業別','營業收入-當月營收','營業收入-上月營收'])

    # 使用 rename() 方法更改 DataFrame (df2) 中欄位的名稱。
    # 將 '營業收入-上月營收' 更名為 '上月營收'，'營業收入-當月營收' 更名為 '當月營收'。
    # 修改後的欄位名稱會儲存到新的 DataFrame (df3) 中。
    df3 = df2.rename(columns={
        '營業收入-上月營收':'上月營收',
        '營業收入-當月營收':'當月營收'
        })

    # 將 DataFrame (df3) 的內容儲存為名為 '上市公司資料整理.csv' 的 CSV 檔案。
    # encoding='utf-8' 參數確保檔案以 UTF-8 編碼儲存，可以支援中文字。
    df3.to_csv('上市公司資料整理.csv',encoding='utf-8')

    # 將 DataFrame (df3) 的內容儲存為名為 '上市公司資料整理.xlsx' 的 Excel 檔案。
    df3.to_excel('上市公司資料整理.xlsx')

    # 印出 "存檔完成"，表示檔案儲存動作已完成。
    print("存檔完成")

# 確保只有在直接執行這個 Python 檔案時，才會執行 main() 函數。
# 如果這個檔案是被其他檔案 import 匯入時，main() 函數就不會被執行。
if __name__ == '__main__':
    main()
