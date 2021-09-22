# To add a new cell, type '# %%'
# To add a new markdown cell, type '# %% [markdown]'
# %%
import pandas as pd
import openpyxl as px
import datetime
import calendar
import datetime
from dateutil.relativedelta import relativedelta

#----------------------
# 【１】 年月日を入力
#----------------------
invoice_month = input("請求書を作成する年月を入れてください(例：2110)")

#入力された値をdataに変換
dte = datetime.datetime.strptime(invoice_month, "%y%m")
print(dte)



#エラー処理(日付型に変換できないときに表示させたい) 
    # try:
    #     print(dte)
    # except unconverted data remains:
    #     print("正しい値を入力してください。")



#----------------------
# 【６】 日付、お振込み期限
#----------------------


#--------- 日本語でstrftimeが表示できないエラーの回避 -------------------------

import locale
locale.setlocale(locale.LC_CTYPE, "Japanese_Japan.932")

#----------------------------------------------------------------------------


#選択月の末日を習得
def EndOfMonth(dte):
    return dte.replace(day=calendar.monthrange(dte.year, dte.month)[1])

EOM = EndOfMonth(dte)
print(EOM.strftime("%Y年%m月%d日"))


#翌月末を取得
next_month = EOM + relativedelta(months=1)
# print(next_month)





# %%
#----------------------
# 【２】 請求書一覧から顧客CDをDFに取得
#----------------------
customer_df = pd.read_excel("請求一覧創造太郎.xlsx", sheet_name="顧客管理テーブル", index=0, header=1)
display(customer_df)


# %%
#----------------------
# 【３】 inputからその月のシートを取得
#----------------------

# invoice_month = input("請求書を作成する年月を入れてください(例：2110)")

sheetname = "請求一覧" + invoice_month

month_df = pd.read_excel("請求一覧創造太郎.xlsx", sheet_name=sheetname, header=2, index=False)
display(month_df)



# ４ 相手先コードのユニークを作成
sup_cd = month_df["相手先コード"].unique()
print(sup_cd)
sup_cd.sort()
print(sup_cd)

#------- 月一覧のDFからカラムごとにリスト化 --------------

odr_no = month_df["受注No."].to_list
customer = month_df["相手先コード"].to_list
product_name1 = month_df["商品名１"].to_list
product_name2 = month_df["商品名２"].to_list
en = month_df["金額（税抜き）"].to_list
bikou = month_df["備考"].to_list

#-------------------------------------------------------
print(odr_no)

print(df["受注No."][0])


# %%
#-------------------------------------
# 【７】 各会社ごとにExcelに出力
#-------------------------------------

# month_df[month_df["相手先コード"] == 1001]

import openpyxl as px


wb = px.load_workbook("請求書創造太郎.xlsx")
ws = wb["請求書元"]


j = 0

for i in range(1,10,2):
    if j == 4:
        break
    # for j in range(0, len(sup_cd)):
    ws.cell(row= i+17, column=1).value = df["受注No."][j]
    j += 1

wb.save("2021年test月分_請求書【会社名様】.xlsx")


#相手先コードが一致する行を順番に出力
    # for customer_cd in sup_cd:
    #     df = month_df[month_df["相手先コード"] == customer_cd] #相手先コードが一致するものを
    #     print("-----------------")
    #     print(df["受注No."])
    #     print(df["商品名１"])




    #         for cell in rows:
    #     for i in range(16,24):
    #             ws.cell(row = i+2, column=1, value=str(df["受注No."])) 


    # wb.save("2021年11月分_請求書【会社名様】.xlsx")


# %%
import openpyxl as px

ivl_path = "C:/Users/rxxx2/OneDrive/05.ハロトレ/赤柗/01.実制作/code/請求一覧創造太郎.xlsx"
tpl_path = "C:/Users/rxxx2/OneDrive/05.ハロトレ/赤柗/01.実制作/code/請求書創造太郎.xlsx"


wb = px.load_workbook(tpl_path)
ws = wb["請求書元"]


li = [210601, 210701, 210702]
j = 0

print(len(li))


for i in range(1,10,2):
    for j in range(0, len(li)):
        ws.cell(row = i+17, column = 1, value=li[j])
        j += 1


wb.save("2021年10月分_請求書【会社名様】.xlsx")


# %%
#日付の出力

import openpyxl as px

wb = px.load_workbook("請求書創造太郎.xlsx")
ws = wb["請求書元"]

#----------- 出力セルの定義 ----------------

date_loc = "H1"                # 日付
ivno_loc = "H2"                # 請求書No.
Co_loc = "A6"                  # 会社名
sum_loc = "G28"                # 合計金額
dl_loc = "B35"                 # 振込期限

#------------------------------------------


# 出力コード

ws[date_loc] = EOM.strftime("%Y年%m月%d日") # 日付
# ws["H2"] =                              # 請求書No.
# ws["A6"] = xxxxx + "御中"               # 会社名
# ws["G28"] =                             # 合計金額
ws[dl_loc] = "お振込み期限　　：　　" + next_month.strftime("%Y年%m月%d日")


wb.save("2021年test月分_請求書【会社名様】.xlsx")


# %%



