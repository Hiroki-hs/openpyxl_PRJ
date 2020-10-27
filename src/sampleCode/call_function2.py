import openpyxl as xp

#関数ごと読み込み
wb = xp.load_workbook("kinou.xlsx")
#値のみ読み込み
wb2 = xp.load_workbook("kinou.xlsx", data_only=True)


#----------------------------------------------------------------
#　相対位置を設定して、コピーする
#----------------------------------------------------------------

#名前から詳細への相対位置（下方向）
relative_position = {"aクラブ":5, "サポートネット":6, "とらぶる解消":7, "なると":8, "テスト":9, "Cristiano":10}
#詳細から件数への相対位置（右方向）
relative_kensu = {"全件":2, "OK":3, "NG":4, "NGOK":5, "NT":6, "QA":7, "未実施":8, "残件":9}

#ws = wb.active("機種別")
#ws_paste = wb.create_sheet("paste")

for i in relative_kensu.values():
    print(i)
else:
    pass