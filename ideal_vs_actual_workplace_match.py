# import pandas as pd

# # アンケートデータの読み込み（C列とD列に数値がある）
# df = pd.read_excel('test.xlsx', usecols=[2, 3], names=['会社側', 'スタッフ側'])
# print(df.head().to_string(index=False))
# # 一致率の計算関数
# def calculate_match_rate(row):
#     return 1 - abs(row['会社側'] - row['スタッフ側']) / 4

# # 各項目ごとに一致率を計算
# df['一致率'] = df.apply(calculate_match_rate, axis=1)

# # 結果をExcelファイルに保存
# df.to_excel('一致率結果.xlsx', index=False)
# print("一致率が計算され、Excelファイルに保存されました。")


import pandas as pd
import numpy as np

# アンケートデータの読み込み（A列, B列: 設問, C列: 求職者の回答, D列: ワーカーの回答）
df = pd.read_excel('test.xlsx', usecols=[0,1,2,3], names=['会社側設問', 'スタッフ側設問', '会社側回答', 'スタッフ側回答'])

#前処理関数
def preprocess_answers(row):
    try:
        row['求職者の回答'] = float(row['求職者の回答'])
    except ValueError:
        row['求職者の回答'] = None
    try:
        row['スタッフの回答'] = float(row['スタッフの回答'])
    except ValueError:
        row['スタッフの回答'] = None
    return row

# 前処理を適用
df = df.apply(preprocess_answers, axis=1)
# 一致率の計算関数
def calculate_match_rate(row):
    try:
        # 両方の値が数値であることを確認
        # if np.isnan(row['会社側']) or np.isnan(row['スタッフ側']):
        #     return None
        return 1 - abs(row['会社側回答'] - row['スタッフ側回答']) / 4
    except:
        return None

# 各項目ごとに一致率を計算
df['一致率'] = df.apply(calculate_match_rate, axis=1)

# 一致率の平均を計算（数値のみで平均を取る）
average_match_rate = df['一致率'].dropna().mean()

# 結果をExcelファイルに保存
output_file = '一致率結果.xlsx'
df.to_excel(output_file, index=False)

# 平均一致率も出力
with pd.ExcelWriter(output_file, mode='a', engine='openpyxl') as writer:
    average_df = pd.DataFrame([{'平均一致率': average_match_rate}])
    average_df.to_excel(writer, sheet_name='Summary', index=False)

print(f"一致率が計算され、Excelファイルに保存されました。平均一致率は{average_match_rate:.3f}です。")








# import pandas as pd

# # アンケートデータの読み込み
# df = pd.read_excel('test.xlsx', usecols=[0,1,2,3], names=['会社側設問', 'スタッフ側設問', '会社側', 'スタッフ側'])

# # データの確認
# print(df.head().to_string(index=False))

# # 一致率の計算関数
# def calculate_match_rate(row):
#     return 1 - abs(row['会社側'] - row['スタッフ側']) / 4

# # 各項目ごとに一致率を計算
# df['一致率'] = df.apply(calculate_match_rate, axis=1)

# # 一致率の平均を計算
# average_match_rate = df['一致率'].mean()

# # 結果をExcelファイルに保存
# output_file = '一致率結果.xlsx'
# df.to_excel(output_file, index=False)

# # 平均一致率も出力
# with pd.ExcelWriter(output_file, mode='a', engine='openpyxl') as writer:
#     average_df = pd.DataFrame([{'平均一致率': average_match_rate}])
#     average_df.to_excel(writer, sheet_name='Summary', index=False)

# print(f"一致率が計算され、Excelファイルに保存されました。平均一致率は{average_match_rate:.2f}です。")
