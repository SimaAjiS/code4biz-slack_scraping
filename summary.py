import glob

import openpyxl
import pandas as pd

from scraping import excel_tabling


def main(month):
    file_path = f'src/2022-{month:02}'
    print(file_path)

    files = glob.glob(f'{file_path}/*_c4b_slack.xlsx')

    # ベース用データフレームの準備
    df_base = pd.read_excel('src/template.xlsx', index_col=None, header=0)

    for file in files:
        # 追加用データフレームの準備
        wb = openpyxl.load_workbook(file, data_only=True)
        ws = wb.active

        data = []
        for i in range(2, ws.max_row):
            datum = {
                '検索結果数': ws[f'A{i}'].value,
                '抽出日時': ws[f'B{i}'].value,
                'thread_ts': ws[f'C{i}'].value,
                '投稿日': ws[f'D{i}'].value,
                '投稿時間': ws[f'E{i}'].value,
                '投稿チャンネル': ws[f'F{i}'].value,
                '投稿者': ws[f'G{i}'].value,
                '投稿メッセージ': ws[f'H{i}'].value,
                'リンク': ws[f'I{i}'].value,
                'リンク2': ws[f'H{i}'].value,
                'リンク3': ws[f'J{i}'].value,
                'リンク4': ws[f'K{i}'].value,
                'リンク5': ws[f'L{i}'].value,
                'リンク6': ws[f'M{i}'].value
            }
            data.append(datum)

        _df = pd.DataFrame(data)
        _df = _df.dropna(how='all')

        # 2つのデータフレームを縦に結合し新たなベース用データフレームとする
        df_base = pd.concat([df_base, _df])

    return df_base, file_path


if __name__ == '__main__':
    # 集計する月を数字入力（2022年3月～2022年12月）
    month = 3

    df_base, file_path = main(month)

    df_base.to_excel(f'{file_path}/月集計.xlsx', index=False)
    excel_tabling(f'{file_path}/月集計.xlsx')
