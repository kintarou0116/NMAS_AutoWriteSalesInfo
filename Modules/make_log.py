import csv
import datetime


def insert(id, status, body):
    LOG_PATH = './logdata/data_insert_id.csv'

    # ログの書き込み
    try:
        with open(LOG_PATH, 'a', encoding='UTF-8', newline='') as f:
            writer = csv.writer(f)
            writer.writerow(
                [id, status, body, datetime.datetime.now()])

    # ファイルが存在しない場合
    except FileNotFoundError:
        # フォルダ作成
        try:
            os.makedirs('./logdata')
        except FileExistsError:
            pass

        # ファイルを新規作成し、1行目に見出しを追加
        with open(LOG_PATH, mode='w') as f:
            writer = csv.writer(f)
            writer.writerow(["[GMailId]", "[Status]",
                             "[BodyData]", "[InsertTime]"])
        with open(LOG_PATH, 'a', encoding='UTF-8', newline='') as f:
            writer = csv.writer(f)
            writer.writerow(
                [id, status, body, datetime.datetime.now()])
