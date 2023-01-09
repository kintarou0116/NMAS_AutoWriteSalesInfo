import gmail
import extract
import make_log
import re

# メールアドレスを受け取り、該当するサービスのメールから情報を抽出し、入金通知に必要な情報のリストを返す
# ユーザーの画面の、ラベル＞新規作成から、「処理済み」という名前のラベルを作成してください


def get_info(mail_adress):

    # 最後に返したいリスト
    payment_info = []

    # APIの認証を行う
    creds = gmail.authorize()
    service = gmail.build_service(creds)

    # 検索条件につかう、件名に含まれる文字
    if mail_adress == "service-jp@paypal.com":
        subject_keyword = "支払いの受領通知"
    elif mail_adress == "noreply@coiney.com":
        subject_keyword = "お支払い完了のお知らせ"
    elif mail_adress == "invoicing@messaging.squareup.com":
        subject_keyword = "請求書の支払いを完了しました"
    else:
        print("does not exist such adress")
        return

    # 「処理済み」ラベルのidを取得
    label_id = gmail.get_label_id(service, "処理済み")

    # 検索条件の設定
    query = ''
    query += "has:nouserlabels" + " "
    query += "from:" + mail_adress + " "
    query += "subject:" + subject_keyword + " "

    # クエリにヒットしたメールをすべて取得
    messages = gmail.get_messages(service, userid="me",
                                  query=query)
    if messages == None:
        print("no mails")
        return

    # 取得したメールに対して個別に操作
    for message in messages:

        msg = gmail.get_message(service, message, userid="me")

        # メールのid,日付、本文を取得
        mail_id = message['id']
        date = gmail.get_message_date(msg)
        content = gmail.get_message_body(msg)

        # 情報のリストを返す関数の呼び出し
        info = extract.info(mail_id, mail_adress, date, content)

        # infoに欠損値がない場合の処理
        if all(info.values()):
            # 抽出情報をリストに追加
            payment_info.append(info)

            # メールに処理済みラベルをはる
            try:
                gmail.add_label(service, msg, label_id)
            except:
                make_log.insert(mail_id, 'ERROR-0300', body)
                print("can't add label")

            # ログファイルの作成
            make_log.insert(mail_id, 'SUCCESS', content)

    return payment_info
