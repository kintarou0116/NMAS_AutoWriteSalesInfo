# -*- coding: utf-8 -*-
import re
import datetime
import make_log


# メイン処理
# メールアドレスで場合分けして、サービスごとに通知したい情報を抽出
# 辞書型のinfoで返す


def info(mail_id, mail_adress, date, content):

    info = {}

    # 正規表現のパターンをそれぞれ定義
    # コメントになっているところは参照例です

    # 2021/04/25 8:47:41 JST
    # XX XX様(xxx@xxx.com)から
    TRANSACTION_PATTERN_PAYPAL = r'(?P<year>[0-9]{4})/(?P<month>[0-9]{1,2})/(?P<date>[0-9]{1,2})\s(?P<hour>[0-9]{1,2}):(?P<minute>[0-9]{1,2}):(?P<second>[0-9]{1,2})'
    PAYMENT_PATTERN_PAYPAL = r'NINJAPAN株式会社様\r\n\r\n(?P<name>[^0-9]+)様\(.+?\)から'

    # ご利用日 : 2021/04/22 23:34
    # ご利用カード : VISA\r\nカード番号(下4桁) : xxxx\r\n承認番号 : xxxxxx\r\n取引番号 : XXXX-XXXX
    TRANSACTION_PATTERN_COINEY = r'ご利用日\s:\s(?P<year>[0-9]{4})/(?P<month>[0-9]{1,2})/(?P<date>[0-9]{1,2})\s(?P<hour>[0-9]{1,2}):(?P<minute>[0-9]{1,2})'
    PAYMENT_PATTERN_COINEY = r'ご利用カード\s:\s(?P<card_name>[a-zA-Z]+)(\n|\r\n|\r)カード番号\(下4桁\)\s:\s(?P<card_number>[0-9]{4})(\n|\r\n|\r)承認番号\s:\s(?P<approval_number>[0-9]+)(\n|\r\n|\r)取引番号\s:\s(?P<transaction_number>[0-9]{4}-[0-9]{4})'

    # 2021年4月20日 19:43
    # お客さま\r\nXX XX 様
    TRANSACTION_PATTERN_SQUARE = r'(?P<year>[0-9]{4})年(?P<month>[0-9]{1,2})月(?P<date>[0-9]{1,2})日\s(?P<hour>[0-9]{1,2}):(?P<minute>[0-9]{1,2})'
    PAYMENT_PATTERN_SQUARE = r'お客さま(\n|\r\n|\r)(?P<name>[^0-9]+)様'

    # サービスごとに情報を取得する関数を呼び出し、infoを作成
    if mail_adress == "service-jp@paypal.com":
        info['メール受信日時'] = get_mail_time(date)
        info['取引日時'] = get_transaction_time(
            mail_id, content, TRANSACTION_PATTERN_PAYPAL)
        info['支払者情報'] = get_payer_name(
            mail_id, content, PAYMENT_PATTERN_PAYPAL)
    elif mail_adress == "noreply@coiney.com":
        info['メール受信日時'] = get_mail_time(date)
        info['取引日時'] = get_transaction_time(
            mail_id, content, TRANSACTION_PATTERN_COINEY)
        info['支払者情報'] = get_card_info(mail_id, content, PAYMENT_PATTERN_COINEY)
    elif mail_adress == "invoicing@messaging.squareup.com":
        info['メール受信日時'] = get_mail_time(date)
        info['取引日時'] = get_transaction_time(
            mail_id, content, TRANSACTION_PATTERN_SQUARE)
        info['支払者情報'] = get_payer_name(
            mail_id, content, PAYMENT_PATTERN_SQUARE)
    else:
        print("does not exist such adress")

    return info


# メール受信日時の取得
def get_mail_time(date):
    # 文字列からタイムスタンプに変換し、タイムゾーンを消去
    date_stamp = datetime.datetime.strptime(
        date, '%a, %d %b %Y %H:%M:%S %z').replace(tzinfo=None)

    # /表記の文字列に変換
    mail_time = date_stamp.strftime('%Y/%m/%d %H:%M:%S')
    return mail_time


#　取引日時の取得
def get_transaction_time(mail_id, content, pattern):
    result = re.search(pattern, content)

    # 秒まである場合とない場合に分けてフォーマットを調整
    if result:

        if len(result.groupdict()) == 5:
            get_transaction_time = '{}/{}/{} {}:{}'.format(result.group('year'), result.group('month'), result.group(
                'date'), result.group('hour'), result.group('minute'))
        elif len(result.groupdict()) == 6:
            get_transaction_time = '{}/{}/{} {}:{}:{}'.format(result.group('year'), result.group('month'), result.group(
                'date'), result.group('hour'), result.group('minute'), result.group('second'))
        else:
            make_log.insert(mail_id, 'ERROR-0201', content)
            print("can't make transaction data")
        return get_transaction_time
    else:
        make_log.insert(mail_id, 'ERROR-0101', content)
        print("can't find transaction data")


# 支払者情報(カード)を取得
def get_card_info(mail_id, content, pattern):
    result = re.search(pattern, content)
    if result:
        try:
            card_info = '[{}]-[{}]-[{}]-[{}]'.format(result.group('card_name'), result.group(
                'card_number'), result.group('approval_number'), result.group('transaction_number'))
        except:
            make_log.insert(mail_id, 'ERROR-0203', content)
            print("can't make card data")
        else:
            return card_info
    else:
        make_log.insert(mail_id, 'ERROR-0103', content)
        print("can't find card data")


# 支払者情報(指名)を取得
def get_payer_name(mail_id, content, pattern):
    result = re.search(pattern, content)

    if result:
        try:
            payer_name = result.group('name')
        except:
            make_log.insert(mail_id, 'ERROR-0202', content)
            print("can't make payer data")
        else:
            return payer_name
    else:
        make_log.insert(mail_id, 'ERROR-0102', content)
        print("can't find payer data")
