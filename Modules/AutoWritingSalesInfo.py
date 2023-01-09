import base64
import os.path
import shutil
from datetime import date, datetime, time, timedelta, timezone
from unicodedata import name

from distutils.log import error
import ErrorInfo as NAMSError
import GoogleApiProc as GPM
import NmasAwsConfig
import NotificationSlackBot
from Consts import NMASConstants as nmas_c


class SalesDataInfo:
    """
    セールスデータをまとめるクラス
    ベース情報はAbuild就活の基本セールスデータ
    """
    def __init__(self, path = nmas_c.FOLDER_PATH_ABUILDC, query = nmas_c.STR_ABUILD_C_MAIL_GET_QUERY, category = nmas_c.STR_ABUILDC_CATEGORY):
        self.work_log_folder_path = nmas_c.FOLDER_PATH_LOGS + path
        self.get_gmail_query = query
        self.data_category = category


# メイン処理ルーチン
def main():
    try:
        print('StartTime - ' + datetime.now().strftime('%Y/%m/%d %H:%M:%S.%f'))
        # GMailのAPIサービスインスタンスを取得（認証済み）
        gmail = GPM.gmail
        
        # SpreadSheetsのAPIサービスインスタンスを取得（認証済み）
        sheets = GPM.sheets
        
        # 就活セールスデータリスト生成
        sales_data_info_list = [
            SalesDataInfo(nmas_c.FOLDER_PATH_ABUILDC, nmas_c.STR_ABUILD_C_MAIL_GET_QUERY, nmas_c.STR_ABUILDC_CATEGORY),
            ]
        
        # 就活セールスデータ入力処理
        for sales_data_info in sales_data_info_list:
            print(nmas_c.FORMAT_INS_PROC_START.format(sales_data_info.data_category))
            #log_folder_path = sales_data_info.work_log_folder_path
            gmail_results = gmail.users().messages().list(userId=nmas_c.STR_ME, q=sales_data_info.get_gmail_query, maxResults=10).execute()
            messageIds = gmail_results.get(nmas_c.STR_MESSAGES, [])
            if not messageIds:
                print(nmas_c.MESSAGE_MAIL_NOT_FOUND)
            else:
                target_message_data_dict = get_target_message_data_dict(messageIds, gmail, sales_data_info.data_category, sales_data_info.work_log_folder_path)
                # 正常メールデータをデータリストとして変換処理する
                for search_id, target_message_data in target_message_data_dict.items():
                    try:
                        # セールスデータシート出力用にデータを加工
                        target_message = target_message_data[0]
                        body_data = target_message_data[1]
                        data_list = get_insert_abuild_sales_data(target_message, sales_data_info.data_category)
                        
                        # 20210323_予約確認メールが2通飛んでくる問題への対応
                        # 取得時点のデータを挿入IDファイルに記載
                        insert_id_file_path = sales_data_info.work_log_folder_path + nmas_c.FILE_NAME_DATA_INSERT_ID_CSV
                        timestamp = datetime.now().strftime("%Y/%m/%d %H:%M:%S.%f")
                        insert_data = nmas_c.FORMAT_LOG_LINE.format(search_id, nmas_c.STR_GET, body_data, timestamp)
                        with open(insert_id_file_path, mode='a') as f:
                            f.write(insert_data + nmas_c.STR_LF)
                        
                        # セールスデータシートに加工したデータを出力
                        data_dict = {search_id: [body_data, {'values':[data_list[0]], 'majorDimension':'ROWS'}, data_list[1]]}
                        write_sales_data_proc(data_dict, sheets, sales_data_info.work_log_folder_path)
                    except Exception as e:
                        # 正常メールデータではなかったことを示すためのエラー出力処理
                        print(e)
                        print(e.args)
                        write_insertdata_errorlog(sales_data_info.work_log_folder_path, search_id, target_message, nmas_c.STR_ERROR9900)
                        continue
                
                #ret_dict.setdefault(search_id, [target_massage[nmas_c.KEY_PAYLOAD][nmas_c.KEY_BODY][nmas_c.KEY_DATA], {'values':[data_list[0]], 'majorDimension':'ROWS'}, data_list[1]])
                
            print(nmas_c.MESSAGE_INS_PROC_END)

        print('EndTime - ' + datetime.now().strftime('%Y/%m/%d %H:%M:%S.%f'))
        
    except NAMSError.MMASWritingSalesInfoError as e:
        print(e)
        print(e.args)
        NotificationSlackBot.slack_bot_notification(e.args, cannnel_id='C01SMPPDU8Z')
        print('program_end')
   

def get_target_message_data_dict(messageIds, gmail, category, logFolderPath):
    """
    AbuildセールスデータをGmailのメールから取得し、
    
    引数
    ---------
    messageIds : list
        メールIDとスレッドIDが格納されたlist
    gmail : Resrouce Object
        GMailのサービスを利用するためのメソッドを持つResrouce
    category = '' : str
        メールデータの種別を分ける文字列
    logFolderPath = '' : str
        ログファイルが格納されるパス
    
    戻り値
    ---------
    ret_dict : dictionary
        {'メールID':
            [エンコード済みメール本文,
                {'values':書き込み辞書データ,
                    'majorDimension':'ROWS'},
                通知用文字列リスト]}
    """
    try:
        ret_dict = {}
        for massageId in messageIds:
            # 書き込み済みのメールIDの取得
            check_logdata_dict = get_inserted_logdata(logFolderPath)

            # 既に書き込み済み、または返信メールのメールIDは処理をしないため、ループをcontinueする
            search_id = massageId['id']
            if search_id in check_logdata_dict:
                check_data = check_logdata_dict[search_id][nmas_c.CSV_INSERT_ID_CSV_HEADER_STATUS] 
                if check_data == nmas_c.STR_SUCCESS:
                    success_text = nmas_c.FORMAT_INSERT_PROCESS_CANCEL_LINE_TEXT.format(search_id, check_data)
                    print(success_text)
                    continue
                elif check_data == nmas_c.STR_ERROR0101:
                    error_text = [nmas_c.FORMAT_INSERT_PROCESS_ERROR0101_LINE_TEXT.format(search_id, check_data)]
                    print(error_text)
                    NotificationSlackBot.slack_bot_notification(error_text, cannnel_id='C01SMPPDU8Z')
                    continue
                elif check_data == nmas_c.STR_ERROR0102:
                    error_text = [nmas_c.FORMAT_INSERT_PROCESS_ERROR0102_LINE_TEXT.format(search_id, check_data)]
                    print(error_text)
                    NotificationSlackBot.slack_bot_notification(error_text, cannnel_id='C01SMPPDU8Z')
                    continue
                elif check_data == nmas_c.STR_ERROR9900:
                    error_text = nmas_c.FORMAT_INSERT_PROCESS_ERROR9900_LINE_TEXT.format(search_id, check_data)
                    print(error_text)
                    continue

            # GMailAPI経由で指定されたメールIDのデータを取得
            print(nmas_c.FORMAT_INSERT_PROCESS_START_LINE_TEXT.format(search_id))
            target_massage = gmail.users().messages().get(userId=nmas_c.STR_ME, id=search_id).execute()
            if not target_massage:
                print(nmas_c.FORMAT_INSERT_PROCESS_ERROR_UNMATCH_LINE_TEXT.format(search_id))
            else:
                # メール本文と固定データを含んだ挿入データlist型オブジェクトを取得
                size = int(get_payload_mail_data_size(target_massage))
                if size <= 0:
                    # error-0101で挿入IDファイルに記載
                    write_insertdata_errorlog(logFolderPath, search_id, target_massage, nmas_c.STR_ERROR0101)
                    continue
                
                # 20210323_予約確認メールが2通飛んでくる問題への対応
                # 完全一致データが取得されていたり、書き込まれていた場合はエラーデータとして除外する
                error0102 = False
                body_data = get_payload_mail_data(target_massage)
                for key in check_logdata_dict:
                    if body_data == check_logdata_dict[key][nmas_c.CSV_INSERT_ID_CSV_HEADER_BODYDATA]:
                        error0102 = True
                if error0102:
                    write_insertdata_errorlog(logFolderPath, search_id, target_massage, nmas_c.STR_ERROR0102)
                    continue
                
                # 正常メールデータをデータリストとして変換処理する
                ret_dict.setdefault(search_id, [target_massage, body_data])
        
        return ret_dict
    except Exception as e:
        print(e)
        print(e.args)
        raise NAMSError.GetSalesDataProcError(e)

def write_insertdata_errorlog(logFolderPath, searchId, targetMassage, errorType):
    """ 
    挿入IDファイルにエラー情報を書き込み、メールデータをCSVファイルに出力する

    引数
    ---------
    logFolderPath : str
        ログファイルの格納先フォルダパス
    searchId : str
        GmailID
    targetMassage : dictionary
        メールデータを格納した辞書型オブジェクト
    errorType : str
        エラー種別文字列
    """
    # 返信メールだった場合は処理をせず、返信メールデータをCSV形式で保存
    timestamp = datetime.now().strftime("%Y/%m/%d %H:%M:%S.%f")
    path_w = logFolderPath + nmas_c.FORMAT_FILE_NAME_ERRORMAIL_FULL_CSV.format(str(searchId))
    with open(path_w, mode='w', errors='replace') as f:
        f.write(nmas_c.CSV_HEADER_ERRORMAIL_FULL_CSV + nmas_c.STR_LF)
        recursive_csv_file_write(targetMassage, f, type(targetMassage))
    
    # 挿入IDファイルに返信メールのメールIDをERRORとして明記
    insert_id_file_path = logFolderPath + nmas_c.FILE_NAME_DATA_INSERT_ID_CSV
    body_data = get_payload_mail_data(targetMassage)
    insert_data = nmas_c.FORMAT_LOG_LINE.format(searchId, errorType, body_data, timestamp)
    with open(insert_id_file_path, mode='a') as f:
        f.write(insert_data + nmas_c.STR_LF)

    # エラーログSlack通知フラグがONの場合はエラー内容をSlack通知する
    message_list = [searchId, errorType, body_data, timestamp]
    NotificationSlackBot.slack_bot_notification(message_list, cannnel_id='C01SMPPDU8Z')
    

def get_payload_mail_data(message):
    return message[nmas_c.KEY_PAYLOAD][nmas_c.KEY_PARTS][0][nmas_c.KEY_BODY][nmas_c.KEY_DATA]
def get_payload_mail_data_size(message):
    return message[nmas_c.KEY_PAYLOAD][nmas_c.KEY_PARTS][0][nmas_c.KEY_BODY][nmas_c.KEY_SIZE]

def get_inserted_logdata(folderPath):
    """ 
    重複するメールIDを検査するためのデータをログファイルから1行データの辞書型で取得

    引数
    ---------
    folderPath : str
        ログファイルの格納先フォルダパス
    戻り値
    ---------
    ret_dict : dictionary
        {'メールID':{'Status':'処理判定文字列','BodyData':'Base64でエンコーディングされたメールデータ','InsertTime'
    """
    try:
        ret_dict = {}
        filepath = folderPath + nmas_c.FILE_NAME_DATA_INSERT_ID_CSV
        if os.path.isfile(filepath):
            with open(filepath, mode='r') as f:
                ids = [s.strip() for s in f.readlines()]
                for id_data in ids:
                    data_dict = {}
                    append_data = id_data.split(',')
                    if 0 < len(append_data):
                        data_dict.setdefault(nmas_c.CSV_INSERT_ID_CSV_HEADER_STATUS, append_data[1])
                        data_dict.setdefault(nmas_c.CSV_INSERT_ID_CSV_HEADER_BODYDATA, append_data[2])
                        data_dict.setdefault(nmas_c.CSV_INSERT_ID_CSV_HEADER_INSERTTIME, append_data[3])
                        ret_dict.setdefault(append_data[0], data_dict)
        else:
            with open(filepath, mode='w') as f:
                print(nmas_c.MESSAGE_FILE_CREATE + filepath)
                header_text = Formats.brakets(nmas_c.CSV_INSERT_ID_CSV_HEADER_GMAILID) + nmas_c.STR_COMMA
                header_text += Formats.brakets(nmas_c.CSV_INSERT_ID_CSV_HEADER_STATUS) + nmas_c.STR_COMMA
                header_text += Formats.brakets(nmas_c.CSV_INSERT_ID_CSV_HEADER_BODYDATA) + nmas_c.STR_COMMA
                header_text += Formats.brakets(nmas_c.CSV_INSERT_ID_CSV_HEADER_INSERTTIME)
                f.write(header_text + nmas_c.STR_LF)
        
        return ret_dict
    except Exception as e:
        print(e)
        print(e.args)
        raise NAMSError.IdsLogReadingError(e)
    
def write_sales_data_proc(insertDataList, sheets, insertIdFileFolderPath):
    """
    抽出したセールス管理データをセールス管理スプレッドシートへ書き込む
    
    引数
    ---------
    insertDataList : dictionary
        セールスに関わるメールデータが格納されたdict
    sheets : Resrouce Object
        Google SpreadSheetsのサービスを利用するためのメソッドを持つResrouce
    insertIdFileFolderPath = '' : str
        挿入IDファイルが格納されるパス
    """
    insert_id_file_path = insertIdFileFolderPath + nmas_c.FILE_NAME_DATA_INSERT_ID_CSV
    for insert_id in insertDataList:
        # データ書き込み処理          
        # responce = sheets.spreadsheets().values().append(spreadsheetId=GPM.UPDATE_SHEETS_ID, range=nmas_c.STR_ABUILD_SALES_INFO_SEETS, valueInputOption='USER_ENTERED', insertDataOption='INSERT_ROWS', body=insertDataList[insert_id][1]).execute()
        timestamp = datetime.now().strftime("%Y/%m/%d %H:%M:%S.%f")
        
        # 挿入IDファイル上に上書きが必要なデータがあるかどうか確認
        update_index = -1
        read_lines = []
        with open(insert_id_file_path, mode='r') as f:
            read_lines = [s.strip() for s in f.readlines()]
            for id_index, id_data in enumerate(read_lines):
                check_data = id_data.split(nmas_c.STR_COMMA)
                if check_data[0] == insert_id:
                    print(nmas_c.MESSAGE_CHECK_DATA + id_data)
                    update_index = id_index

        # 挿入IDファイル上に挿入されたID情報を入力
        # 挿入IDファイル上に挿入か、上書きした際は実行結果を画面に情報を出力
        insert_data = nmas_c.FORMAT_LOG_LINE.format(insert_id, nmas_c.STR_SUCCESS, insertDataList[insert_id][0], timestamp)
        with open(insert_id_file_path, mode='w') as f:
            check_insert = True
            for line_count, line_data in enumerate(read_lines):
                if update_index == line_count:
                    print(nmas_c.MESSAGE_WRITING_UPDATE_DATA + insert_id)
                    f.write(insert_data + nmas_c.STR_LF)
                    check_insert = False
                else:
                    f.write(line_data + nmas_c.STR_LF)

            if check_insert:
                print(nmas_c.MESSAGE_WRITING_INSERT_DATA + insert_id)
                f.write(insert_data + nmas_c.STR_LF)
        
        # Slackボットによる通知
        NotificationSlackBot.slack_bot_notification(insertDataList[insert_id][2])

def get_insert_abuild_sales_data(abuildMassage, category):
    """
    メール本文からセールス管理へ書き込む情報を抽出し、list型に格納して渡す
    
    引数
    ---------
    abuildMassage : dictionary
        セールスに関わるメールデータが格納されたdict
    category : str
        セールスの分類分けを示す文字列
    
    戻り値
    ---------
    ret_list : list
        セールス管理シートの新規行1セルに相当するデータが要素として格納される
    """
    try:
        # 変数宣言
        ret_list = []
        notif_list = []
        
        # メール本文をBase64にてデコード
        defore_body_data = get_payload_mail_data(abuildMassage)
        body_data = base64.b64decode(defore_body_data, altchars=b'-_').decode()
        body_data_split = body_data.splitlines()
        
        # メール文面のフォーマット最初の行インデックスを取得
        start_index = None
        for idx, val in enumerate(body_data_split):
            if val.startswith(nmas_c.STR_MAIL_INFO_START_LINE):
                start_index = idx
                break
        if start_index == None:
            raise NAMSError.MailInfoFormatError(Exception)
        
        # メール文面の最終行のインデックスを取得
        last_index = None
        for idx, val in enumerate(body_data_split):
            if val.startswith(nmas_c.STR_MAIL_INFO_END_LINE):
                last_index = idx
                break
        if last_index == None:
            raise NAMSError.MailInfoFormatError(Exception)  
        
        # 日時データの取得
        internal_date = int("{:.10}".format(abuildMassage['internalDate']))
        received_dt = datetime.fromtimestamp(internal_date)
        
        # 格納用定数宣言
        weekday = ['月', '火', '水', '木', '金', '土', '日']
        str_month = str(received_dt.month)
        str_day = str(received_dt.day)
        str_application_time = received_dt.strftime('%H:%M')
        str_days_symbol = weekday[received_dt.weekday()]

        # 固定値をリストに格納
        # 属性　「就活」リスト
        syukatsu_category_list = [
            nmas_c.STR_ABUILDC_CATEGORY
            ]
        
        # 属性[0]
        print(category)
        if category == nmas_c.STR_ABUILDREFERRAL_CATEGORY:
            ret_list.append(nmas_c.STR_BLANK) # 後ほど処理されるため空白
        elif category in syukatsu_category_list:
            ret_list.append(nmas_c.STR_ABUILDC_CATEGORY)
        else:
            ret_list.append(category)                  
        
        ret_list.append(nmas_c.STR_BLANK)          # 流入経路詳細[]
        ret_list.append(nmas_c.STR_DEFAULT_STEP)   # ステップ[2]
        ret_list.append(nmas_c.STR_FALSE)          # 電話[3]
        ret_list.append(nmas_c.STR_FALSE)          # メール[4]
        ret_list.append(nmas_c.STR_FALSE)          # LINE@[5]
        ret_list.append(nmas_c.STR_BLANK)          # 氏名[6]
        ret_list.append(nmas_c.STR_BLANK)          # 担当[7]
        ret_list.append(nmas_c.STR_BLANK)          # 形式[8]
        ret_list.append(nmas_c.STR_BLANK)          # 年齢[9]
        ret_list.append(str_month)                 # 申込月[10]
        ret_list.append(str_day)                   # 申込日[11]
        ret_list.append(nmas_c.STR_MONTH_NF)       # 月[12]
        ret_list.append(nmas_c.STR_DAY_NF)         # 日程調整用[13]
        ret_list.append(nmas_c.STR_BLANK)          # 初回日時[14]
        ret_list.append(nmas_c.STR_BLANK)          # 次回日時[15]
        ret_list.append(nmas_c.STR_BLANK)          # 再々リクローズ[16]
        ret_list.append(nmas_c.STR_BLANK)          # 所属[17]
        ret_list.append(nmas_c.STR_BLANK)          # メアド[18]
        ret_list.append(nmas_c.STR_BLANK)          # 電話番号[19]
        ret_list.append(nmas_c.STR_FALSE)          # call[20]
        ret_list.append(nmas_c.STR_BLANK)          # call-time[21]
        ret_list.append(str_days_symbol)           # 申込曜日[22]
        ret_list.append(str_application_time)      # 申込時間[23]
        ret_list.append(nmas_c.STR_BLANK)          # 進捗状況・次のアクション[24]
        ret_list.append(nmas_c.STR_BLANK)          # 面談前事前情報[25]
        ret_list.append(nmas_c.STR_BLANK)          # 初回担当者[26]
        ret_list.append(nmas_c.STR_BLANK)          # スコアシート[27]
        ret_list.append(nmas_c.STR_BLANK)          # 初回所感[28]
        ret_list.append(nmas_c.STR_BLANK)          # 閲覧LP[29]
        ret_list.append(nmas_c.STR_FALSE)          # 情報記入[30]
        ret_list.append(nmas_c.STR_FALSE)          # 契約書[31]
        ret_list.append(nmas_c.STR_BLANK)          # 志望業界・志望企業[32]
        ret_list.append(nmas_c.STR_BLANK)          # 契約日[33]
        ret_list.append(nmas_c.STR_BLANK)          # 入金予定日[34]
        ret_list.append(nmas_c.STR_BLANK)          # 入金日[35]
        ret_list.append(nmas_c.STR_BLANK)          # 備考・反省点・改善点等[36]
        ret_list.append(nmas_c.STR_BLANK)          # 紹介者・団体[37]
        ret_list.append(nmas_c.STR_BLANK)          # 調整列[38]
        ret_list.append(nmas_c.STR_BLANK)          # 調整列[39]

        # メール本文内の情報を取得
        md_market_root_data = get_mail_data_attribute(body_data_split,nmas_c.STR_MAILDATA_KEY_MARKET_ROOT_DATA)
        md_fullname = get_mail_data_attribute(body_data_split,nmas_c.STR_MAILDATA_KEY_FULLNAME)
        md_age = get_mail_data_attribute(body_data_split,nmas_c.STR_MAILDATA_KEY_AGE)
        md_first_counseling_date = get_mail_data_attribute(body_data_split,nmas_c.STR_MAILDATA_KEY_FIRST_COUNSELING_DATE)
        md_counseling_course = get_mail_data_attribute(body_data_split,nmas_c.STR_MAILDATA_KEY_COUSELING_COURSE)
        md_university_data = get_mail_data_attribute(body_data_split,nmas_c.STR_MAILDATA_KEY_UNIVERSITY_DATA)
        md_mailaddress = get_mail_data_attribute(body_data_split,nmas_c.STR_MAILDATA_KEY_MAILADDRESS)
        md_phonenumber = get_mail_data_attribute(body_data_split,nmas_c.STR_MAILDATA_KEY_PHONENUMBER)
        md_responsible_timezone = get_mail_data_attribute(body_data_split,nmas_c.STR_MAILDATA_KEY_RESPONSIBLE_TIMEZONE)
        
        # 顧客管理シートへ書き込む内容をメールデータに合わせて変換
        # 流入経路
        ret_list[1] = md_market_root_data

        # 氏名
        ret_list[6] = md_fullname
        ret_list.append(md_fullname)

        # 年齢
        ret_list[9] = md_age
        ret_list.append(md_age)

        # 初回日時
        dt = datetime.strptime(md_first_counseling_date, '%Y/%m/%d')- datetime(1899, 12, 31)
        serial = dt.days + 1
        firstdate_nf = nmas_c.FORMAT_FIRSTDATE_NF.format(serial)
        ret_list[14] = firstdate_nf
        ret_list.append(firstdate_nf)

        # 形式
        ret_list[8] = md_counseling_course
        ret_list.append(md_counseling_course)

        # 所属
        ret_list[17] = md_university_data
        ret_list.append(md_university_data)

        # 電話番号
        ret_list[18] = md_mailaddress
        ret_list.append(md_mailaddress)

        # メールアドレス
        ret_list[19] = md_phonenumber
        ret_list.append(md_phonenumber)

        # call_time
        ret_list[21] = md_responsible_timezone
        ret_list.append(md_responsible_timezone)

        # 卒業年度
        # 区切り文字が" "か"/"で不明なためどちらにも対応
        check_grad_year = [s for s in md_university_data if s.endswith(nmas_c.STR_SLASH)]
        if 0 < len(check_grad_year):
          ret_list.append(check_grad_year[len(check_grad_year) - 1])
        
        # 紹介予約が必要になった際に解析
        # schooling_years = ''
        # referral_flag = False
        # referral_course = ''
        # referraler = ''
        # if nmas_c.STR_ABUILDREFERRAL_CATEGORY == category:
        #     referral_flag = True # 紹介予約のフラグをON

        # if nmas_c.STR_ABUILDLINECASE_CATEGORY == category:
        #     notif_list.append(nmas_c.FORMAT_NOTIFICATION_START_TITLE.format(nmas_c.STR_ABUILDLINECASE_SHORT_CATEGORY))
        # elif referral_flag:
        #     # 紹介予約限定処理
        #     print(body_data_split[last_index - 2])
        #     print(body_data_split[last_index - 1])
        #     referral_course = nmas_c.REFERRAL_SERVICE_DICTIONARY[body_data_split[last_index - 2]]
        #     referraler = body_data_split[last_index - 1]
        #     ret_list[0] = referral_course
        #     notif_list.append(nmas_c.FORMAT_REFERRAL_NOTIFICATION_START_TITLE.format(referral_course))
        #     last_index = last_index - 2
        # else:
        #     notif_list.append(nmas_c.FORMAT_NOTIFICATION_START_TITLE.format(category))
        
        # Slack通知用リスト生成
        for index, line in enumerate(body_data_split):
            if start_index <= index and index <= last_index:
                if line in nmas_c.STR_LF:
                    # 改行のみは通知処理行として、含めない
                    continue
                line_text = line
                if 0 <= line_text.find(nmas_c.STR_MAILDATA_KEY_COUSELING_COURSE):
                    # 【受講形式】対面（新宿校舎）の時は太字にする
                    if line_text not in nmas_c .STR_TYPE_ONLINE:
                        notif_list.append(Formats.asterisk_brakets(line_text) + ' <!subteam^S037N6B1D42>')
                    else:
                        notif_list.append(line_text)
                else:
                    notif_list.append(line_text)
        
        return [ret_list, notif_list]

    except Exception as e:
        raise e

def get_mail_data_attribute(args_list, attribute_name=''):
    if attribute_name != '':
        return [s for s in args_list if s.startswith(attribute_name)][0].replace(attribute_name, nmas_c.STR_BLANK)
    else:
        return nmas_c.STR_BLANK
        

def recursive_csv_file_write(foringMessage, targetFile, currentType=None, currentText=''):
    """
    List型やDictinary型に対応するための再帰的CSVファイル書き込み処理
    
    引数
    ---------
    foringMessage : dictionary, list
        書き込みたい内容を格納したイミュータブルなデータ
    targetFile : file
        書き込み先fileオブジェクト
    currentType = None : type
        再帰的に呼ばれた際の親オブジェクトのタイプ
    currentText='' : str
        再帰的に呼ばれた際の親オブジェクトまでを格納した文字列情報
    """
    count = 0
    for line in foringMessage:
        # 表示する一行データがList型だった場合は取得した要素を格納
        obj_line = ''
        if type(foringMessage) is not list:
            obj_line = foringMessage[line]
        else:
            obj_line = line
        line_type = type(obj_line)
        
        # ヘッダーに当たる前セル情報を加味して、インデックス名を表示
        str_head = currentText
        if currentType is list:
            str_head += Formats.double_quotation_brakets(str(count)) + nmas_c.STR_COMMA
        else:
            str_head += Formats.double_quotation_brakets(line) + nmas_c.STR_COMMA
        
        # 型による処理スイッチング
        if line_type is str:
            targetFile.write(str_head + Formats.double_quotation_brakets(obj_line))
            targetFile.write(nmas_c.STR_LF)
        elif line_type is int:
            targetFile.write(str_head + Formats.double_quotation_brakets(str(obj_line)))
            targetFile.write(nmas_c.STR_LF)
        elif line_type is list:
            now_header_text = str_head
            recursive_csv_file_write(obj_line, targetFile, line_type, now_header_text)
        elif line_type is dict:
            now_header_text = str_head
            recursive_csv_file_write(obj_line, targetFile, line_type, now_header_text)
        
        count += 1

class Formats:
    """
    テキストの形式を整えるメソッドが格納された静的クラス
    """
    @staticmethod
    def brakets(param_str):
        if type(param_str) is str:
            return nmas_c.FORMAT_BRAKETS.format(param_str)
        else:
            return nmas_c.FORMAT_BRAKETS.format(str(param_str))
    
    @staticmethod
    def convtext_filename(param_str):
        if type(param_str) is str:
            return nmas_c.FORMAT_FILE_NAME_CONV_TXT.format(param_str)
        else:
            return nmas_c.FORMAT_FILE_NAME_CONV_TXT.format(str(param_str))
    
    @staticmethod
    def double_quotation_brakets(param_str):
        if type(param_str) is str:
            return nmas_c.FORMAT_QUOTATION_BRAKETS.format(param_str)
        else:
            return nmas_c.FORMAT_QUOTATION_BRAKETS.format(str(param_str))

    @staticmethod
    def asterisk_brakets(param_str):
        if type(param_str) is str:
            return nmas_c.FORMAT_ASTERISK_BRAKETS.format(param_str)
        else:
            return nmas_c.FORMAT_ASTERISK_BRAKETS.format(str(param_str))

if __name__ == "__main__":
    main()
