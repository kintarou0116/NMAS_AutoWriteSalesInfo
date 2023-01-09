from datetime import datetime, date, time, timezone, timedelta
import os.path
import base64
import shutil

from  Consts import NMASConstants as nmas_c
import GoogleApiProc as GPM
import ErrorInfo as NAMSError
import NotificationSlackBot

# メイン処理ルーチン
def main():
    try:
        print('StartTime - ' + datetime.now().strftime('%Y/%m/%d %H:%M:%S.%f'))
        # GMailのAPIサービスインスタンスを取得（認証済み）
        gmail = GPM.gmail
        
        # SpreadSheetsのAPIサービスインスタンスを取得（認証済み）
        sheets = GPM.sheets
        
        # メール取得処理
        proc_start_message = nmas_c.FORMAT_INS_PROC_START.format(nmas_c.STR_ORIENTATION_NOTIFICATION)
        print(proc_start_message)
        log_folder_path = nmas_c.FOLDER_PATH_LOGS + nmas_c.FOLDER_PATH_ORIENTATION_NOTIFICATION
        sheets_write_abuild_sales_data(gmail, sheets, nmas_c.STR_ABUILD_ORIENTATION_NOTIFICATION_MAIL_GET_QUERY, nmas_c.STR_ORIENTATION_NOTIFICATION, log_folder_path)
        print(nmas_c.MESSAGE_INS_PROC_END)

        print('EndTime - ' + datetime.now().strftime('%Y/%m/%d %H:%M:%S.%f'))
        
    except NAMSError.MMASWritingSalesInfoError as e:
        print(e)
        print(e.args)
        

def sheets_write_abuild_sales_data(gmail, sheets, getQuery, category, logFolderPath):
    """
    Abuildセールスデータをセールス管理スプレッドシートへ書き出し

    引数
    ---------
    gmail : Resrouce Object
        GMailのサービスを利用するためのメソッドを持つResrouce
    sheets : Resrouce Object
        Google SpreadSheetsのサービスを利用するためのメソッドを持つResrouce
    getQuery : str
        Abuidセールスに関わるメールを抽出するためのクエリ文字列
    category : str
        メールデータの種別を分ける文字列
    logFolderPath : str
        ログファイルが格納されるパス
    """
    gmail_results = gmail.users().messages().list(userId=nmas_c.STR_ME, q=getQuery, maxResults=10).execute()
    messages = gmail_results.get(nmas_c.STR_MESSAGES, [])
    if not messages:
        print(nmas_c.MESSAGE_MAIL_NOT_FOUND)
    else:
        get_sales_data_proc(messages, gmail, category, logFolderPath)
    
def get_sales_data_proc(messages, gmail, category, logFolderPath):
    """
    AbuildセールスデータをGmailのメールから取得し、
    
    引数
    ---------
    messages : list
        メールデータが分割されて格納されたlist
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
        for massage in messages:
            # 書き込み済みのメールIDの取得
            check_logdata_dict = get_inserted_logdata(logFolderPath)

            # 既に書き込み済み、または返信メールのメールIDは処理をしないため、ループをcontinueする
            search_id = massage['id']
            if search_id in check_logdata_dict:
                check_data = check_logdata_dict[search_id][nmas_c.CSV_INSERT_ID_CSV_HEADER_STATUS] 
                if check_data == nmas_c.STR_SUCCESS:
                    print(nmas_c.FORMAT_INSERT_PROCESS_CANCEL_LINE_TEXT.format(search_id, check_data))
                    continue
                elif check_data == nmas_c.STR_ERROR0101:
                    print(nmas_c.FORMAT_INSERT_PROCESS_ERROR0101_LINE_TEXT.format(search_id, check_data))
                    continue
                elif check_data == nmas_c.STR_ERROR0102:
                    print(nmas_c.FORMAT_INSERT_PROCESS_ERROR0102_LINE_TEXT.format(search_id, check_data))
                    continue
                elif check_data == nmas_c.STR_ERROR9900:
                    print(nmas_c.FORMAT_INSERT_PROCESS_ERROR9900_LINE_TEXT.format(search_id, check_data))
                    continue

            # GMailAPI経由で指定されたメールIDのデータを取得
            print(nmas_c.FORMAT_INSERT_PROCESS_START_LINE_TEXT.format(search_id))
            data_list = []
            target_massage = gmail.users().messages().get(userId=nmas_c.STR_ME, id=search_id).execute()
            if not target_massage:
                # 検索無し
                print(nmas_c.FORMAT_INSERT_PROCESS_ERROR_UNMATCH_LINE_TEXT.format(search_id))
            else:
                # メール本文と固定データを含んだ挿入データlist型オブジェクトを取得
                size = int(target_massage[nmas_c.KEY_PAYLOAD][nmas_c.KEY_BODY][nmas_c.KEY_SIZE])
                if size <= 0:
                    # error-0101で挿入IDファイルに記載
                    write_insertdata_errorlog(logFolderPath, search_id, target_massage, nmas_c.STR_ERROR0101)
                    continue
                
                # 20210323_予約確認メールが2通飛んでくる問題への対応
                # 完全一致データが取得されていたり、書き込まれていた場合はエラーデータとして除外する
                error0102 = False
                body_data = target_massage[nmas_c.KEY_PAYLOAD][nmas_c.KEY_BODY][nmas_c.KEY_DATA]
                for key in check_logdata_dict:
                    if body_data == check_logdata_dict[key][nmas_c.CSV_INSERT_ID_CSV_HEADER_BODYDATA]:
                        error0102 = True
                if error0102:
                    write_insertdata_errorlog(logFolderPath, search_id, target_massage, nmas_c.STR_ERROR0102)
                    continue
                
                # 正常メールデータをデータリストとして変換処理する
                try:
                    notif_list = get_insert_abuild_sales_data(target_massage, category)
                    # Slackボットによる通知
                    # 本番環境
                    #NotificationSlackBot.slack_bot_notification(notif_list, 'C02V5S29HMH')
                    # 開発環境
                    NotificationSlackBot.slack_bot_notification(notif_list, 'C01SMPPDU8Z')
                    
                    # 20210323_予約確認メールが2通飛んでくる問題への対応
                    # 取得時点のデータを挿入IDファイルに記載
                    insert_id_file_path = logFolderPath + nmas_c.FILE_NAME_DATA_INSERT_ID_CSV
                    timestamp = datetime.now().strftime("%Y/%m/%d %H:%M:%S.%f")
                    insert_data = nmas_c.FORMAT_LOG_LINE.format(search_id, nmas_c.STR_SUCCESS, body_data, timestamp)
                    with open(insert_id_file_path, mode='a') as f:
                        f.write(insert_data + nmas_c.STR_LF)
                except Exception as e:
                    # 正常メールデータではなかったことを示すためのエラー出力処理
                    print(e)
                    print(e.args)
                    write_insertdata_errorlog(logFolderPath, search_id, target_massage, nmas_c.STR_ERROR9900)
                    continue
                
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
    body_data = targetMassage[nmas_c.KEY_PAYLOAD][nmas_c.KEY_BODY][nmas_c.KEY_DATA]
    insert_data = nmas_c.FORMAT_LOG_LINE.format(searchId, errorType, body_data, timestamp)
    with open(insert_id_file_path, mode='a') as f:
        f.write(insert_data + nmas_c.STR_LF)

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
        #responce = sheets.spreadsheets().values().append(spreadsheetId=GPM.UPDATE_SHEETS_ID, range=nmas_c.STR_ABUILD_SALES_INFO_SEETS, valueInputOption='USER_ENTERED', insertDataOption='INSERT_ROWS', body=insertDataList[insert_id][1]).execute()
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
        
        # メール本文をBase64にてデコード
        defore_body_data = abuildMassage[nmas_c.KEY_PAYLOAD][nmas_c.KEY_BODY][nmas_c.KEY_DATA]
        body_data = base64.b64decode(defore_body_data, altchars=b'-_').decode()
        body_data_split = body_data.splitlines()
        
        # メール文面のフォーマット最初の行インデックスを取得
        start_index = None
        for idx, val in enumerate(body_data_split):
            if val.startswith(nmas_c.STR_ORIENTATION_MAIL_FORMAT_START_BORDER):
                start_index = idx
                break
        if start_index == None:
            raise NAMSError.MailInfoFormatError(Exception)
        
        # メール文面の最終行のインデックスを取得
        last_index = None
        cnt = 0
        for idx, val in enumerate(body_data_split):
            if val.startswith(nmas_c.STR_NOTIFICATION_END_TITLE):
                last_index = idx
                break
        if last_index == None:
            raise NAMSError.MailInfoFormatError(Exception)  
        
        # メールデータからSlack通知に必要な情報を抽出
        notif_list = []
        print(body_data_split[start_index])
        notif_list.append(nmas_c.FORMAT_NOTIFICATION_TITLE.format('TR-OT'))
        print(body_data_split[start_index + 1])
        notif_list.append(body_data_split[start_index + 1])
        print(body_data_split[start_index + 2])
        notif_list.append(body_data_split[start_index + 2])
        print(body_data_split[start_index + 3])
        notif_list.append(body_data_split[start_index + 3])
        if body_data_split[start_index + 4] == nmas_c.STR_MAIL_INFO_COACHING_AREA_BORDER: # エラーになる文面は処理しない
            # passのところにはいる処理
            print(nmas_c.STR_MAIL_INFO_COACHING_AREA_BORDER)
            notif_list.append(nmas_c.STR_MAIL_INFO_COACHING_AREA_BORDER)
            for idx in range(5, 12):
                print('[coaching[' + str(idx-4) + ']' + str(body_data_split[start_index + idx].encode('cp932', "ignore")))
                notif_list.append(body_data_split[start_index + idx])
            print(nmas_c.STR_MAIL_INFO_COACHING_AREA_BORDER)
            notif_list.append(nmas_c.STR_MAIL_INFO_COACHING_AREA_BORDER)
            # passのところにはいる処理
        print(body_data_split[last_index])
        notif_list.append(nmas_c.STR_NOTIFICATION_END_TITLE)
        
        return notif_list

    except Exception as e:
        raise e

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
