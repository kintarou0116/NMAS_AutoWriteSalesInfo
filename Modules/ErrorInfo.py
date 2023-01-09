class MMASWritingSalesInfoError(Exception):
    """NAMS Module WritingSalesInfo で例外が発生しました。"""
    def __str__(self):
        return 'NAMS Module WritingSalesInfo で例外が発生しました。'

class MMASGetinsertAbuildSalesDataError(Exception):
    """NAMS Module get_insert_abuild_sales_data で例外が発生しました。"""
    def __str__(self):
        return 'NAMS Module  get_insert_abuild_sales_data で例外が発生しました。'

class GoogleApiProcError(MMASWritingSalesInfoError):
    """GoogleAPIサービスインスタンスの取得に失敗しました。"""
    def __str__(self):
        return 'GoogleAPIサービスインスタンスの取得に失敗しました。'

class ConfigError(MMASWritingSalesInfoError):
    """NMAS Module Config で例外が発生しました。"""
    def __str__(self):
        return 'NMAS Module Config で例外が発生しました。'

class IdsLogReadingError(MMASWritingSalesInfoError):
    """挿入IDファイルから情報を取得する際に例外が発生しました。"""
    def __str__(self):
        return '挿入IDファイルから情報を取得する際に例外が発生しました。'

class AbuildSalesInfoMailNotMatchError(MMASWritingSalesInfoError):
    """GMAILから取得したいメールが1通も検出されませんでした。"""
    def __str__(self):
        return 'GMAILから取得したいメールが1通も検出されませんでした。'

class GetSalesDataProcError(MMASWritingSalesInfoError):
    """メールから情報を取得する際に例外が発生しました。"""
    def __str__(self):
        return 'メールから情報を取得する際に例外が発生しました。'

class MailInfoFormatError(MMASWritingSalesInfoError):
    """メール内の情報フォーマットが正しくありません。"""
    def __str__(self):
        return 'メール内の情報フォーマットが正しくありません。'

class SlackNotificationError(MMASWritingSalesInfoError):
    """Slack通知をする際にエラーになりました。"""
    def __str__(self):
        return 'Slack通知をする際にエラーになりました。'
