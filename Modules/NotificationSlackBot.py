from slacker import Slacker
import slackbot_settings
import ErrorInfo as NAMSError

def slack_bot_notification(notifcationTextList):
    print('------Bot起動------')
    slack = Slacker(slackbot_settings.API_TOKEN)
    slack.chat.post_message(
        channel=slackbot_settings.CHANNEL_ID,
        text= '\n'.join(notifcationTextList),
        as_user=True
    )
    print('------Bot停止------')

def slack_bot_notification(notifcationTextList, cannnel_id=''):
    print('------Bot起動------')
    try:
        slack = Slacker(slackbot_settings.API_TOKEN)
        if cannnel_id == '':
            cannnel_id = slackbot_settings.CHANNEL_ID
        slack.chat.post_message(
            channel=cannnel_id,
            text= '\n'.join(notifcationTextList),
            as_user=True
        )
    except Exception as e:
        print(e)
        print(e.args)
        raise NAMSError.SlackNotificationError(e)
    print('------Bot停止------')
