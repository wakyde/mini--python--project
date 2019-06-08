# encoding:utf-8
import itchat
import requests


# 登录微信
itchat.auto_login(hotReload=True)
# 获取微信好友发的消息  根据发的消息回复
apiUrl='http://www.tuling123.com/openapi/api'
def get_info(message):

    data = {
            'key': "64e6d2f3a8e4437d96c3045f48f48187",
            'info': message,
            'userid': 'robot'
        }

    try:
        get_reply = requests.post(apiUrl,data=data).json()
        info = get_reply['text']
        print("robot reply:%s" % info)
        return info
    except:
        return

# 回复给微信好友
@itchat.msg_register(itchat.content.TEXT)
def auto_reply(msg):
    defaultReply = "我知道了"
    Nick_name = itchat.search_friends(name='芦星浩')
    real_Name = Nick_name[0]['UserName']
    # 打印好友回复的信息
    print("message:%s"%msg['Text'])
    # 调用图灵接口
    reply = get_info(msg['Text'])
    if msg['FromUserName'] == real_Name:
        itchat.send(reply, toUserName=real_Name)
itchat.run()