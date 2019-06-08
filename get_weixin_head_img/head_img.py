# encoding:utf-8
import itchat

# 微信登录
itchat.auto_login(hotReload=True)

friends = itchat.get_friends()[0:]
for i in friends:
    img = itchat.get_head_img(userName=i['UserName'])
    path = "F:\\Python\Homework\get_weixin_head_img\images\\"+ i['NickName'] + '.jpg'
    print("正在下载:%s" % i['NickName'])
    
    try:
        with open(path, 'wb') as f:
            f.write(img)
    except Exception as e:
        print(repr(e))
itchat.run()

