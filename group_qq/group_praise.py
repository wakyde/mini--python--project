# encoding: utf-8

import os
import time
# try:
#     from common import debug, config
# except ImportError:
#     print('请在项目根目录中运行脚本')
#     exit(-1)

VERSION = "1.1.1"
countpeople = 0

# debug_switch = False    # debug 开关，需要调试的时候请改为：True
# config = config.open_accordant_config()

def touch_praise():
    global countpeople
    for touch_place in range(375, 1875, 125):
        tempstr = 'adb shell input tap 500 ' + ' ' + str(touch_place)
        os.system(tempstr)
        for touch_num in range(0, 10):
            os.system('adb shell input tap 980 625')
        countpeople = countpeople + 1
        print('%d people have been praised' % countpeople)
        time.sleep(2)
        os.system('adb shell input swipe 400 1500 800 1500')

    os.system('adb shell input swipe 500 1800 500 312 1565')

def main():
    i = 0
    print('程序版本号：{}'.format(VERSION))
    # debug.dump_device_info()
    while i < 4700:
        i = i + 1
        touch_praise()

if __name__ == '__main__':
    main()