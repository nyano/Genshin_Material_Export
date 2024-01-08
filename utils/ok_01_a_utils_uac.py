#
'''该模块用于获得UAC授权
创建时间：2023-12-27
修改时间：2023-12-27
'''
'''调用该库的模块

'''


import pyuac
import sys


def get_uac():
    if not pyuac.isUserAdmin():
        print("非管理员权限")
        try:
            pyuac.runAsAdmin()
        except:
            print("管理员权限获取失败")
            sys.exit(0)

