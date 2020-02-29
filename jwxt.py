# coding=utf-8
import os
import requests

# 网络图片爬取
url='http://jwgl.nepu.edu.cn/verifycode.servlet'
root = '验证码//'
path =root
try:
    if not os.path.exists(root):
        os.mkdir(root)
        print('创建文件夹')
    for i in range(100):
        if not os.path.exists(path+str(i)+'.jpg'):
            r = requests.get(url)
            r.raise_for_status()
            # print(r.status_code)
            with open(path+str(i)+'.jpg', 'wb') as f:
                f.write(r.content)
                f.close()
                print(str(i)+'文件保存成功')
        else:
            print('文件已存在')
except:
    print('爬取失败')
