import urllib.request, re

url = urllib.request.urlopen("http://photo.sina.com.cn")
source = url.read()
# 将中文字符解码成utf-8的形式
source = source.decode('utf-8')

res = re.search(r'(<img src=\")(.*)(\" alt)', source)
link = res.groups()[1]

link_jpg = urllib.request.urlopen(link)
f = open("test.jpg", 'wb')
f.write(link_jpg.read())
f.close()
