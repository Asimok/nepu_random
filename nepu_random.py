from random import shuffle

name = ["马琦", "郭星怡", "柏乐", "陈丹", "陈子煊", "孔书晗", "李鑫", "桑怡琳", "王昊",
        "唐薪媛", "林琳", "刘永", "崔若楠", "崔智鑫", '向波', "辛千一"]
shuffle(name)
# for i in name:
#     print(i)
for i, j in enumerate(name):
    name[i] = (i + 1).__str__() + ':' + j.__str__()
# print(name)
print("---述职随机顺序---")
for i in name:
    print(i)
