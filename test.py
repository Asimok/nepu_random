str='向波崔若楠陈子煊陈丹陈心如陈丹桑怡琳桑怡琳崔智鑫马琦，王昊林琳，郭星怡，陈'
result={'林琳':2,'郭星怡':1}
for key in sorted(result, key=result.__getitem__, reverse=True):
    print(key)
    print()
# print(str.replace(' ',''))
# # print(''.join(str.split()))