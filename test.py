# str='向波崔若楠陈子煊陈丹陈心如陈丹桑怡琳桑怡琳崔智鑫马琦，王昊林琳，郭星怡，陈'
import operator

result={'林琳':1,'郭星怡':17}

sorted_x=sorted(result.items(), key=lambda item:item[1], reverse=True)
print(dict(sorted_x))


    # print()
# print(str.replace(' ',''))
# # print(''.join(str.split()))


# key='马琦(物工17-1)'
# s=key[0:key.index('(')]
# print(s)