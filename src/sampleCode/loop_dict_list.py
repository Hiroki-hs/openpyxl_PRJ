mydict = {"L":"Lemon", "O":"Orage", "G":"Grapes"}


#keyを全て抽出

print("--------------keys--------------")
for mykey in mydict.keys():
    print(mykey)
    
    
#valueを全て抽出
print("")
print("--------------values--------------")
for myval in mydict.values():
    print(myval)
    
#keyとvalueを全て抽出
print("")
print("--------------items--------------")
for myitems in mydict.items():
    print(myitems)