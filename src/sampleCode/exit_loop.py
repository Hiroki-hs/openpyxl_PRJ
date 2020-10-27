#1:フラグを用意する方法
break_loop = False

for i in range(3):
    if i == 0:
        print("------------------外側ループ------------------")
    
    for j in range(3):
        if j == 0 and i == 0:
            print("------------------内側ループ------------------")
        
        print(i, j)

        if i == 1 and j == 1:
            print("☆breakフラグをTrueに☆")
            break_loop = True
            print("----------内側ループをbreakする----------")
            break

    if break_loop:
        print("----------外側ループをbreakする----------")
        break

#-----------------------------------------------------------------------------
print("")
print("------------------------ for/while ~ else: -----------------------------------------------------")
#for/while ~ elseを使う方法
#for/whileは、ループ内の処理が全部終わった場合にelse節へ入る
h = [[0, 1, 2], [3, 4, 5], [6, 7, 8]]

for i in range(3):
    for j in range(3):
        print('--- inner:({}, {})={}'.format(i, j, h[i][j]))
        if h[i][j] == 4:
            print('Find:{}'.format(h[i][j]))
            break
        
    else:   #forループが満了すると、このelse節が実行される
        print('continue outer')
        continue
    
    print('break outer')    #forループがbreakするとここに飛ぶ
    break