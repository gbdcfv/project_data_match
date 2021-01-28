# -*- coding: UTF-8 -*

list1 = [-51079, -41624, -41601, -30250, 64707, -51612]
x = []
for i in list1:
    x.append(abs(i))
n = max(x)
print(list1[x.index(n)])
