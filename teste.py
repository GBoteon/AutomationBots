list = [1, 2, 3]
list1 = [4, 5, 6]
print(list)
x = 0

for x in range(len(list1)):
    list.append(list1[x])
print(list)