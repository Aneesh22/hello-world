import random
import time
list =[]
x=0
for i in range(0,10):
    while x in list:
        x=random.randint(500,1000)
    list.append(x)
print sorted(list, key=int)
time.sleep(5)

