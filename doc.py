import sys

f = open('C:/aneesh/hello.txt','a+')
lines = f.readlines()
count=0
totalcount=0
for line in lines:
    totalcount=totalcount+1
    if 'dell' in line:
        count=0
    if 'rahul' in line:
        f.write('qwerty')
        store=count
    count=count+1
print count
print store-1
print totalcount




