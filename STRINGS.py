f = open('C:/aneesh/hello.txt','r+')
lines = f.readlines()
str=''
for line in lines:
    for ch in line:
        str = str+ch
        if ch ==" ":
            print str
            str=''
f.close()
f = open('C:/aneesh/hello.txt','r+')
lines = f.readlines()
str=''
for line in lines:
    for ch in line:
        str = str+ch
        if ch ==" ":
            print str
            str=''
f.close()