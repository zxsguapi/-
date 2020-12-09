str2=""
cir=6
while(cir>0):
    str=input()
    str2=str2+str+','
    cir=cir-1

list=str2.split(",")
list=list[:len(list)-1]
for i   in list:
    print(i)
print(len(list))

