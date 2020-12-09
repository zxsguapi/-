str2=""
str=input()
str2=str2+str+','


list=str2.split(",")
list=list[:len(list)-1]
print(len(list))
for i in list:
    print(i)
