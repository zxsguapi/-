str2=""
str3=""

while(True):
    str=input()
    str3=input()
    str2=str2+str+','
    if(str3=="execute"):
        break
print(str2)