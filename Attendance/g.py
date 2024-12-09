n=int(input("n="))
val=2
for i in range(1,n+1):
    if i%2==0:
        print(i*2,end=" ")
    else:
        print(val,end=" ")
        val=val+2