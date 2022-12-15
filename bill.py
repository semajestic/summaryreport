def billableHrs(timein, timeout):
    myin = timein.split(":")
    myout = timeout.split(":")
    inhr = int(myin[0])
    inmin = int(myin[1])
    outhr = int(myout[0])
    outmin = int(myout[1])
    hrs = outhr-inhr
    print("hrs:{}".format(hrs))
    
    #if inhr>=5 and (inhr<=7 and inmin)


while True:
    timein = input("in: ")
    timeout = input("out: ")
    print("timein:{} timeout:{}".format(timein,timeout))

    print(billableHrs(timein,timeout))