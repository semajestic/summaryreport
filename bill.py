def billableHrs(timein, timeout):
    myin = timein.split(":")
    myout = timeout.split(":")
    inhr = int(myin[0])
    inmin = int(myin[1])
    outhr = int(myout[0])
    outmin = int(myout[1])
    hrs = outhr-inhr
    print("hrs:{}".format(hrs))
    
    # if inhr>=5 and inhr<7:
    #     truein = 7
    # else if inhr==7 and inmin<=15:
    #     truein = 7
    # else if inhr==7 and inmin>15 and inmin<=59:
    #     truein = 7
    # else if inhr==8:
    #     truein = 8
    # else if inhr>8 and inhr<17:
    #     truein = 9 #case for no timein si am shift
    # else if inhr>=17 and inhr<19:
    #     truein = 19
    # else if inhr==19 and inmin<=15:
    #     truein = 19
    # else if inhr==7 and inmin>15 and inmin<=59:
    #     truein = 7
    # else if inhr==8:
    #     truein = 8
    # else if inhr>8 and inhr<17:
    #     truein = 9 #case for no timein si am shift

# def billableHrs2(timein, timeout):
#     myin = timein.split(":")
#     myout = timeout.split(":")
#     inhr = int(myin[0])
#     inmin = int(myin[1])
#     outhr = int(myout[0])
#     outmin = int(myout[1])
#     hrs = outhr-inhr
#     print("hrs:{}".format(hrs))
    
#     if inhr>=5 and inhr<8:
#         shift = 1 #am shift
#         if inhr==7 and inmin>15 and inmin<=59:
#             late = 1
#         else:
#             late = 0
#     else if inhr>=17 and inhr<20:
#         shift = 2 #pm shift
#         if inhr==19 and inmin>15 and inmin<=59:
#             late = 1
#         else:
#             late = 0
#     else:
#         return 0
    
#     if shift==1:
#         if outhr>=19 and outhr<=23:
#             return 12-late
#         else if outhr>=13 and outhr<19:
#             time=outhr-7-late #halfday atleaest 6 hrs 
#             if time >=6:
#                 return 6
#             else:
#                 return 0
#         else:
#             return 0
#     else if shift==2:
#         if outhr >=7 and outhr<=11:
#             return 12-late
#         else if outhr>=1 and outhr<7:
#             time=outhr-19-late+24 #halfday atleaest 6 hrs 
#             if time >=6:
#                 return 6
#             else:
#                 return 0
#         else:
#             return 0


while False:
    timein = input("in: ")
    timeout = input("out: ")
    print("timein:{} timeout:{}".format(timein,timeout))

    print(billableHrs(timein,timeout))

for i in range(10):
    if i==5:
        continue
    print(i)