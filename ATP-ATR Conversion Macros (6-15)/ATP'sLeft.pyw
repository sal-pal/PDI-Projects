from os import listdir


#Retrieving PN's of ATR's left in the ATR Queue.    
with open("C:\Users\palomis\Desktop\Count.txt","r") as f:
    raw_list = f.readlines()
    
queue_list = []
for line in raw_list:
    queue_list.append(line.split(" ")[0])



#Retrieving the PN's of completed ATR's. 
folder_list = []
for f in listdir("P:\MIS\Sal\ATP-ATR conversion\ATR_Docs"):
    if f[-3:] == "csv":
        if f[-6:-4] == "NC":
            folder_list.append(f.strip("_NC.csv"))
            
        else:
            for char in f:
                if char == "_":
                    under_index = f.index(char)
                    folder_list.append(f.strip(f[under_index:]))



#Cecking which completed ATR's are in the Queue. 
notdone = []
for item in queue_list:
    if item not in folder_list:
        notdone.append(item)



#Retrieving the PN's of ATR's on the skip list. 
with open("C:\Users\palomis\Desktop\skippedATRs.txt",'r') as f:
    skipped = f.readlines()
    no_newline = []
    for i in skipped:
        if i[-2:] == " \n":
            numb = i.replace(" \n","")
            no_newline.append(numb)

        elif i[-1] == "\n":
            numb = i.replace("\n","")
            no_newline.append(numb)



#Determining which ATR's are not on the skip list, then writing them to ATRsLeft.txt 
with open('C:\Users\palomis\Desktop\ATRsLeft.txt','w') as f:
    count = 0
    for i in notdone:
        if i not in no_newline:
            f.write(i + '\n')
            count += 1
    if count == 0:
        f.write("There are no more ATR's left in the Queue.)   
    else:
        f.write("\n")
        f.write("\n")
        f.write("The number of ATR's left is %d" % count)
   
   
