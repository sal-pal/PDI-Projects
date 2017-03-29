from os import listdir


with open("C:\Users\palomis\Desktop\FilePaths.txt","w") as f:
    for name in listdir("C:\Users\palomis\Desktop\ATR_Docs"):
        f.write("C:/Users/palomis/Desktop/ATR_Docs/" + name + "\n")

        
