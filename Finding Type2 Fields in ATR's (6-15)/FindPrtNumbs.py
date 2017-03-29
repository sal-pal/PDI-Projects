""" Find which ATR's have a section containing both: an "Examination of Product" title and a "Pass/Fail" field type. """




from csv import reader
from os import listdir, path


#Needed to create function to access variables belonging at higher scope
def PartNumber(name,row):
    if row[1] == '2' and row[3] in examin_prod:
        #If desired conditions met, remove irrelevant chars and write to CorrectNumbs.txt.
        scriptpath = path.dirname(__file__)
        filename = path.join(scriptpath, "CorrectNumbs.txt")

        with open(filename,"a") as f:
            if name[-6:-5] == "NC":
                part_number = name.strip("_NC.csv")
                f.write(part_number + "\n")
                return 1 
            else:
                for char in name:
                    if char == "_":
                        under_index = name.index(char)
                        part_number = name.strip(name[under_index:])
                        f.write(part_number + "\n")
                        return 1

    elif row[1] == "2" and row[3] in examin_prod and row[5] in numbers and row[7] in numbers:
        #If desired conditions not met, strip irrelevant chars and write to IncorrectNumbs.txt.
        scriptpath = path.dirname(__file__)
        filename = path.join(scriptpath, "IncorrectNumbs.txt")

        with open(filename,"a") as f:
            if name[-6:-5] == "NC":
                part_number = name.strip("_NC.csv")
                f.write(part_number + "\n")
                return 1 
            else:
                for char in name:
                    if char == "_":
                        under_index = name.index(char)
                        part_number = name.strip(name[under_index:])
                        f.write(part_number + "\n")
                        return 1 


                     

                
numbers = []
for i in range(10):
    numbers.append(str(i))


examin_prod = ["Examination of Product","Examination Of Product","examination of product"]

#Iterate through each csv 
for name in listdir("C:\Users\palomis\Desktop\LocallySavedCSV"):
    #Cryptic way to open a file, but only way to do it.
    scriptpath = path.dirname(__file__)
    filename = path.join(scriptpath + "\ATR_Docs", name)
    with open(filename,"r") as csv:
        #Iterate through each row and check if Pass/Fail Field Type and desired title
        #are found.
        reader_obj = reader(csv)
        for row in reader_obj:
            outcome = PartNumber(name,row)
            if outcome == 1:
                break
            







           








