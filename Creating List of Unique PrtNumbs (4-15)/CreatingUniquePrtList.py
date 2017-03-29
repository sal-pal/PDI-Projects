"""Create unique list of part numbers"""


uniquePrts = []

with open("C:\Users\palomis\Desktop\UniquePrts\UniqueTrimPrts.txt","r") as f:
    for partNumb in f:
        if partNumb not in uniquePrts:
            uniquePrts.append(partNumb)



with open("C:\Users\palomis\Desktop\UniquePrts\UniqueFinalPrts.txt","a") as f:
    for prt in uniquePrts:
         f.write(prt)

        
        
    
            
    
            
        
