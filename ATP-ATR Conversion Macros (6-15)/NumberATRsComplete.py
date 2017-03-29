from os import listdir


csv_quant = 0
for i in listdir("P:\MIS\Sal\ATP-ATR conversion\ATR_Docs"):
	if i[-3:] == "csv":
		csv_quant += 1

print csv_quant 
