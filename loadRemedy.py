import openpyxl, pandas as pd, numpy as numpy
import re, json



##nameWb = input("Enter the file of Prep Sheet")
##sheetname = input("Enter the name of Sheet contain the data")

##if (isEmpty())
#wb = openpyxl.load_workbook("prepSheet.xlsx")

## We would get the right name of the months for this information


#sheet = wb.get_sheet_by_name("DeployApril16")

df1 = pd.read_excel("prepSheet.xlsx", "DeployApril16", na_value=["Na"]) 

dataF = pd.DataFrame(df1)


#Import the Json file that would be use for the location
'''try:
	pass
except Exception, e:
	raise
else:
	pass
finally:
	pass

'''	
#INITIALIZING THE NECESSARY FILES

with open('location.json') as json_data:
	dataLoc =json.load(json_data)

keys = list(dataLoc.keys())

print(dataLoc["GOP"])

#Rename the different colum of the prep Sheet to reflect what is on Remedy

dataF1 = dataF.rename(columns={"IS tag(New)":"IS Tag", "Device S/No":"Serial Number", "Contact":"Phone Number", "Date":"Date Deployed", "Remark":"CurrentState"})


#Drop the columns that are not used on RemedySheet

dataF1.drop(["Day", "Comment", "Serial Number", "IS-Tag(Old)"], axis=1, inplace=True, errors='ignore')

## UserID change to First Name and the Second Name


## Site, location, floor and region

def getLoc(data):
	#p = re.compile('Go.*p.*', re.IGNORECASE)
	count = 0
	
	for key in keys:
		count = count +1
		
		if (re.compile(str(key)+".*", re.IGNORECASE)).match(str(data))!=None:
				# I could get the right location of the system
				# Using the re to remove the for example
				# Go.*P could become Gop
			#print(dataLoc["Golden Plaza"])    #

			site = dataLoc[key]["Site"]  #There could be an improvement here by making the reg exp take the match and find the correct site name
			location = dataLoc[key]["Location"]
			region = dataLoc[key]["Region"]

			return [site, location, region]

		elif(count>=len(keys)):
			return ["None", "None", "None"]
		
		else:continue

#df = pd.DataFrame({'ID':['1','Zaria Connect','gop'], 'col_1': [0,2,3], 'col_2':[1,4,5]})


#df['Site'], df["Location"], df["Region"] = zip(*df["ID"].map(getLoc))

#print(df)

print(type(dataF1["Location"]))

dataF1['Site'], dataF1["ActualLocation"], dataF1["Region"] = zip(*dataF1["Location"].map(getLoc))

print(dataF1)


#Cleaning the dataframe

#1. Location and chance ActualLocation to Location
#2. Day




## Hostname


#p = re.compile ('Go*.p.*', re.IGNORECASE)

#print(type(df1['Device']))

#print(dataF1)

#print(sheet))