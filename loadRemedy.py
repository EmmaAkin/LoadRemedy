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




## UserID change to First Name and the Second Name

def firstName(data):
	name = data.split()
	if len(name)>1:
		return name[0], name[1]
	else:
		return name[0], name[0]




dataF['FirstName'], dataF['Last Name'] =zip(*dataF['User ID'].map(firstName))
#Drop the columns that are not used on RemedySheet

dataF.drop(["Day", "User ID", "Comment", "Serial Number", "IS-Tag(Old)"], axis=1, inplace=True, errors='ignore')

#Rename the different colum of the prep Sheet to reflect what is on Remedy

dataF1 = dataF.rename(columns={"IS tag(New)":"IS Tag", "Device S/No":"Serial Number", "Contact":"Phone Number", "Date":"Date Deployed", "Remark":"CurrentState"})


##Hostname is the combination of the site abbreviation (Determined by the location for instance for location from connect store start with CS and connection 
## points CP else it contain the right Abbreviation) and the IS tag number and the Type of the device
## for a laptop it would have LT and DT for desktop. For device that are neither laptop or desktop do not have a hostname

#using a formula  --http://stackoverflow.com/questions/21263020/pandas-update-value-if-condition-in-3-columns-are-met
#
def hostname(rowData):
	
	# Get the value of the row and remove the numbers



## Site, location, floor and region

def getLoc(data):
	gop = re.compile('Go.*p.*|Fa.*m', re.IGNORECASE)
	count = 0

	if gop.match(str(data))!=None:
		site = "Golden Plaza"
		location = "Falomo"
		region = "Lagos"
		# We can use the department to determine the floor that the User would be on
		floor = re.findall(r'\d+',str(data))
		print(type(re.findall(r'\d+',str(data))))
		return [site,floor, location,region]
	
	for key in keys:
		count = count +1


		
		if (re.compile(str(key)+".*", re.IGNORECASE)).match(str(data))!=None:
				# I could get the right location of the system
				# Using the re to remove the for example
				# Go.*P could become Gop
			#print(dataLoc["Golden Plaza"])    

			site = dataLoc[key]["Site"]  #There could be an improvement here by making the reg exp take the match and find the correct site name
			location = dataLoc[key]["Location"]
			region = dataLoc[key]["Region"]
            
			return [site,"Na", location, region]

		elif(count>=len(keys)):
			return ["Na", "Na", "Na", "Na"]
		
		else:continue

#df = pd.DataFrame({'ID':['1','Zaria Connect','gop'], 'col_1': [0,2,3], 'col_2':[1,4,5]})


#df['Site'], df["Location"], df["Region"] = zip(*df["ID"].map(getLoc))

#print(df)

print(type(dataF1["Location"]))

dataF1['Site'], dataF1["Floor"], dataF1["ActualLocation"], dataF1["Region"] = zip(*dataF1["Location"].map(getLoc))

print(dataF1)


#Cleaning the dataframe

#1. Location and chance ActualLocation to Location
#2. Day




## Hostname


#p = re.compile ('Go*.p.*', re.IGNORECASE)

#print(type(df1['Device']))

#print(dataF1)

#print(sheet))