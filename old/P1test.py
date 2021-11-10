import sys, os, time, datetime
from Dymo import Dymoprint as DP
#Since none of this job is really logical we end up with a bunch of conditional statements to work through it

mydir = os.path.realpath(os.path.join(os.getcwd(), os.path.dirname(__file__)))
DL_Stool = f"{mydir}\Documents to Print\DL_Stool.docx"
PMC_Stool = f"{mydir}\Documents to Print\PMC_Stool.docx"
DL_Blank = f"{mydir}\Documents to Print\DL_blank.docx"
PMC_Blank = f"{mydir}\Documents to Print\PMC_blank.docx"
PAL_Blank = f"{mydir}\Documents to Print\PAL_blank.docx"

templateL = fr"{mydir}\Dymo\1by3template.label"
newL = fr"{mydir}\Dymo\New.label"

class Account:
    def __init__(self, name, username, password, stoolDoc, DLreq):
        self.name = name 
        self.username = username
        self.password = password
        self.stool = stoolDoc
        self.blankreq = DLreq
#printing document function, takes string arg of the full filepath of the file
def printdoc(filepath):
    os.startfile(filepath, "print")
    time.sleep(2)


def label_create(search_for, replace_with):

    f = open(templateL, 'r')
    filedata = f.read()
    f.close

    newdata = filedata.replace(search_for, replace_with)

    f = open(newL, 'w')
    f.write(newdata)
    f.close()
#initialize classes for each billing type
DL = Account("DL", "1485", "lab999", DL_Stool, DL_Blank)
PAL = Account("PAL", "1485", "lab999", DL_Stool, PAL_Blank)
PMC = Account("PMC", "1248", "PMC1234", PMC_Stool, PMC_Blank)
NewPrac = Account("New Practice", "1322", "nppassword", DL_Stool, PAL_Blank)

#todays date mm-dd-yy
date = datetime.date.today().strftime('%m/%d/%y')

#User enters patient F and L name, dob, doctor
firstname = input("Patient's First Name:  ")
lastname = input("Patient's Last Name:  ")
dob = input("Patient's date of birth (dd/mm/yyyy):  ")
doctor = input("Physician's last name:  ")

#combine l and f name
fullname = (lastname + ", " + firstname)

#this is where we create the new label as New.label will maybe later have each patient load to its own label so it can be reprinted 
label_create("Lastname, Firstname\ndd/mm/yyyy\ndd/mm/1yyy\nProvider Name", ("%s\n%s\n%s\n%s" % (fullname, dob, date, doctor)))

print("\nYou entered: %s, %s: %s with Dr: %s\n" % (firstname, lastname, dob, doctor))

print("Account options: ")
billTypes = {"0": "None Listed", "1": "DL", "2":  "PMC", "3":  "Self Pay","4": "PAL", "5": "New Practice"}

for i in billTypes:
    print(i,": ",  billTypes[i])

billSelect = input("\nPlease enter the number corresponding to the listing on the top of the order form:  ")

for key, val in billTypes.items():
    if billSelect == key:
        print("\nYou selected %s!" % (val, ))
        billMethod = val
        
#A couple rules to narrow down accounts to one of the four types and check for DLs that have BCBS (which end up being PAL no matter what)
if billMethod == "Self Pay":
    billMethod = "PMC"

if billMethod == "DL":
    BCTF = input("\nIs the insurance type BlueCross Blue Sheild? y/n: ")
    BCTF = BCTF.lower()
    if BCTF == "y":
        print("\nThis is supposed to be a PAL! Changing Billing Type to PAL.")
        billMethod = "PAL"
    else:
        print("\nGreat, continuing to requisition printing...")
        pass
#Print statement for None Listed to instruct user to take chart to billing manager
if billMethod == "None Listed":
    print("\nIllegal Operation!, Please return chart to billing manager to retreive this information!")
    quit = input("\nPress any key to quit...")
    sys.exit()

print("\nPlease make a copy of all docs attached to the front of the chart including insurance verification if given")

 
if billMethod == "DL":
    account = DL

if billMethod == "PAL":
    account = PAL

if billMethod == "PMC":
    account = PMC

if billMethod == "New Practice":
    account = NewPrac

#print(account.username)
#print(account.password)
#print(account.name)

DL_Testlist = {"588": "DAT Combo", "5150": "Advanced IBA", "6000": "OxLDL", "5600": "Advanced Oxidative Stress", \
               "5900": "Basic Oxidatibe Stress (urine only)","6020": "Airborne Allergy", "1600": "Advanced Adrenal Stress Test",\
               "0": "Write in test", "11": "Stool" }
DL_Selected = []
print("Please Select Your DunwoodyLabs Test by typing in the test code:")

for i in DL_Testlist:
    print(i, " :", DL_Testlist[i])


repeat = "y"    
while repeat == "y":
    sel = input("\nCode: ")
    
    if sel in DL_Testlist:
        DL_Selected.append(sel)
    else:
        print("that is not on the list...try again")
                
    repeat = input("\nAdd another test? y/n: ")
    repeat = repeat.lower()
    
print(DL_Selected)

#Print stool req based on account
if "11" in DL_Selected:
    print("Printing stool requisition to default printer...")
    #printdoc(account.stool)

if "588" in DL_Selected:
    print("Printing Dunwoody requisition to default printer...")
    #printdoc(account.blankreq)

baselabelnum = 0
extralabelum = 0
for test in DL_Selected:
    baselabelnum += 1
print(baselabelnum)



extralabelnum = int(input("Printing %s labels, would you like to add more? How many(type 0 for none)?: " % (baselabelnum)))

finalnum = (baselabelnum + extralabelnum)

print(finalnum)
'''
while extralabelnum >0:  
    if extralabelum < 15:
        print("printing dymo patient labels!")
        #DP.dymoprint(finalnum)
    else:
        print("Cant print more than 15 labels at once...for your own safety try again")
        extralabelnum = int(input("Printing %s labels, would you like to add more? How many(type 0 for none)?: " % (baselabelnum)))
        continue
   ''' 

input("Press any key to end")

    



