import csv

def dictmake(sheet, target_dict):
    with open(sheet, 'r') as cvsfile:
        reader = csv.reader(cvsfile)
        next(reader) #skips the header of cvs file
        #dict ={rows[0]:[rows[1], rows[2]] for rows in reader}
        for rows in reader:
            target_dict[rows[0]] = [rows[1], rows[2]] 
            
def listmake(sheet, target_list):
    with open(sheet, 'r') as cvsfile:
        reader = csv.reader(cvsfile)
        next(reader) 
        for rows in reader:
            for item in rows:
                target_list.append(item)


def fulldictmake(sheet, target_dict):
	csvlist = ['out_list.csv', 'L1_list.csv', 'kit_list.csv']
	with open(sheet, 'r') as cvsfile:
		reader = csv.reader(cvsfile)
		next(reader)
		for book in csvlist:
			for rows in reader:
				target_dict[rows[0]+":"+rows[1]] = rows[2] 

if __name__ == "__main__":
    out_dict={}
    dl_dict = {}
    kit_dict = {}
    full_dict= {}
    doclist =[]
    
    dictmake('out_list.csv', out_dict)
    print(out_dict)
    print("*"*30)
    dictmake('L1_list.csv', dl_dict)
    print(dl_dict)
    print("*"*30)
    dictmake('kit_list.csv', kit_dict)
    print(kit_dict)
    print("*"*30)
    fulldictmake('kit_list.csv', full_dict)
    print(full_dict)
    print("*"*30)
    listmake('doc_list.csv', doclist)
    print(doclist)
	
