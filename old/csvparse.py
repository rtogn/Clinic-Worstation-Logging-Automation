import csv

def dictcreate(sheet, target_dict):
    with open(sheet, 'r') as cvsfile:
        reader = csv.reader(cvsfile)
        next(reader) #skips the header of cvs file
        #dict ={rows[0]:[rows[1], rows[2]] for rows in reader}
        for rows in reader:
            target_dict[rows[0]] = [rows[1], rows[2]] 

#for testing
if __name__ == "__main__":
    out_dict={}
    testparse('out_list.csv', out_dict)
    print(out_dict)

