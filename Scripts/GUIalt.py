from tkinter import *
from tkinter import ttk
#import tkinter as tk

class patient():
    def __init__(self):
        self.billmethod = None
        self.lname = None
        self.fname = None
        self.dob = None
        self.doc = None
        self.gender = None
        self.testl = []

#sets each user input to the values in the patient class

#parses through list and sets each item as one member of a column in the GUI
def setgrid(list, col):
    row = 0
    
    
    for item in list:
        item.grid(row=row, column = col)
        row +=1 
        

                 
def gui():
    
    root = Tk()
    root.title("Patient Entry")
    #root.geometry("300x300")
    

    spacer1 = Label( root, text="   ")
    #billing method selected as drop down list
    billlist= ['DL', 'PAL', 'New Practice', 'PMC', 'Self Pay', 'Employee']
    Bill = Label(root, text = "Billing Method:  ")
    E0 = ttk.Combobox(root, state='readonly')
    E0['values']=billlist

    #patient info as user tpyed boxes
    Lname = Label( root, text="Last name:  ")
    E1 = Entry(root, bd =3)

    Fname = Label( root, text="First Name:  ")
    E2 = Entry(root, bd =3)

    DOB = Label( root, text="Date of birth:  ")
    E3 = Entry(root, bd =3)

    #doctors selected as dropdown list
    mylist = ['James', 'Dana']

    Doc = Label( root, text="Physician:  ")
    E4 = ttk.Combobox(root)
    E4['values']=mylist
    
    
    def printboxes():
        pat.billmethod = E0.get()
        pat.lname = E1.get()
        pat.firstname = E2.get()
        pat.dob = E3.get()
        pat.doc = E4.get()
        pat.gender = buttonstr.get()
        root.destroy() 


    #Gender selection as radio dot slection
    Gender= Label(root, text='Select gender:')
    buttonstr =  StringVar()

    buttonM = Radiobutton(root, text='Male', variable=buttonstr, value='M')
    buttonF = Radiobutton(root, text='Female', variable=buttonstr, value='F')
    
    
   # test1 = '111'
    test2='222'
    test3='333'
    test4='444'
    test5='5555'
    test6='6666'
    test7='7777'
    testl = {'5150': 'cat', 'test2': 'dog', 'test3':'astra', 'test4':'btsta', 'test5': 'sadsad', 'test6':'adsada', 'test7':'sdasa'}
    
    # {'.txt':0, '.py':0, '.c':0, '.jpeg':0}
    
    #
    

    
    
    
    
    
    
    
    submit = Button(root, text ="            Submit           ", command = printboxes, bd =3)

    col0 =[spacer1,Bill, Lname, Fname, DOB, Doc, Gender]
    col1= [spacer1,E0, E1, E2, E3, E4, buttonM, buttonF]
        
    setgrid(col0, 2)
    setgrid(col1, 3)      


    buttonM.select() #buttonF is in the col1 list since it happened to lign up better
    
    #function to create grid of tests dynamically. testlistname sets the label
    class testc:
        def __init__(self, testname, sname):
            self.name = self
            self.keyname = sname
            self.valname = testname
            self.var = StringVar()
            self.check = Checkbutton(root, text=(sname + ": " + testname), variable=self.var)
            #self.grid(row = row, column =col)
            
    def testgrid(dicty, row, testlistname):
        testl = Label( root, text= f"{testlistname} Tests:  ")
        testl.grid(row=row, column = 0)
        row +=1
        #checkl =  StringVar()
        col = 0
        '''
        for testkey, testval in dicty.items():
            testkey = testc(testval, testkey))
            print(testkey.valname + "this is it")
            testkey.check
            testkey.check.grid(row=row, column = col)
            col += 1
            #to keep from having an infinite length list catch if len > 6 and rest to 1 then go to next row.
            if col == 6:
                col = 0
                row += 1
                continue 
                '''
        for value,key in enumerate(dicty,1):
            #dicty[key] = IntVar() # set all values of the dict to intvars
            # set the variable of the checkbutton to the value of our dictionary so that our dictionary is updated each time a button is checked or unchecked
            c = Checkbutton(root, text=f"{key}: {dicty[key]}", variable= dicty[key]) 
            c.grid(row=row, column = col)
            col += 1
            #to keep from having an infinite length list catch if len > 6 and rest to 1 then go to next row.
            if col == 6:
                col = 0
                row += 1
                continue 

        #def query_include(self):
            for key, value in dicty.items():
                if value.get():
                    print (key)
                    
        def query_exclude(self):
            for key, value in dicty.items():
                if not value.get():
                    print (key)

                  
    
    
    testgrid(testl, 9, "Testy")
    submit.grid(column=3) 
    root.mainloop()
    print(vars(pat))

if __name__ == "__main__":
    pat = patient() #create pat as class patient()
    gui()
    

