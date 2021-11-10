import os
import time

def printdoc(filepath):
    os.startfile(filepath, "print")
    time.sleep(2)
    
mydir = os.path.realpath(os.path.join(os.getcwd(), os.path.dirname(__file__)))
mydoc = f"{mydir}\DL_Stool.docx"

printdoc(mydoc)


#webbrowser.open(f"{mydir}\DL_Stool.pdf")
#webbrowser.open(f"{mydir}\PMC_Stool.pdf")

