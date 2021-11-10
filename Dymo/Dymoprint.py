#!/usr/bin/env python

import sys
from os import path
from win32com.client import Dispatch

def dymoprint(numlabel):
        curdir = None
        if getattr(sys, 'frozen', False):
                # frozen
                curdir = path.dirname(sys.executable)
        else:
                # unfrozen
                curdir = path.dirname(path.abspath(__file__))

        mylabel = path.join(curdir,'New.label')


        try:

                labelCom = Dispatch('Dymo.DymoAddIn')
                labelText = Dispatch('Dymo.DymoLabels')
                isOpen = labelCom.Open(mylabel)
                selectPrinter = 'DYMO LabelWriter 450'
                labelCom.SelectPrinter(selectPrinter)


                labelCom.StartPrintJob()
                labelCom.Print(numlabel,False)  #this is where you change how many labels
                labelCom.EndPrintJob()
        except:
                print('An error occurred during printing.')
                sys.exit(1)

        sys.exit(0)

if __name__ == "__main__":
        dymoprint(1)

