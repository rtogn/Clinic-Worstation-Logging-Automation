import tkinter as tk

class GUI():
    def __init__(self, button_dict):
        self.root = tk.Tk()
        self.root.title("File Type")
        
        self.button_dict = button_dict
        row = len(self.button_dict) + 1
        
        for i,key in enumerate(self.button_dict,1):
            self.button_dict[key] = tk.IntVar() # set all values of the dict to intvars
            # set the variable of the checkbutton to the value of our dictionary so that our dictionary is updated each time a button is checked or unchecked
            c = tk.Checkbutton(self.root, text=key, variable=self.button_dict[key]) 
            c.grid(row=i, sticky=tk.W)
        
        include = tk.Button(self.root, text='Include',
                            command=self.query_include)
        include.grid(row=row, sticky=tk.W)
        
        exclude = tk.Button(self.root, text='Exclude',
                            command=self.query_exclude)
        exclude.grid(row=row, sticky=tk.E, padx=50)
        
        quit = tk.Button(self.root, text='Quit', command=self.root.quit)
        quit.grid(row=row+1, sticky=tk.W)
        
    def query_include(self):
        for key, value in self.button_dict.items():
            if value.get():
                x = key
                print (x)
                
    def query_exclude(self):
        for key, value in self.button_dict.items():
            if not value.get():
                y= key
                print (y)
        
    def main(self):
        self.root.mainloop()
        
if __name__ == "__main__":
    button_dict = {'.txt':0, '.py':0, '.c':0, '.jpeg':0}
    gui = GUI(button_dict)
    gui.main()