

def label_create(search_for, replace_with):

    f = open("template.label", 'r')
    filedata = f.read()
    f.close

    newdata = filedata.replace(search_for, replace_with)

    f = open("New.label", 'w')
    f.write(newdata)
    f.close()
fname = 'cddaat, pork'
dob = '11/11/19dd55'
dos = '08/13/2018'
phys = 'robinson'

x= "Lastname, Firstname\ndd/mm/yyyy\ndd/mm/1yyy\nProvider Name"
y= ("%s\n%s\n%s\n%s" % (fname, dob, dos, phys))

label_create(x, y)


