import pdfrw, os, io, textwrap
from reportlab.pdfgen import canvas
#user = os.getlogin()

def run(lastname, firstname, dob, dos, doctor, dl_list, out_list, kit_list, template_req, file_tag):
    canvas_data = get_overlay_standard_req(lastname, firstname, dob, dos, doctor, dl_list, out_list, kit_list)
    form = merge(canvas_data, template_path= template_req)
    save(form, filename= '%s, %s, %s.pdf' % (lastname, firstname, file_tag))


def get_overlay_standard_req (lastname, firstname, dob, dos, doctor, dl_list, out_list, kit_list) -> io.BytesIO:
    data = io.BytesIO()
    pdf = canvas.Canvas(data)
    pdf.drawString(x=78, y=325, text= lastname) #Last
    pdf.drawString(x=350, y=325, text= firstname) #First
    pdf.drawString(x=215 , y=344, text = dob)#DOB
    pdf.drawString(x=253 , y=463, text = dos)#DOS
    pdf.drawString(x=78 , y=407, text = doctor) #physician
    if len(dl_list) > 102:
        wrap_text = textwrap.wrap(dl_list, width = 100)
        pdf.drawString(x=12 , y=580, text = wrap_text[0])
        pdf.drawString(x=12, y=565, text =wrap_text[1])
    else:
        pdf.drawString(x=12 , y=580, text = dl_list)#L1 Test List
    if len(out_list) > 102:
        wrap_text = textwrap.wrap(out_list, width = 102)
        pdf.drawString(x=12 , y=535, text = wrap_text[0])
        pdf.drawString(x=12, y=520, text =wrap_text[1])
    else:
        pdf.drawString(x=12 , y=535, text = out_list) #AAAA Testlist
    pdf.drawString(x=12 , y=490, text = kit_list) #Kits Given to patient
    pdf.save()
    data.seek(0)
    return data
'''
Perhaps-> have cordinate for each test check box. Compare tests entered to tests on req,
somehow set cords for each test being set T/F. I could iterate cord vals for each item into a function but that seems a bit much

combined_list = list(set(input_testlist) & (set(req_testlist)))  <- creates a set of items in both lists, converts to a list because awkward


food: x: y:
iba: x: y:
for item in combined_list:
    {item_name global val} == True  <-default to False
    


            
    
    
'''

def merge(overlay_canvas: io.BytesIO, template_path: str) -> io.BytesIO:
    template_pdf = pdfrw.PdfReader(template_path)
    overlay_pdf = pdfrw.PdfReader(overlay_canvas)
    for page, data in zip(template_pdf.pages, overlay_pdf.pages):
        overlay = pdfrw.PageMerge().add(data)[0]
        pdfrw.PageMerge(page).add(overlay).render()
    form = io.BytesIO()
    pdfrw.PdfWriter().write(form, template_pdf)
    form.seek(0)
    return form


def save(form: io.BytesIO, filename: str):
    with open(filename, 'wb') as f:
        f.write(form.read())

if __name__ == '__main__':
    run("smith", "john", '08/08/78', '09/04/18', 'buttface', '[588C: DAT,5150: IBA]',\
    "[5094: CS, 5600: een, 5600: Cardio Profile, 5150: FH,:anotherprofileCardio Profile, 5150: Female Hormones, 5140: Male Hormones, 5093: Comprehensive thyroid proooofiel,5111:anotherprofile]",'[5200: Comp Stool, 4001: NeuroTransmitter, 1600: AAT]','./template_L1.pdf', 'L1req')
