# ================================================================ Import Module ==================================================================
from tkinter import *
from tkinter import ttk, filedialog
import pandas as pd
import pyqrcode
from PIL import Image, ImageDraw, ImageFont, ImageTk

# ==================================================================== main functions =============================================================
font = ImageFont.truetype("projectMaterial\\OpenSans-SemiBold.ttf",size = 8)

def createQrCode():
    df = pd.read_excel(f'''{sheet_ent.get()}''')
    for index, values in df.iterrows():
        rollno = values["ID"]
        name = values["Name"]
        branch = values["Branch"]
        phone = values["Phone"]
        address = values["Address"]
        email = values["Email"]
        function = values['Function']
        dob = values['D.O.B']
        session = values['Session']
        validity = values['Validity']
        father = values['F_Name']
        blood = values['Blood']
        
        data = f"ID: {rollno} |Name: {name} |D.O.B: {dob} |Department: {branch} |Father Name: {father} |Blood Group: {blood} |Phone: {phone} |Address: {address} |Email: {email} |Function: {function} |Session: {session} |Validity: {validity}"
        
        qr_image = pyqrcode.create(data)
        qr_image.png(f"QR\\{rollno}.png")

def createIdCard(data):

    process_lbl = Label(main_frm, text = "Generated Successfully!", font = ("Calibri", "10", "bold"), bg = "dark slate grey", fg ="white")
    process_lbl.place(x=220, y=350)

    idCard = Image.open("projectMaterial\\f_template.jpg")
    
    pic = Image.open(f"photos\\{data['ID']}.jpg").resize((79, 90), Image.ANTIALIAS)
    idCard.paste(pic, (17,67,96,157))
    
    logo = Image.open("projectMaterial\\SSGI Logo.jpg").resize((141,39),Image.ANTIALIAS)
    idCard.paste(logo, (8,8,149,47))
    
    excellence = Image.open("projectMaterial\\service.jpg").resize((41,41), Image.ANTIALIAS)
    idCard.paste(excellence, (174,8,215,49))

    qrCode = Image.open(f"QR\\{data['ID']}.png")
    idCard.paste(qrCode,(110,66))
    

    draw = ImageDraw.Draw(idCard)
    draw.text((31,48), "www.srisaiuniversity.in", font = font, fill = "white")
    draw.text((119,164), f'''{str(data["ID"])}''', font = font, fill = "black")
    draw.text((119,177), f'''{data["Name"]}''', font = font, fill = "black")
    draw.text((119,191), f'''{data["Branch"]}''', font = font, fill = "black")
    draw.text((119,205), f'''{data["Function"]}''', font = font, fill = "black")
    draw.text((119,219), f'''{data["D.O.B"]}''', font = font, fill = "black")
    draw.text((119,233), f'''{data["F_Name"]}''', font = font, fill = "black")
    draw.text((119,275), f'''{str(data["Phone"])}''', font = font, fill = "black")
    draw.text((119,247), f'''{data["Address"]}''', font = font, fill = "black")
    draw.text((119,289), f'''{data["Blood"]}''', font = font, fill = "black")
    draw.text((70,315), f'''{data["Validity"]}''', font = font, fill = "red")
    
    if f'''{data["Function"]}''' == 'Teacher' :
        draw.text((223,53), "\n\n S\n S\n G\n I\n\n\n S\n T\n A\n F\n F", font = font, fill = "white")
    elif f'''{data["Function"]}''' == 'Student':
        draw.text((223,53), "\n\n S\n S\n G\n  I\n\n\n S\n T\n U\n D\n E\n N\n T", font = font, fill = "white")
    return idCard

def singleIdCard():
    process_lbl = Label(main_frm, text = "Generated Successfully!", font = ("Calibri", "10", "bold"), bg = "dark slate grey", fg ="white")
    process_lbl.place(x=220, y=350)
    df = pd.read_excel(f'''{sheet_ent.get()}''')
    ID = single_ent.get()
    for index, row in df.iterrows():
        if ID in str(row['ID']):
            singleList = []
            singleList.append(row['ID'])#0
            singleList.append(row['Name'])#1
            singleList.append(row['Branch'])#2
            singleList.append(row['Function'])#3
            singleList.append(row['D.O.B'])#4
            singleList.append(row['F_Name'])#5
            singleList.append(row['Phone'])#6
            singleList.append(row['Address'])#7
            singleList.append(row['Blood'])#8
            singleList.append(row['Session'])#9
            singleList.append(row['Email'])#10
            singleList.append(row['Validity'])#11

            def singleQR():
                rollno = singleList[0]
                name = singleList[1]
                branch = singleList[2]
                function = singleList[3]
                dob = singleList[4]
                father = singleList[5]
                phone = singleList[6]
                address = singleList[7]
                blood = singleList[8]
                session = singleList[9]
                email = singleList[10]
                validity = singleList[11]
                
                data = f"ID: {rollno} |Name: {name} |D.O.B: {dob} |Department: {branch} |Father Name: {father} |Blood Group: {blood} |Phone: {phone} |Address: {address} |Email: {email} |Function: {function} |Session: {session} |Validity: {validity}"     
                qr_image = pyqrcode.create(data)
                qr_image.png(f"QR\\{singleList[0]}.png", scale = 1)
            singleQR()

            idCard = Image.open("projectMaterial\\f_template.jpg")
    
            pic = Image.open(f"photos\\{singleList[0]}.jpg").resize((79, 90), Image.ANTIALIAS)
            idCard.paste(pic, (17,67,96,157))
        
            logo = Image.open("projectMaterial\\SSGI Logo.jpg").resize((141,39),Image.ANTIALIAS)
            idCard.paste(logo, (8,8,149,47))
        
            excellence = Image.open("projectMaterial\\service.jpg").resize((41,41), Image.ANTIALIAS)
            idCard.paste(excellence, (174,8,215,49))

            excellence = Image.open("projectMaterial\\Signature.jpg").resize((70,40), Image.ANTIALIAS)
            idCard.paste(excellence, (158,289,228,329))
    
            qrCode = Image.open(f"QR\\{singleList[0]}.png")
            idCard.paste(qrCode,(110,66))
        
            draw = ImageDraw.Draw(idCard)
            draw.text((31,48), "www.srisaiuniversity.in", font = font, fill = "white")
            draw.text((119,164), f'''{singleList[0]}''', font = font, fill = "black")
            draw.text((119,177), f'''{singleList[1]}''', font = font, fill = "black")
            draw.text((119,191), f'''{singleList[2]}''', font = font, fill = "black")
            draw.text((119,205), f'''{singleList[3]}''', font = font, fill = "black")
            draw.text((119,219), f'''{singleList[4]}''', font = font, fill = "black")
            draw.text((119,233), f'''{singleList[5]}''', font = font, fill = "black")
            draw.text((119,275), f'''{singleList[6]}''', font = font, fill = "black")
            draw.text((119,247), f'''{singleList[7]}''', font = font, fill = "black")
            draw.text((119,289), f'''{singleList[8]}''', font = font, fill = "black")
            draw.text((70,315), f'''{singleList[11]}''', font = font, fill = "red")
        
            if f'''{singleList[3]}''' == 'Teacher':
                draw.text((223,53), "\n\n S\n S\n G\n I\n\n\n S\n T\n A\n F\n F", font = font, fill = "white")
            elif f'''{singleList[3]}''' == 'Student':
                draw.text((223,53), "\n\n S\n S\n G\n  I\n\n\n S\n T\n U\n D\n E\n N\n T", font = font, fill = "white")
            idCard.save(f'''Single Card\\{singleList[0]}.jpg''')


def main():
    process_lbl = Label(main_frm, text = "processing...", font = ("Calibri", "10", "bold"), bg = "dark slate grey", fg ="white")
    process_lbl.place(x=220, y=350)
    if combo_options.get() == "Complete Sheet":
        df = pd.read_excel(f'''{sheet_ent.get()}''')
        records = df.to_dict('records') 
        
        createQrCode()  
        for record in records:
            card = createIdCard(record)
            card.save(f"Complete Cards\\{record['ID']}.jpg")
    elif combo_options.get() == "Single ID Card":
        singleIdCard()

#================================================================= Extra Buttons ================================================================

def singleCard():
    path1 = filedialog.askopenfilename(initialdir="Single Card")
    img = ImageTk.PhotoImage(Image.open(path1))

    demoImg_lbl = Label(demo_frm,image=img,bd = 1)
    demoImg_lbl.preview_image = img
    demoImg_lbl.place(x= 80, y = 150)
    
def completeSheet():
    path2 = filedialog.askopenfilename(initialdir="Complete Cards")
    img = ImageTk.PhotoImage(Image.open(path2))

    demoImg_lbl = Label(demo_frm,image=img,bd = 1)
    demoImg_lbl.preview_image = img
    demoImg_lbl.place(x= 80, y = 150)

    demoImg_lbl = Label(demo_frm,image=img,bd = 1)
    demoImg_lbl.main_img = img
    demoImg_lbl.place(x= 80, y = 150)

def preview():
    preview_image = Image.open("projectMaterial\\preview.jpg")
    resize_image = preview_image.resize((238, 334))
    img = ImageTk.PhotoImage(resize_image)

    demoImg_lbl = Label(demo_frm,image=img,bd = 1)
    demoImg_lbl.preview_image = img
    demoImg_lbl.place(x= 80, y = 150)

def template():
    image3 = Image.open("projectMaterial\\f_template.jpg")
    resize_image = image3.resize((238, 334))
    img3 = ImageTk.PhotoImage(resize_image)

    demoImg_lbl = Label(demo_frm,image=img3,bd = 1)
    demoImg_lbl.image = img3
    demoImg_lbl.place(x= 80, y = 150)

def browseSheet():
    path3 = filedialog.askopenfilename()
    sheet_ent.insert(0, path3)

#================================================================== Create Tkinter Object ===========================================================

root = Tk()
root.geometry("1366x786+0+0")
root.title('SSGI ID Card Generator')
root.iconbitmap(r'projectMaterial\\icon.ico')

#=====================================================================All frames======================================================================

title_frm = Frame(root,bg = "dark slate grey", bd = 1, relief = RIDGE)
title_frm.pack(side=TOP, fill =X)

demo_frm = Frame(root,bg = "dark slate grey", bd = 1, relief = GROOVE)
demo_frm.place(x = 980, y = 110, width = 380, height = 590)

instruction_frm = Frame(root,bg = "dark slate grey", bd = 1, relief = GROOVE)
instruction_frm.place(x = 10, y = 110, width = 380, height = 590)

main_frm = Frame(root,bg = "dark slate grey", bd = 1, relief = GROOVE)
main_frm.place(x = 400, y = 110, width = 570, height = 590)

#================================================================ Images + Title  + combo/dropdown =============================================================

site_lbl = Label(title_frm, text = "www.srisaiuniversity.in",bg = "dark slate grey", fg ="white")
site_lbl.place(x = 90, y = 80)

image1 = Image.open("projectMaterial\\SSGI Logo.jpg")
resize_image = image1.resize((290, 70))
img1 = ImageTk.PhotoImage(resize_image)

logo_lbl = Label(title_frm,image=img1,bd = 3)
logo_lbl.image = img1
logo_lbl.place(x = 20, y = 8)

image2 = Image.open("projectMaterial\\service.jpg")
resize_image = image2.resize((80, 80))
img2 = ImageTk.PhotoImage(resize_image)

logo_lbl = Label(title_frm,image=img2,bd = 3)
logo_lbl.image = img2
logo_lbl.place(x = 1262, y = 8)

image3 = Image.open("projectMaterial\\f_template.jpg")
resize_image = image3.resize((238, 334))
img3 = ImageTk.PhotoImage(resize_image)

demoImg_lbl = Label(demo_frm,image=img3,bd = 1)
demoImg_lbl.image = img3
demoImg_lbl.place(x= 80, y = 150)

demoTitle_lbl = Label(demo_frm, text = "Current Template Used", font = ("Calibri", "20", "underline"), bg = "dark slate grey", fg ="white")
demoTitle_lbl.place(x= 70, y= 50)

title_lbl = Label(title_frm, text = "ID-Card Generation System", font = ("Calibri", "20", "bold"), bg = "dark slate grey", fg ="white")
title_lbl.pack(pady = 30)

mainTitle_lbl = Label(main_frm, text = "Fill Information",font = ("Calibri", "20", "underline"), bg = "dark slate grey", fg ="white")
mainTitle_lbl.place(x= 200, y= 50)

#================================================================= Instructions======================================================================

instruction_lbl = Label(instruction_frm, text = "INSTRUCTIONS",font = ("Calibri", "20", "underline"), bg = "dark slate grey", fg ="white")
instruction_lbl.place(x= 90, y= 50)

lbl_1 = Label(instruction_frm, text = "For Single ID-Card:",font = ("Calibri", "15", "italic"), bg = "dark slate grey", fg ="white")
lbl_1.place(x= 10, y= 150)
lbl1 = Label(instruction_frm, text = "1. Browse EXCEL SHEET containing records. or\nEnter full path of that Excel Sheet.",font = ("Calibri", "10", "italic"), bg = "dark slate grey", fg ="white")
lbl1.place(x= 10, y= 180)
lbl2 = Label(instruction_frm, text = "2. Now, Select Single ID Card.",font = ("Calibri", "10", "italic"), bg = "dark slate grey", fg ="white")
lbl2.place(x= 10, y= 220)
lbl3 = Label(instruction_frm, text = "3. Enter ID number (the ID must be present in the Excel Sheet).",font = ("Calibri", "10", "italic"), bg = "dark slate grey", fg ="white")
lbl3.place(x= 10, y= 240)
lbl4 = Label(instruction_frm, text = "4. Click Submit Button.",font = ("Calibri", "10", "italic"), bg = "dark slate grey", fg ="white")
lbl4.place(x= 10, y= 260)
lbl5 = Label(instruction_frm, text = "5. Check Single ID-Card Folder.",font = ("Calibri", "10", "italic"), bg = "dark slate grey", fg ="white")
lbl5.place(x= 10, y= 280)

lbl_1 = Label(instruction_frm, text = "For Complete Sheet ID-Cards:",font = ("Calibri", "15", "italic"), bg = "dark slate grey", fg ="white")
lbl_1.place(x= 10, y= 320)
lbl1 = Label(instruction_frm, text = "1. Browse EXCEL SHEET containing records. or\nEnter full path of that Excel Sheet.",font = ("Calibri", "10", "italic"), bg = "dark slate grey", fg ="white")
lbl1.place(x= 10, y= 360)
lbl2 = Label(instruction_frm, text = "2. Now, Select Complete Sheet.",font = ("Calibri", "10", "italic"), bg = "dark slate grey", fg ="white")
lbl2.place(x= 10, y= 400)
lbl3 = Label(instruction_frm, text = "3. Leave empty the 'Enter ID Number' field.",font = ("Calibri", "10", "italic"), bg = "dark slate grey", fg ="white")
lbl3.place(x= 10, y= 420)
lbl4 = Label(instruction_frm, text = "4. Click Submit Button.",font = ("Calibri", "10", "italic"), bg = "dark slate grey", fg ="white")
lbl4.place(x= 10, y= 440)
lbl5 = Label(instruction_frm, text = "5. Check Complete Sheet Folder.",font = ("Calibri", "10", "italic"), bg = "dark slate grey", fg ="white")
lbl5.place(x= 10, y= 460)

#================================================================== Workable frame ===============================================================

sheet_lbl = Label(main_frm, text = "Excel Sheet Name ",font = ("Calibri", "10", "bold"),bg = "dark slate grey", fg ="white")
sheet_lbl.place(x=100, y= 150)
sheet_ent = Entry(main_frm, font = ("Calibri", "10", "bold"))
sheet_ent.place(x=220, y= 150, width = 190)
browse_btn = Button(main_frm, background = "orange",foreground = "White",activeforeground = "blue", activebackground = "white", text= "Browse", command = browseSheet, font=("Calibri", "9", "bold"))
browse_btn.place(x=420, y=150, width = 50, height = 20)

Choice_lbl = Label(main_frm,text ="Select! Drop Down",font = ("Calibri", "10", "bold"),bg = "dark slate grey", fg ="white")
Choice_lbl.place(x = 100, y= 190)
combo_options = ttk.Combobox(main_frm, font = ("Calibri", "10", "bold"))
combo_options['values'] = ("Single ID Card", "Complete Sheet")
combo_options.place(x=220, y= 190, width = 250)

single_lbl = Label(main_frm, text ="Enter ID Number ",font = ("Calibri", "10", "bold"),bg = "dark slate grey", fg ="white")
single_lbl.place(x = 100, y= 230)
single_ent = Entry(main_frm, font = ("Calibri", "10", "bold"))
single_ent.place(x=220, y= 230, width = 250)

submit_btn = Button(main_frm, background = "orange",foreground = "White",activeforeground = "blue", activebackground = "white", text= "Generate", command = main, font=("Calibri", "10", "bold"))
submit_btn.place(x=105, y=280, width = 365)



singleCardFolder_btn = Button(main_frm, background = "grey",foreground = "gold",activeforeground = "blue", activebackground = "white", text= "Single ID-Cards",command = singleCard, font=("Calibri", "10", "bold"))
singleCardFolder_btn.place(x=105, y=470, width = 365)

allCardFolder_btn = Button(main_frm, background = "grey",foreground = "gold",activeforeground = "blue", activebackground = "white", text= "Complete Sheet ID-Cards",command = completeSheet, font=("Calibri", "10", "bold"))
allCardFolder_btn.place(x=105, y=510, width = 365)

preview_btn = Button(demo_frm, background = "grey",foreground = "gold",activeforeground = "blue", activebackground = "white", text= "Demo",command = preview, font=("Calibri", "10", "bold"))
preview_btn.place(x=80, y=510, width = 110)

template_btn = Button(demo_frm, background = "grey",foreground = "gold",activeforeground = "blue", activebackground = "white", text= "Tempalte",command = template, font=("Calibri", "10", "bold"))
template_btn.place(x=210, y=510, width = 110)

root.mainloop()