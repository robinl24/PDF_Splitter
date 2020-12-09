import os
from tkinter import *
from tkinter.scrolledtext import ScrolledText
from tkinter import filedialog
from tkinter import ttk 
import shutil
import PyPDF2_IMPORT #PdfFileReader,PdfFileWriter
#from PyPDF2 import PdfFileReader,PdfFileWriter
import time
import tkinter as tk                     
from tkinter import ttk 



def isResume(filePath):
    pdfFileObject = open(filePath, 'rb')
    pdfReader = PyPDF2_IMPORT.PdfFileReader(pdfFileObject)
    count = pdfReader.numPages
    page = pdfReader.getPage(0)

    text= (page.extractText())
    lowerText=str(text.lower())
    #list of words/phrases to look for in the pdfText (make all lowercase)
    matches=['email', 'e-mail', '@','gpa']
    #number of char in the reduced content view
    if len(text) <=250:
        return False
    elif any([x in lowerText for x in matches]) and len(text) >=400:
        return True
    else:
        return False

def split_pdf():
    #deleted pages
    del_Pages=[]
    #stores first page
    previousFile=''

    path=file_path.strip()
    
    pdf = PyPDF2_IMPORT.PdfFileReader(path, strict=False)

    tagName= PDF_Tag.get().strip()
    


    progress.config(text='Error')
    
    

    #Tracks when to create a new folder
    numFolder=1
    folderStorage=0
    for page in range(0,pdf.getNumPages()):

        #creates folder to store files in (for zip)
        if folderStorage==0:
            folderLocation=path_folder + '/' + tagName + '_' + str(numFolder)
            try:  
                os.mkdir(folderLocation)  
            except OSError as error:  
                print(error) 


        pdf_writer = PyPDF2_IMPORT.PdfFileWriter()
        pdf_writer.addPage(pdf.getPage(page))
        output_filename = '{}/{}_page_{}.pdf'.format(folderLocation,
            tagName, page+1)
        with open(output_filename, 'wb') as out:
            pdf_writer.write(out)
            fileSize=os.stat(output_filename).st_size

        #store first page (potential Cover)
        if page==0:
            previousFile=output_filename
            #checks whether the file is a valid Resume based on file size then isResume
        #time.sleep(3)
        if page != 0 and fileSize >=50000:
            
            folderStorage+=fileSize
            
        elif page != 0 and fileSize <50000 and isResume(output_filename) == False: #biggest blank page size was from Anderson FTMBA Class of 2021 page 230 (44kb) 
            del_Pages.append(page+1)
            #print(output_filename)
            os.remove(output_filename)
            
            if page==1:
                del_Pages=[1,2]
                #print(previousFile)
                os.remove(previousFile)
        else:
            folderStorage+=fileSize

        #file size limit of 20MB reached for folder
        if folderStorage >= 19500000:
            shutil.make_archive(folderLocation, 'zip', folderLocation)
            numFolder+=1
            folderStorage=0
        
        

        #print('Created: {}'.format(output_filename))
    
    #zip final folder
    shutil.make_archive(folderLocation, 'zip', folderLocation)
    progress.config(text='Task Completed!  Auto Deleted Pages:' + str(del_Pages))


def browse_button():
    # Allow user to select file
    global file_path
    file_path = filedialog.askopenfilename()
    
    lbl1.config(text=file_path)
    

def browse_button2():
    # Allow user to select file
    global path_folder
    path_folder = filedialog.askdirectory()
    
    lbl2.config(text=path_folder)

#LEGACY-------

def Leg_browse_button():
    # Allow user to select file
    global Leg_file_path
    Leg_file_path = filedialog.askopenfilename()
    
    Leg_lbl1.config(text=Leg_file_path)

def Leg_browse_button2():
    # Allow user to select file
    global Leg_path_folder
    Leg_path_folder = filedialog.askdirectory()
    
    Leg_lbl2.config(text=Leg_path_folder)

def split_pdf_LEGACY():

    #start page for pdf
    startPage=0
    if Leg_PDF_Start_Page.get().strip() != "":
        if int(Leg_PDF_Start_Page.get().strip()) >= 0:
            startPage=int(Leg_PDF_Start_Page.get().strip())


    path=Leg_file_path.strip()
    
    pdf = PyPDF2_IMPORT.PdfFileReader(path, strict=False)

    Leg_tagName= Leg_PDF_Tag.get().strip()


    numFolder=1
    folderStorage=0
    
    for page in range(startPage,pdf.getNumPages()):

        #makes a folder to zip later
        if folderStorage==0:
            folderLocation=Leg_path_folder + '/' + Leg_tagName + '_' + str(numFolder)
            try:  
                os.mkdir(folderLocation)  
            except OSError as error:  
                print(error) 
        
        pdf_writer = PyPDF2_IMPORT.PdfFileWriter()
        pdf_writer.addPage(pdf.getPage(page))
        output_filename = '{}/{}_page_{}.pdf'.format(folderLocation,
            Leg_tagName, page+1)
        with open(output_filename, 'wb') as out:
            pdf_writer.write(out)
            folderStorage+=os.stat(output_filename).st_size
            
            #file size limit of 20MB reached for folder
            if folderStorage >= 19500000:
                shutil.make_archive(folderLocation, 'zip', folderLocation)
                numFolder+=1
                folderStorage=0
        
        

        #print('Created: {}'.format(output_filename))
    
    #zip final folder
    shutil.make_archive(folderLocation, 'zip', folderLocation)
    Leg_progress.config(text='Task Completed!')

    


root = tk.Tk() 
root.title('Qualtrics - PDF Splitter')
mygreen = "light gray"
myred = "red"

style = ttk.Style()

style.theme_create( "theme", parent="alt", settings={
        "TNotebook": {"configure": {"tabmargins": [2, 5, 2, 0] } },
        "TNotebook.Tab": {
            "configure": {"padding": [5, 1], "background": myred },
            "map":       {"background": [("selected", mygreen)],
                          "expand": [("selected", [1, 1, 1, 0])] } } } )

style.theme_use("theme")


tabControl = ttk.Notebook(root) 
  
tab1 = ttk.Frame(tabControl) 
tab2 = ttk.Frame(tabControl) 
root.configure(background='light gray')
#version
tabControl.add(tab1, text ='2.1') 
tabControl.add(tab2, text ='Legacy') 
tabControl.pack(expand = 1, fill ="both") 

Label (tab1, text="Welcome to the PDF Splitter/Uploader\nNOTE: you may need to print >> save file as pdf if there is an error", background='light gray', foreground='black', font='none 20 bold' ) .grid(row=1, column=0,sticky=W)

Label (tab1, text="Please select the location of your pdf file:", background='light gray', foreground='black', font='none 12 bold' ) .grid(row=2, column=0,sticky=W)
lbl1 = Label(master=tab1,text='                                                   ')
lbl1.grid(row=3, column=0)
browse1 = Button(tab1, text="Browse",  command=browse_button)
browse1.grid(row=3, column=1)

Label (tab1, text="Please enter the name of the tag/school name:", background='light gray', foreground='black', font='none 12 bold' ) .grid(row=4, column=0,sticky=W)
PDF_Tag=Entry(tab1)
PDF_Tag.grid(row=5, column=0, sticky=W)

Label (tab1, text="Please select the location to SAVE pdf files", background='light gray', foreground='black', font='none 12 bold' ) .grid(row=6, column=0,sticky=W)
lbl2 = Label(master=tab1,text='                                                   ')
lbl2.grid(row=7, column=0)
browse2 = Button(tab1, text="Browse",  command=browse_button2)
browse2.grid(row=7, column=1)


Label (tab1, text="", background='light gray', foreground='black', font='none 12 bold' ).grid(row=10, column=0,sticky=W)
progress=Label (tab1, text="", background='light gray', foreground='black', font='none 12 bold' )
progress.grid(row=11, column=0,sticky=W)
Label (tab1, text="", background='light gray', foreground='black', font='none 12 bold' ).grid(row=12, column=0,sticky=W)


Button(tab1, text='SUBMIT', width=8, command=split_pdf).grid(row=20,column=0, sticky=W)


#LEGACY------------------------------------

Label (tab2, text="[LEGACY]Welcome to the PDF Splitter/Uploader\nRecommended for use when pages that have resumes are deleted", background='light gray', foreground='black', font='none 20 bold' ) .grid(row=1, column=0,sticky=W)

Label (tab2, text="Please select the location of your pdf file:", background='light gray', foreground='black', font='none 12 bold' ) .grid(row=2, column=0,sticky=W)
Leg_lbl1 = Label(master=tab2,text='                                                   ')
Leg_lbl1.grid(row=3, column=0)
Leg_browse1 = Button(tab2, text="Browse",  command=Leg_browse_button)
Leg_browse1.grid(row=3, column=1)

Label (tab2, text="Please enter the name of the tag/school name:", background='light gray', foreground='black', font='none 12 bold' ) .grid(row=4, column=0,sticky=W)
Leg_PDF_Tag=Entry(tab2)
Leg_PDF_Tag.grid(row=5, column=0, sticky=W)

Label (tab2, text="Please select the location to SAVE pdf files", background='light gray', foreground='black', font='none 12 bold' ) .grid(row=6, column=0,sticky=W)
Leg_lbl2 = Label(master=tab2,text='                                                   ')
Leg_lbl2.grid(row=7, column=0)
Leg_browse2 = Button(tab2, text="Browse",  command=Leg_browse_button2)
Leg_browse2.grid(row=7, column=1)

Label (tab2, text="Add page number of where the resumes begin, excluding cover pages(optional):", background='light gray', foreground='black', font='none 12 bold' ) .grid(row=8, column=0,sticky=W)
Leg_PDF_Start_Page=Entry(tab2)
Leg_PDF_Start_Page.grid(row=9, column=0, sticky=W)

Label (tab2, text="", background='light gray', foreground='black', font='none 12 bold' ).grid(row=10, column=0,sticky=W)
Leg_progress=Label (tab2, text="", background='light gray', foreground='black', font='none 12 bold' )
Leg_progress.grid(row=11, column=0,sticky=W)
Label (tab2, text="", background='light gray', foreground='black', font='none 12 bold' ).grid(row=12, column=0,sticky=W)


Button(tab2, text='SUBMIT', width=8, command=split_pdf_LEGACY).grid(row=20,column=0, sticky=W)


root.mainloop()   