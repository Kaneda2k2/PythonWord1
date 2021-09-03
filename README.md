# PythonWord1
# UTILIZACIION DE WORD EN PYTHON - # APUNTES
##
##
##import docx
##doc = docx.Document('demo.docx')
##print(len(doc.paragraphs))
##print(doc.paragraphs[0].text)
##print(len(doc.paragraphs[1].runs))










##################Getting Full Text from a docx File

###!readDocx.py
##import docx
##
##
##def getText(filename):  # abre el documento
##    doc=docx.Document(filename)   #abrir el documento
##    fullText=[]    # matriz donde se va a llenar
##    for para in doc.paragraphs:   # por cada parrafo en el documento
##        fullText.append(para.text)  # lo sumas
##    return "\n".join(fullText)      # lo devuelves
##
##
##print(getText('demo.docx'))


####################################################
##llamandolo desde un externo
##import readDocx
##print(readDocx.getText("demo.docx"))
#########################################
####ajustar texto   getText()
#example se puede añadir doble salto en el join \n\n
#LISTADO DE STRING VALUES WORDS STYLES
##The string values for the default Word styles are as follows:
##'Normal'
##'Heading 5'
##'List Bullet'
##'List Paragraph'
##'Body Text'
##'Heading 6'
##'List Bullet 2'
##'MacroText'
##'Body Text 2'
##'Heading 7'
##'List Bullet 3'
##'No Spacing'
##'Body Text 3'
##'Heading 8'
##'List Continue'
##'Quote'
##'Caption'
##'Heading 9'
##'List Continue 2'
##'Subtitle'
##'Heading 1'
##'Intense Quote'
##'List Continue 3'
##'TOC Heading'
##'Heading 2'
##'List'
##'List Number '
##'Title
##'Heading 3'
##'List 2'
##'List Number 2'
##'Heading 4'
##'List 3'
##'List Number 3'
####################################################################


##import docx
##
##wb=docx.Document()
############################Run object text attributes
##Attribute Description
##bold The text appears in bold.
##italic The text appears in italic.
##underline The text is underlined.
##strike The text appears with strikethrough.
##double_strike The text appears with double strikethrough.
##all_caps The text appears in capital letters.
##small_caps The text appears in capital letters, with lowercase
##letters two points smaller.
##shadow The text appears with a shadow.
##outline The text appears outlined rather than solid.
##rtl The text is written right-to-left.
##imprint The text appears pressed into the page.
##emboss The text appears raised off the page in relief.

###################################################################
##########################REEE STYLANDO
##import docx
##doc=docx.Document("demo.docx")
##doc.paragraphs[0].text   # imprime la cabezera
##print(doc.paragraphs[0].style) #the exact id may be different
##doc.paragraphs[0].style="Normal"
##doc.paragraphs[1].text
##print(doc.paragraphs[1].runs[0].text,doc.paragraphs[1].runs[1].text,doc.paragraphs[1].runs[2].text,
##doc.paragraphs[1].runs[3].text)
##doc.paragraphs[1].runs[3].underline=True
##doc.save("restyled.docx")

############################WRITING WORDS DOCUMENTSSSSSSSSSS

##
##import docx
##
##doc=docx.Document()
##doc.add_paragraph("Hello world")
##doc.add_paragraph("Hello world")
##doc.save("helloworld.docx")


###########################################################
##import docx
##doc=docx.Document()
##doc.add_paragraph("Hello world")
##paraObj1=doc.add_paragraph("This is a second paragraph")
##paraObj2=doc.add_paragraph("This is a second paragraph")
##paraObj3=doc.add_paragraph("this is a second paragraph")
##paraObj1.add_run("This text is being added to the second paragraph")  ## lo añade en la segnda hoja
##doc.save("multipleParagraphs.docx")
###########################################################
##############################################################
#add_paragraph() #add_run() accept and optional second argument
#doc.add_paragraph("Hello, World!", "Title")


#######################ADDIING HEADINGS
##import docx
###doc=docx.Document()
##doc=docx.Document()
##doc.add_heading("Header 0",0)
##doc.add_heading("Header 1",1)
##doc.add_heading("Header 2",2)
##doc.add_heading("Header 3",3)
##doc.add_heading("Header 4",4)
##doc.add_heading("Header 5",5)
##doc.save("headings.docx ")
##

##import docx
############################################ CREAR PAGINAS
##doc=docx.Document()
##doc.add_paragraph("This is on the first page!")
##doc.paragraphs[0].runs[0].add_break(docx.enum.text.WD_BREAK.PAGE)
##doc.add_paragraph("This is on the second page!")
##doc.save("twoPage.docx")




########################ADDING PICTURES
##import docx
##
##doc=docx.Document()
##doc.add_picture("zophie.png",width=docx.shared.Inches(1),height=docx.shared.Cm(4))
##doc.save("aqui.docx")


##########







#########################CREATING PDFS for Word Document


########this script runs on windows only,and you must have word installed.
##################SIN ACTUALIZAR
### This script runs on Windows only, and you must have Word installed.
##import win32com.client # install with "pip install pywin32==224"
##import docx
##wordFilename = 'your_word_document.docx'
##pdfFilename = 'your_pdf_filename.pdf'
##doc = docx.Document()
### Code to create Word document goes here.
##doc.save(wordFilename)
##wdFormatPDF = 17 # Word's numeric code for PDFs.
##wordObj = win32com.client.Dispatch('Word.Application')
##
##docObj = wordObj.Documents.Open(wordFilename)
##docObj.SaveAs(pdfFilename, FileFormat=wdFormatPDF)
##docObj.Close()
##wordObj.Quit()
