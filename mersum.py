# -*- coding: utf-8 -*-

import sys
import docx
import datetime
from dateutil.relativedelta import relativedelta
from docxcompose.composer import Composer
from docxtpl import DocxTemplate

#import template document
template = DocxTemplate("template.docx")
master = docx.Document("template.docx")
master._body.clear_content() # getting the same file and erasing contents to keep the style
composer = Composer(master)

# get arguments
amount = int(sys.argv[1])
day = int(sys.argv[2])
month = int(sys.argv[3])
year = int(sys.argv[4])

for i in range(amount):
    update = datetime.date(year,month,day) + relativedelta(months=i) #starting date + add month
    
    context = {                 #dictionary to replace words written as: {{WORD}}   
        "date" : update.strftime("%d/%m/%y") # format
        }
    
    template.render(context)    # updating the template
    composer.append(template)   # adding filled template to master
    if i < amount - 1:          # page break between each page except last one 
        master.add_page_break()
    
composer.save("out.docx")       
print("Done!")

