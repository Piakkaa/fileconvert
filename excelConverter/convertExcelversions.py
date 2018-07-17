
# coding: utf-8

# In[88]:


import xlrd
from openpyxl.workbook import Workbook


# In[89]:


book1=xlrd.open_workbook('payrollreporta32l.xls',encoding_override="iso8859_11")


# In[90]:


sheet0=book1.sheet_by_index(0)


# In[92]:


sheet0.cell_value(9,2)


# In[93]:


def open_xls_as_xlsx(filename):
    # first open using xlrd
    book = xlrd.open_workbook(filename,encoding_override="iso8859_11")
    index = 0
    nrows, ncols = 0, 0
    while nrows * ncols == 0:
        sheet = book.sheet_by_index(index)
        nrows = sheet.nrows
        ncols = sheet.ncols
        index += 1

    # prepare a xlsx sheet
    book1 = Workbook()
    sheet1 = book1.get_active_sheet()

    for row in range(1, nrows+1):
        for col in range(1, ncols+1):
            sheet1.cell(row=row, column=col).value = sheet.cell_value(row-1, col-1)

    return book1


# In[94]:


def savefile(fileinput,fileoutput):
    output=open_xls_as_xlsx(fileinput)
    output.save(fileoutput)


# In[95]:


savefile('payrollreporta32l.xls','example1.xlsx')

