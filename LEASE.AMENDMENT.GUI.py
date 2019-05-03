from __future__ import print_function
from mailmerge import MailMerge
from tkinter import *



class LeaseGUI:

    def __init__(self,parent):
        self.parent=parent
        input=StringVar

        # CREATES OPTIONS DROP DOWN LIST
        OPTIONS=[
            'The Tenant wishes to lease, and Landlord wishes to rent, ',
            'The Tenant wishes to return, and Landlord wishes to accept, '
            ]
        self.options_selection=StringVar()
        self.options_selection.set(OPTIONS[0])

        #CREATES THE ENTRY FIELDS
        self.entry_1 = Entry(parent, textvariable=input)
        self.entry_2 = Entry(parent, textvariable=input)
        self.entry_3 = Entry(parent, textvariable=input)
        self.entry_4 = Entry(parent, textvariable=input)
        self.entry_5 = OptionMenu(parent,self.options_selection,*OPTIONS)
        self.entry_6 = Entry(parent, textvariable=input)
        self.entry_7 = Entry(parent, textvariable=input)
        self.entry_8 = Entry(parent, textvariable=input)
        self.entry_9 = Entry(parent, textvariable=input)
        self.entry_10 = Entry(parent, textvariable=input)

        self.entry_1.grid(row=0, column=1)
        self.entry_2.grid(row=1, column=1)
        self.entry_3.grid(row=2, column=1)
        self.entry_4.grid(row=3, column=1)
        self.entry_5.grid (row=4,column=1)
        self.entry_6.grid(row=5,column=1)
        self.entry_7.grid(row=6,column=1)
        self.entry_8.grid(row=7,column=1)
        self.entry_9.grid(row=8,column=1)

        #LABELS THE ENTRY FIELD
        self.label_1 = Label(parent, text='Company')
        self.label_2 = Label(parent, text='Lease Number')
        self.label_3 = Label(parent, text='Lease Start')
        self.label_4 = Label(parent, text='Effective Date')
        self.label_5 = Label(parent,text='Change')
        self.label_6 = Label(parent,text='Item')
        self.label_7 = Label(parent,text='Unit')
        self.label_8 = Label(parent,text='Price')
        self.label_9 = Label(parent,text='Total Price')

        self.label_1.grid(row=0, sticky=E)
        self.label_2.grid(row=1, sticky=E)
        self.label_3.grid(row=2, sticky=E)
        self.label_4.grid(row=3, sticky=E)
        self.label_5.grid(row=4, sticky=E)
        self.label_6.grid(row=5, sticky=E)
        self.label_7.grid(row=6, sticky=E)
        self.label_8.grid(row=7, sticky=E)
        self.label_9.grid(row=8, sticky=E)

        #CREATES BUTTON TO RUN INPUT FUNCTION
        self.b1 = Button(parent, text='Save Inputs', command=self.getInput)
        self.b1.grid(row=10, column=1)

    #FUNCTION THAT WILL POPULATE WORD TEMPLATE WITH ENTRIES AND DROP DOWN OPTION
    #WILL SAVE LEASE AMENDMENT WITH COMPANY NAME AND LEASE AMENDMENT NUMBER
    def getInput(self):
        self.e1=self.entry_1.get()
        self.e2=self.entry_2.get()
        self.e3=self.entry_3.get()
        self.e4=self.entry_4.get()
        self.e5=self.options_selection.get()
        self.e6 = self.entry_6.get()
        self.e7 = self.entry_7.get()
        self.e8 = self.entry_8.get()
        self.e9 = self.entry_9.get()
        self.e10 = self.entry_10.get()
        self.template='lease_amendment_template.docx'
        self.document=MailMerge(self.template)
        print(self.document.get_merge_fields())
        self.document.merge(
            Company=self.e1,
            LeaseNumber=self.e2,
            LeaseStart=self.e3,
            EffectiveDate=self.e4,
            Change=self.e5,
            Item=self.e6,
            Unit=self.e7,
            Price=self.e8,
            TotalPrice=self.e9,
            NewRent=self.e10)
        self.document.write(str(self.e1) + ' Lease Amendment ' + str(self.e2)+ '.docx')


#CREATES THE GUI
if __name__ == '__main__':
    root=Tk()
    top=LeaseGUI(root)
    root.mainloop()