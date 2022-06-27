import openpyxl
from tkinter import *
from tkinter import messagebox
def bookRegister():       # helper function of ADD BOOK
    
    #messagebox is imported to display message boxes
    #pandas library is imported to create dataframe
    #openpyxl library is imported to write / read excel sheets

    title = Info1.get()
    author =Info2.get()
    date = Info3.get()
    bktype =Info4.get()
    bookid =Info5.get()
    
    if(title=='' or author=='' or date=='' or bktype=='' or bookid==''):
        messagebox.showwarning("Invalid","All the fields are required")
        return
        #if some or all the entries are empty, warning is shown, and it returns to the calling function
        #else the data is added to the excel sheet
    else:
        wb = openpyxl.load_workbook(r"C:\Users\Admin\Downloads\Library Books.xlsx") #excel sheet is loaded in the workbook
        sheet = wb.active                        
        issuer_name= "none"
        serial=newRowLoc=sheet.max_row 
        bktype=bktype.lower()
        about_book=[serial,title,author,date,bktype,bookid,issuer_name] #a list is created of all the data entries of book
        sheet.append(about_book)             # the list is added at the end of excel file retaining old data
        wb.save(r"C:\Users\Admin\Downloads\Library Books.xlsx")          #saves the excel file
    
        messagebox.showinfo('Success',"Book added successfully") #message is displayed after successful addition of data to excel sheet
        #all the previous entries are deletd after the data is saved to file
        Info5.delete(0, END) 
        Info4.delete(0, END)
        Info3.delete(0, END)
        Info2.delete(0, END)
        Info1.delete(0, END)   
def addBook():     #add book function
    
    global Info1 ,Info2,Info3, Info4,Info5,Info6
    root=Tk()
    root.title("Add books to Library ")
    root.geometry("1000x500")
    if len(root.winfo_children()) > 1:
        root.winfo_children()[1].destroy()
    
    frame1 = Frame(root,bg="#f7f1e3") #frame is used  for arranging the position of other widgets
    frame1.pack(expand=True,fill=BOTH) #expand is set to true so that frame1 expands to fill any space not otherwise used in frame1's parent.
    #fill=BOTH fills ( both horizontally and vertically ) any extra space allocated to it
    
    #all the variables holds a string 
    title=StringVar()
    author=StringVar()
    date=StringVar()
    bktype=StringVar()
    bookid=StringVar()
        
    #heading is given to the frame ,also headingFrame1 is used to hold the heading which is above frame1
    headingFrame1 = Frame(frame1,bg="#6bd5e8",bd=5) 
    headingFrame1.place(relx=0.25,rely=0.03,relwidth=0.5,relheight=0.12)
    headingLabel = Label(headingFrame1, text="Add Books", bg="#6bd5e8", fg='#151c52', font=('Courier',19,'bold'))
    headingLabel.place(relx=0,rely=0, relwidth=1, relheight=1)
    
    #frame to hold all the blank space and their labels 
    labelFrame = Frame(frame1,bg="#6bd5e8")
    labelFrame.place(relx=0.1,rely=0.20,relwidth=0.8,relheight=0.6)
        
    #relx and rely: Horizontal and vertical offset as a float between 0.0 and 1.0, as a fraction of the height and width of the parent widget.
    #relheight, relwidth - Height and width as a float between 0.0 and 1.0, as a fraction of the height and width of the parent widget
    #relx, rely, relheight, relwidth are used for every label and entry and also for two buttons
    # Title
    lb1 = Label(labelFrame,text="Title : ", bg="#6bd5e8", fg='black',font=("calibre",14,'normal'))
    lb1.place(relx=0.05,rely=0.10, relheight=0.08)
        
    Info1 = Entry(labelFrame,justify='center',textvariable=title,font=("calibre",14,'normal'))
    Info1.place(relx=0.4,rely=0.10, relwidth=0.56, relheight=0.09)
        
    # Book Author
    lb2 = Label(labelFrame,text="Author : ", bg="#6bd5e8", fg='black',font=("calibre",14,'normal'))
    lb2.place(relx=0.05,rely=0.25, relheight=0.08)
        
    Info2 = Entry(labelFrame,textvariable=author,font=("calibre",14,'normal'),justify='center')
    Info2.place(relx=0.4,rely=0.25, relwidth=0.56, relheight=0.09)
    
    #Addition Date
    lb3 = Label(labelFrame,text="Addition Date : ", bg="#6bd5e8", fg='black',font=("calibre",14,'normal'))
    lb3.place(relx=0.05,rely=0.40, relheight=0.08)
        
    Info3 = Entry(labelFrame,textvariable=date,font=("calibre",14,'normal'),justify='center')
    Info3.place(relx=0.4,rely=0.40, relwidth=0.56, relheight=0.09)
    
    #type of book
    lb4 = Label(labelFrame,text="Type of book : ", bg="#6bd5e8", fg='black',font=("calibre",14,'normal'))
    lb4.place(relx=0.05,rely=0.55, relheight=0.08)
        
    Info4 = Entry(labelFrame,textvariable=bktype,font=("calibre",14,'normal'),justify='center')
    Info4.place(relx=0.4,rely=0.55, relwidth=0.56, relheight=0.09)
    
    # Book ID
    lb5 = Label(labelFrame,text="Book ID : ", bg="#6bd5e8", fg='black',font=("calibre",14,'normal'))
    lb5.place(relx=0.05,rely=0.70, relheight=0.08)
        
    Info5 = Entry(labelFrame,textvariable=bookid,font=("calibre",14,'normal'),justify='center')
    Info5.place(relx=0.4,rely=0.70, relwidth=0.56, relheight=0.09)
        
    #two buttons are on frame1 
    #bg is background colour
    #fg is forground colour
    #text is displayed over the buttons
    
    #Submit Button
    SubmitBtn= Button(frame1,text="SUBMIT",bg='#d4d4d4', fg='black',command= bookRegister,font=("calibre",12,'normal'))
    SubmitBtn.place(relx=0.42,rely=0.85, relwidth=0.18,relheight=0.08)  
    #if the user clicks Submit button, bookRegister() function is called. this function adds the data to the excel sheet.
    root.mainloop()
def main():
    addBook()
if __name__=="__main__":
    main()