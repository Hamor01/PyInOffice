import tkinter
from tkinter import ttk #tree View
from docxtpl import DocxTemplate
import datetime
from tkinter import messagebox

def clear_item():
    qty_spinbox.delete(0, tkinter.END) #0 means delete from start to .END
    qty_spinbox.insert(0, "1")
    desc_entry.delete(0, tkinter.END)
    unitPrice_spinbox.delete(0, tkinter.END)
    unitPrice_spinbox.insert(0, "0.0")

invoice_list =[]

def add_item():
    qty = int(qty_spinbox.get())
    desc = desc_entry.get()
    unitPrice = float(unitPrice_spinbox.get())
    line_total = qty*unitPrice
    invoice_item = [qty, desc, unitPrice, line_total]
    tree.insert('', 0, values=invoice_item)
    clear_item()

    invoice_list.append(invoice_item)

def new_invoice():
    first_name_entry.delete(0, tkinter.END)
    last_name_entry.delete(0, tkinter.END)
    phone_entry.delete(0, tkinter.END)
    clear_item()
    tree.delete(*tree.get_children())
    invoice_list.clear()

def generate_invoice():
    doc = DocxTemplate("invoice_template.docx")
    name = first_name_entry.get()+last_name_entry.get()
    phone = phone_entry.get()
    subtotal = sum(item[3] for item in invoice_list)
    salestax = 0.1
    total = subtotal*(1-salestax)


    doc.render({'name': name,
                "phone": phone,
                "invoice_list": invoice_list,
                "subtotal":subtotal,
                "salestax": str(salestax*100)+"%",
                "total":total})

    doc_name = "new_invoice" + name + datetime.datetime.now().strftime("%Y-%m-%d-%H%M%S")+ ".docx"
    doc.save(doc_name)

    
    new_invoice()

    messagebox.showinfo("Invoice Complete", "Customer Invoice Generated Successful!")


window = tkinter.Tk()
window.title("GeHuB Invoice Generator")

frame = tkinter.Frame(window) #creating the widget
frame.pack(padx=30, pady=20) #positions your widget in the centre of your screen

first_name_label = tkinter.Label(frame, text = 'First Name')
first_name_label.grid(row=0,column=0) #grid allows us to organize and place contents
last_name_label = tkinter.Label(frame, text = 'Last Name')
last_name_label.grid(row=0,column=1)

#NAME TEXTBOXES
first_name_entry = tkinter.Entry(frame)
last_name_entry = tkinter.Entry(frame)
first_name_entry.grid(row=1,column=0)
last_name_entry.grid(row=1,column=1)

#PHONE LABEL & ENTRY
phone_label = tkinter.Label(frame, text="Phone")
phone_label.grid(row=0,column=2)
phone_entry = tkinter.Entry(frame)
phone_entry.grid(row=1,column=2)

#QUANTITY LABEL & ENTRY
qty_label = tkinter.Label(frame, text="Quantity")
qty_label.grid(row=2,column=0)
qty_spinbox = tkinter.Spinbox(frame, from_=1, to=100)
qty_spinbox.grid(row=3, column=0)

#DESCRIPTIONS LABEL & ENTRY
desc_label = tkinter.Label(frame, text="Description")
desc_label.grid(row=2,column=1)
desc_entry = tkinter.Entry(frame)
desc_entry.grid(row=3, column=1)

#UNIT PRICE LABEL & ENTRY
unitPrice_label = tkinter.Label(frame, text="Unit Price")
unitPrice_label.grid(row=2,column=2)
unitPrice_spinbox = tkinter.Spinbox(frame, from_=0.0, to=500, increment= 0.5)
unitPrice_spinbox.grid(row=3, column=2)

#add item button
add_item_button = tkinter.Button(frame, text="Add Item", command= add_item)
add_item_button.grid(row=4,column=2, pady=5)

#CREATING THE TREE VIEW
columns = ('qty','desc','unitPrice','total')
tree = ttk.Treeview(frame, columns=columns, show="headings")
tree.heading('qty', text='Quantity')
tree.heading('desc', text='Description')
tree.heading('unitPrice', text='Unit Price')
tree.heading('total', text='Total')

tree.grid(row=5, column=0,columnspan=3,padx=20, pady=10) #end of treeview

save_invoice_button = tkinter.Button(frame, text="Generate Invoice", command=generate_invoice)
save_invoice_button.grid(row=6,column=0,columnspan=3, sticky="news", padx=20, pady=5) #news= north,east,west, and south
new_invoice_button = tkinter.Button(frame, text="New Invoice", command=new_invoice)
new_invoice_button.grid(row=7,column=0,columnspan=3, sticky="news", padx=20, pady=5)


window.mainloop()
