import win32com.client as win32
from tkinter import filedialog
from tkinter import ttk
from tkinter import *
import csv


# this function allow the user to import a list of email contacts from a csv file and appends the values to a list

def get_data():
    try:
        filename = filedialog.askopenfilename()
        with open(filename, newline='') as file:
            reader = csv.reader(file)
            distro = [row for row in reader]
        count = 0
        for item in distro:
            tree.insert(parent='', index='end', iid=count, text='', values=(item))
            count += 1
    except FileNotFoundError:
        pass
    except:
        get_data_error_window = Tk()
        l1 = Label(get_data_error_window, text="An error has occured!")
        l1.grid(row=0, column=0)
        get_data_error_window.mainloop()
    return


  
email_list = []
 

# this function get the current user selection from the treeview, appends it to email list and then creats an email message with the respective fields filled out.
# email list is cleared af function end. user selects 1 at a time from tree view.

 
def draft_emails():
    try:
        curItem = tree.focus()
        cur_value = tree.item(curItem)
        value = cur_value["values"]
        email_list.append(value)
        outlook = win32.Dispatch('Outlook.Application')
        for company, contact_email, cc_email in email_list:
            message = outlook.CreateItem(0)
            message.To = contact_email
            message.CC = cc_email
            message.Subject = company + " - Inquiry"
            message.Display()
            email_list.clear()
    except:
        draft_email_error_window = Tk()
        l1 = Label(draft_email_error_window, text="An error has occured!")
        l1.grid(row=0, column=0)
        draft_email_error_window.mainloop()
    return

# Create the root window
root = Tk()
root.title("Vendor Email Inquiries")
root.geometry("600x400")
root["background"] = "#457b9d"

#adding style
style = ttk.Style()
#pick a theme
style.theme_use("clam")

# configure treeview colors
style.configure("Treeview",
                background="#e5e5e5",
                foreground="#e5e5e5",
                rowheigt=25,
                fieldbackground="#e5e5e5"
                )

#change selected color
# style.map('Treeview',
#           background=[("selected", "green")])

# create treeview
tree = ttk.Treeview(root)

# make tree columns
tree["columns"] = ("Company Name", "Contact Email", "CC Email")

#format tree columsn # # 0 is a "phantom" column.
tree.column("#0", width=0, stretch=NO)
tree.column("Company Name", anchor=W)
tree.column("Contact Email", anchor=W)
tree.column("CC Email", anchor=W,)

#create tree headings:
tree.heading("#0", text="", anchor=W)
tree.heading("Company Name", text="Company Name", anchor=CENTER)
tree.heading("Contact Email", text="Contact Email", anchor=CENTER)
tree.heading("CC Email", text="SM Email", anchor=CENTER)




btn = Button(root, text="Import email contacts", command=get_data, bg="white")
btn2 = Button(root, text="Draft email", command=draft_emails, bg="white")
btn3 = Button(root, text="Close Application", command=root.destroy, bg="white")


# Placing the buttons and treeview
btn.pack(pady=12)
tree.pack()
btn2.pack(pady=12)
btn3.pack(pady=20)


root.mainloop()
