import tkinter as tk
from tkinter import ttk
import csv
import getpass
from datetime import datetime
import msvcrt
import time
import random


def get_default_user():
    return getpass.getuser()

# Lock to reduce file corrupttion
def lock_file(file, retries=5, delay=0.1):
    for i in range(retries):
        try:
            msvcrt.locking(file.fileno(), msvcrt.LK_LOCK, 1)
            return True
        except IOError:
            time.sleep(delay + random.random() * delay)
    return False

def unlock_file(file):
    msvcrt.locking(file.fileno(), msvcrt.LK_UNLCK, 1)
    


def submit_order():
    user = get_default_user()
    submit_date = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    order_no = order_no_entry.get()
    skus = sku_entry.get()
    action_taken = action_taken_var.get()
    notes = notes_entry.get("1.0", tk.END).strip()

    if not (order_no and skus):
        show_toast("Error", "Please fill in all required fields.")
        return

    with open('returns.csv', 'a+', newline='') as file:
        try:
            if lock_file(file):
                writer = csv.writer(file)
                writer.writerow([user, submit_date, order_no, skus, action_taken, notes])
            else:
                show_toast("Error", "Failed to Submit log, Please Try Again.")
                return
        finally:
            unlock_file(file)

    show_toast("Success", "Record added successfully.")
    
def update_order():
    # Get the selected item
    selected_item = treeview.focus()
    if selected_item:
        # Get the updated values from the text boxes
        user = user_entry.get("1.0", tk.END).strip()
        date = submit_entry.get("1.0", tk.END).strip()
        order_no = order_no_entry.get().strip()
        skus = sku_entry.get().strip()
        action_taken = action_taken_var.get()
        notes = notes_entry.get("1.0", tk.END).strip()
        ic = get_default_user()
        ad_notes = ad_notes_entry.get("1.0", tk.END).strip()
        update_date = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        # Update the values in the Treeview
        treeview.item(selected_item, values=(user, order_no, date, skus, action_taken, notes, ic, ad_notes, update_date, "", ""))

        # Update the CSV file
        with open('returns.csv', 'r') as file:
            rows = list(csv.reader(file))
        with open('returns.csv', 'w', newline='') as file:
            writer = csv.writer(file)
            for row in rows:
                if row and row[2] == order_no:  # Check if the order number matches
                    writer.writerow([user, date, order_no, skus, action_taken, notes, ic, ad_notes, update_date, "", ""])
                else:
                    writer.writerow(row)
        
        show_toast("Success", "Record updated successfully.")
        


#---^^ Saving & Writing  //  Logging & Updating vv-----------------------------



def show_toast(title, message):
    popup = tk.Tk()
    popup.title(title)
    label = ttk.Label(popup, text=message)
    label.pack(side="top", fill="x", pady=10)
    popup.after(2000, popup.destroy)
    popup.mainloop()

def research_mode():
    global treeview_enabled
    if treeview_enabled:
        file_menu.entryconfig("Disable Research Mode", label="Enable Research Mode")
        toggle_visibility(submit_label)
        toggle_visibility(submit_entry)
        toggle_visibility(user_label)
        toggle_visibility(user_entry)
        toggle_visibility(submit_button)
        toggle_visibility(update_button)
        toggle_visibility(ad_notes_label)
        toggle_visibility(ad_notes_entry)
        root.geometry("365x255")  # Adjust window size after removing Treeview
    else:
        file_menu.entryconfig("Enable Research Mode", label="Disable Research Mode")
        toggle_visibility(submit_label)
        toggle_visibility(submit_entry)
        toggle_visibility(user_label)
        toggle_visibility(user_entry)
        toggle_visibility(submit_button)
        toggle_visibility(update_button)
        toggle_visibility(ad_notes_label)
        toggle_visibility(ad_notes_entry)
        root.geometry("2195x390")  # Adjust window size after adding Treeview
    treeview_enabled = not treeview_enabled
    
def update_treeview():
    # Clear the existing items in the Treeview
    treeview.delete(*treeview.get_children())
    
    # Open the CSV file and read the data
    with open('returns.csv', 'r') as file:
        csv_reader = csv.DictReader(file)
        for row in csv_reader:
            # Extract data based on the headers
            user = row.get("User", "")
            date = row.get("Date", "")
            order = row.get("Order", "")
            skus = row.get("Skus", "")
            action = row.get("Action", "")
            notes = row.get("Notes", "")
            ic = row.get("IC", "")
            additional_notes = row.get("Additional Notes", "")
            update_date = row.get("Update Date", "")
            
            # Check if the value is blank and replace it with an empty string
            user = user if user else ""
            date = row.get("Date", "")
            order = row.get("Order", "")
            skus = skus if skus else ""
            action = action if action else ""
            notes = notes if notes else ""
            ic = ic if ic else ""
            additional_notes = additional_notes if additional_notes else ""
            update_date = update_date if update_date else ""
            
            # Insert the data into the Treeview
            treeview.insert("", "end", values=(user, date, order, skus, action, notes, ic, additional_notes, update_date))

def filter_view(event):
    # Get the filter text from the order number entry
    filter_text = order_no_entry.get().lower()


    def update_view():
        treeview.delete(*treeview.get_children())

        with open('returns.csv', 'r') as file:
            csv_reader = csv.DictReader(file)
            for row in csv_reader:
                if filter_text in row["Order"].lower():
                    # Extract data based on the headers
                    user = row.get("User", "")
                    submit_date = row.get("Date", "")
                    order = row.get("Order", "")
                    skus = row.get("Skus", "")
                    action = row.get("Action", "")
                    notes = row.get("Notes", "")
                    ic = row.get("IC", "")
                    additional_notes = row.get("Additional Notes", "")
                    update_date = row.get("Update Date","")

                    # Check if the value is blank and replace it with an empty string
                    user = user if user else ""
                    submit_date = submit_date if submit_date else ""
                    order = order if order else ""
                    skus = skus if skus else ""
                    action = action if action else ""
                    notes = notes if notes else ""
                    ic = ic if ic else ""
                    additional_notes = additional_notes if additional_notes else ""
                    update_date = update_date if update_date else ""

                    # Insert the data into the Treeview
                    treeview.insert("", "end", values=(user, submit_date, order, skus, action, notes, ic, additional_notes, update_date))

    # Update the treeview after a short delay (50ms)
    root.after(50, update_view)


def view_to_box(event):
    # Get the selected item
    item = treeview.focus()
    if item:
        # Get the values of the selected item
        values = treeview.item(item, "values")
        # Populate the text boxes with the values
        user_entry.delete("1.0", tk.END)
        user_entry.insert(tk.END, values[0])
        submit_entry.delete("1.0", tk.END)
        submit_entry.insert(tk.END, values[2])
        order_no_entry.delete(0, tk.END)
        order_no_entry.insert(0, values[1])
        sku_entry.delete(0, tk.END)
        sku_entry.insert(0, values[3])
        action_taken_var.set(values[4])
        notes_entry.delete("1.0", tk.END)
        notes_entry.insert(tk.END, values[5])
        ic_entry.delete("1.0", tk.END)
        ic_entry.insert(tk.END, values[6])
        ad_notes_entry.delete("1.0", tk.END)
        ad_notes_entry.insert(tk.END, values[7])


def toggle_visibility(widget):
    if widget.winfo_viewable():
        widget.grid_remove()
    else:
        widget.grid()

    

#-- vv The interface vv--------------------------------------------------------



root = tk.Tk()
root.title("Return Log")
root.iconbitmap("face.ico")
root.geometry("365x255")


# Menu
menu_bar = tk.Menu(root)
root.config(menu=menu_bar)
file_menu = tk.Menu(menu_bar, tearoff=0)
treeview_enabled = False

menu_bar.add_cascade(label="Research", menu=file_menu)
file_menu.add_command(label="Enable Research Mode", command=research_mode)
file_menu.add_command(label="Update Research", command=update_treeview)

# widgets
tk.Label(root, text="Order No", anchor="e").grid(row=0, column=0, padx=10, pady=5)
order_no_entry = tk.Entry(root)
order_no_entry.grid(row=0, column=1, padx=10, pady=5)

tk.Label(root, text="SKU(s)", anchor="e").grid(row=1, column=0, padx=10, pady=5)
sku_entry = tk.Entry(root)
sku_entry.grid(row=1, column=1, padx=10, pady=5)

tk.Label(root, text="Action Taken", anchor="e").grid(row=2, column=0, padx=10, pady=5)
action_taken_var = tk.StringVar(value="Return Processed")
action_taken_dropdown = ttk.Combobox(root, textvariable=action_taken_var, values=["Return Processed", "Already Returned", "Misship", "Overgoods", "Other"])
action_taken_dropdown.grid(row=2, column=1, padx=10, pady=5)

tk.Label(root, text="Notes", anchor="e").grid(row=3, column=0, padx=10, pady=5)
notes_entry = tk.Text(root, height=5, width=30)
notes_entry.grid(row=3, column=1, padx=10, pady=5)

user_label = tk.Label(root, text="User")
user_label.grid(row=4, column=0, padx=10, pady=5)
user_label.grid_remove()
user_entry = tk.Text(root, height=1, width=30)
user_entry.grid(row=4, column=1, padx=10, pady=5)
user_entry.grid_remove()


submit_label = tk.Label(root, text="Submit Date")
submit_label.grid(row=5, column=0, padx=10, pady=5)
submit_label.grid_remove()
submit_entry = tk.Text(root, height=1, width=30)
submit_entry.grid(row=5, column=1, padx=10, pady=5)
submit_entry.grid_remove()

ad_notes_label = tk.Label(root, text="Additional Notes")
ad_notes_label.grid(row=6, column=0, padx=10, pady=5)
ad_notes_label.grid_remove()
ad_notes_entry = tk.Text(root, height=5, width=30)
ad_notes_entry.grid(row=6, column=1, padx=10, pady=5)
ad_notes_entry.grid_remove()

ic_entry = tk.Text(root, height = 5, width =30)
ic_entry.grid(row=6, column=1, padx=10, pady=5)
ic_entry.grid_remove()

submit_button = tk.Button(root, text="Submit", command=submit_order)
submit_button.grid(row=7, column=0, columnspan=2, padx=10, pady=10)

update_button = tk.Button(root, text="Update", command=update_order)
update_button.grid(row=7, column=0, columnspan=2, padx=10, pady=10)
update_button.grid_remove()

# Treeview
columns = ("User", "Submit Date", "Order", "Skus", "Action", "Notes", "IC", "Additional Notes", "Update Date")
treeview = ttk.Treeview(root, columns=columns, show="headings")
for col in columns:
    treeview.heading(col, text=col)
treeview.grid(row=0, column=2, rowspan=15, padx=10, pady=5)
order_no_entry.bind("<KeyRelease>", filter_view)
treeview.bind("<ButtonRelease-1>", view_to_box)




root.mainloop()
