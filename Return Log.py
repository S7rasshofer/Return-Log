import tkinter as tk
from tkinter import ttk
import tkinter.font as tkfont
from plyer import notification
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
    order_no = order_no_entry.get("1.0", tk.END).strip()
    skus = sku_entry.get("1.0", tk.END).strip()
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
        customer = customer_entry.get("1.0", tk.END).strip()
        user = user_entry.get("1.0", tk.END).strip()
        date = submit_entry.get("1.0", tk.END).strip()
        order_no = order_no_entry.get("1.0", tk.END).strip()
        skus = sku_entry.get("1.0", tk.END).strip()
        action_taken = action_taken_var.get()
        notes = notes_entry.get("1.0", tk.END).strip()
        ic = get_default_user()
        ad_notes = ad_notes_entry.get("1.0", tk.END).strip()
        rejection = rejection_var.get()
        update_date = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        # Update the values in the Treeview
        treeview.item(selected_item, values=(user, order_no, date, skus, action_taken, notes, ic, ad_notes, customer, rejection, update_date, "", ""))

        # Update the CSV file
        with open('returns.csv', 'r') as file:
            rows = list(csv.reader(file))
        with open('returns.csv', 'w', newline='') as file:
            writer = csv.writer(file)
            for row in rows:
                if row and row[2] == order_no:  # Check if the order number matches
                    writer.writerow([user, date, order_no, skus, action_taken, notes, ic, ad_notes, customer, rejection, update_date, "", ""])
                else:
                    writer.writerow(row)
        
        show_toast("Success", "Record updated successfully.")

        


#---^^ Saving & Writing  //  Logging & Updating vv-----------------------------



def show_toast(title, message):
    notification.notify(
        title="Return Log",
        message=message,
        timeout=10  # duration in seconds
    )

def research_mode():
    global treeview_enabled
    if treeview_enabled:
        file_menu.entryconfig("Disable Research Mode", label="Enable Research Mode")
        toggle_visibility(customer_label)
        toggle_visibility(customer_entry)
        toggle_visibility(submit_label)
        toggle_visibility(submit_entry)
        toggle_visibility(user_label)
        toggle_visibility(user_entry)
        toggle_visibility(submit_button)
        toggle_visibility(update_button)
        toggle_visibility(ad_notes_label)
        toggle_visibility(ad_notes_entry)
        toggle_visibility(reject_taken_dropdown)
        toggle_visibility(reject_label)
        toggle_visibility(treeview)
        root.geometry("365x255")  # Adjust window size after removing Treeview
    else:
        file_menu.entryconfig("Enable Research Mode", label="Disable Research Mode")
        toggle_visibility(customer_label)
        toggle_visibility(customer_entry)
        toggle_visibility(submit_label)
        toggle_visibility(submit_entry)
        toggle_visibility(user_label)
        toggle_visibility(user_entry)
        toggle_visibility(submit_button)
        toggle_visibility(update_button)
        toggle_visibility(ad_notes_label)
        toggle_visibility(ad_notes_entry)
        toggle_visibility(reject_taken_dropdown)
        toggle_visibility(reject_label)
        toggle_visibility(treeview)
        root.geometry("2600x455")  # Adjust window size after adding Treeview
        auto_size_treeview_columns(treeview)
    treeview_enabled = not treeview_enabled
    auto_size_treeview_columns(treeview)
    update_treeview()


def update_treeview():
    # Clear the existing items in the Treeview
    treeview.delete(*treeview.get_children())
    
    # Reapply the current theme settings
    

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
            customer = row.get("Customer", "")
            rejection = row.get("Rejection", "")
            update_date = row.get("Update Date", "")
            
            # Insert the data into the Treeview
            treeview.insert("", "end", values=(user, date, order, skus, action, notes, ic, additional_notes, customer, rejection, update_date))
    
    apply_custom_settings(theme_var.get())
    auto_size_treeview_columns(treeview)

def auto_size_treeview_columns(treeview):
    # Auto-size the columns based on the content
    for col in treeview["columns"]:
        treeview.column(col, width=tkfont.Font().measure(col.title()), stretch=True)
        children = treeview.get_children()
        if children:
            max_width = max(tkfont.Font().measure(treeview.set(child, col)) for child in children)
            treeview.column(col, width=max_width + 10)  # Add a little padding
        else:
            treeview.column(col, width=tkfont.Font().measure(col.title()) + 10)  # Default width


def sort_treeview_by_submit_date():
    # Sort the Treeview by the "Submit Date" column in descending order
    treeview_sort_column = "Submit Date"
    treeview_sort_order = "descending"  # or "ascending" for ascending order
    treeview.sort(treeview_sort_column, order=treeview_sort_order)

def filter_view(event):
    widget = event.widget

    # Determine which entry is being typed into
    if widget == order_no_entry:
        filter_text = order_no_entry.get("1.0", tk.END).strip().lower()
        filter_by = "Order"
    elif widget == customer_entry:
        filter_text = customer_entry.get("1.0", tk.END).strip().lower()
        filter_by = "Customer"
    else:
        return

    # Clear out the entries to their default settings
    if filter_by == "Order":
        customer_entry.delete("1.0", tk.END)
    elif filter_by == "Customer":
        order_no_entry.delete("1.0", tk.END)

    user_entry.delete("1.0", tk.END)
    submit_entry.delete("1.0", tk.END)
    sku_entry.delete("1.0", tk.END)
    action_taken_var.set("Return Processed")
    notes_entry.delete("1.0", tk.END)
    ic_entry.delete("1.0", tk.END)
    ad_notes_entry.delete("1.0", tk.END)
    rejection_var.set("Return Processed")  # Reset rejection_var
    
    #sort_treeview_by_submit_date()
    #auto_size_treeview_columns(treeview)
    
    def update_view():
        # Clear the existing items in the Treeview
        treeview.delete(*treeview.get_children())

        with open('returns.csv', 'r') as file:
            csv_reader = csv.DictReader(file)
            for row in csv_reader:
                if filter_text in row[filter_by].lower():
                    # Extract data based on the headers
                    user = row.get("User", "")
                    submit_date = row.get("Date", "")
                    order = row.get("Order", "")
                    skus = row.get("Skus", "")
                    action = row.get("Action", "")
                    notes = row.get("Notes", "")
                    ic = row.get("IC", "")
                    additional_notes = row.get("Additional Notes", "")
                    customer = row.get("Customer", "")
                    rejection = row.get("Rejection", "")  
                    update_date = row.get("Update Date", "")
                    

                    # Check if the value is blank and replace it with an empty string
                    user = user if user else ""
                    submit_date = submit_date if submit_date else ""
                    order = order if order else ""
                    skus = skus if skus else ""
                    action = action if action else ""
                    notes = notes if notes else ""
                    ic = ic if ic else ""
                    additional_notes = additional_notes if additional_notes else ""
                    customer = customer if customer else ""
                    rejection = rejection if rejection else ""
                    update_date = update_date if update_date else ""
                    
                    treeview.insert("", "end", values=(user, submit_date, order, skus, action, notes, ic, additional_notes, customer, rejection, update_date))

    # Update the treeview after a short delay (200ms)
    root.after(200, update_view)
    auto_size_treeview_columns(treeview)
    apply_custom_settings(theme_var.get())
    sort_treeview_by_submit_date()

def view_to_box(event):
    # Get the selected item
    item = treeview.focus()
    if item:
        # Get the values of the selected item
        values = treeview.item(item, "values")
        # Populate the text boxes with the values
        user_entry.delete("1.0", tk.END)
        user_entry.insert("1.0", values[0])
        submit_entry.delete("1.0", tk.END)
        submit_entry.insert("1.0", values[1])
        order_no_entry.delete("1.0", tk.END)
        order_no_entry.insert("1.0", values[2])
        sku_entry.delete("1.0", tk.END)
        sku_entry.insert("1.0", values[3])
        action_taken_var.set(values[4])
        notes_entry.delete("1.0", tk.END)
        notes_entry.insert("1.0", values[5])
        ic_entry.delete("1.0", tk.END)
        ic_entry.insert("1.0", values[6])
        customer_entry.delete("1.0", tk.END)
        customer_entry.insert("1.0", values[8])
        ad_notes_entry.delete("1.0", tk.END)
        ad_notes_entry.insert("1.0", values[7])



def toggle_visibility(widget):
    if widget.winfo_viewable():
        widget.grid_remove()
    else:
        widget.grid()




#-- vv Theme and Styles vv-----------------------------------------------------



settings = {
    "Light": {"background": "#FFFFFF", "foreground": "#000000", "entry_background": "#F0F0F0", "entry_foreground": "#000000",
              "treeview_background": "#FFFFFF", "treeview_foreground": "#000000", "header_background": "#F0F0F0", "header_foreground": "#000000",
              "odd_row_background": "#F0F0F0", "odd_row_foreground": "#000000", "even_row_background": "#FFFFFF", "even_row_foreground": "#000000"},
    "Dark": {"background": "#1E1E1E", "foreground": "#FFFFFF", "entry_background": "#333333", "entry_foreground": "#FFFFFF",
             "treeview_background": "#1E1E1E", "treeview_foreground": "#FFFFFF", "header_background": "#333333", "header_foreground": "#FFFFFF",
             "odd_row_background": "#333333", "odd_row_foreground": "#FFFFFF", "even_row_background": "#1E1E1E", "even_row_foreground": "#FFFFFF"},
    "Cyber": {"background": "#0F0F0F", "foreground": "#33FF33", "entry_background": "#0F0F0F", "entry_foreground": "#33FF33",
              "treeview_background": "#333333", "treeview_foreground": "#FFFFFF", "header_background": "#555555", "header_foreground": "#FFFFFF",
              "odd_row_background": "#333333", "odd_row_foreground": "#FFFFFF", "even_row_background": "#444444", "even_row_foreground": "#FFFFFF"},
    "Core Blue": {"background": "#003B64", "foreground": "#FFFFFF", "entry_background": "#00234E", "entry_foreground": "#FFFFFF",
                  "treeview_background": "#003B64", "treeview_foreground": "#FFFFFF", "header_background": "#00234E", "header_foreground": "#FFFFFF",
                  "odd_row_background": "#00234E", "odd_row_foreground": "#FFFFFF", "even_row_background": "#003B64", "even_row_foreground": "#FFFFFF"}
}






def change_theme(event=None):
    try:
        selected_theme = theme_var.get()
        print(f"Selected theme: {selected_theme}")
        if selected_theme in settings:
            apply_custom_settings(selected_theme)
        else:
            style.theme_use(selected_theme)
    except Exception as e:
        print(f"Error in change_theme: {e}")


def apply_custom_settings(theme):
    if theme in settings:
        theme_settings = settings[theme]
        root.configure(bg=theme_settings["background"])
        for widget in root.winfo_children():
            if isinstance(widget, tk.Label):
                widget.configure(bg=theme_settings["background"], fg=theme_settings["foreground"])
            elif isinstance(widget, ttk.Combobox):
                style.configure('Custom.TCombobox', fieldbackground=theme_settings["entry_background"], foreground=theme_settings["entry_foreground"])
                widget.configure(style='Custom.TCombobox')
            elif isinstance(widget, tk.Entry) or isinstance(widget, tk.Text):
                widget.configure(bg=theme_settings["entry_background"], fg=theme_settings["entry_foreground"])
            elif isinstance(widget, ttk.Button):
                style.configure('Custom.TButton', background=theme_settings["background"], foreground=theme_settings["foreground"])
                widget.configure(style='Custom.TButton')

        # Update Treeview colors and appearance
        style.configure('Custom.Treeview', background=theme_settings["treeview_background"], foreground=theme_settings["treeview_foreground"], fieldbackground=theme_settings["background"])
        style.configure('Custom.Treeview.Heading', background=theme_settings["header_background"], foreground=theme_settings["header_foreground"])
        style.configure('Custom.Treeview', rowheight=25)  # Adjust row height as needed
        treeview.tag_configure('evenrow', background=theme_settings["even_row_background"], foreground=theme_settings["even_row_foreground"])
        treeview.tag_configure('oddrow', background=theme_settings["odd_row_background"], foreground=theme_settings["odd_row_foreground"])
        for i, item in enumerate(treeview.get_children()):
            if i % 2 == 0:
                treeview.item(item, tags=('evenrow',))
            else:
                treeview.item(item, tags=('oddrow',))






#-- vv The interface vv--------------------------------------------------------



root = tk.Tk()
root.title("Return Log")
root.iconbitmap("face.ico")
root.geometry("365x255")

style = ttk.Style()
initial_theme = "Cyber"  


# Menu
menu_bar = tk.Menu(root)
root.config(menu=menu_bar)
file_menu = tk.Menu(menu_bar, tearoff=0)
treeview_enabled = False

menu_bar.add_cascade(label="Research", menu=file_menu)
file_menu.add_command(label="Enable Research Mode", command=research_mode)
file_menu.add_command(label="Update Research", command=update_treeview)

theme_menu = tk.Menu(menu_bar, tearoff=0)
menu_bar.add_cascade(label="Themes", menu=theme_menu)

theme_var = tk.StringVar(value=initial_theme)
for theme in settings.keys():
    theme_menu.add_radiobutton(label=theme, variable=theme_var, command=change_theme)

# Widgets
customer_label = tk.Label(root, text="Customer", anchor="e")
customer_label.grid(row=0, column=0, padx=10, pady=5)
customer_label.grid_remove()
customer_entry = tk.Text(root, height=1, width=30)
customer_entry.grid(row=0, column=1, padx=10, pady=5)
customer_entry.grid_remove()

tk.Label(root, text="Order No", anchor="e").grid(row=1, column=0, padx=10, pady=5)
order_no_entry = tk.Text(root, height=1, width=30)
order_no_entry.grid(row=1, column=1, padx=10, pady=5)



tk.Label(root, text="SKU(s)", anchor="e").grid(row=2, column=0, padx=10, pady=5)
sku_entry = tk.Text(root, height=1, width=30)
sku_entry.grid(row=2, column=1, padx=10, pady=5)

tk.Label(root, text="Action Taken", anchor="e").grid(row=3, column=0, padx=10, pady=5)
action_taken_var = tk.StringVar(value="Return Processed")
action_taken_dropdown = ttk.Combobox(root, width=30, textvariable=action_taken_var, values=["Return Processed", "Already Returned", "Misship", "Overgoods", "Other"])
action_taken_dropdown.grid(row=3, column=1, padx=10, pady=5)

tk.Label(root, text="Notes", anchor="e").grid(row=4, column=0, padx=10, pady=5)
notes_entry = tk.Text(root, height=5, width=30)
notes_entry.grid(row=4, column=1, padx=10, pady=5)

user_label = tk.Label(root, text="User")
user_label.grid(row=5, column=0, padx=10, pady=5)
user_label.grid_remove()
user_entry = tk.Text(root, height=1, width=30)
user_entry.grid(row=5, column=1, padx=10, pady=5)
user_entry.grid_remove()

submit_label = tk.Label(root, text="Submit Date")
submit_label.grid(row=6, column=0, padx=10, pady=5)
submit_label.grid_remove()
submit_entry = tk.Text(root, height=1, width=30)
submit_entry.grid(row=6, column=1, padx=10, pady=5)
submit_entry.grid_remove()

ad_notes_label = tk.Label(root, text="Additional Notes")
ad_notes_label.grid(row=7, column=0, padx=10, pady=5)
ad_notes_label.grid_remove()
ad_notes_entry = tk.Text(root, height=5, width=30)
ad_notes_entry.grid(row=7, column=1, padx=10, pady=5)
ad_notes_entry.grid_remove()

ic_entry = tk.Text(root, height=5, width=30)
ic_entry.grid(row=8, column=1, padx=10, pady=5)
ic_entry.grid_remove()

reject_label = tk.Label(root, text="Reject?", anchor="e")
reject_label.grid(row=9, column=0, padx=10, pady=5)
reject_label.grid_remove()
rejection_var = tk.StringVar(value="")
reject_taken_dropdown = ttk.Combobox(root, width=30, textvariable="Return Processed", values=["OOP","Serial","Hazmat"])
reject_taken_dropdown.grid(row=9, column=1, padx=10, pady=5)
reject_taken_dropdown.grid_remove()

submit_button = ttk.Button(root, text="Submit", command=submit_order, style='Custom.TButton')
submit_button.grid(row=10, column=0, columnspan=2, padx=10, pady=10)

update_button = ttk.Button(root, text="Update", command=update_order, style='Custom.TButton')
update_button.grid(row=10, column=0, columnspan=2, padx=10, pady=10)
update_button.grid_remove()



# Treeview
columns = ("User", "Submit Date", "Order", "Skus", "Action", "Notes", "IC", "Additional Notes", "Customer", "Rejection", "Update Date")
treeview = ttk.Treeview(root, columns=columns, show="headings")
for col in columns:
    treeview.heading(col, text=col)
treeview.grid(row=0, column=2, rowspan=30, columnspan=len(columns), sticky="nsew", padx=10, pady=5)


order_no_entry.bind("<KeyRelease>", filter_view)
treeview.bind("<ButtonRelease-1>", view_to_box)
treeview.grid_remove()

order_no_entry.bind("<KeyRelease>", filter_view)
customer_entry.bind("<KeyRelease>", filter_view)

style.theme_use('default')
apply_custom_settings(initial_theme)
auto_size_treeview_columns(treeview)
root.mainloop()