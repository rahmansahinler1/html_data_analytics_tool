import tkinter as tk
from tkinter import simpledialog

class CustomInputDialog(simpledialog.Dialog):
    def __init__(self, parent, title, labels):
        self.labels = labels
        super().__init__(parent, title)
    
    def body(self, master):
        for i, label_text in enumerate(self.labels):
            tk.Label(master, text=label_text).grid(row=i, column=0, columnspan=2)
            
        self.result = tk.StringVar()
        self.entry = tk.Entry(master, textvariable=self.result)
        self.entry.grid(row=4, column=0, columnspan=2)
        return self.entry

    def apply(self):
        self.strings = self.result.get()

class MotorNumberSelectionDiaglog:
    def __init__(self, master, options):
        self.master = master  
        self.options = options
        self.result = None

        self.dialog = tk.Toplevel(master)
        self.dialog.geometry('300x540')
        self.dialog.title("Motor Number Selection")
        self.dialog.resizable(False, False)

        self.listbox = tk.Listbox(self.dialog, font=("Arial", 10), fg="black", bg="white", selectmode=tk.MULTIPLE)
        self.listbox.place(x=19, y=40, width=250, height=460)
        self.scrollbar = tk.Scrollbar(self.dialog)
        self.scrollbar.place(x=271, y=40, height=460, width=12)
        self.listbox.config(yscrollcommand=self.scrollbar.set)
        self.scrollbar.config(command=self.listbox.yview, orient="vertical")
        
        self.label = tk.Label(self.dialog, text='Select the motor numbers that you want to see in the plot')
        self.label.config(font=("Arial", 7), fg="white", bg="#00437A")
        self.label.place(x=19, y=10, width=250, height=20)
        
        for option in options:
            self.listbox.insert(tk.END, option)
            
        self.select_all_button = tk.Button(self.dialog, text="Select All", command=self.toggle_select_all)
        self.select_all_button.place(x=19, y=510, height=20, width=120)

        self.ok_button = tk.Button(self.dialog, text="Confirm Selection", command=self.select)
        self.ok_button.place(x=149, y=510, height=20, width=120)

        self.all_selected = False

    def select(self):
        selected_indices = self.listbox.curselection()
        self.selected_options = [self.options[i] for i in selected_indices]
        self.dialog.destroy()
    
    def toggle_select_all(self):
        if self.all_selected:
            # Deselect all items
            self.listbox.selection_clear(0, tk.END)
            self.all_selected = False
            self.select_all_button.config(text="Select All")
        else:
            # Select all items
            self.listbox.select_set(0, tk.END)
            self.all_selected = True
            self.select_all_button.config(text="Deselect All")
