import pandas as pd
import numpy as np

import tkinter as tk
from tkinter import *
from tkinter import messagebox
from tkinter import ttk, filedialog
from tkinter.messagebox import askyesno
from tkinter import filedialog
from tkcalendar import Calendar

from PIL import Image, ImageTk
from bs4 import BeautifulSoup

import os
import datetime

from gui.window import Window
from gui.dialogues import MotorNumberSelectionDiaglog, CustomInputDialog

class App(tk.Tk):
    def __init__(self):
        def insert_html():
            self.btn_create_database["state"] = "normal"

            # Define file path to be read by the tool
            file_type = (('html files', '*.html'), ('html files', '*.htm'), ('All files', '*.*'))
            html_files = filedialog.askopenfilenames(title='Open files', initialdir=os.getcwd(), filetypes=file_type)
            html_names, paths = [], []
            outstr = [None] * len(html_files)

            if len(html_files) == 0:
                return

            for i in range(len(html_files)):
                html_path = html_files[i].strip()
                slash_index = html_path.rfind("/")
                html_name = html_path[slash_index+1:]

                html_names.append(html_name)
                paths.append(html_path)

                # Merge HTML files and create final dataframe
                index = open(paths[i], "r").read()
                raw_html_file = BeautifulSoup(index, 'lxml')
                if i == 0:
                    raw_html_file_str = str(raw_html_file)
                    rb_index = raw_html_file_str.rfind(">")
                    raw_html_file_str = raw_html_file_str[0:rb_index-37]
                    outstr[i] = raw_html_file_str

                elif i != 0 and i != len(outstr) - 1:
                    raw_html_file = str(raw_html_file.find_all("tr")[1:])
                    raw_html_file = raw_html_file.replace("[", "")
                    raw_html_file = raw_html_file.replace("]", "")
                    raw_html_file = raw_html_file.replace(",", "\n")
                    raw_html_file_str = raw_html_file
                    outstr[i] = raw_html_file_str

                else:
                    raw_html_file = str(raw_html_file.find_all("tr")[1:])
                    raw_html_file = raw_html_file.replace("[", "")
                    raw_html_file = raw_html_file.replace("]", "")
                    raw_html_file = raw_html_file.replace(",", "")
                    rb_index = raw_html_file_str.rfind(">")
                    raw_html_file_str = raw_html_file[0:rb_index] + "\n</tbody></table>\n</div>\n</body>\n</html>"
                    outstr[i] = raw_html_file_str

            outstr = "\n".join(outstr)

            with open('./{}.html'.format('html_merged'), mode='w', encoding='utf-8') as code:
                code.write(outstr)
            df = pd.read_html('.\html_merged.html', header=0)[0]

            replaced_characters = ['-', '_', '/']
            for each_char in replaced_characters:
                df['Motornummer'] = df['Motornummer'].str.replace(each_char, '')

            # Show html file names on the tool
            for i in range(0, len(html_names)):
                self.tree.insert('', index=i, values=html_names[i])

            # Delete unnecessary rows
            filter_in_values = ["EEF-MOD-F", "EEF-MOD", "EWP"]
            df['Versuchsbeschreibung'] = df['Versuchsbeschreibung'].str.upper()
            df = df[df['Versuchsbeschreibung'].isin(filter_in_values)].reset_index(drop=True)
            self.df = df

        def check_existing_lines(txt_file_df, df_to_save):
            size_df_to_save = df_to_save.shape[0]
            df_to_save_dropped = df_to_save.copy()

            for ctr_row in range(size_df_to_save):
                # Filtering based on channels:
                pst_value = df_to_save.loc[ctr_row, "PST"]
                motornummer_value = df_to_save.loc[ctr_row, "Motornummer"]
                zeitstempel_value = df_to_save.loc[ctr_row, "Zeitstempel"]
                laufstrecke_value = df_to_save.loc[ctr_row, "Laufstrecke"]

                # Filter available database based on Zeitstempel and Laufstrecke
                txt_file_df_filtered = txt_file_df[(txt_file_df["Zeitstempel"] == zeitstempel_value) &
                                                   (txt_file_df["Laufstrecke"] == laufstrecke_value)]

                # Check the indices of PST and Motornummer
                if pst_value in txt_file_df_filtered["PST"].to_list() and motornummer_value in txt_file_df_filtered["Motornummer"].to_list():
                    if txt_file_df_filtered["PST"].to_list().index(pst_value) == txt_file_df_filtered["Motornummer"].to_list().index(motornummer_value):
                        df_to_save_dropped = df_to_save_dropped.drop(ctr_row)

            return df_to_save_dropped

        def create_database():
            answer_create_new_db = askyesno(title="Create New Database",
                              message='Would you like to create a new database file?')

            if answer_create_new_db:
                txt_path = filedialog.asksaveasfilename(
                    filetypes=[("txt file", ".txt")],
                    defaultextension=".txt",
                    initialdir=os.getcwd())
                try:
                    if os.path.isfile(txt_path):
                        os.remove(txt_path)
                        self.df.to_csv(txt_path, header=True, index=False, sep='\t', mode='a')
                        messagebox.showinfo("Info", f"Exporting is done!\n\nFile Path: {txt_path}")
                    else:
                        self.df.to_csv(txt_path, header=True, index=False, sep='\t', mode='a')
                        messagebox.showinfo("Info", f"Exporting is done!\n\nFile Path: {txt_path}")
                except:
                    messagebox.showerror("Error", f"Error occurred while saving database!")
            else:
                file_type = (('text files', '*.txt'), ('All files', '*.*'))
                txt_path = filedialog.askopenfilenames(title='Open files', initialdir=os.getcwd(), filetypes=file_type)

                if len(txt_path) == 0:
                    return

                txt_file_df = pd.read_csv(txt_path[0], header=0, sep='\t')
                df_filtered = check_existing_lines(txt_file_df, self.df)

                if df_filtered.shape[0] == 0:
                    messagebox.showinfo("Info", f"All information in the selected HTML files is already found in the database.")
                else:
                    try:
                        df_filtered.to_csv(txt_path[0], header=False, index=False, sep='\t', mode='a')
                        messagebox.showinfo("Info", f"Exporting is done!\n\nFile Path: {txt_path[0]}")
                    except:
                        messagebox.showinfo("Info", f"The database file is already open by different application!\n\nFile Path: {txt_path}")

        def delete_file_name_list_sel_rows():
            for selected_item in self.tree.selection():
                self.tree.delete(selected_item)

        def delete_file_name_list_all_rows():
            self.tree.delete(*self.tree.get_children())
            self.btn_create_database["state"] = "disabled"

        def delete_date_list_sel_rows():
            for selected_item in self.milestone_tree.selection():
                self.milestone_tree.delete(selected_item)

        super().__init__()
        self.proj_dir = os.getcwd()
        self.memory = pd.read_json(f'{self.proj_dir}\\utils\\memory.json')  # Read memory when called

        self.geometry('800x675+592+150')
        self['background'] = '#f7f7f7'
        self.iconbitmap(f'{self.proj_dir}\\utils\\avl_logo.ico')
        self.title('Main Page')
        self.resizable(False, False)

        # Adjustment variables
        row_span = 2
        column_span = 3
        x_pad = 3
        y_pad = 2
        ix_pad = 3
        iy_pad = 2

        # Header image
        image2 = ImageTk.PhotoImage(Image.open(f"{self.proj_dir}\\utils\\Tool_Header.png"))
        label_photo = Label(image=image2)
        label_photo.image2 = image2
        label_photo.grid(row=0, column=0, rowspan=row_span, columnspan=column_span+10)

        # writing current htmls into blank frame. Insert, delete and delete all functions will be added.
        column = ('1',)
        self.tree = ttk.Treeview(self, columns=column, show='headings')
        self.tree.heading('1', text='File Name')
        self.tree.bind('<<TreeviewSelect>>')
        self.tree.grid(row=5, column=0, columnspan=column_span+9, padx= 1, sticky=NSEW)

        # add a scrollbar into html list
        scrollbar = ttk.Scrollbar(self, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscroll=scrollbar.set)
        scrollbar.grid(row=5, column=12, sticky='NSE')

        # row:10
        self.btn = Button(self, text="Insert HTML File", bg='white', command=insert_html)
        self.btn.grid(row=10, column=0, rowspan=row_span, columnspan=column_span, ipadx=ix_pad + 20, ipady=iy_pad,
                      padx=x_pad + 10, pady=y_pad + 5, sticky=NSEW)
        self.btn = Button(self, text="Delete Selected File", command=delete_file_name_list_sel_rows, bg='white',
                          relief='raised', highlightcolor='white')
        self.btn.grid(row=10, column=9, rowspan=row_span, columnspan=column_span,ipadx=ix_pad + 20, ipady=iy_pad,
                      padx=x_pad + 10, pady=y_pad+5, sticky=NSEW)
        self.btn = Button(self, text="Import Database", bg='white', command= lambda: self.import_database())
        self.btn.grid(row=10, column=3, rowspan=row_span, columnspan=column_span, ipadx=ix_pad + 20, ipady=iy_pad,
                      padx=x_pad + 10, pady=y_pad + 5, sticky=NSEW)

        self.btn = Button(self, text="Delete All Files", bg='white', command=delete_file_name_list_all_rows)
        self.btn.grid(row=10, column=6, rowspan=row_span, columnspan=column_span, ipadx=ix_pad + 20, ipady=iy_pad,
                      padx=x_pad + 10, pady=y_pad + 5, sticky=NSEW)

        # row:15
        self.btn_create_database = Button(text='Create Database', bg='white', fg='#005799', state='disabled',
                                          command=create_database)
        self.btn_create_database.grid(row=15, column=0, rowspan=row_span, columnspan=column_span, ipadx=ix_pad + 20, ipady=iy_pad,
                                      padx=x_pad + 10, pady=y_pad + 5, sticky=NSEW)

        self.btn_create_table = Button(text='Create Table in XLSX', bg='white', fg='#005799', state='disabled',
                                       command=lambda: Window.export_xlsx(self))
        self.btn_create_table.grid(row=15, column=3, rowspan=row_span, columnspan=column_span, ipadx=ix_pad + 20, ipady=iy_pad,
                                   padx=x_pad + 10, pady=y_pad + 5, sticky=NSEW)
        self.btn_plot_laufstrecke = Button(text='Plot Zyklus', bg='white', fg='#005799', state='disabled',
                                           command=lambda: Window.create_plot(self))
        self.btn_plot_laufstrecke.grid(row=15, column=6, rowspan=row_span, columnspan=column_span, ipadx=ix_pad + 20, ipady=iy_pad,
                                       padx=x_pad + 10, pady=y_pad + 5, sticky=NSEW)
        self.btn_export_ppt = Button(text='Export to PPT', bg='white', fg='#005799', state='normal',
                                     command=lambda: Window.export_pptx(self))
        self.btn_export_ppt.grid(row=15, column=9, rowspan=row_span, columnspan=column_span, ipadx=ix_pad + 20, ipady=iy_pad,
                                 padx=x_pad + 10, pady=y_pad + 5, sticky=NSEW)

        # row: 20
        curr_date = datetime.date.today()
        self.calendar = Calendar(selectmode='day', year=curr_date.year, month=curr_date.month, day=curr_date.day, pady=y_pad)
        self.calendar.grid(row=20, column=1, rowspan=row_span, columnspan=column_span, ipadx=ix_pad, ipady=iy_pad,
                           padx=x_pad, pady=y_pad, sticky=NSEW)

        # writing current htmls into blank frame. Insert, delete and delete all functions will be added. 
        columns = ['milestone_name', 'milestone_date']
        self.milestone_tree = ttk.Treeview(self, columns=columns, show='headings')
        self.milestone_tree.heading('milestone_name', text='Milestone Name')
        self.milestone_tree.heading('milestone_date', text='Milestone Date')
        self.milestone_tree.bind('<<TreeviewSelect>>')
        self.milestone_tree.grid(row=20, column=6, columnspan=column_span+3, padx= 1, sticky=NSEW)
        
        # initialize milestone tree based on memory
        for _, row in self.memory.iterrows():
            milestone_name = row['Name']
            milestone_date = row['Date'].strftime('%m/%d/%Y')
            self.milestone_tree.insert('', tk.END, values=[milestone_name, milestone_date])

        # row:40
        self.btn_create_plan = Button(text='Plot Zeitplan', bg='white', fg='#005799', state='disabled',
                                      command=lambda: Window.create_plan(self))
        self.btn_create_plan.grid(row=40, column=6, rowspan=row_span, columnspan=column_span - 1, ipadx=ix_pad + 5, ipady=iy_pad,
                                  padx=x_pad, pady=y_pad + 5, sticky=NSEW)

        self.btn = Button(text='Delete Date', bg='white', fg='#005799',
                          command=delete_date_list_sel_rows)
        self.btn.grid(row=40, column=8, rowspan=row_span, columnspan=column_span - 1, ipadx=ix_pad + 5, ipady=iy_pad,
                      padx=x_pad, pady=y_pad + 5, sticky=NSEW)
        
        self.btn_update_memory = Button(text='Push to Memory', bg='white', fg='#005799', state='normal',
                                      command=lambda: Window.update_memory(self))
        self.btn_update_memory.grid(row=40, column=10, rowspan=row_span, columnspan=column_span - 1, ipadx=ix_pad + 5, ipady=iy_pad,
                                  padx=x_pad, pady=y_pad + 5, sticky=NSEW)

        self.btn = Button(text='Select Date', bg='white', fg='#005799',
                          command=lambda: Window.select_date(self))
        self.btn.grid(row=40, column=1, rowspan=row_span, columnspan=column_span, ipadx=ix_pad + 20, ipady=iy_pad,
                      padx=x_pad + 10, pady=y_pad + 5, sticky=NSEW)

    def import_database(self):
        def convert_excel_date(column: pd.DataFrame):
            def convert_date(date_str: str):
                try:
                    return pd.to_datetime(date_str.split(' ')[0])
                except (ValueError, TypeError):
                    try:
                        return pd.to_datetime('1899-12-30') + pd.to_timedelta(int(date_str), unit='d')
                    except (ValueError, TypeError):
                        return ''
            return column.apply(convert_date)
        
        def to_cw(column: pd.DataFrame):
            cw = []
            def take_year_week(date):
                year, week = date.isocalendar().year, date.isocalendar().week
                if week < 10:
                    cw.append(f'{year}-0{week}')
                else:
                    cw.append(f'{year}-{week}')
            column.apply(take_year_week)
            return cw
            
        # Read database and get information
        file_type = (('text files', '*.txt'), ('All files', '*.*'))
        db_path = filedialog.askopenfilenames(title='Open files', initialdir=os.getcwd(), filetypes=file_type)

        if len(db_path) == 0:  # if cancel button is pressed
            return

        # When database is imported, make the following buttons as enabled.
        self.btn_create_table["state"] = "normal"
        self.btn_plot_laufstrecke["state"] = "normal"
        self.btn_export_ppt["state"] = "normal"
        self.btn_create_plan["state"] = "normal"
        self.btn_create_plan['state'] = 'normal'

        # Import database
        self.db_df = pd.read_csv(db_path[0], sep='\t', on_bad_lines='skip', header=0)
        
        # Turn date column into proper format and delete empty rows if motornumber or date is empty
        self.db_df['Zeitstempel'] = convert_excel_date(self.db_df['Zeitstempel'])
        self.db_df = self.db_df.dropna(subset=['Motornummer', 'Zeitstempel'])
        
        # Delete the rows if 'Laufstrecke' and 'Soll-Laufstrecke' has non-numeric values
        self.db_df['Laufstrecke'] = pd.to_numeric(self.db_df['Laufstrecke'], errors='coerce')
        self.db_df['Solllaufstrecke'] = pd.to_numeric(self.db_df['Solllaufstrecke'], errors='coerce')
        self.db_df = self.db_df.dropna(subset=['Laufstrecke', 'Solllaufstrecke'])
        
        # Add Calendar Week Column and sort dataframe by it       
        self.db_df['KW'] = pd.DataFrame(to_cw(self.db_df['Zeitstempel']), index=self.db_df.index)
        self.db_df = self.db_df.sort_values(by='Zeitstempel')  #sort values by KW
        self.db_df['Zeitstempel'] = self.db_df['Zeitstempel'].astype(str)

        self.db_df_uniques_keep_first = self.db_df.drop_duplicates(subset=['Motornummer'], keep='first') # get start date when motornummer is firstly seen
        self.db_df_uniques_keep_last = self.db_df.drop_duplicates(subset=['Motornummer'], keep='last') # get end date when motornummer is lastly seen

        self.db_df["Start_date"] = np.nan
        self.db_df["Start_KW"] = np.nan
        
        # Check dtypes for critical columns
        self.db_df['Laufstrecke'] = self.db_df['Laufstrecke'].astype('float64')
        self.db_df['Solllaufstrecke'] = self.db_df['Solllaufstrecke'].astype('int64')
        
        # Reset index
        self.db_df = self.db_df.reset_index(drop=True)
        
        # Get start and end date & kw and add to dataframe
        for ctr_row in range(len(self.db_df)):
            motornummmer = self.db_df.loc[ctr_row, "Motornummer"]

            start_date = self.db_df_uniques_keep_first.loc[self.db_df_uniques_keep_first["Motornummer"] == motornummmer, "Zeitstempel"].item()
            start_cw = self.db_df_uniques_keep_first.loc[self.db_df_uniques_keep_first["Motornummer"] == motornummmer, "KW"].item()

            end_date = self.db_df_uniques_keep_last.loc[self.db_df_uniques_keep_last["Motornummer"] == motornummmer, "Zeitstempel"].item()
            end_cw = self.db_df_uniques_keep_last.loc[self.db_df_uniques_keep_last["Motornummer"] == motornummmer, "KW"].item()

            self.db_df.loc[ctr_row, "Start_date"] = start_date
            self.db_df.loc[ctr_row, "Start_KW"] = start_cw

            self.db_df.loc[ctr_row, "End_date"] = end_date
            self.db_df.loc[ctr_row, "End_KW"] = end_cw

        # dataframe generation for "create table" functionality
        # Create global cw list
        first_global_cw = self.db_df.drop_duplicates(subset=['KW'])["KW"].values[0]
        last_global_cw = self.db_df.drop_duplicates(subset=['KW'])["KW"].values[-1]

        global_cw = []
        year_first = int(first_global_cw.split("-")[0])
        cw_first = int(first_global_cw.split("-")[1])
        year_last = int(last_global_cw.split("-")[0])
        cw_last = int(last_global_cw.split("-")[1])
        if year_first < year_last:
            num_added_cw = 52 - cw_first + cw_last + 1
        elif year_first == year_last:
            num_added_cw = cw_last - cw_first + 1
        else:
            return 0  # Hata mesajı ekle

        cw_added = cw_first
        for _ in range(num_added_cw):
            if cw_added < 10:
                cw_added_str = "0" + str(cw_added)
                year_added_str = str(year_first)
            elif cw_added >= 10 and cw_added <= 52:
                cw_added_str = str(cw_added)
                year_added_str = str(year_first)
            else:
                cw_added_str = '0' + str(cw_added % 52)
                year_added_str = str(year_last)
            global_cw.append(year_added_str + "-" + cw_added_str)
            cw_added += 1
        self.global_cw_df = pd.DataFrame(global_cw, columns=["KW"])

        # Weekly representation using last values
        df_table_list = []
        df_table = None
        for k, group in enumerate(self.db_df.groupby('Motornummer')):
            modified_group = pd.DataFrame(columns=['KW', 'PST', 'Motornummer', 'Laufstrecke', 'Solllaufstrecke'])
            modified_group["KW"] = global_cw

            group = group[1].drop_duplicates(subset=['KW'], keep='last')
            group = group[['KW', 'PST', 'Motornummer', 'Laufstrecke', 'Solllaufstrecke']].reset_index(drop=True)

            for ctr_row in range(len(modified_group)):
                curr_week = modified_group.loc[ctr_row, "KW"]
                if curr_week in group["KW"].values:
                    row_val = group.where(group["KW"] == curr_week).dropna().reset_index(drop=True)
                    modified_group.loc[ctr_row, :] = row_val.loc[0, :]
            if k == 0:
                df_table = modified_group
            else:
                modified_group = modified_group.drop(columns="KW")
                df_table = pd.concat([df_table, modified_group], axis=1)
            df_table_list.append(modified_group)

        self.df_table = df_table
        self.df_table_list = df_table_list

        # Dataframe generation for create plan functionality
        df_plan_list_last_values = []
        for k, group in enumerate(self.db_df.groupby('Motornummer')):
            modified_group = pd.DataFrame(columns=['KW', 'Zeitstempel', 'PST', 'Motornummer', 'Laufstrecke', 'Solllaufstrecke'])
            modified_group["KW"] = global_cw

            group = group[1].drop_duplicates(subset=['KW'], keep='last')
            group = group[['KW', 'Zeitstempel', 'PST', 'Motornummer', 'Laufstrecke', 'Solllaufstrecke']].reset_index(drop=True)

            for ctr_row in range(len(modified_group)):
                curr_week = modified_group.loc[ctr_row, "KW"]
                if curr_week in group["KW"].values:
                    row_val = group.where(group["KW"] == curr_week).dropna().reset_index(drop=True)
                    modified_group.loc[ctr_row, :] = row_val.loc[0, :]

            df_plan_list_last_values.append(modified_group)

        self.df_plan_list_last_values = df_plan_list_last_values.copy()

        # Weekly representation using first values
        df_plan_list_first_values = []
        for k, group in enumerate(self.db_df.groupby('Motornummer')):
            modified_group = pd.DataFrame(columns=['KW', 'Zeitstempel', 'PST', 'Motornummer', 'Laufstrecke', 'Solllaufstrecke'])
            modified_group["KW"] = global_cw

            group = group[1].drop_duplicates(subset=['KW'], keep='first')  ### Here is the only difference from above part
            group = group[['KW', 'Zeitstempel', 'PST', 'Motornummer', 'Laufstrecke', 'Solllaufstrecke']].reset_index(drop=True)

            for ctr_row in range(len(modified_group)):
                curr_week = modified_group.loc[ctr_row, "KW"]
                if curr_week in group["KW"].values:
                    row_val = group.where(group["KW"] == curr_week).dropna().reset_index(drop=True)
                    modified_group.loc[ctr_row, :] = row_val.loc[0, :]

            df_plan_list_first_values.append(modified_group)

        self.df_plan_list_first_values = df_plan_list_first_values.copy()

        # Create messagebox on the screen
        messagebox.showinfo("Info", f"Database file is imported!\n\nDatabase path: {db_path[0]}")
    
    def get_motor_numbers(self, options):
        dialog = MotorNumberSelectionDiaglog(self, options)
        self.wait_window(dialog.dialog)
        if dialog.selected_options:
            return dialog.selected_options
        else:
            messagebox.showerror('Please enter at least one motor number.')
    
    def get_milestone_name(self):
        title = 'Select Milestone Name'
        labels = ['Please enter you milestone name:', 
                  "Note: For default name, you can leave it empty."]
        dialog = CustomInputDialog(self, title, labels)
        milestone_name = dialog.strings
        if milestone_name:
            return milestone_name
        else:
            return ''
    
    def get_milestone(self):
        dict_milestone = {}
        for child in self.milestone_tree.get_children():
            # Milestone data from tree
            milestone_name = self.milestone_tree.item(child)['values'][0]
            milestone_date = self.milestone_tree.item(child)['values'][1]
            dict_milestone[milestone_name] = milestone_date
        return dict_milestone
    