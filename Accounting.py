# !/usr/bin/env python
# coding: utf-8

"""
    Accounting program

    Program written, interpreted and created by Gabriele Lo Cascio.

    Contacts:
        LinkedIn: https://www.linkedin.com/in/gabriele-locascio
        GitHub: https://github.com/Gabro29
        Fiverr: https://it.fiverr.com/gabro_29?up_rollout=true
        YouTube: https://www.youtube.com/channel/UCkGvbGqYzDi3lfgtbQ_pngg
        PayPal: https://www.paypal.com/paypalme/Gabro29
        Instagram: https://www.instagram.com/ga8ro


    Versions:
        - Alpha version released on 09/12/2022
        - Version 1.00 released on 01/15/2022
        - Version 1.50 released on 03/20/2022
        - Verison 3.00 released on 12/09/2022

    Background:
        This SoftWare was created in order to replace that of Francesco Paolo Lo Cascio which was provided courtesy
        of his previous employer Enzo Gulizzi. It was a program written in MS-Dos, no longer supported by Win10
        and therefore no longer in use. In addition, the convention of naming revenue has been maintained with 'HAVING'
        and exits with "GIVING" to avoid misunderstandings from Lo Cascio Senior.

    How it works:
        The program consists of a main class called 'Cashier', due to the direct link with the concept of accounting,
        to which nine other classes are added as a corollary (MenuP, Input, PrintStorage, PrintDay, PrintSingle,
         EditValue, Reset, UndoLastImport, Trasmission) developed as you go along to the needs.

    Cashier:
        Main class that acts as a master container for the nine frames of which the software is composed,
        although they can be added other frames and other classes according to various needs. Furthermore, has a menu
        with three functions:
            -File:
                ° For importing an Excel file.
            -Options:
                ° Client List:
                    Editing the list of stored names;
                ° Print options:
                    Change the separator used for export of the txt file;
                ° Create Backup:
                    Create a local Backup in a specific folder.
            -Help:
                ° Contact me:
                    Provides the manufacturer's email address;
                ° About:
                    Show a disclaimer for using SoftWare.

    MenuP:
        It is the main hub from which you can navigate to the others screenshots. And, vice versa, it can be
        reached by each screen.

    Input:
        Screen of daily use, allows you to enter the entries (HAVING) and exits (GIVING) and it is possible to
        carry out this procedure is done manually by filling in the appropriate fields either by importing an Excel
        spreadsheet. Also, it is possible select the date of the entered data. Finally, every time an account is
        entered and the total balance is updated and the one of the single account is shown in such a way as to correct
        any errors before you even save your work.

    PrintStorage:
        Display the list of all the present accounts and the respective totals with the addition at the bottom of the
        total balance sheet. It is also possible export the entire archive of accounts to txt or just
        those having a particular prefix.

    PrintDay:
        Along the lines of PrintStorage, this class operates at same way. What it sets out to do is go back
        over time, then by selecting a particular date it is possible to view the total balance for that day.
        Very useful when looking for errors.
        N.B:
            See EditValue and Reset for further information.

    PrintSingle:
        This class is in charge of analyzing every single account in the following ways:
            -List of relative movements;
            -Bar chart by date of movements;
            -Selective filtering of data by date range;
            -Txt export of movements;
            -Balance related to the account (Debit, Credit, ZERO).

    EditValue:
        As the name suggests, this class deals with the modification and removal of the data entered.
        First you look for an account, and then you operate the desired
        modification.
        N.B:
            It also works in combination with PrintDay, in fact in case you notice any errors in data entry
            by clicking on the respective line you can switch to edit screen and repair the mistake made.
            Finally, you can return to PrintDay and continue the review or to MenuP.

    Reset:
        Such a class is very minimalist and takes care of the removal of all those accounts whose difference
        between HAVING and GIVING is zero.
        N.B:
            If the reset is done then the balance for some dates may change because they are being deleted some
            of the accounts they have only at the time of removal a nil balance, but previously it was not.

    UndoLastImport:
        This screen searches in the dataset the most recently inserted data and allows you to remove them, maybe
        just in case you make some mistake and want ;to postpone it immediately.

    Transmission:
        This class deals with the migration of an account having one name to another. In case you want to change
        a series of entries in bulk. For example transfer at the end of the month the profit to an account previously
        created. So much in the end you can always use the date or data analysis.
"""

# Adding Imports
from tkinter import font as tkfont
from ttkthemes import ThemedStyle
from tkinter import Label, NO, W, CENTER, Frame, Entry, Listbox, E, TclError, END, Toplevel, Tk, Menu, Button, \
    LabelFrame, Scrollbar, Text, ttk, messagebox
from tkcalendar import DateEntry
from datetime import date
from datetime import datetime
from tkinter.filedialog import asksaveasfile, askopenfile
from pandas import read_csv, set_option, DataFrame, Series, concat, read_excel, to_datetime
from numpy import array, vstack
from math import fsum
from matplotlib.pyplot import figure, show, setp
from seaborn import set_style
from matplotlib import use

from os import getcwd
from shutil import copyfile
from threading import Thread
from smtplib import SMTP_SSL
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import email.mime.application

# Custom Name For File
bill_names_file = "bill_names.txt"
try:
    txt = open(f"{bill_names_file}", "xt")
    txt.close()
except FileExistsError:
    pass
with open(f"{bill_names_file}", "r") as file:
    bill_names = [line.rstrip() for line in file]

dataset_file = "dataframe.csv"
try:
    my_csv = open(f"{dataset_file}", "xt")
    my_csv.close()
    with open(f"{dataset_file}", "w") as f:
        f.write("Name,Description,Amount,Date")
except FileExistsError:
    pass

global_data_in_out = read_csv(f"{dataset_file}", engine="c")
icon_app = "icon.ico"

# Adding Global Variables

set_option('precision', 2)

count_d = 0
count_a = 0
name_day = ""
desc_day = ""
give_day = 0
have_day = 0
date_day = ""
name_in_the_box = ''
data_in = list()
data_out = list()
date_on_excel_file = str()
starting = True
styles = "clearlooks"


class SecondWindow(Toplevel):
    """Templete class for TopLevel windows"""

    def __init__(self, parent, title, geometry, propagate, b):
        super().__init__(parent)
        self.title(title)
        self.geometry(geometry)
        self.propagate(propagate)
        self.resizable(0, 0)
        self.iconbitmap(f"{icon_app}")


# Utility Functions

def info_app(self):
    """Terms and conditions"""

    info_window = SecondWindow(self, "Info App", "720x140", False, (0, 0))
    style = ThemedStyle(info_window)
    style.theme_use(f"{styles}")

    info_label = Label(info_window,
                       text="This program was created for personal purposes by Gabro.\n"
                            "Disclosure of this program is prohibited\n"
                            "unless specifically requested by the manufacturer.\n"
                            "See the 'Help' menu for more info.",
                       font=("Spectral", 15),
                       foreground="black")

    info_label.pack(pady=10)


def contact_me(self):
    """Give mail address of productor"""

    contact_window = SecondWindow(self, "Mail Contact", "300x100", False, (0, 0))
    style = ThemedStyle(contact_window)
    style.theme_use(f"{styles}")

    label_contact = Text(contact_window, height=3, font=("Spectral", 15))
    label_contact.insert(1.0,
                         "    You can contact the producer\n               by email at:\n          gabri729@gmail.com")
    label_contact.pack()
    label_contact.configure(state="disabled")


def scandata():
    """Get current Date"""

    Mesi = {"January": "Gennaio", "February": "Febbraio", "March": "Marzo", "April": "Aprile", "May": "Maggio",
            "June": "Giugno", "July": "Luglio", "August": "Agosto", "September": "Settembre",
            "October": "Ottobre",
            "November": "Novembre", "December": "Dicembre"}
    oggi = date.today()
    mese = oggi.strftime("%B")
    dat = (oggi.strftime("%d"), Mesi[mese], oggi.strftime("%Y"))

    return dat


def send_mail():
    """Send csv file via mail"""

    global dataset_file

    porta = 465
    password = "your_password"
    smtp_server = "smtp.gmail.com"
    sender_email = "your_email"
    receiver_email = "your_email"

    # html to include in the body section
    giorno, mese, anno = scandata()
    time = datetime.now()
    current_time = time.strftime("%H:%M:%S")
    html = f"""BackUp File {giorno}/{mese}/{anno} Ore {current_time}"""

    # Creating message.
    msg = MIMEMultipart('alternative')
    msg['Subject'] = "BackUp"
    msg['From'] = sender_email
    msg['To'] = receiver_email

    # The MIME types for text/html
    HTML_Contents = MIMEText(html, 'html')

    with open(f"{dataset_file}", "rb") as myfile:
        attach = email.mime.application.MIMEApplication(myfile.read(), _subtype="csv")
    attach.add_header('Content-Disposition', 'attachment', filename=f"{dataset_file}")

    # Attachment and HTML to body message.
    msg.attach(attach)
    msg.attach(HTML_Contents)
    try:
        with SMTP_SSL(smtp_server, porta) as server:
            server.login(sender_email, password)
            server.sendmail(msg['From'], msg['To'], msg.as_string())
    except Exception as e:
        with open("exception.txt", "w") as file:
            file.write(f"{e}")


# Main Class Cashier

class Cashier(Tk):

    def __init__(self, *args, **kwargs):
        Tk.__init__(self, *args, **kwargs)
        self.title("Accounting program")
        self.geometry(f"{800}x{670}+{100}+{70}")
        self.propagate(True)
        self.resizable(1, 1)
        global icon_app
        self.iconbitmap(f"{icon_app}")
        self.title_font = tkfont.Font(family='Helvetica', size=18, weight="bold", slant="italic")

        # Overlap frames and then raise each one per time
        container = Frame(self)
        container.pack(side="top", fill="both", expand=True)
        container.grid_rowconfigure(0, weight=1)
        container.grid_columnconfigure(0, weight=1)

        self.frames = {}

        self.task(container)

    @staticmethod
    def on_start():
        """Backup dataset"""
        global dataset_file
        copyfile(f"{dataset_file}", fr"Backup\{dataset_file}")

    def task(self, container):
        """Main function for loading"""

        global starting

        # Add Menu
        self.menubar = Menu(self)
        self.add_menu()
        self.config(menu=self.menubar)

        for F in (Input, PrintStorage, PrintDay, PrintSingle, EditValue,
                  Reset, UndoLastImport, Trasmission, MenuP):
            page_name = F.__name__
            frame = F(parent=container, controller=self)
            self.frames[page_name] = frame
            frame.grid(row=0, column=0, sticky="nsew")

        self.show_frame("MenuP")

        starting = False

        try:
            import pyi_splash
            pyi_splash.close()
        except ModuleNotFoundError:
            pass

    def show_frame(self, page_name):
        """Show frame by the given page name"""

        global name_day

        if page_name in ("PrintDay", "PrintSingle", "Trasmission"):
            name_day = ""
            self.geometry("900x670")
        elif page_name == "EditValue":
            self.geometry("900x670")
        elif page_name in ("Reset", "UndoLastImport"):
            self.geometry("800x200")
        else:
            self.geometry("800x670")

        frame = self.frames[page_name]
        frame.tkraise()
        frame.event_generate("<<ShowFrame>>")

    def add_menu(self):
        """Add menu on main window"""

        filemenu = Menu(self.menubar, tearoff=0, font=("Lucinda Console", 10))
        filemenu.add_command(label="Open...", command=lambda: Thread(target=self.open_file).start())
        filemenu.add_separator()
        filemenu.add_command(label="Exit", command=self.quit)
        self.menubar.add_cascade(label="File", menu=filemenu)

        optionmenu = Menu(self.menubar, tearoff=0, font=("Lucinda Console", 10))
        optionmenu.add_command(label="Bill Name List", command=self.edit_bill_name_file)
        optionmenu.add_separator()
        optionmenu.add_command(label="Print Options", command=self.setting_stampa)
        optionmenu.add_separator()
        optionmenu.add_command(label="Do Backup", command=self.do_backup)
        self.menubar.add_cascade(label="Options", menu=optionmenu)

        helpmenu = Menu(self.menubar, tearoff=0, font=("Lucinda Console", 10))
        helpmenu.add_command(label="Contact Me", command=lambda: contact_me(self))
        helpmenu.add_separator()
        helpmenu.add_command(label="About...", command=lambda: info_app(self))
        self.menubar.add_cascade(label="Help", menu=helpmenu)

    def setting_stampa(self):
        """Custom your txt file"""

        print_window = SecondWindow(self, "Print Options", "800x308", False, (0, 0))
        style = ThemedStyle(stampa_window)
        style.theme_use(f"{styles}")

        separator_title = Label(print_window, text="Separator", font=("Spectral", 15), foreground="black",
                                relief="ridge")
        separator_title.place(relx=0.13, rely=0.06)

        exstension = Label(print_window, text="Exstension", font=("Spectral", 15), foreground="black",
                           relief="ridge")
        exstension.place(relx=0.59, rely=0.06)

        # ComboBox sep
        self.separator_selection = ttk.Combobox(print_window, values=["Underscore", "Line", "Dot", "Tabular"],
                                                font=("Lucinda Console", 15), foreground="black", state="readonly")
        self.separator_selection.place(relx=0.04, rely=0.15)
        self.separator_selection.bind("<<ComboboxSelected>>", self.see_anteprima)
        self.separator_dict = {"______": 0, "------": 1, "......": 2, "         ": 3}
        with open("frills.dat", "r") as file:
            for line in file:
                if line.split("=")[0] == "Separator":
                    my_separatore = line.split("=")[1][:-1]
                elif line.split("=")[0] == "Exstension":
                    my_estensione = line.split("=")[1][:-1]
        self.separator_selection.current(self.separator_dict[my_separatore])

        # ComboBox ext
        self.extension_selection = ttk.Combobox(print_window, values=["txt", "csv", "dat"],
                                                font=("Lucinda Console", 15), foreground="black", state="readonly")
        self.extension_selection.place(relx=0.5, rely=0.15)
        self.extension_selection.bind("<<ComboboxSelected>>", self.see_anteprima)
        self.extension_dict = {"txt": 0, "csv": 1, "dat": 2}
        self.extension_selection.current(self.extension_dict[my_estensione])

        preview = LabelFrame(print_window, text="Preview", font=("Times", 20),
                             foreground="black")
        preview.place(height=100, width=750, relx=0.04, rely=0.3)

        self.show_preview = Label(preview, font=("Lucinda Console", 25),
                                  foreground="black")
        self.show_preview.place(relx=0.05, rely=0.05)
        self.sep_dict = {"Underscore": "______", "Line": "------", "Dot": "......", "Tabular": "         "}
        self.show_preview.config(text=f"Name{self.sep_dict[self.separator_selection.get()]}Amount")

        save_button = Button(print_window, text="Save",
                             fg="black", bg="sandy brown", relief="raised",
                             activebackground="light gray", font=("Lucinda Console", 20))
        save_button.place(relx=0.42, rely=0.75)

    def do_backup(self):
        """Open Popup and do backup"""

        self.on_start()
        messagebox.showinfo("Copy of File", "Backup done in directory 'Backup'")

    def open_file(self):
        """Open Excel file on Input screen"""

        global data_in
        global data_out
        global date_on_excel_file

        error_entrate = [False, None]
        error_uscite = [False, None]

        open_xls = askopenfile(initialdir=getcwd(), title="Import Excel File",
                               mode="r", filetypes=[("Excel File", ".xls"), ("Excel File", ".xlsx")])

        if open_xls is not None:
            data_cashier = read_excel(open_xls.name)
            raw_date_time = (str(data_cashier[data_cashier.columns[0]][0]).split(" ")[0]).split("-")
            try:
                date_on_excel_file = f"{raw_date_time[2]}-{raw_date_time[1]}-{raw_date_time[0]}"
            except IndexError:
                date_on_excel_file = None

            for index, col in enumerate(data_cashier.columns):
                if index < 6:
                    data_cashier.drop(index, inplace=True)
                if index >= 4:
                    data_cashier.drop(col, axis=1, inplace=True)

            data_cashier.columns = ["Nomi_Entrate", "Importo_Entrate", "Nomi_Uscite", "Importo_Uscite"]

            # Entrate
            for major_index in range(6, data_cashier.Nomi_Entrate.size):
                if f"{data_cashier.Nomi_Entrate[major_index]}" != "nan":
                    try:
                        desc = str(data_cashier.Nomi_Entrate[major_index]).split("(")[1].replace(")", "")
                    except IndexError:
                        desc = ''

                    raff_name = ''
                    for letter in str(data_cashier.Nomi_Entrate[major_index]).split("(")[0].split(' '):
                        if letter != '':
                            raff_name += letter + ' '
                    raff_name = (raff_name[:-1]).upper()
                    raff_desc = ''
                    for letter in desc.split(' '):
                        if letter != '':
                            raff_desc += letter + ' '
                    raff_desc = (raff_desc[:-1])

                    try:
                        float(data_cashier.Importo_Entrate[major_index])
                    except Exception:
                        error_entrate[0] = True
                        error_entrate[1] = major_index + 8
                    else:
                        if str(data_cashier.Importo_Entrate[major_index]) != 'nan':
                            data_in.append([raff_name, raff_desc, float(data_cashier.Importo_Entrate[major_index])])
                        if "," in str(raff_name):
                            error_entrate[0] = True
                            error_entrate[1] = major_index + 8

            # Uscite
            for major_index in range(6, data_cashier.Nomi_Uscite.size):
                desc = ''
                if f"{data_cashier.Nomi_Uscite[major_index]}" != "nan":
                    try:
                        desc = str(data_cashier.Nomi_Uscite[major_index]).split("(")[1].replace(")", "")
                    except IndexError:
                        desc = ''

                    raff_name = ''
                    for letter in str(data_cashier.Nomi_Uscite[major_index]).split("(")[0].split(' '):
                        if letter != '':
                            raff_name += letter + ' '
                    raff_name = (raff_name[:-1]).upper()
                    raff_desc = ''

                    for letter in desc.split(' '):
                        if letter != '':
                            raff_desc += letter + ' '
                    raff_desc = (raff_desc[:-1])

                    try:
                        float(data_cashier.Importo_Uscite[major_index])
                    except Exception:
                        error_uscite[0] = True
                        error_uscite[1] = major_index + 8
                    else:
                        if str(data_cashier.Importo_Uscite[major_index]) != 'nan':
                            data_out.append([raff_name, raff_desc, float(data_cashier.Importo_Uscite[major_index])])
                        if "," in str(raff_name):
                            error_uscite[0] = True
                            error_uscite[1] = major_index + 8

            if error_entrate[0]:
                messagebox.showerror("Attenzione", "Controllare il File da importare!\n"
                                                   f"Errore nella colonna dell' AVERE\n"
                                                   f"alla linea {error_entrate[1]}")
            elif error_uscite[0]:
                messagebox.showerror("Attenzione", "Controllare il File da importare!\n"
                                                   f"Errore nella colonna del DARE\n"
                                                   f"alla linea {error_uscite[1]}")
            elif not error_entrate[0] and not error_uscite[0]:
                self.show_frame("Input")

    def edit_bill_name_file(self):
        """Edit bill_name_file window"""

        edit_file_window = SecondWindow(self, "Edit Bill Name File", "480x720", False, (0, 0))
        style = ThemedStyle(edit_file_window)
        style.theme_use(f"{styles}")

        title_label = Label(edit_file_window, text="Remove single bill name "
                                                   "or entire file", font=("Spectral", 15),
                            foreground="black", relief="ridge")
        title_label.pack(side="top", fill="x", pady=10)

        tree_style = ttk.Style()
        tree_style.configure("mystyle.Treeview.Heading",
                             font=('Spectral', 18, 'italic'))  # Modify the font of the headings
        tree_style.layout("mystyle.Treeview",
                          [('mystyle.Treeview.treearea', {'sticky': 'nswe'})])  # Remove the borders
        tree_style.configure('Treeview', rowheight=40)

        self.tree = ttk.Treeview(edit_file_window, style="mystyle.Treeview", selectmode="extended")

        self.tree["columns"] = "Nome"
        self.tree.column("#0", width=0, stretch=NO)
        self.tree.column("Nome", anchor=W, width=250)

        self.tree.heading("#0", text="", anchor=CENTER)
        self.tree.heading("Nome", text="Nome", anchor=W)

        # Create a list of all the menu items
        global bill_names

        for count, record in enumerate(bill_names):
            self.tree.insert(parent="", index="end", iid=count, text="", values=(record,))

        tree_scrolly = Scrollbar(edit_file_window, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=tree_scrolly.set)
        tree_scrolly.pack(side="right", fill="y")
        self.tree.pack()

        rm_sel = Button(edit_file_window, text="Remove selected", command=self.rmp_from_file,
                        fg="black", bg="Silver", relief="raised",
                        activebackground="light gray", font=("Times", 15), pady=10)
        rm_sel.place(relx=0.27, rely=0.71, width=200, height=30)

        rm_all = Button(edit_file_window, text="Remove All", command=self.rmall_from_file,
                        fg="black", bg="sandy brown", relief="raised",
                        activebackground="light gray", font=("Times", 10), pady=30)
        rm_all.place(relx=0.35, rely=0.81, width=120, height=30)

    def rmp_from_file(self):
        """Remove person in tree"""

        global bill_names

        selected = self.tree.focus()
        selected = self.tree.item(selected)
        bill_names.remove(selected["values"][0])
        selected = self.tree.selection()
        self.tree.delete(int(str(selected).split("(")[1][1]))

        with open(f"{bill_names}", "w") as file:
            for c in bill_names:
                file.write(f"{c}\n")

    def rmall_from_file(self):
        """Clean up txt file and tree view"""

        with open(f"{bill_names}", "w") as file:
            pass
        for child in self.tree.get_children():
            self.tree.delete(child)

    def see_anteprima(self, sel):
        """See preview"""

        if self.extension_selection.get() == "csv":
            self.separator_selection.config(state="disable")
            self.show_anteprima.config(text="Name,Amount")
        else:
            self.separator_selection.config(state="readonly")
            self.show_anteprima.config(text=f"Name{self.sep_dict[self.separator_selection.get()]}Amount")

            with open("frills.dat", "w") as file:
                file.write(f"Separator={self.sep_dict[self.separator_selection.get()]}\n")
                file.write(f"Exstension={self.extension_selection.get()}\n")
                file.write("Dir=")


# Main Menu


class MenuP(Frame):

    def __init__(self, parent, controller):
        Frame.__init__(self, parent)
        self.style = ThemedStyle(self)
        self.style.theme_use(f"{styles}")

        self.controller = controller

        self.menu_selection = LabelFrame(self, text="Main Menu", font=("Times", 15),
                                         foreground="black")
        self.menu_selection.place(height=600, width=650, relx=0.03, rely=0.03)

        input_button = Button(self.menu_selection, text="Input", bg="silver", relief="raised",
                              font=("Spectral", 15), command=lambda: controller.show_frame("Input"),
                              activebackground="sandy brown")

        print_button = Button(self.menu_selection, text="Archive Print", bg="silver", relief="raised",
                              font=("Spectral", 15), command=lambda: controller.show_frame("PrintStorage"),
                              activebackground="sandy brown")

        print_day_button = Button(self.menu_selection, text="Print Inputs Of The Day",
                                  bg="silver", relief="raised", font=("Spectral", 15),
                                  command=lambda: controller.show_frame("PrintDay"),
                                  activebackground="sandy brown")

        print_single_button = Button(self.menu_selection, text="Print Single Bill Name", bg="silver", relief="raised",
                                     font=("Spectral", 15), command=lambda: controller.show_frame("PrintSingle"),
                                     activebackground="sandy brown")

        edit_button = Button(self.menu_selection, text="Edit Values", bg="silver", relief="raised",
                             font=("Spectral", 15), command=lambda: controller.show_frame("EditValue"),
                             activebackground="sandy brown")

        reset_button = Button(self.menu_selection, text="Reset", bg="silver", relief="raised",
                              font=("Spectral", 15), command=lambda: controller.show_frame("Reset"),
                              activebackground="sandy brown")

        undo_button = Button(self.menu_selection, text="Undo Last Import", bg="silver", relief="raised",
                             font=("Spectral", 15), command=lambda: controller.show_frame("UndoLastImport"),
                             activebackground="sandy brown")

        trasmissione_button = Button(self.menu_selection, text="Trasmission", bg="silver",
                                     relief="raised",
                                     font=("Spectral", 15), command=lambda: controller.show_frame("Trasmission"),
                                     activebackground="sandy brown")

        input_button.place(height=30, width=500, relx=0.1, rely=0.1)
        print_button.place(height=30, width=500, relx=0.1, rely=0.2)
        print_day_button.place(height=30, width=500, relx=0.1, rely=0.3)
        print_single_button.place(height=30, width=500, relx=0.1, rely=0.4)
        edit_button.place(height=30, width=500, relx=0.1, rely=0.5)
        reset_button.place(height=30, width=500, relx=0.1, rely=0.6)
        undo_button.place(height=30, width=500, relx=0.1, rely=0.7)
        trasmissione_button.place(height=30, width=500, relx=0.1, rely=0.8)

        try:
            Thread(target=send_mail).start()
        except Exception as e:
            with open("exception.txt", "w") as file:
                file.write(f"{e}")


# Input Screen


class Input(Frame):

    def __init__(self, parent, controller):
        Frame.__init__(self, parent)
        self.style = ThemedStyle(self)
        self.style.theme_use(f"{styles}")

        self.controller = controller
        label = Label(self, text="Input", font=("Spectral", 15), foreground="black", relief="ridge")
        label.pack(side="top", fill="x", pady=10)

        button = Button(self, text="Back to main menu", command=lambda: self.back_to_menu(controller),
                        fg="black", bg="sandy brown", relief="raised",
                        activebackground="light gray", font=("Lucinda Console", 10))
        button.pack()

        self.container = LabelFrame(self)
        self.container.place(height=308, width=720, relx=0.05, rely=0.13)

        self.in_out_tabs = ttk.Notebook(self.container)

        self.mesi_ordine = {"Gennaio": 1, "Febbraio": 2, "Marzo": 3, "Aprile": 4, "Maggio": 5, "Giugno": 6,
                            "Luglio": 7, "Agosto": 8, "Settembre": 9, "Ottobre": 10, "Novembre": 11, "Dicembre": 12}

        # Create Frames
        self.give_to_me_tab = Frame(self.in_out_tabs, width=740, height=308)  # , bg="#d9ead3")
        self.give_to_you_tab = Frame(self.in_out_tabs, width=740, height=308)  # , bg="#d5c6d7")

        self.in_out_tabs.add(self.give_to_me_tab, text="GIVING (E)")
        self.in_out_tabs.add(self.give_to_you_tab, text="HAVING (I)")

        self.in_out_tabs.place(relx=0, rely=0)
        self.in_out_tabs.bind("<<NotebookTabChanged>>", lambda x: self.change_tab_color(x))

        self.style_tab = ttk.Style()
        self.style_tab.map('TNotebook.Tab', background=[('selected', 'red'), ('active', 'lightgreen')])

        # Insert tree in tab
        self.tree_have()
        self.tree_give()

        # Create the Entries
        self.name_entry = Entry(self, font=("Lucinda Console", 15))
        self.name_entry.place(relx=0.05, rely=0.6, width=160, height=30)
        self.name_entry.bind('<KeyRelease>', self.see_balance)

        self.description_entry = Entry(self, font=("Lucinda Console", 15))
        self.description_entry.place(relx=0.28, rely=0.6, width=250, height=30)

        self.amount_entry = Entry(self, font=("Lucinda Console", 15))
        self.amount_entry.place(relx=0.62, rely=0.6, width=110, height=30)

        # Create a Listbox widget to display the list of items
        self.suggestion = Listbox(self, font=("Lucinda Console", 15), relief="flat")
        self.suggestion.place(relx=0.05, rely=0.64, width=160, height=100)
        self.suggestion.bind("<Double-Button-1>", self.select_suggestion)
        self.suggestion.bind("<Return>", self.select_suggestion)

        global bill_names

        # Add values to combobox
        self.update_suggestion_list(bill_names)

        # Add buttons
        add_button = Button(self, text="Add", command=lambda: Thread(target=self.add_person).start(),
                            fg="black", bg="Silver", relief="raised",
                            activebackground="light gray", font=("Times", 15))
        add_button.place(relx=0.8, rely=0.6, width=110, height=30)
        add_button.bind("<Return>", self.enter_add_button)

        clear_button = Button(self, text="Clear", command=self.clear_boxes,
                              fg="black", bg="light gray", relief="raised",
                              activebackground="#d9ead3", font=("Times", 15))
        clear_button.place(relx=0.5, rely=0.7, width=110, height=30)

        remove_button = Button(self, text="Remove", command=self.remove_person,
                               fg="black", bg="sandy brown", relief="raised",
                               activebackground="light gray", font=("Times", 15))
        remove_button.place(relx=0.68, rely=0.7, width=110, height=30)

        edit_button = Button(self, text="Edit", command=self.edit_person,
                             fg="black", bg="wheat1", relief="raised",
                             activebackground="light gray", font=("Times", 15))
        edit_button.place(relx=0.85, rely=0.7, width=110, height=30)

        undo_import = Button(self, text="Undo\nImport", command=self.undo,
                             fg="black", bg="#a8bed0", relief="raised",
                             activebackground="Silver", font=("Times", 12))
        undo_import.place(relx=0.3, rely=0.7, width=110, height=40)

        self.on_start()

        self.bind("<<ShowFrame>>", self.on_show_frame)

    def on_show_frame(self, event):
        """ON Raise"""

        global data_in
        global data_out
        global count_d
        global count_a
        global starting
        global bill_names
        global bill_names_file
        global global_data_in_out
        global date_on_excel_file
        global dataset_file

        if not starting:
            global_data_in_out = read_csv(f"{dataset_file}", engine="c")
            if len(data_in) != 0 or len(data_out) != 0:

                for i_exit in range(len(data_out)):
                    name = str(data_out[i_exit][0])
                    raff_name = ''

                    for letter in name.split(' '):
                        if letter != '':
                            raff_name += letter + ' '

                    if (raff_name[:-1]).upper() not in bill_names:
                        bill_names.append(name)
                        bill_names.sort()
                        with open(f"{bill_names_file}", "w") as file:
                            for c in bill_names:
                                file.write(f"{c.upper()}\n")

                    if count_d % 2 == 0:
                        self.tree_give_tme.insert(parent="", index="end", iid=count_d, text="",
                                                  values=data_out[i_exit],
                                                  tags=("evenrow",))
                    else:
                        self.tree_give_tme.insert(parent="", index="end", iid=count_d, text="",
                                                  values=data_out[i_exit],
                                                  tags=("oddrow",))
                    count_d += 1

                for i_entry in range(len(data_in)):

                    name = str(data_in[i_entry][0])
                    raff_name = ''

                    for letter in name.split(' '):
                        if letter != '':
                            raff_name += letter + ' '

                    if (raff_name[:-1]).upper() not in bill_names:
                        bill_names.append(name)
                        bill_names.sort()
                        with open(f"{bill_names_file}", "w") as file:
                            for c in bill_names:
                                file.write(f"{c.upper()}\n")

                    if count_a % 2 == 0:
                        self.tree_give_tyou.insert(parent="", index="end", iid=count_a, text="",
                                                   values=data_in[i_entry],
                                                   tags=("evenrow",))
                    else:
                        self.tree_give_tyou.insert(parent="", index="end", iid=count_a, text="",
                                                   values=data_in[i_entry],
                                                   tags=("oddrow",))
                    count_a += 1

                self.diff_totali()
                self.date_entry.destroy()

                # Check if the Date is spoecified in Excel file
                if date_on_excel_file is None:
                    giorno, mese, anno = scandata()
                    date_on_excel_file = f"{giorno}-{self.mesi_ordine[mese]}-{anno}"
                    raff_month = self.mesi_ordine[mese]
                elif str(date_on_excel_file.split("-")[1])[0] == "0":
                    raff_month = str(date_on_excel_file.split("-")[1])[1:]
                else:
                    raff_month = date_on_excel_file.split("-")[1]

                self.date_entry = DateEntry(self, selectmode="day", font="Lucinda 10", locale="it_IT",
                                            showweeknumbers=False,
                                            showothermonthdays=True,
                                            year=int(date_on_excel_file.split("-")[2]),
                                            month=int(raff_month),
                                            day=int(date_on_excel_file.split("-")[0]),
                                            background="dark slate gray", date_pattern="DD/MM/YYYY",
                                            bordercolor="dark slate gray", selectbackground="SlateGray4",
                                            headersbackground="light gray", normalbackground="light gray",
                                            foreground='white',
                                            normalforeground='black', headersforeground='black',
                                            weekendbackground="IndianRed1")
                self.date_entry.place(relx=0.81, rely=0.09)
                self.date_entry.bind("<<DateEntrySelected>>", self.select_data)
            else:
                self.diff_totali()
                self.clear_boxes()
                self.name_entry.focus()

    def on_start(self):
        """Build Frame"""
        # Let user select the day
        giorno, mese, anno = scandata()

        self.date_entry = DateEntry(self, selectmode="day", font="Lucinda 10", locale="it_IT", showweeknumbers=False,
                                    showothermonthdays=True,
                                    year=int(anno), month=self.mesi_ordine[mese], day=int(giorno),
                                    background="dark slate gray", date_pattern="DD/MM/YYYY",
                                    bordercolor="dark slate gray", selectbackground="SlateGray4",
                                    headersbackground="light gray", normalbackground="light gray", foreground='white',
                                    normalforeground='black', headersforeground='black', weekendbackground="IndianRed1")
        self.date_entry.place(relx=0.81, rely=0.09)
        self.date_entry.bind("<<DateEntrySelected>>", self.select_data)
        self.today = True

        # See stored bill
        live_attributes_d = Label(self, text="GIVING (E)", font=("Spectral", 15), foreground="black", relief="ridge")
        live_attributes_d.place(relx=0.01, rely=0.81, width=110, height=30)

        live_attributes_a = Label(self, text="HAVING (I)", font=("Spectral", 15), foreground="black", relief="ridge")
        live_attributes_a.place(relx=0.01, rely=0.86, width=110, height=30)

        self.live_attributes_t = Label(self, text="BALANCE", font=("Spectral", 11), foreground="black", relief="ridge")
        self.live_attributes_t.place(relx=0.01, rely=0.915, width=110, height=40)

        self.give_live_label = Label(self, font=("Spectral", 15), foreground="black", relief="ridge")
        self.give_live_label.place(relx=0.15, rely=0.81, width=130, height=30)

        self.have_live_label = Label(self, font=("Spectral", 15), foreground="black", relief="ridge")
        self.have_live_label.place(relx=0.15, rely=0.86, width=130, height=30)

        self.balance_live_label = Label(self, font=("Spectral", 15), foreground="black", relief="ridge")
        self.balance_live_label.place(relx=0.15, rely=0.915, width=130, height=40)

        tot_label = Label(self, text="Total", font=("Spectral", 18), foreground="black", relief="ridge")
        tot_label.config(anchor=CENTER)
        tot_label.place(height=38, width=100, relx=0.46, rely=0.835)

        self.give_label = Label(self, text="", font=("Spectral", 15),
                                foreground="black",
                                relief="ridge")
        self.give_label.config(anchor=CENTER)
        self.give_label.place(height=38, width=150, relx=0.6, rely=0.835)

        self.have_label = Label(self, text="", font=("Spectral", 15),
                                foreground="black",
                                relief="ridge")
        self.have_label.config(anchor=CENTER)
        self.have_label.place(height=38, width=150, relx=0.79, rely=0.835)

        self.diff_label = Label(self, text="",
                                font=("Spectral", 15),
                                foreground="black", relief="ridge")
        self.diff_label.config(anchor=CENTER)
        self.diff_label.place(height=40, width=150, relx=0.69, rely=0.90)

        self.diff_label_t = Label(self, text="Difference", font=("Spectral", 18), foreground="black", relief="ridge")
        self.diff_label_t.config(anchor=CENTER)
        self.diff_label_t.place(height=40, width=150, relx=0.46, rely=0.90)

        self.suggestion.bind("<Tab>", self.see_balance)

    def change_tab_color(self, color):
        """Change color of the tab"""

        if self.in_out_tabs.index("current") == 1:
            self.name_entry.focus()
            self.style_tab = ttk.Style()
            self.style_tab.map('TNotebook.Tab', background=[('selected', 'lightgreen'), ('active', 'red')])
        elif self.in_out_tabs.index("current") == 0:
            self.name_entry.focus()
            self.style_tab = ttk.Style()
            self.style_tab.map('TNotebook.Tab', background=[('selected', 'red'), ('active', 'lightgreen')])

    def tree_give(self):
        """Add GIVING Tree in tab"""

        tree_style = ttk.Style()
        tree_style.configure("mystyle.Treeview", highlightthickness=0, bd=0,
                             font=('Lucinda Console', 16))  # Modify the font of the body
        tree_style.configure("mystyle.Treeview.Heading",
                             font=('Spectral', 18, 'italic'), width=1, pady=20)  # Modify the font of the headings
        tree_style.layout("mystyle.Treeview", [('mystyle.Treeview.treearea', {'sticky': 'nswe'})])  # Remove the borders
        tree_style.configure('Treeview', rowheight=40)

        self.tree_give_tme = ttk.Treeview(self.give_to_me_tab, style="mystyle.Treeview", selectmode="extended")

        self.tree_give_tme["columns"] = ("Name", "Description", "Amount")
        self.tree_give_tme.column("#0", width=0, stretch=NO)
        self.tree_give_tme.column("Name", anchor=W, width=250, stretch=NO)
        self.tree_give_tme.column("Description", anchor=CENTER, width=280, stretch=NO)
        self.tree_give_tme.column("Amount", anchor=E, width=145, stretch=NO)

        self.tree_give_tme.heading("#0", text="", anchor=CENTER)
        self.tree_give_tme.heading("Name", text="Name", anchor=W)
        self.tree_give_tme.heading("Description", text="Description", anchor=CENTER)
        self.tree_give_tme.heading("Amount", text="Amount", anchor=E)

        self.tree_give_tme.tag_configure("oddrow", background="white")
        self.tree_give_tme.tag_configure("evenrow", background="IndianRed3")

        global count_d
        count_d = 0
        for record in []:
            if count_d % 2 == 0:
                self.tree_give_tme.insert(parent="", index="end", iid=count_d, text="", values=record,
                                          tags=("evenrow",))
            else:
                self.tree_give_tme.insert(parent="", index="end", iid=count_d, text="", values=record, tags=("oddrow",))
            count_d += 1

        tree_give_tme_scrolly = Scrollbar(self.container, orient="vertical",
                                          command=self.tree_give_tme.yview)
        self.tree_give_tme.configure(yscrollcommand=tree_give_tme_scrolly.set)
        tree_give_tme_scrolly.pack(side="right", fill="y")

        self.tree_give_tme.bind("<<TreeviewSelect>>", self.display_edit_person)
        self.tree_give_tme.place(width=680, height=300, relx=0, rely=0)

    def tree_have(self):
        """Add HAVING Tree in tab"""

        tree_style = ttk.Style()
        tree_style.configure("mystyle.Treeview", highlightthickness=0, bd=0,
                             font=('Lucinda Console', 16))  # Modify the font of the body
        tree_style.configure("mystyle.Treeview.Heading",
                             font=('Spectral', 18, 'italic'))  # Modify the font of the headings
        tree_style.layout("mystyle.Treeview", [('mystyle.Treeview.treearea', {'sticky': 'nswe'})])  # Remove the borders
        tree_style.configure('Treeview', rowheight=40)

        self.tree_give_tyou = ttk.Treeview(self.give_to_you_tab, style="mystyle.Treeview", selectmode="extended")

        self.tree_give_tyou["columns"] = ("Name", "Description", "Amount")
        self.tree_give_tyou.column("#0", width=0, stretch=NO)
        self.tree_give_tyou.column("Name", anchor=W, width=250, stretch=NO)
        self.tree_give_tyou.column("Description", anchor=CENTER, width=280, stretch=NO)
        self.tree_give_tyou.column("Amount", anchor=E, width=145, stretch=NO)

        self.tree_give_tyou.heading("#0", text="", anchor=CENTER)
        self.tree_give_tyou.heading("Name", text="Name", anchor=W)
        self.tree_give_tyou.heading("Description", text="Description", anchor=CENTER)
        self.tree_give_tyou.heading("Amount", text="Amount", anchor=E)

        self.tree_give_tyou.tag_configure("oddrow", background="white")
        self.tree_give_tyou.tag_configure("evenrow", background="lightgreen")

        global count_a
        count_a = 0
        for record in []:
            if count_a % 2 == 0:
                self.tree_give_tyou.insert(parent="", index="end", iid=count_a, text="", values=record,
                                           tags=("evenrow",))
            else:
                self.tree_give_tyou.insert(parent="", index="end", iid=count_a, text="", values=record,
                                           tags=("oddrow",))
            count_a += 1

        tree_give_tyou_scrolly = Scrollbar(self.container, orient="vertical",
                                           command=self.tree_give_tyou.yview)  # command means update the yaxis view of the widget
        self.tree_give_tyou.configure(yscrollcommand=tree_give_tyou_scrolly.set)
        tree_give_tyou_scrolly.pack(side="right", fill="y")  # make the scrollbar fill the y axis of the Treeview widget

        self.tree_give_tyou.bind("<<TreeviewSelect>>", self.display_edit_person)
        self.tree_give_tyou.place(width=680, height=300, relx=0, rely=0)

    def add_person(self):
        """Add record on tree"""

        global count_d
        global count_a
        global name_in_the_box

        try:
            float(self.amount_entry.get())
        except Exception:
            messagebox.showerror("Attention", "Insert a vaid number in the field "
                                              "or insert zero to left it empty.")
        else:
            if name_in_the_box != '':
                name = str(self.name_entry.get())
                raff_name = ''

                for letter in name.split(' '):
                    if letter != '':
                        raff_name += letter + ' '

                if self.in_out_tabs.index("current") == 0:
                    if count_d % 2 == 0:
                        self.tree_give_tme.insert(parent="", index="end", iid=count_d, text="",
                                                  values=((raff_name[:-1]).upper(), self.description_entry.get(),
                                                          self.amount_entry.get()),
                                                  tags=("evenrow",))
                    else:
                        self.tree_give_tme.insert(parent="", index="end", iid=count_d, text="",
                                                  values=((raff_name[:-1]).upper(), self.description_entry.get(),
                                                          self.amount_entry.get()),
                                                  tags=("oddrow",))
                    count_d += 1

                elif self.in_out_tabs.index("current") == 1:
                    if count_a % 2 == 0:
                        self.tree_give_tyou.insert(parent="", index="end", iid=count_a, text="",
                                                   values=((raff_name[:-1]).upper(), self.description_entry.get(),
                                                           self.amount_entry.get()),
                                                   tags=("evenrow",))
                    else:
                        self.tree_give_tyou.insert(parent="", index="end", iid=count_a, text="",
                                                   values=((raff_name[:-1]).upper(), self.description_entry.get(),
                                                           self.amount_entry.get()),
                                                   tags=("oddrow",))
                    count_a += 1

                self.update_clienti()
                self.diff_totali()
                self.clear_boxes()
                self.name_entry.focus()

    def remove_person(self):
        """Remove person in tree"""

        global count_a
        global count_d

        try:
            if self.in_out_tabs.index("current") == 0:
                selected_tme = self.tree_give_tme.selection()
                if int(str(selected_tme).split("(")[1][1]) < count_d:
                    pass
                else:
                    count_d -= 1
                self.tree_give_tme.delete(selected_tme)

            elif self.in_out_tabs.index("current") == 1:
                selected_tyou = self.tree_give_tyou.selection()
                if int(str(selected_tyou).split("(")[1][1]) < count_a:
                    pass
                else:
                    count_a -= 1
                self.tree_give_tyou.delete(selected_tyou)

            self.diff_totali()
        except TclError:
            pass

    def edit_person(self):
        """Edit just added value in tree"""

        try:
            float(self.amount_entry.get())
        except Exception:
            messagebox.showerror("Attention", "Insert a vaid number in the field "
                                              "or insert zero to left it empty.")
        else:
            if self.name_entry.get() != "":
                if self.in_out_tabs.index("current") == 0:
                    selected = self.tree_give_tme.focus()
                    self.tree_give_tme.item(selected, text="", values=(
                        self.name_entry.get(), self.description_entry.get(), self.amount_entry.get()))
                elif self.in_out_tabs.index("current") == 1:
                    selected = self.tree_give_tyou.focus()
                    self.tree_give_tyou.item(selected, text="", values=(
                        self.name_entry.get(), self.description_entry.get(), self.amount_entry.get()))
                self.diff_totali()

        self.clear_boxes()

    def clear_boxes(self):
        """Clear up entry boxes"""

        self.name_entry.delete(0, END)
        self.description_entry.delete(0, END)
        self.amount_entry.delete(0, END)
        self.give_live_label.config(text="")
        self.have_live_label.config(text="")
        self.balance_live_label.config(text="")

    def display_edit_person(self, sel):
        """Fill Entry boxes with tree selection"""

        self.clear_boxes()

        if self.in_out_tabs.index("current") == 0:
            selected = self.tree_give_tme.focus()
            selected_row = self.tree_give_tme.item(selected, "values")
            self.name_entry.insert(0, selected_row[0])
            self.description_entry.insert(0, selected_row[1])
            self.amount_entry.insert(0, selected_row[2])
        else:
            selected = self.tree_give_tyou.focus()
            selected_row = self.tree_give_tyou.item(selected, "values")
            self.name_entry.insert(0, selected_row[0])
            self.description_entry.insert(0, selected_row[1])
            self.amount_entry.insert(0, selected_row[2])

        self.see_balance(0)

    def suggestion_key(self, key):
        """Keyboard suggestion base on value list"""

        global bill_names

        if self.name_entry.get() == '':
            self.update_suggestion_list(bill_names)
        else:
            self.update_suggestion_list([item for item in bill_names if self.name_entry.get().lower() in item.lower()])

    def update_suggestion_list(self, data):
        """Keyboard suggestion base on previous function"""

        # Clear the Combobox
        self.suggestion.delete(0, END)
        # Add values to the combobox
        for value in data:
            self.suggestion.insert(END, value)

    def select_suggestion(self, key):
        """Enter the suggestion on Entry box"""

        self.name_entry.delete(0, "end")
        for i in self.suggestion.curselection():
            if " " in self.suggestion.get(i):
                self.name_entry.insert(0, self.suggestion.get(i))
            else:
                self.name_entry.insert(0, self.suggestion.get(i))
        self.see_balance(0)

    def enter_add_button(self, key):
        """Add bill to tree via Return button"""

        self.add_person()

    def update_clienti(self):
        """Update clienti.txt base on new input"""

        global bill_names
        global bill_names_file

        name = str(self.name_entry.get())
        raff_name = ''

        for letter in name.split(' '):
            if letter != '':
                raff_name += letter + ' '

        if (raff_name[:-1]).upper() not in bill_names:
            bill_names.append(self.name_entry.get())
            bill_names.sort()
            with open(f"{bill_names_file}", "w") as file:
                for c in bill_names:
                    file.write(f"{c.upper()}\n")

    def back_to_menu(self, controller):
        """Return to main menu"""

        global data_out
        global data_in
        global count_a
        global count_d

        if count_d != 0 or count_a != 0:
            self.save_to_csv()
            count_d = 0
            count_a = 0

        self.delete_entries()

        data_in.clear()
        data_out.clear()
        controller.show_frame("MenuP")

    def delete_entries(self):
        """Clean up treeview"""

        for row in self.tree_give_tme.get_children():
            self.tree_give_tme.delete(row)
        for row in self.tree_give_tyou.get_children():
            self.tree_give_tyou.delete(row)

    def save_to_csv(self):
        """Update csv file with new entries"""

        global global_data_in_out
        global dataset_file

        names = list()
        descriptions = list()
        imports = list()

        # GIVE TREE
        for item in self.tree_give_tme.get_children():
            t_i = ""
            the_item = str(self.tree_give_tme.item(item).values()).split(",")
            names.append(str(the_item[2][3:len(the_item[2]) - 1]).upper())
            descriptions.append(f"{the_item[3][2:len(the_item[3]) - 1]}/")
            for i in the_item[4][1:len(the_item[4]) - 1]:
                if i in ("\n", "'"):
                    pass
                else:
                    t_i += i
            imports.append(-float(t_i))

        if self.today:
            giorno, mese, anno = scandata()
            data = [f"{giorno}-{self.mesi_ordine[mese]}-{anno}" for i in names]
        else:
            data = [f"{self.giorno}-{self.mesi_ordine[self.num_mesi[self.mese]]}-{self.anno}" for i in names]

        if len(names) != 0:
            for element in range(len(names)):
                for letter in range(len(names[element])):
                    if names[element][letter] == " " and letter == len(names[element]) - 1:
                        names[element] = names[element][:-1]

            tree_data_frame_0 = DataFrame(names)
            tree_data_frame_0["1"] = Series(descriptions)
            tree_data_frame_0["2"] = Series(imports)
            tree_data_frame_0["3"] = Series(data)
            tree_data_frame_0.columns = ["Name", "Description", "Amount", "Date"]

        names.clear()
        descriptions.clear()
        imports.clear()

        # AVERE TREE
        for item in self.tree_give_tyou.get_children():
            t_i = ""
            the_item = str(self.tree_give_tyou.item(item).values()).split(",")
            names.append(str(the_item[2][3:len(the_item[2]) - 1]).upper())
            descriptions.append(f"{the_item[3][2:len(the_item[3]) - 1]}/")
            for i in the_item[4][1:len(the_item[4]) - 1]:
                if i in ("\n", "'"):
                    pass
                else:
                    t_i += i
            imports.append(float(t_i))

        if self.today:
            giorno, mese, anno = scandata()
            data = [f"{giorno}-{self.mesi_ordine[mese]}-{anno}" for i in names]
        else:
            data = [f"{self.giorno}-{self.mesi_ordine[self.num_mesi[self.mese]]}-{self.anno}" for i in names]
        if len(names) != 0:
            for element in range(len(names)):
                for letter in range(len(names[element])):
                    if names[element][letter] == " " and letter == len(names[element]) - 1:
                        names[element] = names[element][:-1]

            tree_data_frame_1 = DataFrame(names)
            tree_data_frame_1["1"] = Series(descriptions)
            tree_data_frame_1["2"] = Series(imports)
            tree_data_frame_1["3"] = Series(data)
            tree_data_frame_1.columns = ["Name", "Description", "Amount", "Date"]

        prev_data_frame = global_data_in_out

        try:
            join = concat([prev_data_frame, tree_data_frame_0, tree_data_frame_1])
        except Exception:
            try:
                try:
                    join = concat([prev_data_frame, tree_data_frame_0])
                except Exception:
                    join = concat([prev_data_frame, tree_data_frame_1])
            except Exception:
                join = concat([prev_data_frame])
        join.to_csv(f"{dataset_file}", index=False)

    def select_data(self, key):
        """Select work day"""

        self.giorno = str(self.date_entry.get_date()).split("-")[2]
        self.mese = str(self.date_entry.get_date()).split("-")[1]
        self.anno = str(self.date_entry.get_date()).split("-")[0]
        self.num_mesi = {"01": "Gennaio", "02": "Febbraio", "03": "Marzo", "04": "Aprile", "05": "Maggio",
                         "06": "Giugno", "07": "Luglio", "08": "Agosto", "09": "Settembre", "10": "Ottobre",
                         "11": "Novembre", "12": "Dicembre"}
        self.giorno_di_lavoro = f"('{self.giorno}', '{self.num_mesi[self.mese]}', '{self.anno}')"
        self.today = False

    def see_balance(self, b):
        """Print on screen the SALDO"""

        self.suggestion_key(b)

        global name_in_the_box
        global global_data_in_out

        name_in_the_box = str(self.name_entry.get()).upper()

        if global_data_in_out[global_data_in_out["Name"] == str('SIG ' + name_in_the_box)].size > 0:
            self.name_entry.delete(0, END)
            self.name_entry.insert(0, str('SIG ' + name_in_the_box))
        else:
            self.name_entry.delete(0, END)
            self.name_entry.insert(0, name_in_the_box)

        # Dare based on better file
        give = round(fsum(global_data_in_out[global_data_in_out["Name"] == name_in_the_box][
                              global_data_in_out[global_data_in_out["Name"] == name_in_the_box]["Amount"] < 0][
                              "Amount"]), 2)
        # Avere based on better file
        have = round(fsum(global_data_in_out[global_data_in_out["Name"] == name_in_the_box][
                              global_data_in_out[global_data_in_out["Name"] == name_in_the_box]["Amount"] > 0][
                              "Amount"]), 2)

        self.give_live_label.config(text=f'{-1 * give}')
        self.have_live_label.config(text=f'{have}')

        if round(fsum([have, give]), 2) > 0:
            self.live_attributes_t.config(text="BALANCE\nGIVING")
            self.balance_live_label.config(text=f"{round(fsum([give, have]), 2)}")
        elif round(fsum([have, give]), 2) < 0:
            self.live_attributes_t.config(text="BALANCE\nHAVING")
            self.balance_live_label.config(text=f"{round(fsum([give, have]), 2) * -1}")
        elif round(fsum([have, give]), 2) == 0.0:
            self.live_attributes_t.config(text="BALANCE\nZERO")
            self.balance_live_label.config(text=f"{0}")

    def diff_totali(self):
        """Calculate the total of the balance"""

        global global_data_in_out

        imports_dare = list()
        for item in self.tree_give_tme.get_children():
            t_i = ""
            the_item = str(self.tree_give_tme.item(item).values()).split(",")
            for i in the_item[4][1:len(the_item[4]) - 1]:
                if i in ("\n", "'"):
                    pass
                else:
                    t_i += i
            imports_dare.append(-float(t_i))

        imports_avere = list()
        for item in self.tree_give_tyou.get_children():
            t_i = ""
            the_item = str(self.tree_give_tyou.item(item).values()).split(",")
            for i in the_item[4][1:len(the_item[4]) - 1]:
                if i in ("\n", "'"):
                    pass
                else:
                    t_i += i
            imports_avere.append(float(t_i))

        give = fsum([fsum(global_data_in_out[global_data_in_out["Amount"] < 0]["Amount"]), fsum(imports_dare)])
        self.give_label.config(text=str(round(give, 2) * -1))
        have = fsum([fsum(global_data_in_out[global_data_in_out["Amount"] > 0]["Amount"]), fsum(imports_avere)])
        self.have_label.config(text=str(round(have, 2)))
        if -0.001 <= round(fsum([have, give]), 2) <= 0.001:
            self.diff_label.config(text="0")
        else:
            self.diff_label.config(text=str(round(fsum([have, give]), 2)))

    def undo(self):
        """Remove rows imported by xls file"""

        global data_out
        global data_in
        global count_d
        global count_a

        if len(data_out) != 0 or len(data_in) != 0:
            self.delete_entries()
            data_in.clear()
            data_out.clear()
            count_d = 0
            count_a = 0
            self.diff_totali()
            self.name_entry.focus()


#  PrintStorage Screen


class PrintStorage(Frame):

    def __init__(self, parent, controller):
        Frame.__init__(self, parent)
        self.style = ThemedStyle(self)
        self.style.theme_use(f"{styles}")

        self.controller = controller

        label = Label(self, text="Archive Print", font=("Spectral", 15), foreground="black", relief="ridge")
        label.pack(side="top", fill="x", pady=10)

        button = Button(self, text="Back to main menu", command=lambda: controller.show_frame("MenuP"),
                        fg="black", bg="sandy brown", relief="raised",
                        activebackground="light gray", font=("Lucinda Console", 10))
        button.pack()

        tot_label = Label(self, text="Total" + "\t", font=("Spectral", 18), foreground="black", relief="ridge")
        tot_label.config(anchor=CENTER)
        tot_label.pack(side="bottom", fill="x", pady=40)

        export_button = Button(self, text="Credit Export", fg="black", bg="light gray", relief="ridge",
                               activebackground="silver", font=("Spectral", 15),
                               command=self.save_credit)
        export_button.place(height=30, width=170, relx=0.74, rely=0.078)

        export_all_button = Button(self, text="Archive Export", fg="black", bg="silver", relief="ridge",
                                   activebackground="wheat1", font=("Spectral", 15),
                                   command=self.save_all)
        export_all_button.place(height=30, width=170, relx=0.74, rely=0.02)

        refresh_button = Button(self, text="Update", fg="black", bg="wheat1", relief="ridge",
                                activebackground="silver", font=("Spectral", 15),
                                command=self.update_tree)
        refresh_button.place(height=30, width=170, relx=0.05, rely=0.078)

        self.stampa_container = LabelFrame(self)
        self.stampa_container.place(height=500, width=785, relx=0.01, rely=0.13)

        self.stampa_storage_tree([])

        self.give_label = Label(self, text="", font=("Spectral", 15),
                                foreground="black",
                                relief="ridge")
        self.give_label.config(anchor=CENTER)
        self.give_label.place(height=38, width=150, relx=0.6, rely=0.885)

        self.have_label = Label(self, text="", font=("Spectral", 15),
                                foreground="black",
                                relief="ridge")
        self.have_label.config(anchor=CENTER)
        self.have_label.place(height=38, width=150, relx=0.79, rely=0.885)

        self.diff_label = Label(self, text="",
                                font=("Spectral", 15),
                                foreground="black", relief="ridge")
        self.diff_label.config(anchor=CENTER)
        self.diff_label.place(height=35, width=150, relx=0.65, rely=0.94)

        self.diff_label_t = Label(self, text="Difference", font=("Spectral", 18), foreground="black", relief="ridge")
        self.diff_label_t.config(anchor=CENTER)
        self.diff_label_t.place(height=35, width=150, relx=0.46, rely=0.94)

    def speed_task(self):
        """Support func for loading"""

        global global_data_in_out
        global dataset_file
        global_data_in_out = read_csv(f"{dataset_file}", engine="c")
        size_of_unique = global_data_in_out.Name.unique().size
        raff_rows = array([0, 0, 0, 0])
        for index, bill_name in enumerate(global_data_in_out.sort_values(by=["Name"])["Name"].unique()):
            give = round(fsum(global_data_in_out[global_data_in_out["Name"] == bill_name][
                                  global_data_in_out[global_data_in_out["Name"] == bill_name]["Amount"] < 0][
                                  "Amount"]), 2)

            have = round(fsum(global_data_in_out[global_data_in_out["Name"] == bill_name][
                                  global_data_in_out[global_data_in_out["Name"] == bill_name]["Amount"] > 0][
                                  "Amount"]), 2)

            description = global_data_in_out[global_data_in_out["Name"] == bill_name][
                global_data_in_out[global_data_in_out["Name"] == bill_name]["Description"].astype(str) != "nan"][
                "Description"].astype(str).sum()

            if round(fsum([have, give]), 2) > 0:
                raff_rows = vstack(
                    [raff_rows,
                     [bill_name, str(description).lower(), 0, round(fsum([give, have]), 2)]])
            elif round(fsum([have, give]), 2) < 0:
                raff_rows = vstack(
                    [raff_rows,
                     [bill_name, str(description).lower(), round(fsum([give, have]), 2) * -1, 0]])

            self.load_label.config(text=f"Loading {round((100 * index) / size_of_unique, 1)}%")
            self.loading_window.update_idletasks()

        self.speed_diff_totali()
        self.loading_window.destroy()

        self.stampa_storage_tree(raff_rows)

    def update_tree(self):
        """Update new entry tree"""

        self.loading_window = SecondWindow(self, "Wait", "300x50", False, (0, 0))
        stile_load = ThemedStyle(self.loading_window)
        stile_load.theme_use(f"{styles}")
        self.loading_window.title("Wait")

        self.load_label = Label(self.loading_window, text="Reading the file", font=("Spectral", 15),
                                foreground="black", relief="ridge")
        self.load_label.pack(side="top", fill="x", pady=10)

        self.loading_window.after(1, Thread(target=self.speed_task).start())
        self.loading_window.mainloop()

    def speed_diff_totali(self):
        """Support func for GIVE - HAVE"""

        self.give = fsum(global_data_in_out[global_data_in_out["Amount"] < 0]["Amount"])
        self.have = fsum(global_data_in_out[global_data_in_out["Amount"] > 0]["Amount"])
        self.give_label.config(text=str(round(self.give, 2)))
        self.have_label.config(text=str(round(self.have, 2)))
        self.diff_label.config(text=str(round(fsum([self.have, self.give]), 2)))

    def stampa_storage_tree(self, data):
        """Add Tree in tab"""

        tree_style = ttk.Style()
        tree_style.configure("mystyle.Treeview", highlightthickness=0, bd=0,
                             font=('Lucinda Console', 16))  # Modify the font of the body
        tree_style.configure("mystyle.Treeview.Heading",
                             font=('Spectral', 18, 'italic'), width=1, pady=20)  # Modify the font of the headings
        tree_style.layout("mystyle.Treeview",
                          [('mystyle.Treeview.treearea', {'sticky': 'nswe'})])  # Remove the borders
        tree_style.configure('Treeview', rowheight=40)

        self.tree_stampa = ttk.Treeview(self.stampa_container, style="mystyle.Treeview")

        self.tree_stampa["columns"] = ("Name", "Description", "GIVE", "HAVE")
        self.tree_stampa.column("#0", width=0, stretch=NO)
        self.tree_stampa.column("Name", anchor=W, width=230, stretch=NO)
        self.tree_stampa.column("Description", anchor=CENTER, width=235, stretch=NO)
        self.tree_stampa.column("GIVE", anchor=CENTER, width=150, stretch=NO)
        self.tree_stampa.column("HAVE", anchor=CENTER, width=150, stretch=NO)

        self.tree_stampa.heading("#0", text="", anchor=CENTER)
        self.tree_stampa.heading("Name", text="Name", anchor=W)
        self.tree_stampa.heading("Description", text="Description", anchor=CENTER)
        self.tree_stampa.heading("GIVE", text="GIVE", anchor=CENTER)
        self.tree_stampa.heading("HAVE", text="HAVE", anchor=CENTER)

        self.tree_stampa.tag_configure("oddrow", background="IndianRed3")
        self.tree_stampa.tag_configure("evenrow", background="PaleGreen1")

        try:
            for contatore, record in enumerate(data):
                if contatore == 0:
                    pass
                elif record[2] == "0":
                    self.tree_stampa.insert(parent="", index="end", iid=contatore, text="",
                                            values=(record[0], record[1], record[2], record[3]),
                                            tags=("evenrow",))
                else:
                    self.tree_stampa.insert(parent="", index="end", iid=contatore, text="",
                                            values=(record[0], record[1], record[2], record[3]),
                                            tags=("oddrow",))
        except IndexError:
            for contatore, record in enumerate([]):
                self.tree_stampa.insert(parent="", index="end", iid=contatore, text="",
                                            values=(record[0], record[1], record[2], record[3]),
                                            tags=("evenrow",))

        tree_stampa_scrolly = Scrollbar(self.stampa_container, orient="vertical",
                                        command=self.tree_stampa.yview)
        self.tree_stampa.configure(yscrollcommand=tree_stampa_scrolly.set)
        tree_stampa_scrolly.place(relx=0.988, rely=0, anchor="n", height=495)
        self.tree_stampa.place(width=780, height=495, relx=0, rely=0)

    def save_credit(self):
        """Esporta la colonna del credito"""

        try:
            with open("frills.dat", "r") as file:
                for line in file:
                    if line.split("=")[0] == "Separator":
                        separator = line.split("=")[1][:-1]
                    elif line.split("=")[0] == "Exstension":
                        exstension = line.split("=")[1]
                    else:
                        dire = line.split("=")[1]

            names = list()
            give_list = list()
            have_list = list()
            for item in self.tree_stampa.get_children():
                t_g = ""
                t_h = ""
                the_item = str(self.tree_stampa.item(item).values()).split(",")
                if "SIG" in str(the_item[2][3:len(the_item[2]) - 1]):
                    names.append(the_item[2][3:len(the_item[2]) - 1])
                    for i in the_item[4][1:len(the_item[4]) - 1]:
                        if i in ("\n", "'"):
                            pass
                        else:
                            t_g += i
                    give_list.append(t_g)

                    for i in the_item[5][1:len(the_item[5]) - 1]:
                        if i in ("\n", "'"):
                            pass
                        else:
                            t_h += i
                    have_list.append(t_h)
            if exstension == "csv":
                pass
            else:
                self.separa_dict = {"______": "_", "------": "-", "......": ".", "         ": " "}

                try:
                    file = asksaveasfile(initialfile='credit',
                                         initialdir=dire,
                                         title="Save Credit File",
                                         defaultextension=".txt")
                except UnboundLocalError:
                    file = asksaveasfile(initialfile='credit',
                                         initialdir="C:",
                                         title="Save Credit File",
                                         defaultextension=".txt")

                with open("frills.dat", "w") as f:
                    for i in range(3):
                        if i == 0:
                            f.write(f"Separator={separator}\n")
                        elif i == 1:
                            f.write(f"Exstension={exstension}\n")
                        elif i == 2:
                            f.write(f"Dir={file.name.split(f'bill_{names[0]}.txt')[0]}")

                t_give = list()
                t_have = list()
                for line in range(len(names)):

                    try:
                        t_give.append(float(give_list[line]))
                    except ValueError:
                        pass

                    try:
                        t_have.append(float(have_list[line]))
                    except ValueError:
                        pass

                    first_name = f"{names[line]}" + str(self.separa_dict[separator] * (25 - len(names[line])))

                    if give_list[line] == "":
                        sep_halign = str(self.separa_dict[separator]) * (40 - len(first_name) + len(have_list[line]))
                        have_halign = str(self.separa_dict[separator]) * (14 - len(have_list[line]))
                        file.write(f"{first_name}{sep_halign}{have_halign}{have_list[line]}\n")
                    elif have_list[line] == "0":
                        file.write(f"{first_name}{give_list[line]}\n")
                    elif have_list[line] == "0" and give_list[line] == "":
                        pass

                Give = "Give" + str(self.separa_dict[separator] * (25 - len("Give")))
                file.write("\n")
                file.write(f"{Give}{round(fsum(t_give), 2)}\n")
                Have = "Have" + str(self.separa_dict[separator] * (25 - len("Have")))
                sep_halign = str(self.separa_dict[separator]) * (40 - len(Have) + len(str(fsum(t_have))))
                have_halign = str(self.separa_dict[separator]) * (14 - len(str(fsum(t_have))))
                file.write(f"{Have}{sep_halign}{have_halign}{round(fsum(t_have), 2)}\n")

                Diff = "Difference" + str(self.separa_dict[separator] * (25 - len("Difference")))
                file.write("\n")
                file.write(f"{Diff}{round(fsum([fsum(t_have), -fsum(t_give)]), 2)}")


        except AttributeError:
            messagebox.showerror("Attention", "Something has happened contact the productor!")
        else:
            messagebox.showinfo("Export", "Export done!")

    def save_all(self):
        """Export all database in txt file"""

        try:

            with open("frills.dat", "r") as file:
                for line in file:
                    if line.split("=")[0] == "Separator":
                        separator = line.split("=")[1][:-1]
                    elif line.split("=")[0] == "Exstension":
                        exstension = line.split("=")[1]
                    else:
                        dire = line.split("=")[1]

            names = list()
            give_list = list()
            have_list = list()
            for item in self.tree_stampa.get_children():
                t_g = ""
                t_h = ""
                the_item = str(self.tree_stampa.item(item).values()).split(",")

                names.append(the_item[2][3:len(the_item[2]) - 1])
                for i in the_item[4][1:len(the_item[4]) - 1]:
                    if i in ("\n", "'"):
                        pass
                    else:
                        t_g += i
                give_list.append(t_g)

                for i in the_item[5][1:len(the_item[5]) - 1]:
                    if i in ("\n", "'"):
                        pass
                    else:
                        t_h += i
                have_list.append(t_h)

            self.separa_dict = {"______": "_", "------": "-", "......": ".", "         ": " "}

            try:
                file = asksaveasfile(initialfile='archive.txt',
                                     initialdir=dire,
                                     title="Save Archive File",
                                     defaultextension=".txt")
            except UnboundLocalError:
                file = asksaveasfile(initialfile='archive.txt',
                                     initialdir="C:",
                                     title="Save Archive File",
                                     defaultextension=".txt")

            with open("frills.dat", "w") as f:
                for i in range(3):
                    if i == 0:
                        f.write(f"Separator={separator}\n")
                    elif i == 1:
                        f.write(f"Exstension={exstension}\n")
                    elif i == 2:
                        f.write(f"Dir={file.name.split(f'bill_{names[0]}.txt')[0]}")

            t_give = list()
            t_have = list()
            for line in range(len(names)):

                try:
                    t_give.append(float(give_list[line]))
                except ValueError:
                    pass

                try:
                    t_have.append(float(have_list[line]))
                except ValueError:
                    pass

                first_name = f"{names[line]}" + str(self.separa_dict[separator] * (25 - len(names[line])))

                if give_list[line] == "":
                    sep_halign = str(self.separa_dict[separator]) * (40 - len(first_name) + len(have_list[line]))
                    have_halign = str(self.separa_dict[separator]) * (14 - len(have_list[line]))
                    file.write(f"{first_name}{sep_halign}{have_halign}{have_list[line]}\n")
                elif have_list[line] == "0":
                    file.write(f"{first_name}{give_list[line]}\n")
                elif have_list[line] == "0" and give_list[line] == "":
                    pass

            Give = "Give" + str(self.separa_dict[separator] * (25 - len("Give")))
            file.write("\n")
            file.write(f"{Give}{self.give}\n")
            Have = "Have" + str(self.separa_dict[separator] * (25 - len("Have")))
            sep_halign = str(self.separa_dict[separator]) * (40 - len(Have) + len(str(self.have)))
            have_halign = str(self.separa_dict[separator]) * (14 - len(str(self.have)))
            file.write(f"{Have}{sep_halign}{have_halign}{self.have}\n")

            Diff = "Difference" + str(self.separa_dict[separator] * (25 - len("Difference")))
            file.write("\n")
            file.write(f"{Diff}{round(fsum([fsum(t_have), -fsum(t_give)]), 2)}")

        except AttributeError:
            messagebox.showerror("Attention", "Something has happened contact the productor!")
        else:
            messagebox.showinfo("Export", "Export done!")


# PrintDay Screen


class PrintDay(Frame):

    def __init__(self, parent, controller):
        Frame.__init__(self, parent)
        self.style = ThemedStyle(self)
        self.style.theme_use(f"{styles}")

        self.controller = controller

        label = Label(self, text="Print Inputs Of The Day", font=("Spectral", 15), foreground="black",
                      relief="ridge")
        label.pack(side="top", fill="x", pady=10)

        button = Button(self, text="Back to main menu", command=lambda: controller.show_frame("MenuP"),
                        fg="black", bg="sandy brown", relief="raised",
                        activebackground="light gray", font=("Lucinda Console", 10))
        button.pack()

        self.on_start()

    def on_start(self):
        """Build Frame"""

        self.search_label = Label(self, text="Select a date", font=("Spectral", 10), foreground="black",
                                  relief="ridge")
        self.search_label.config(anchor=CENTER)
        self.search_label.place(height=40, width=160, relx=0.01, rely=0.08)

        self.container = LabelFrame(self)
        self.container.place(height=560, width=880, relx=0.01, rely=0.15)

        self.mesi_ordine = {"Gennaio": 1, "Febbraio": 2, "Marzo": 3, "Aprile": 4, "Maggio": 5, "Giugno": 6,
                            "Luglio": 7, "Agosto": 8, "Settembre": 9, "Ottobre": 10, "Novembre": 11, "Dicembre": 12}

        giorno, mese, anno = self.scandata()
        self.date_entry = DateEntry(self, selectmode="day", font="Lucinda 10", locale="it_IT", showweeknumbers=False,
                                    showothermonthdays=True,
                                    year=int(anno), month=self.mesi_ordine[mese], day=int(giorno),
                                    background="dark slate gray", date_pattern="DD/MM/YYYY",
                                    bordercolor="dark slate gray", selectbackground="SlateGray4",
                                    headersbackground="light gray", normalbackground="light gray", foreground='white',
                                    normalforeground='black', headersforeground='black', weekendbackground="IndianRed1")
        self.date_entry.place(relx=0.84, rely=0.11)
        self.date_entry.bind("<<DateEntrySelected>>", self.on_selected_day)

        export_day_button = Button(self, text="Export List", fg="black", bg="silver", relief="ridge",
                                   activebackground="wheat1", font=("Spectral", 15),
                                   command=self.save_day)
        export_day_button.place(height=30, width=170, relx=0.63, rely=0.09)

        refresh_button = Button(self, text="Update", fg="black", bg="wheat1", relief="ridge",
                                activebackground="silver", font=("Spectral", 12),
                                command=self.update_tree)
        refresh_button.place(height=30, width=100, relx=0.23, rely=0.09)

        self.diff_label = Label(self, relief="ridge")
        self.diff_label.config(anchor=CENTER)
        self.diff_label.place(height=35, width=150, relx=0.65, rely=0.94)

        self.diff_label_t = Label(self, text="Difference", font=("Spectral", 18), foreground="black", relief="ridge")
        self.diff_label_t.config(anchor=CENTER)
        self.diff_label_t.place(height=35, width=150, relx=0.46, rely=0.94)

    def task(self):
        """Support function for loading"""

        global global_data_in_out
        global dataset_file
        global_data_in_out = read_csv(f"{dataset_file}", engine="c")
        try:
            size_of_df = global_data_in_out.Name.unique().size
            dataframe_filtered_by_data = global_data_in_out[
                global_data_in_out['Date'] == f'{self.giorno}-{int(self.mese)}-{self.anno}']
            if dataframe_filtered_by_data.size != 0:
                self.raff_rows = array([0, 0, 0])
                for index, iname in enumerate(dataframe_filtered_by_data["Name"].unique()):
                    self.loading(iname)
                    self.load_label.config(text=f"Loading {round(100 * (index / 2) / size_of_df, 1)}%")
                    self.loading_window.update_idletasks()

                size_of_df = global_data_in_out.size
                start = dataframe_filtered_by_data.index[-1] + 1
                end = global_data_in_out.index[-1] + 1

                # Delete all value aftrer selected day
                for index in range(start, end):
                    global_data_in_out.drop(index, inplace=True)
                    self.load_label.config(text=f"Loading {round(100 * (index / 2) / size_of_df, 1)}%")
                    self.loading_window.update_idletasks()

                self.loading_window.destroy()
                self.stampa_tree_day(dataframe_filtered_by_data.sort_values(by=["Name"]))

            else:
                self.stampa_tree_day([])
                self.loading_window.destroy()
        except AttributeError:
            self.loading_window.destroy()
            messagebox.showerror("Attention", "Select a date first!")

    def loading(self, bill_name):
        """Loading function"""

        global global_data_in_out

        give = round(fsum(global_data_in_out[global_data_in_out["Name"] == bill_name][
                              global_data_in_out[global_data_in_out["Name"] == bill_name]["Amount"] < 0][
                              "Amount"]), 2)

        have = round(fsum(global_data_in_out[global_data_in_out["Name"] == bill_name][
                              global_data_in_out[global_data_in_out["Name"] == bill_name]["Amount"] > 0][
                              "Amount"]), 2)

        description = global_data_in_out[global_data_in_out["Name"] == bill_name][
            global_data_in_out[global_data_in_out["Name"] == bill_name]["Description"].astype(str) != "nan"][
            "Description"].astype(str).sum()

        if round(fsum([have, give]), 2) > 0:
            self.raff_rows = vstack(
                [self.raff_rows,
                 [bill_name, str(description).lower(), round(fsum([give, have]), 2)]])
        elif round(fsum([have, give]), 2) < 0:
            self.raff_rows = vstack(
                [self.raff_rows,
                 [bill_name, str(description).lower(), round(fsum([give, have]), 2)]])

    def update_tree(self):
        """Update tree base on selected day"""

        self.loading_window = SecondWindow(self, "Wait", "300x50", False, (0, 0))
        stile_load = ThemedStyle(self.loading_window)
        stile_load.theme_use(f"{styles}")
        self.loading_window.title("Wait")

        self.load_label = Label(self.loading_window, text="Reading File", font=("Spectral", 15),
                                foreground="black", relief="ridge")
        self.load_label.pack(side="top", fill="x", pady=10)

        self.loading_window.after(1, Thread(target=self.task).start())
        self.loading_window.mainloop()

    def stampa_tree_day(self, raff_rows):
        """Add Tree in tab"""

        global global_data_in_out

        tree_style = ttk.Style()
        tree_style.configure("mystyle.Treeview", highlightthickness=0, bd=0,
                             font=('Lucinda Console', 16))  # Modify the font of the body
        tree_style.configure("mystyle.Treeview.Heading",
                             font=('Spectral', 18, 'italic'), width=1, pady=20)  # Modify the font of the headings
        tree_style.layout("mystyle.Treeview",
                          [('mystyle.Treeview.treearea', {'sticky': 'nswe'})])  # Remove the borders
        tree_style.configure('Treeview', rowheight=40)

        self.tree_day = ttk.Treeview(self.container, style="mystyle.Treeview", selectmode="extended")

        self.tree_day["columns"] = ("Name", "Description", "GIVE", "HAVE")
        self.tree_day.column("#0", width=0, stretch=NO)
        self.tree_day.column("Name", anchor=W, width=250, stretch=NO)
        self.tree_day.column("Description", anchor=CENTER, width=240, stretch=NO)
        self.tree_day.column("GIVE", anchor=CENTER, width=180, stretch=NO)
        self.tree_day.column("HAVE", anchor=CENTER, width=180, stretch=NO)

        self.tree_day.heading("#0", text="", anchor=CENTER)
        self.tree_day.heading("Name", text="Name", anchor=W)
        self.tree_day.heading("Description", text="Description", anchor=CENTER)
        self.tree_day.heading("GIVE", text="GIVE", anchor=CENTER)
        self.tree_day.heading("HAVE", text="HAVE", anchor=CENTER)

        self.tree_day.tag_configure("oddrow", background="IndianRed3")
        self.tree_day.tag_configure("evenrow", background="PaleGreen1")

        try:
            for index, row in raff_rows.iterrows():
                if f"{row[1]}" == "nan":
                    row1 = ""
                else:
                    row1 = row[1]
                if row[2] < 0:
                    self.tree_day.insert(parent="", index="end", iid=index, text="",
                                         values=[row[0], row1, row[2] * -1, 0],
                                         tags=("oddrow",))
                else:
                    self.tree_day.insert(parent="", index="end", iid=index, text="", values=[row[0], row1, 0, row[2]],
                                         tags=("evenrow",))

            self.give = fsum(global_data_in_out[global_data_in_out["Amount"] < 0]["Amount"])
            self.have = fsum(global_data_in_out[global_data_in_out["Amount"] > 0]["Amount"])
            self.difference = round(fsum([self.have, self.give]), 2)
            self.diff_label.config(text=str(self.difference),
                                   font=("Spectral", 15),
                                   foreground="black")
            self.search_label.config(text="Search Completed", font=("Spectral", 10))
        except AttributeError:
            self.diff_label.config(text="")
            self.search_label.config(text="There are no inputs\n"
                                          "for selected day", font=("Spectral", 10))

        tree_day_scrolly = Scrollbar(self.container, orient="vertical",
                                     command=self.tree_day.yview)
        self.tree_day.configure(yscrollcommand=tree_day_scrolly.set)
        tree_day_scrolly.place(relx=0.988, rely=0, anchor="n", height=530)
        self.tree_day.place(width=888, height=530, relx=0, rely=0)
        self.tree_day.bind("<<TreeviewSelect>>", lambda x: self.edit_selected_person(x))

    def on_selected_day(self, day):
        """Get selected day on calendar"""

        self.giorno = str(self.date_entry.get_date()).split("-")[2]
        self.mese = str(self.date_entry.get_date()).split("-")[1]
        self.anno = str(self.date_entry.get_date()).split("-")[0]

        self.num_mesi = {"01": "Gennaio", "02": "Febbraio", "03": "Marzo", "04": "Aprile", "05": "Maggio",
                         "06": "Giugno", "07": "Luglio", "08": "Agosto", "09": "Settembre", "10": "Ottobre",
                         "11": "Novembre", "12": "Dicembre"}

        self.selected_date = f"{self.giorno}/{self.mese}/{self.anno[2:]}"
        self.raff_selected_date = f"{self.giorno}_{self.mese}_{self.anno[2:]}"
        self.update_tree()

    def edit_selected_person(self, sel):
        """Go into Edit mode for selected row"""

        ask_for_edit = messagebox.askyesno('Edit', 'Edit selected Bill?')

        if ask_for_edit:
            global name_day
            global desc_day
            global give_day
            global have_day
            global date_day

            self.num_mesi = {"01": "Gennaio", "02": "Febbraio", "03": "Marzo", "04": "Aprile", "05": "Maggio",
                             "06": "Giugno", "07": "Luglio", "08": "Agosto", "09": "Settembre", "10": "Ottobre",
                             "11": "Novembre", "12": "Dicembre"}

            selected = self.tree_day.focus()
            selected_row = self.tree_day.item(selected, "values")

            name_day = selected_row[0]
            desc_day = selected_row[1]
            give_day = selected_row[2]
            have_day = selected_row[3]

            self.controller.show_frame("EditValue")

    def scandata(self):
        """Get current Date"""

        self.Mesi = {"January": "Gennaio", "February": "Febbraio", "March": "Marzo", "April": "Aprile", "May": "Maggio",
                     "June": "Giugno", "July": "Luglio", "August": "Agosto", "September": "Settembre",
                     "October": "Ottobre",
                     "November": "Novembre", "December": "Dicembre"}
        oggi = date.today()
        mese = oggi.strftime("%B")
        dat = (oggi.strftime("%d"), self.Mesi[mese], oggi.strftime("%Y"))

        return dat

    def save_day(self):
        """Export all database in txt file"""

        try:
            with open("frills.dat", "r") as file:
                for line in file:
                    if line.split("=")[0] == "Separator":
                        separator = line.split("=")[1][:-1]
                    elif line.split("=")[0] == "Exstension":
                        exstension = line.split("=")[1]
                    else:
                        dire = line.split("=")[1]

            names = list()
            give_list = list()
            have_list = list()
            desc = list()
            for item in self.tree_day.get_children():
                t_g = ""
                t_h = ""
                the_item = str(self.tree_day.item(item).values()).split(",")

                names.append(the_item[2][3:len(the_item[2]) - 1])
                for i in the_item[4][1:len(the_item[4]) - 1]:
                    if i in ("\n", "'"):
                        pass
                    else:
                        t_g += i
                give_list.append(t_g)

                for i in the_item[5][1:len(the_item[5]) - 1]:
                    if i in ("\n", "'"):
                        pass
                    else:
                        t_h += i
                have_list.append(t_h)

                if the_item[3][2:len(the_item[3]) - 1].isalpha():
                    desc.append(the_item[3][2:len(the_item[3]) - 1])
                else:
                    desc.append("")

            self.separa_dict = {"______": "_", "------": "-", "......": ".", "         ": " "}

            try:
                file = asksaveasfile(initialfile=f'print_of_day_{self.raff_selected_date}.txt', initialdir=dire,
                                     title="Save File",
                                     defaultextension=".txt")
            except UnboundLocalError:
                file = asksaveasfile(initialfile=f'print_of_day_{self.raff_selected_date}.txt', initialdir="C:",
                                     title="Save File",
                                     defaultextension=".txt")

            with open("frills.dat", "w") as f:
                for i in range(3):
                    if i == 0:
                        f.write(f"Separator={separator}\n")
                    elif i == 1:
                        f.write(f"Exstension={exstension}")
                    elif i == 2:
                        f.write(f"Dir={file.name.split(f'bill_{names[0]}.txt')[0]}")

            for line in range(len(names)):

                first_name = f"{names[line]}" + str(self.separa_dict[separator] * (25 - len(names[line])))

                if give_list[line] == "":
                    if desc[line] == "":
                        sep_halign = str(self.separa_dict[separator]) * (40 - len(first_name) + len(have_list[line]))
                        avere_halign = str(self.separa_dict[separator]) * (14 - len(have_list[line]))
                        file.write(f"{first_name}{sep_halign}{avere_halign}{have_list[line]}\n")
                    else:
                        sep_halign = str(self.separa_dict[separator]) * (40 - len(first_name) + len(have_list[line]))
                        avere_halign = str(self.separa_dict[separator]) * (14 - len(have_list[line]))
                        file.write(f"{first_name}{sep_halign}{avere_halign}{have_list[line]}---({desc[line]})\n")
                elif have_list[line] == "":
                    if desc[line] == "":
                        file.write(f"{first_name}{give_list[line]}\n")
                    else:
                        file.write(f"{first_name}{give_list[line]}---({desc[line]})\n")

            Give = "Give" + str(self.separa_dict[separator] * (25 - len("Give")))
            file.write("\n")
            file.write(f"{Give}{self.give}\n")
            Have = "Have" + str(self.separa_dict[separator] * (25 - len("Have")))
            sep_halign = str(self.separa_dict[separator]) * (40 - len(Have) + len(str(self.have)))
            avere_halign = str(self.separa_dict[separator]) * (14 - len(str(self.have)))
            file.write(f"{Have}{sep_halign}{avere_halign}{self.have}\n")

            Diff = "Difference" + str(self.separa_dict[separator] * (25 - len("Difference")))
            file.write("\n")
            file.write(f"{Diff}{self.difference}")

        except AttributeError:
            messagebox.showerror("Attention", "Something has happened contact the productor!")
        else:
            messagebox.showinfo("Export", "Export done!")


# PrintSingle Screen


class PrintSingle(Frame):

    def __init__(self, parent, controller):
        Frame.__init__(self, parent)
        self.style = ThemedStyle(self)
        self.style.theme_use(f"{styles}")

        self.controller = controller

        label = Label(self, text="Print Single Bill Name", font=("Spectral", 15), foreground="black", relief="ridge")
        label.pack(side="top", fill="x", pady=10)

        button = Button(self, text="Back to main menu", command=lambda: controller.show_frame("MenuP"),
                        fg="black", bg="sandy brown", relief="raised",
                        activebackground="light gray", font=("Lucinda Console", 10))
        button.pack()

        self.on_start()

        self.bind("<<ShowFrame>>", self.on_show_frame)

    def on_show_frame(self, event):
        """ON  Raise"""
        global starting

        if not starting:
            self.name_entry.focus()

    def on_start(self):
        """Build Frame"""

        search_button = Button(self, text="Search", command=self.enter_search,
                               fg="black", bg="Silver", relief="raised",
                               activebackground="light gray", font=("Times", 15))
        search_button.place(relx=0.6, rely=0.15, width=110, height=30)

        export_single_button = Button(self, text="Export List", fg="black", bg="silver", relief="ridge",
                                      activebackground="wheat1", font=("Spectral", 15),
                                      command=self.save_single)
        export_single_button.place(height=30, width=170, relx=0.74, rely=0.1)

        graph_button = Button(self, text="Show Plot", fg="black", bg="wheat1", relief="ridge",
                              activebackground="silver", font=("Spectral", 15),
                              command=lambda: Thread(target=self.plot()).start())
        graph_button.place(height=30, width=170, relx=0.74, rely=0.15)

        date_button = Button(self, text="Print by Date", fg="black", bg="#a8bed0", relief="ridge",
                             activebackground="silver", font=("Spectral", 15),
                             command=lambda: Thread(target=self.select_date_for_single()).start())
        date_button.place(height=30, width=170, relx=0.74, rely=0.2)

        self.search_label = Label(self, text="Write in the field",
                                  font=("Spectral", 15),
                                  foreground="black", relief="ridge")
        self.search_label.config(anchor=CENTER)
        self.search_label.place(height=60, width=200, relx=0.08, rely=0.15)

        self.stampa_single_container = LabelFrame(self)
        self.stampa_single_container.place(height=410, width=880, relx=0.01, rely=0.28)

        self.name_entry = Entry(self, font=("Lucinda Console", 15))
        self.name_entry.place(relx=0.41, rely=0.15, width=160, height=30)
        self.name_entry.bind('<KeyRelease>', self.find_name)

        # Create a Listbox widget to display the list of items
        self.suggestion = Listbox(self, font=("Lucinda Console", 15), relief="flat")
        self.suggestion.place(relx=0.41, rely=0.19, width=160, height=50)
        self.suggestion.bind("<Double-Button-1>", self.select_suggestion)
        self.suggestion.bind("<Return>", self.select_suggestion)

        # Create a list of all the menu items
        global bill_names

        # Add values to combobox
        self.update_suggestion_list(bill_names)

        self.mesi_ordine = {"Gennaio": 1, "Febbraio": 2, "Marzo": 3, "Aprile": 4, "Maggio": 5, "Giugno": 6,
                            "Luglio": 7, "Agosto": 8, "Settembre": 9, "Ottobre": 10, "Novembre": 11, "Dicembre": 12}

        self.diff_label = Label(self, font=("Spectral", 15), foreground="black", relief="ridge")
        self.diff_label.config(anchor=CENTER)
        self.diff_label.place(height=35, width=150, relx=0.562, rely=0.908)

        self.diff_label_t = Label(self, text="Differenza", font=("Spectral", 18), foreground="black", relief="ridge")
        self.diff_label_t.config(anchor=CENTER)
        self.diff_label_t.place(height=35, width=150, relx=0.392, rely=0.908)

        self.stampa_single_tree([])

        self.today_from = True
        self.today_to = True

    def suggestion_key(self, key):
        """Keyboard suggestion base on value list"""
        global bill_names

        typing = self.name_entry.get()
        if typing == '':
            data = bill_names
        else:
            data = []
            for item in bill_names:
                if typing.lower() in item.lower():
                    data.append(item)

        self.update_suggestion_list(data)

    def update_suggestion_list(self, data):
        """Keyboard suggestion base on previous function"""

        # Clear the Combobox
        self.suggestion.delete(0, END)
        # Add values to the combobox
        for value in data:
            self.suggestion.insert(END, value)

    def select_suggestion(self, key):
        """Enter the suggestion on Entry box"""

        self.name_entry.delete(0, "end")
        for i in self.suggestion.curselection():
            if " " in self.suggestion.get(i):
                self.name_entry.insert(0, self.suggestion.get(i))
            else:
                self.name_entry.insert(0, self.suggestion.get(i))

    def update_single_tree(self):
        """Update tree for new added values"""

        global global_data_in_out
        global dataset_file
        global_data_in_out = read_csv(f"{dataset_file}", engine="c")

        self.stampa_single_tree(global_data_in_out[global_data_in_out["Name"] == self.name_entry.get()])

        bill_name = self.name_entry.get()
        if global_data_in_out[global_data_in_out["Name"] == self.name_entry.get()].size > 0:
            self.search_label.config(text="Search Completed", font=("Spectral", 15))
            give = round(fsum(global_data_in_out[global_data_in_out["Name"] == bill_name][
                                  global_data_in_out[global_data_in_out["Name"] == bill_name]["Amount"] < 0][
                                  "Amount"]), 2)
            have = round(fsum(global_data_in_out[global_data_in_out["Name"] == bill_name][
                                  global_data_in_out[global_data_in_out["Name"] == bill_name]["Amount"] > 0][
                                  "Amount"]), 2)
            self.difference = round(fsum([have, give]), 2)
            if self.difference > 0:
                self.diff_label_t.config(text="Balance Have")
                self.diff_label.config(text=str(self.difference))
            elif self.difference < 0:
                self.diff_label_t.config(text="Balance Give")
                self.diff_label.config(text=str(self.difference * -1))
            elif self.difference == 0.0:
                self.diff_label_t.config(text="Balance Zero")
                self.diff_label.config(text=f"{0}")
        else:
            self.search_label.config(text="There are no bills\n"
                                          "with this name", font=("Spectral", 15))

    def stampa_single_tree(self, raff_rows):
        """Add Tree in tab"""

        tree_style = ttk.Style()
        tree_style.configure("mystyle.Treeview", highlightthickness=0, bd=0,
                             font=('Lucinda Console', 16))  # Modify the font of the body
        tree_style.configure("mystyle.Treeview.Heading",
                             font=('Spectral', 18, 'italic'), width=1, pady=20)  # Modify the font of the headings
        tree_style.layout("mystyle.Treeview",
                          [('mystyle.Treeview.treearea', {'sticky': 'nswe'})])  # Remove the borders
        tree_style.configure('Treeview', rowheight=40)

        self.tree_stampa = ttk.Treeview(self.stampa_single_container, style="mystyle.Treeview")

        self.tree_stampa["columns"] = ("Name", "Description", "GIVE", "HAVE", "Date")
        self.tree_stampa.column("#0", width=0, stretch=NO)
        self.tree_stampa.column("Name", anchor=W, width=200, stretch=NO)
        self.tree_stampa.column("Description", anchor=CENTER, width=210, stretch=NO)
        self.tree_stampa.column("GIVE", anchor=CENTER, width=160, stretch=NO)
        self.tree_stampa.column("HAVE", anchor=CENTER, width=160, stretch=NO)
        self.tree_stampa.column("Date", anchor=CENTER, width=125, stretch=NO)

        self.tree_stampa.heading("#0", text="", anchor=CENTER)
        self.tree_stampa.heading("Name", text="Name", anchor=W)
        self.tree_stampa.heading("Description", text="Description", anchor=CENTER)
        self.tree_stampa.heading("GIVE", text="GIVE", anchor=CENTER)
        self.tree_stampa.heading("HAVE", text="HAVE", anchor=CENTER)
        self.tree_stampa.heading("Date", text="Date", anchor=CENTER)

        self.tree_stampa.tag_configure("oddrow", background="white")
        self.tree_stampa.tag_configure("evenrow", background="SlateGray2")

        try:
            for index, row in raff_rows.iterrows():
                if f"{row[1]}" == "nan":
                    row1 = ""
                else:
                    row1 = row[1]
                if row[2] < 0:
                    self.tree_stampa.insert(parent="", index="end", iid=index, text="",
                                            values=[row[0], row1, row[2] * -1, 0, row[3]],
                                            tags=("evenrow",))
                else:
                    self.tree_stampa.insert(parent="", index="end", iid=index, text="",
                                            values=[row[0], row1, 0, row[2], row[3]],
                                            tags=("oddrow",))
        except AttributeError:
            if not starting:
                self.diff_label.config(text="")
                self.search_label.config(text="There are no bills\n"
                                              "with this name", font=("Spectral", 10))

        tree_stampa_scrolly = Scrollbar(self.stampa_single_container, orient="vertical",
                                        command=self.tree_stampa.yview)
        self.tree_stampa.configure(yscrollcommand=tree_stampa_scrolly.set)
        tree_stampa_scrolly.place(relx=0.988, rely=0, anchor="n", height=400)
        self.tree_stampa.place(width=875, height=400, relx=0, rely=0)

    def enter_search(self):
        """Look for person bill via Return button"""

        self.diff_label.config(text="")
        for row in self.tree_stampa.get_children():
            self.tree_stampa.delete(row)
        self.update_single_tree()

    def plot(self, start=-2, end=0):

        global global_data_in_out
        data_to_plot = global_data_in_out[global_data_in_out["Name"] == self.name_entry.get()]

        use("TkAgg")
        set_style('darkgrid')

        palette = ["green" if value > 0 else "red" for value in data_to_plot["Amount"] > 0]

        # Histo label
        fig_graph = figure("Plot Bar", figsize=(11, 8), dpi=100)
        ax = fig_graph.add_subplot(111)
        data_to_plot.Amount.abs().plot(kind="bar", x=global_data_in_out.Date, ax=ax, legend=False, color=palette)
        ax.set_title(f'Sum Up {self.name_entry.get()}')
        ax.get_xaxis().set_visible(False)
        ax.set(xlabel="Days (DD/M/AA)", ylabel='Amount (€)')
        ax.set_ylim(top=int(max(data_to_plot.Amount)) + 20, bottom=0)
        # if end == 0:
        #     end = len(Importo)
        # if start == -1:
        #     start = -1
        # ax.set_xlim(left=start, right=end)
        setp(ax.get_xticklabels(), rotation=360, ha='right')

        count = 0
        data_to_plot.reset_index(inplace=True)
        for bar, frequency in zip(ax.patches, data_to_plot.Amount.abs()):
            text_x = bar.get_x() + bar.get_width() / 2.0
            text_y = bar.get_height()
            if f"{data_to_plot.Description[count]}" == "nan":
                desc = ""
            else:
                desc = data_to_plot.Description[count]
            ax.text(text_x, text_y, f"{data_to_plot.Date[count]}\n{desc}\n{frequency:.2f}", fontsize=8, ha='center',
                    va='bottom')
            count += 1

        # sx_starter = fig_graph.add_axes([0.12, .05, 0.3, 0.02])
        # dx_ender = fig_graph.add_axes([0.12, .03, 0.3, .02])
        # sx_slider = Slider(sx_starter, r'Sinistro', valmin=1, valmax=len(Importo), valinit=1, valstep=1)
        # sx_slider.label.set_size(10)
        # dx_slider = Slider(dx_ender, r'Destro', valmin=1, valmax=len(Importo), valinit=12, valstep=1)
        # dx_slider.label.set_size(10)

        show(block=False)

    def find_name(self, b):
        """Auto complete name typed"""

        self.suggestion_key(b)

        global global_data_in_out

        name_in_the_box = str(self.name_entry.get()).upper()

        if global_data_in_out[global_data_in_out["Name"] == str('SIG ' + name_in_the_box)].size > 0:
            self.name_entry.delete(0, END)
            self.name_entry.insert(0, str('SIG ' + name_in_the_box))
        else:
            self.name_entry.delete(0, END)
            self.name_entry.insert(0, name_in_the_box)

    def save_single(self):
        """Export all database in txt file"""

        try:
            with open("frills.dat", "r") as file:
                for line in file:
                    if line.split("=")[0] == "Separator":
                        separator = line.split("=")[1][:-1]
                    elif line.split("=")[0] == "Extension":
                        extension = line.split("=")[1]
                    else:
                        dire = line.split("=")[1]

            names = list()
            desc = list()
            give_list = list()
            have_list = list()

            for item in self.tree_stampa.get_children():
                t_g = ""
                t_h = ""
                the_item = str(self.tree_stampa.item(item).values()).split(",")

                names.append(f"{the_item[2][3:len(the_item[2]) - 1]}")
                desc.append(the_item[3].replace("'", ""))
                for i in the_item[4][1:len(the_item[4]) - 1]:
                    if i in ("\n", "'"):
                        pass
                    else:
                        t_g += i
                give_list.append(t_g)

                for i in the_item[5][1:len(the_item[5]) - 1]:
                    if i in ("\n", "'"):
                        pass
                    else:
                        t_h += i
                have_list.append(t_h)

            self.separa_dict = {"______": "_", "------": "-", "......": ".", "         ": " "}

            try:
                file = asksaveasfile(initialfile=f'bill_{names[0]}.txt', initialdir=dire,
                                     title=f"Save File of {names[0]}",
                                     defaultextension=".txt")
            except UnboundLocalError:
                file = asksaveasfile(initialfile=f'bill_{names[0]}.txt', initialdir="C:",
                                     title=f"Save File of {names[0]}",
                                     defaultextension=".txt")

            with open("frills.dat", "w") as f:
                for i in range(3):
                    if i == 0:
                        f.write(f"Separator={separator}\n")
                    elif i == 1:
                        f.write(f"Extension={extension}\n")
                    elif i == 2:
                        f.write(f"Dir={file.name.split(f'bill_{names[0]}.txt')[0]}")

            file.write(f"Bill {names[0]}\n")

            global global_data_in_out
            data_to_plot = global_data_in_out[global_data_in_out["Name"] == self.name_entry.get()]
            data_to_plot.reset_index(inplace=True)
            for line in range(len(names)):
                first_name = f"{data_to_plot.Date[line]}" + str(
                    self.separa_dict[separator] * (25 - len(names[line])))

                if give_list[line] == "":
                    sep_halign = str(self.separa_dict[separator]) * (40 - len(first_name) + len(have_list[line]))
                    avere_halign = str(self.separa_dict[separator]) * (14 - len(have_list[line]))
                    if desc[line] != " ":
                        file.write(f"{first_name}{sep_halign}{avere_halign}{have_list[line]}_____{desc[line]}\n")
                    else:
                        file.write(f"{first_name}{sep_halign}{avere_halign}{have_list[line]}\n")
                elif have_list[line] == "":
                    if desc[line] != " ":
                        file.write(f"{first_name}{give_list[line]}_____{desc[line]}\n")
                    else:
                        file.write(f"{first_name}{give_list[line]}\n")

            Diff = "Total" + str(self.separa_dict[separator] * (25 - len("Total")))
            file.write("\n")
            if -0.001 <= self.difference <= 0.001:
                file.write(f"{Diff}{self.difference} (Zero)")
            elif self.difference < 0:
                file.write(f"{Diff}{self.difference * -1} (Credit)")
            elif self.difference > 0:
                file.write(f"{Diff}{self.difference} (Debit)")

        except AttributeError:
            messagebox.showerror("Attention", "Something has happened contact the productor!")
        else:
            messagebox.showinfo("Export", "Export done!")

    def select_date_for_single(self):
        """Open selected bill from-to a given date"""

        global global_data_in_out

        if self.name_entry.get() == '':
            messagebox.showerror("Attention", "Fill the field")

        elif global_data_in_out.Name[global_data_in_out.Name == f"{self.name_entry.get()}"].size != 0:
            self.date_window = SecondWindow(self, "Wait", "300x200", False, (0, 0))
            stile_load = ThemedStyle(self.date_window)
            stile_load.theme_use(f"{styles}")
            self.date_window.title("Select Time Interval")

            self.date_label = Label(self.date_window, text="Select Time Interval", font=("Spectral", 15),
                                    foreground="black", relief="ridge")
            self.date_label.pack(side="top", fill="x", pady=10)

            self.from_label = Label(self.date_window, text="from", font=("Spectral", 15),
                                    foreground="black", relief="flat")
            self.from_label.place(relx=0.4, rely=0.25)

            giorno, mese, anno = scandata()

            self.mesi_ordine = {"Gennaio": 1, "Febbraio": 2, "Marzo": 3, "Aprile": 4, "Maggio": 5, "Giugno": 6,
                                "Luglio": 7, "Agosto": 8, "Settembre": 9, "Ottobre": 10, "Novembre": 11, "Dicembre": 12}

            self.to_label = Label(self.date_window, text="To", font=("Spectral", 15),
                                  foreground="black", relief="flat")
            self.to_label.place(relx=0.41, rely=0.5)

            self.date_entry_from = DateEntry(self.date_window, selectmode="day", font="Lucinda 10", locale="it_IT",
                                             showweeknumbers=False,
                                             showothermonthdays=True,
                                             year=int(anno), month=self.mesi_ordine[mese], day=int(giorno),
                                             background="dark slate gray", date_pattern="DD/MM/YYYY",
                                             bordercolor="dark slate gray", selectbackground="SlateGray4",
                                             headersbackground="light gray", normalbackground="light gray",
                                             foreground='white',
                                             normalforeground='black', headersforeground='black',
                                             weekendbackground="IndianRed1")
            self.date_entry_from.place(relx=0.3, rely=0.4)
            self.date_entry_from.bind("<<DateEntrySelected>>", self.select_data_from)

            self.date_entry_to = DateEntry(self.date_window, selectmode="day", font="Lucinda 10", locale="it_IT",
                                           showweeknumbers=False,
                                           showothermonthdays=True,
                                           year=int(anno), month=self.mesi_ordine[mese], day=int(giorno),
                                           background="dark slate gray", date_pattern="DD/MM/YYYY",
                                           bordercolor="dark slate gray", selectbackground="SlateGray4",
                                           headersbackground="light gray", normalbackground="light gray",
                                           foreground='white',
                                           normalforeground='black', headersforeground='black',
                                           weekendbackground="IndianRed1")
            self.date_entry_to.place(relx=0.3, rely=0.65)
            self.date_entry_to.bind("<<DateEntrySelected>>", self.select_data_to)

            print_button = Button(self.date_window, text="Print", fg="black", bg="#a8bed0", relief="ridge",
                                  activebackground="silver", font=("Spectral", 15),
                                  command=lambda: Thread(target=self.print_single_base_on_date()).start())
            print_button.place(height=25, width=80, relx=0.33, rely=0.81)

            self.date_window.mainloop()

        else:
            messagebox.showerror("Attention", "There are no bills with this name")

    def select_data_from(self, date_from):
        self.giorno = str(self.date_entry_from.get_date()).split("-")[2]
        self.mese = str(self.date_entry_from.get_date()).split("-")[1]
        self.anno = str(self.date_entry_from.get_date()).split("-")[0]
        self.num_mesi = {"01": "Gennaio", "02": "Febbraio", "03": "Marzo", "04": "Aprile", "05": "Maggio",
                         "06": "Giugno", "07": "Luglio", "08": "Agosto", "09": "Settembre", "10": "Ottobre",
                         "11": "Novembre", "12": "Dicembre"}
        self.date_from = (f'{self.giorno}', f'{self.num_mesi[self.mese]}', f'{self.anno}')
        self.today_from = False

    def select_data_to(self, date_to):
        self.giorno = str(self.date_entry_to.get_date()).split("-")[2]
        self.mese = str(self.date_entry_to.get_date()).split("-")[1]
        self.anno = str(self.date_entry_to.get_date()).split("-")[0]
        self.num_mesi = {"01": "Gennaio", "02": "Febbraio", "03": "Marzo", "04": "Aprile", "05": "Maggio",
                         "06": "Giugno", "07": "Luglio", "08": "Agosto", "09": "Settembre", "10": "Ottobre",
                         "11": "Novembre", "12": "Dicembre"}
        self.date_to = (f'{self.giorno}', f'{self.num_mesi[self.mese]}', f'{self.anno}')
        self.today_to = False

    def single_from_date_window(self):

        date_on_the_tree = list()
        desc = list()
        give_list = list()
        have_list = list()

        for item in self.tree_single.get_children():
            t_g = ""
            t_h = ""
            the_item = str(self.tree_single.item(item).values()).split(",")

            desc.append(the_item[2][3: len(the_item[2]) - 1])
            for i in the_item[3][1:len(the_item[3]) - 1]:
                if i in ("\n", "'"):
                    pass
                else:
                    t_g += i
            give_list.append(t_g)

            for i in the_item[4][1:len(the_item[4]) - 1]:
                if i in ("\n", "'"):
                    pass
                else:
                    t_h += i

            have_list.append(t_h)

            date_on_the_tree.append(the_item[5][2:len(the_item[5]) - 2])

        file = asksaveasfile(initialfile=f'bill_{self.name_entry.get()}.txt', initialdir="C:",
                             title=f"Save file of {self.name_entry.get()}",
                             defaultextension=".txt")

        file.write(f"Bill {self.name_entry.get()}\n")

        for line in range(len(desc)):
            first_name = f"{date_on_the_tree[line]}" + str(
                "_" * (25 - len(desc[line])))

            if give_list[line] == "":
                sep_halign = "_" * (40 - len(first_name) + len(have_list[line]))
                avere_halign = "_" * (14 - len(have_list[line]))
                if desc[line] == "":
                    file.write(f"{first_name}{sep_halign}{avere_halign}{have_list[line]}\n")
                else:
                    file.write(f"{first_name}{sep_halign}{avere_halign}{have_list[line]}_____{desc[line]}\n")
            elif have_list[line] == "":
                if desc[line] == "":
                    file.write(f"{first_name}{give_list[line]}\n")
                else:
                    file.write(f"{first_name}{give_list[line]}_____{desc[line]}\n")

        Diff = "Total" + "_" * (25 - len("Total"))
        file.write("\n")

        Dare = [float(val) for val in give_list if val != ""]
        Avere = [float(val) for val in have_list if val != ""]

        if -0.001 <= fsum([fsum(Dare), -fsum(Avere)]) <= 0.001:
            file.write(f"{Diff}{fsum([round(fsum([fsum(Dare), -fsum(Avere)]), 2)])} (Zero)")
        elif fsum([fsum(Dare), -fsum(Avere)]) < 0:
            file.write(f"{Diff}{round(fsum([fsum(Dare), -fsum(Avere)]), 2) * -1} (Credit)")
        elif fsum([fsum(Dare), -fsum(Avere)]) > 0:
            file.write(f"{Diff}{fsum([round(fsum([fsum(Dare), -fsum(Avere)]), 2)])} (Debit)")

        messagebox.showinfo("Export", "Export done!")

    def print_single_base_on_date(self):

        if self.today_from:
            self.date_from = scandata()
        if self.today_to:
            self.date_to = scandata()

        self.date_window.destroy()

        self.print_single = SecondWindow(self, "Prin Bill By Date", "675x427", True, (0, 0))
        stile_load = ThemedStyle(self.print_single)
        stile_load.theme_use(f"{styles}")
        self.print_single.title(
            f"{self.name_entry.get()} From {self.date_from[0]}/{self.date_from[1]}/{self.date_from[2]} To {self.date_to[0]}/{self.date_to[1]}/{self.date_to[2]}")

        self.print_label = Label(self.print_single,
                                 text=f"{self.name_entry.get()} From {self.date_from[0]}/{self.date_from[1]}/{self.date_from[2]} To {self.date_to[0]}/{self.date_to[1]}/{self.date_to[2]}",
                                 font=("Spectral", 15),
                                 foreground="black", relief="ridge")
        self.print_label.pack(side="top", fill="x", pady=10)

        tree_style = ttk.Style()
        tree_style.configure("mystyle.Treeview", highlightthickness=0, bd=0,
                             font=('Lucinda Console', 16))  # Modify the font of the body
        tree_style.configure("mystyle.Treeview.Heading",
                             font=('Spectral', 18, 'italic'), width=1, pady=20)  # Modify the font of the headings
        tree_style.layout("mystyle.Treeview",
                          [('mystyle.Treeview.treearea', {'sticky': 'nswe'})])  # Remove the borders
        tree_style.configure('Treeview', rowheight=40)

        self.tree_single = ttk.Treeview(self.print_single, style="mystyle.Treeview")

        self.tree_single["columns"] = ("Description", "GIVE", "HAVE", "Date")
        self.tree_single.column("#0", width=0, stretch=NO)
        self.tree_single.column("Description", anchor=CENTER, width=205, stretch=NO)
        self.tree_single.column("GIVE", anchor=CENTER, width=155, stretch=NO)
        self.tree_single.column("HAVE", anchor=CENTER, width=155, stretch=NO)
        self.tree_single.column("Date", anchor=CENTER, width=140, stretch=NO)

        self.tree_single.heading("#0", text="", anchor=CENTER)
        self.tree_single.heading("Description", text="Description", anchor=CENTER)
        self.tree_single.heading("GIVE", text="GIVE", anchor=CENTER)
        self.tree_single.heading("HAVE", text="HAVE", anchor=CENTER)
        self.tree_single.heading("Date", text="Date", anchor=CENTER)

        self.tree_single.tag_configure("oddrow", background="white")
        self.tree_single.tag_configure("evenrow", background="SlateGray2")

        global global_data_in_out

        start_day = f'{self.date_from[2]}-{self.mesi_ordine[self.date_from[1]]}-{self.date_from[0]}'
        end_day = f'{self.date_to[2]}-{self.mesi_ordine[self.date_to[1]]}-{self.date_to[0]}'
        raff_rows = global_data_in_out[global_data_in_out["Name"] == self.name_entry.get()]

        for index, value in zip(raff_rows.index, raff_rows['Date']):
            raff_rows.at[
                index, 'Date'] = f'{str(raff_rows["Date"][index]).split("-")[2]}-{str(raff_rows["Date"][index]).split("-")[1]}-{str(raff_rows["Date"][index]).split("-")[0]}'

        raff_rows = raff_rows.loc[raff_rows["Date"].between(start_day, end_day)]

        for index, value in zip(raff_rows.index, raff_rows['Date']):
            raff_rows.at[
                index, 'Date'] = f'{str(raff_rows["Date"][index]).split("-")[2]}-{str(raff_rows["Date"][index]).split("-")[1]}-{str(raff_rows["Date"][index]).split("-")[0]}'

        for index in raff_rows.index:
            if f"{raff_rows.Description[index]}" == "nan":
                row1 = ""
            else:
                row1 = raff_rows.Description[index]
            if raff_rows.Amount[index] < 0:
                self.tree_single.insert(parent="", index="end", iid=index, text="",
                                        values=[row1, raff_rows.Amount[index] * -1, 0, raff_rows.Date[index]],
                                        tags=("evenrow",))
            else:
                self.tree_single.insert(parent="", index="end", iid=index, text="",
                                        values=[row1, 0, raff_rows.Amount[index], raff_rows.Date[index]],
                                        tags=("oddrow",))

        tree_single_scrolly = Scrollbar(self.print_single, orient="vertical",
                                        command=self.tree_single.yview)
        self.tree_single.configure(yscrollcommand=tree_single_scrolly.set)
        tree_single_scrolly.place(relx=0.978, rely=0.11, anchor="n", height=370)
        self.tree_single.place(width=640, height=370, relx=0.01, rely=0.11)

        export_button = Button(self.print_single, text="Export", fg="black", bg="wheat1", relief="ridge",
                               activebackground="silver", font=("Spectral", 15),
                               command=lambda: Thread(target=self.single_from_date_window()).start())
        export_button.place(height=25, width=80, relx=0.86, rely=0.037)

        self.print_single.mainloop()

    def scandata(self):
        """Get current Date"""

        Mesi = {"January": "Gennaio", "February": "Febbraio", "March": "Marzo", "April": "Aprile", "May": "Maggio",
                "June": "Giugno", "July": "Luglio", "August": "Agosto", "September": "Settembre",
                "October": "Ottobre",
                "November": "Novembre", "December": "Dicembre"}
        oggi = date.today()
        self.mese = oggi.strftime("%B")
        dat = (oggi.strftime("%d"), Mesi[self.mese], oggi.strftime("%Y"))

        return dat


# EditValue Screen


class EditValue(Frame):

    def __init__(self, parent, controller):
        Frame.__init__(self, parent)
        self.style = ThemedStyle(self)
        self.style.theme_use(f"{styles}")

        self.controller = controller
        self.removed_index = list()

        label = Label(self, text="Edit Values", font=("Spectral", 15), foreground="black", relief="ridge")
        label.pack(side="top", fill="x", pady=10)

        button = Button(self, text="Back to main menu", command=lambda: self.back_to_menu(controller),
                        fg="black", bg="sandy brown", relief="raised",
                        activebackground="light gray", font=("Lucinda Console", 10))
        button.pack()

        self.search_label = Label(self, text="Search", font=("Spectral", 10), foreground="black", relief="ridge")
        self.search_label.config(anchor=CENTER)
        self.search_label.place(height=30, width=170, relx=0.08, rely=0.14)

        self.name_entry = Entry(self, font=("Lucinda Console", 15))
        self.name_entry.place(relx=0.411, rely=0.14, width=160, height=30)
        self.name_entry.bind('<KeyRelease>', self.find_name)

        self.container = LabelFrame(self)
        self.container.place(height=380, width=880, relx=0.01, rely=0.26)

        self.edit_tree([])
        self.name_entry.focus()

        self.edit_name_entry = Entry(self, font=("Lucinda Console", 15))
        self.edit_name_entry.place(relx=0.01, rely=0.85, width=160, height=30)

        self.description_entry = Entry(self, font=("Lucinda Console", 15))
        self.description_entry.place(relx=0.2, rely=0.85, width=220, height=30)

        self.amount_give_entry = Entry(self, font=("Lucinda Console", 15))
        self.amount_give_entry.place(relx=0.47, rely=0.85, width=140, height=30)

        self.amount_have_entry = Entry(self, font=("Lucinda Console", 15))
        self.amount_have_entry.place(relx=0.66, rely=0.85, width=140, height=30)

        add_button = Button(self, text="Edit", command=self.edit_person,
                            fg="black", bg="wheat1", relief="raised",
                            activebackground="light gray", font=("Times", 15))
        add_button.place(relx=0.85, rely=0.85, width=110, height=30)
        add_button.bind("<Return>", self.enter_add_button)

        swap_button = Button(self, text="Swap", command=self.swap,
                             fg="black", bg="#6d9ac4", relief="raised",
                             activebackground="light gray", font=("Times", 15))
        swap_button.place(relx=0.54, rely=0.93, width=110, height=30)

        clear_button = Button(self, text="Clear", command=self.clear_boxes,
                              fg="black", bg="light gray", relief="raised",
                              activebackground="#d9ead3", font=("Times", 15))
        clear_button.place(relx=0.69, rely=0.93, width=110, height=30)

        remove_button = Button(self, text="Remove", command=self.remove_person,
                               fg="black", bg="sandy brown", relief="raised",
                               activebackground="light gray", font=("Times", 15))
        remove_button.place(relx=0.85, rely=0.93, width=110, height=30)

        save_button = Button(self, text="Save", command=self.save_to_csv,
                             fg="black", bg="#8fce00", relief="raised",
                             activebackground="light gray", font=("Times", 15))
        save_button.place(relx=0.11, rely=0.2, width=110, height=30)

        self.clear_boxes()

        # Create a Listbox widget to display the clienti
        self.suggestion = Listbox(self, font=("Lucinda Console", 15), relief="flat")
        self.suggestion.place(relx=0.411, rely=0.18, width=160, height=50)
        self.suggestion.bind("<Double-Button-1>", self.select_suggestion)
        self.suggestion.bind("<Return>", self.select_suggestion)

        self.on_start()

        self.bind("<<ShowFrame>>", self.on_show_frame)

    def on_show_frame(self, event):
        """Check if window is raised from MenuP or from PrintDay"""

        global name_day
        global starting
        self.removed_index.clear()

        if not starting:
            if name_day == "":
                try:
                    self.go_back_button.destroy()
                except AttributeError:
                    pass
                self.name_entry.focus()
                self.search_label = Label(self, text="Search", font=("Spectral", 10), foreground="black",
                                          relief="ridge")
                self.search_label.config(anchor=CENTER)
                self.search_label.place(height=30, width=170, relx=0.08, rely=0.14)
            else:
                self.search_label.destroy()
                self.go_back_button = Button(self, text="Back to Print", command=self.back_to_stampa,
                                             fg="black", bg="light gray", relief="raised",
                                             activebackground="#d9ead3", font=("Times", 12))
                self.go_back_button.place(height=30, width=170, relx=0.08, rely=0.14)
                self.search_for()

    def on_start(self):
        """Build Frame"""

        global bill_names

        # Add values to combobox
        self.update_suggestion_list(bill_names)

        self.mesi_ordine = {"Gennaio": 1, "Febbraio": 2, "Marzo": 3, "Aprile": 4, "Maggio": 5, "Giugno": 6,
                            "Luglio": 7, "Agosto": 8, "Settembre": 9, "Ottobre": 10, "Novembre": 11, "Dicembre": 12}
        giorno, mese, anno = self.scandata()

        self.num_mesi = {"01": "Gennaio", "02": "Febbraio", "03": "Marzo", "04": "Aprile", "05": "Maggio",
                         "06": "Giugno", "07": "Luglio", "08": "Agosto", "09": "Settembre", "10": "Ottobre",
                         "11": "Novembre", "12": "Dicembre"}
        self.selected_date = f"('{giorno}', '{mese}', '{anno}')"

        self.search_button = Button(self, text="Search", command=self.search_for,
                                    fg="black", bg="Silver", relief="raised",
                                    activebackground="light gray", font=("Spectral", 15))
        self.search_button.place(relx=0.7, rely=0.14, height=30, width=100)
        self.search_button.bind("<Return>", self.enter_search_for)

    def edit_person(self):
        """Edit record on tree"""

        if self.amount_have_entry.get() not in ("0", "") and self.amount_give_entry.get() not in ("0", ""):
            messagebox.showerror("Attention", "You cannot enter values in DARE and AVERE at the same time. "
                                              "Only one column can be changed. To pass the values from one column to "
                                              "another it is necessary to reset one and modify the other.")
        else:
            try:
                float(self.amount_give_entry.get())
                float(self.amount_have_entry.get())
            except Exception:
                pass
            else:
                if self.edit_name_entry.get() != "":
                    self.update_clienti()
                    selected = self.tree_edit.focus()
                    self.tree_edit.item(selected, text="", values=(
                        str(self.edit_name_entry.get()).upper(), self.description_entry.get(),
                        self.amount_give_entry.get(), self.amount_have_entry.get(), self.data_on_row))
                    self.clear_boxes()
                    self.name_entry.focus()
                else:
                    pass

    def remove_person(self):
        """Remove person in tree"""

        selected_tme = self.tree_edit.selection()
        self.tree_edit.delete(selected_tme)
        # Memorize the index because then it has to be eliminated
        self.removed_index.append(selected_tme[0])

    def clear_boxes(self):
        """Clear up entry boxes"""

        self.edit_name_entry.delete(0, END)
        self.description_entry.delete(0, END)
        self.amount_give_entry.delete(0, END)
        self.amount_have_entry.delete(0, END)

    def swap(self):
        """Swap GIVE HAVE cols"""

        dare = self.amount_give_entry.get()
        avere = self.amount_have_entry.get()
        self.amount_give_entry.delete(0, END)
        self.amount_have_entry.delete(0, END)
        self.amount_give_entry.insert(0, avere)
        self.amount_have_entry.insert(0, dare)

    def enter_add_button(self, key):
        """Add field to tree via Return button"""

        self.edit_person()

    def edit_tree(self, data):
        """Edit stored value"""

        tree_style = ttk.Style()
        tree_style.configure("mystyle.Treeview", highlightthickness=0, bd=0,
                             font=('Lucinda Console', 16))  # Modify the font of the body
        tree_style.configure("mystyle.Treeview.Heading",
                             font=('Spectral', 18, 'italic'), width=1, pady=20)  # Modify the font of the headings
        tree_style.layout("mystyle.Treeview",
                          [('mystyle.Treeview.treearea', {'sticky': 'nswe'})])  # Remove the borders
        tree_style.configure('Treeview', rowheight=40)

        self.tree_edit = ttk.Treeview(self.container, style="mystyle.Treeview")

        self.tree_edit["columns"] = ("Name", "Description", "GIVE", "HAVE", "Date")
        self.tree_edit.column("#0", width=0, stretch=NO)
        self.tree_edit.column("Name", anchor=W, width=200, stretch=NO)
        self.tree_edit.column("Description", anchor=CENTER, width=210, stretch=NO)
        self.tree_edit.column("GIVE", anchor=CENTER, width=160, stretch=NO)
        self.tree_edit.column("HAVE", anchor=CENTER, width=160, stretch=NO)
        self.tree_edit.column("Date", anchor=CENTER, width=125, stretch=NO)

        self.tree_edit.heading("#0", text="", anchor=CENTER)
        self.tree_edit.heading("Name", text="Name", anchor=W)
        self.tree_edit.heading("Description", text="Description", anchor=CENTER)
        self.tree_edit.heading("GIVE", text="GIVE", anchor=CENTER)
        self.tree_edit.heading("HAVE", text="HAVE", anchor=CENTER)
        self.tree_edit.heading("Date", text="Date", anchor=CENTER)

        self.tree_edit.tag_configure("oddrow", background="white")
        self.tree_edit.tag_configure("evenrow", background="SlateGray2")

        try:
            for index, row in data.iterrows():
                if f"{row[1]}" == "nan":
                    row1 = ""
                else:
                    row1 = row[1]
                if row[2] < 0:
                    self.tree_edit.insert(parent="", index="end", iid=index, text="",
                                          values=[row[0], row1, row[2] * -1, 0, row[3], index],
                                          tags=("evenrow",))
                else:
                    self.tree_edit.insert(parent="", index="end", iid=index, text="",
                                          values=[row[0], row1, 0, row[2], row[3], index],
                                          tags=("oddrow",))
        except AttributeError:
            pass

        tree_stampa_scrolly = Scrollbar(self.container, orient="vertical",
                                        command=self.tree_edit.yview)
        self.tree_edit.configure(yscrollcommand=tree_stampa_scrolly.set)
        tree_stampa_scrolly.place(relx=0.988, rely=0, anchor="n", height=380)
        self.tree_edit.bind("<<TreeviewSelect>>", self.display_edit_person)
        self.tree_edit.place(width=870, height=372, relx=0, rely=0)

    def display_edit_person(self, sel):
        """Fill Entry boxes with tree selection"""

        self.clear_boxes()

        selected = self.tree_edit.focus()
        selected_row = self.tree_edit.item(selected, "values")
        self.edit_name_entry.insert(0, selected_row[0])
        self.description_entry.insert(0, selected_row[1])
        self.amount_give_entry.insert(0, selected_row[2])
        self.amount_have_entry.insert(0, selected_row[3])
        self.data_on_row = selected_row[4]
        self.edit_name_entry.focus()

    def suggestion_key(self, key):
        """Keyboard suggestion base on value list"""

        global bill_names

        typing = self.name_entry.get()
        if typing == '':
            data = bill_names
        else:
            data = []
            for item in bill_names:
                if typing.lower() in item.lower():
                    data.append(item)

        self.update_suggestion_list(data)

    def update_suggestion_list(self, data):
        """Keyboard suggestion base on previous function"""

        # Clear the Combobox
        self.suggestion.delete(0, END)
        # Add values to the combobox
        for value in data:
            self.suggestion.insert(END, value)

    def select_suggestion(self, key):
        """Enter the suggestion on Entry box"""

        self.name_entry.delete(0, "end")
        for i in self.suggestion.curselection():
            if " " in self.suggestion.get(i):
                self.name_entry.insert(0, self.suggestion.get(i))
            else:
                self.name_entry.insert(0, self.suggestion.get(i))

    def enter_search_for(self, key):
        """Search via button"""

        self.search_for()

    def search_for(self):
        """Search for bill base on date and name"""

        global name_day

        if name_day != "" and self.name_entry.get() == "":
            self.search_label.destroy()
            self.name_look_for = name_day
        elif self.name_entry.get() == "":
            pass
        elif name_day != "" and self.name_entry.get() != "":
            self.name_look_for = self.name_entry.get()
            self.name_entry.delete(0, END)
        else:
            self.name_look_for = self.name_entry.get()
            try:
                self.go_back_button.destroy()
            except AttributeError:
                pass
            self.search_label = Label(self, text="", font=("Spectral", 10), foreground="black", relief="ridge")
            self.search_label.config(anchor=CENTER)
            self.search_label.place(height=30, width=170, relx=0.08, rely=0.14)
            self.name_entry.delete(0, END)

        self.removed_index = list()
        self.update_tree()

    def update_tree(self):
        """Update edit tree"""

        global name_day
        global global_data_in_out
        global dataset_file
        global_data_in_out = read_csv(f"{dataset_file}", engine="c")

        self.edit_tree(global_data_in_out[global_data_in_out["Name"] == self.name_look_for])

        if name_day != "":
            pass
        elif global_data_in_out[global_data_in_out["Name"] == self.name_look_for].size > 0:
            self.search_label.config(text="Search Completed")
        else:
            self.search_label.config(text="There are no matches")

    @staticmethod
    def back_to_menu(controller):
        """Return to main menu"""
        global name_day
        name_day = ""
        controller.show_frame("MenuP")

    def back_to_stampa(self):
        self.controller.show_frame("PrintDay")

    def delete_entries(self):
        """Clean up treeview"""

        for row in self.tree_edit.get_children():
            self.tree_edit.delete(row)

    def update_clienti(self):
        """Update clienti.txt file base on new input"""

        global bill_names
        global bill_names_file
        name = str(self.name_entry.get())
        raff_name = ''

        for letter in name.split(' '):
            if letter != '':
                raff_name += letter + ' '

        if (raff_name[:-1]).upper() not in bill_names:
            bill_names.append(self.name_entry.get())
            bill_names.sort()
            with open(f"{bill_names_file}", "w") as file:
                for c in bill_names:
                    file.write(f"{c.upper()}\n")

    def save_to_csv(self):
        """Update csv file with edited entries"""

        global global_data_in_out
        global dataset_file
        global_data_in_out = read_csv(f"{dataset_file}", engine="c")

        the_item = ""
        name = list()
        description_list = list()
        give_list = list()
        have_list = list()
        edit_tree_counter = 0

        for item in self.tree_edit.get_children():
            t_g = ""
            t_h = ""
            the_item = str(self.tree_edit.item(item).values()).split(",")
            name.append(the_item[2][3:len(the_item[2]) - 1].upper())
            description_list.append(the_item[3][2:len(the_item[3]) - 1])
            for i in the_item[4][1:len(the_item[4])]:
                if i in ("\n", "'"):
                    pass
                else:
                    t_g += i
            if t_g == "":
                t_g = "0"
            give_list.append(-float(t_g))

            for i in the_item[5][1:len(the_item[5])]:
                if i in ("\n", "'"):
                    pass
                else:
                    t_h += i
            if t_h == "":
                t_h = "0"
            have_list.append(float(t_h))

        for the_index, this_index in enumerate(
                global_data_in_out[global_data_in_out["Name"] == self.name_look_for].index):
            # Remove rows
            if f"{this_index}" in self.removed_index:
                global_data_in_out.drop(this_index, inplace=True)
            # Edit rows
            else:
                global_data_in_out.loc[this_index, "Name"] = name[edit_tree_counter]
                global_data_in_out.loc[this_index, "Description"] = description_list[edit_tree_counter].replace("'", "")
                if have_list[edit_tree_counter] == 0:
                    global_data_in_out.loc[this_index, "Amount"] = give_list[edit_tree_counter]
                else:
                    global_data_in_out.loc[this_index, "Amount"] = have_list[edit_tree_counter]
                edit_tree_counter += 1

        global_data_in_out.to_csv(f"{dataset_file}", index=False)

        self.removed_index.clear()
        messagebox.showinfo("Save", "Save Done!")

    def scandata(self):
        """Get current Date"""

        self.Mesi = {"January": "Gennaio", "February": "Febbraio", "March": "Marzo", "April": "Aprile", "May": "Maggio",
                     "June": "Giugno", "July": "Luglio", "August": "Agosto", "September": "Settembre",
                     "October": "Ottobre",
                     "November": "Novembre", "December": "Dicembre"}
        oggi = date.today()
        mese = oggi.strftime("%B")
        dat = (oggi.strftime("%d"), self.Mesi[mese], oggi.strftime("%Y"))

        return dat

    def find_name(self, b):
        """Auto complete name typed"""

        self.suggestion_key(b)

        global global_data_in_out

        name_in_the_box = str(self.name_entry.get()).upper()

        if global_data_in_out[global_data_in_out["Name"] == str('SIG ' + name_in_the_box)].size > 0:
            self.name_entry.delete(0, END)
            self.name_entry.insert(0, str('SIG ' + name_in_the_box))
        else:
            self.name_entry.delete(0, END)
            self.name_entry.insert(0, name_in_the_box)


# Reset Screen


class Reset(Frame):

    def __init__(self, parent, controller):
        Frame.__init__(self, parent)
        self.style = ThemedStyle(self)
        self.style.theme_use(f"{styles}")

        self.controller = controller

        label = Label(self, text="Reset", font=("Spectral", 15), foreground="black", relief="ridge")
        label.pack(side="top", fill="x", pady=10)

        button = Button(self, text="Back to main menu", command=lambda: controller.show_frame("MenuP"),
                        fg="black", bg="sandy brown", relief="raised",
                        activebackground="light gray", font=("Lucinda Console", 10))
        button.pack()

        self.on_start()

    def on_start(self):
        """Build Frame"""

        azzero_label = Label(self, text="Click on Zero"
                                        " delete all accounts whose differences are zero",
                             font=("Helvetica", 15), foreground="black", relief="ridge")
        azzero_label.pack(fill="x", pady=20)

        azzera_button = Button(self, text="Zero", command=self.azzero,
                               fg="black", bg="Silver", relief="raised",
                               activebackground="light gray", font=("Times", 15))
        azzera_button.pack()

    def azzero(self):
        """Ask for Azzero"""

        ask_for_azzero = messagebox.askyesno("Attention", "Are you sure you want to proceed?")

        if ask_for_azzero:
            self.do_azzero()

    def do_azzero(self):
        """Delete rows with GIVE - HAVE == 0"""

        self.loading_window = SecondWindow(self, "Wait", "300x50", False, (0, 0))
        stile_load = ThemedStyle(self.loading_window)
        stile_load.theme_use(f"{styles}")

        self.load_label = Label(self.loading_window, text="Reading File", font=("Spectral", 15),
                                foreground="black", relief="ridge")
        self.load_label.pack(side="top", fill="x", pady=10)

        self.loading_window.after(1, Thread(target=self.task).start())
        self.loading_window.mainloop()

    def task(self):
        """Support func for loading"""

        global global_data_in_out
        global dataset_file
        global_data_in_out = read_csv(f"{dataset_file}", engine="c")
        size_of_unique = global_data_in_out.Name.unique().size
        for index_minor, bill_name in enumerate(global_data_in_out["Name"].unique()):
            give = round(fsum(global_data_in_out[global_data_in_out["Name"] == bill_name][
                                  global_data_in_out[global_data_in_out["Name"] == bill_name]["Amount"] < 0][
                                  "Amount"]), 2)
            have = round(fsum(global_data_in_out[global_data_in_out["Name"] == bill_name][
                                  global_data_in_out[global_data_in_out["Name"] == bill_name]["Amount"] > 0][
                                  "Amount"]), 2)

            if -0.0001 <= round(fsum([have, give]), 2) <= 0.0001:
                for index_on_df in global_data_in_out[global_data_in_out["Name"] == bill_name]["Name"].index:
                    global_data_in_out.drop(index_on_df, inplace=True)

            self.load_label.config(text=f"Loading {round((100 * index_minor) / size_of_unique, 1)}%")
            self.loading_window.update_idletasks()
        global_data_in_out.to_csv(f"{dataset_file}", index=False)
        self.loading_window.destroy()


#  UndoLastImport Screen


class UndoLastImport(Frame):

    def __init__(self, parent, controller):
        Frame.__init__(self, parent)
        self.style = ThemedStyle(self)
        self.style.theme_use(f"{styles}")

        self.controller = controller

        label = Label(self, text="Undo Last Import", font=("Spectral", 15), foreground="black",
                      relief="ridge")
        label.pack(side="top", fill="x", pady=10)

        button = Button(self, text="Back to main menu", command=lambda: controller.show_frame("MenuP"),
                        fg="black", bg="sandy brown", relief="raised",
                        activebackground="light gray", font=("Lucinda Console", 10))
        button.pack()

        self.undo_label = Label(self, text="",
                                font=("Spectral", 15),
                                foreground="black",
                                relief="ridge")
        self.undo_label.config(anchor=CENTER)
        self.undo_label.pack(side="top", fill="x", pady=10)

        undo_button = Button(self, text="Removal", command=self.undo_last,
                             fg="black", bg="silver", relief="raised",
                             activebackground="wheat1", font=("Spectral", 13))
        undo_button.pack()

        self.bind("<<ShowFrame>>", self.on_show_frame)

    def on_show_frame(self, event):
        """Update label every time Frame is showed up"""

        self.giorno, self.mese, self.anno = self.get_last_date_on_file()

        self.undo_label.config(text=f"The last entry took place on {self.giorno}/{self.mese}/{self.anno}",
                               anchor=CENTER)

    def undo_last(self):
        """Loading Screen"""

        ask_for_undo = messagebox.askyesno('Removal', 'Remove the latest imported data?')

        if ask_for_undo:
            self.loading_window = SecondWindow(self, "Wait", "300x50", False, (0, 0))
            stile_load = ThemedStyle(self.loading_window)
            stile_load.theme_use(f"{styles}")
            self.loading_window.title("Wait")

            self.load_label = Label(self.loading_window, text="Reading File", font=("Spectral", 15),
                                    foreground="black", relief="ridge")
            self.load_label.pack(side="top", fill="x", pady=10)

            self.loading_window.after(1, Thread(target=self.task).start())
            self.loading_window.mainloop()

    def task(self):
        """Support function for loading"""

        global global_data_in_out
        global dataset_file
        global_data_in_out = read_csv(f"{dataset_file}", engine="c")
        size_of_df = global_data_in_out.Date.size
        if len(str(self.giorno)) == 1:
            self.giorno = f"0{self.giorno}"
        for index, date_time in enumerate(
                global_data_in_out[global_data_in_out["Date"] == f"{self.giorno}-{self.mese}-{self.anno}"].index):
            global_data_in_out.drop(date_time, inplace=True)
            self.load_label.config(text=f"Loading {round((100 * index) / size_of_df, 1)}%")
            self.loading_window.update_idletasks()

        global_data_in_out.to_csv(f"{dataset_file}", index=False)
        self.loading_window.destroy()
        self.undo_label.config(text="Removal Done", anchor=CENTER)

    @staticmethod
    def get_last_date_on_file():
        """Go to the end of file to get last date"""

        global global_data_in_out
        global dataset_file
        global_data_in_out = read_csv(f"{dataset_file}", engine="c")

        global_data_in_out['Date'] = to_datetime(global_data_in_out['Date'], format='%d-%m-%Y')
        global_data_in_out.sort_values(by='Date', inplace=True)
        last_date = str(global_data_in_out["Date"].tail(1).to_list()[0])
        raff_last_date = last_date.split(" ")[0]
        giorno, mese, anno = int(raff_last_date.split("-")[2]), int(raff_last_date.split("-")[1]), int(
            raff_last_date.split("-")[0])

        return giorno, mese, anno


# Trasmission Screen


class Trasmission(Frame):

    def __init__(self, parent, controller):
        Frame.__init__(self, parent)
        self.style = ThemedStyle(self)
        self.style.theme_use(f"{styles}")

        self.controller = controller

        label = Label(self, text="Trasmission", font=("Spectral", 15), foreground="black",
                      relief="ridge")
        label.pack(side="top", fill="x", pady=10)

        button = Button(self, text="Back to main menu", command=lambda: controller.show_frame("MenuP"),
                        fg="black", bg="sandy brown", relief="raised",
                        activebackground="light gray", font=("Lucinda Console", 10))
        button.pack()

        self.on_start()

        self.bind("<<ShowFrame>>", self.on_show_frame)

    def on_show_frame(self, event):
        """ON  Raise"""
        global starting

        if not starting:
            self.name_entry.delete(0, END)
            self.name_entry.focus()

    def on_start(self):
        """Build Frame"""

        search_button = Button(self, text="Search", command=self.enter_search,
                               fg="black", bg="Silver", relief="raised",
                               activebackground="light gray", font=("Times", 15))
        search_button.place(relx=0.25, rely=0.08, width=110, height=30)

        self.trasmissione_container = LabelFrame(self)
        self.trasmissione_container.place(height=510, width=880, relx=0.01, rely=0.22)

        self.name_entry = Entry(self, font=("Lucinda Console", 13))
        self.name_entry.place(relx=0.05, rely=0.08, width=160, height=30)
        self.name_entry.bind('<KeyRelease>', self.find_name)

        # Create a Listbox widget to display the list of items
        self.suggestion = Listbox(self, font=("Lucinda Console", 13), relief="flat")
        self.suggestion.place(relx=0.05, rely=0.12, width=160, height=60)
        self.suggestion.bind("<Double-Button-1>", self.select_suggestion)
        self.suggestion.bind("<Return>", self.select_suggestion)

        global bill_names

        # Add values to combobox
        self.update_suggestion_list(bill_names)

        self.mesi_ordine = {"Gennaio": 1, "Febbraio": 2, "Marzo": 3, "Aprile": 4, "Maggio": 5, "Giugno": 6,
                            "Luglio": 7, "Agosto": 8, "Settembre": 9, "Ottobre": 10, "Novembre": 11, "Dicembre": 12}
        self.num_to_mesi = {1: "Gennaio", 2: "Febbraio", 3: "Marzo", 4: "Aprile", 5: "Maggio", 6: "Giugno",
                            7: "Luglio", 8: "Agosto", 9: "Settembre", 10: "Ottobre", 11: "Novembre", 12: "Dicembre"}

        self.stampa_trasmissione_tree([])

        self.transfer_entry = Entry(self, font=("Lucinda Console", 13))
        self.transfer_entry.place(relx=0.75, rely=0.08, width=160, height=30)
        self.transfer_entry.bind('<KeyRelease>', self.make_upper_trasmission_entry)
        transfer_button = Button(self, text="Trasferimento", command=self.do_transfer,
                                 fg="black", bg="Silver", relief="raised",
                                 activebackground="light gray", font=("Times", 15))
        transfer_button.place(relx=0.75, rely=0.13, width=160, height=30)

    def suggestion_key(self, key):
        """Keyboard suggestion base on value list"""

        global bill_names

        typing = self.name_entry.get()
        if typing == '':
            data = bill_names
        else:
            data = []
            for item in bill_names:
                if typing.lower() in item.lower():
                    data.append(item)

        self.update_suggestion_list(data)

    def update_suggestion_list(self, data):
        """Keyboard suggestion base on previous function"""

        # Clear the Combobox
        self.suggestion.delete(0, END)
        # Add values to the combobox
        for value in data:
            self.suggestion.insert(END, value)

    def select_suggestion(self, key):
        """Enter the suggestion on Entry box"""

        self.name_entry.delete(0, "end")
        for i in self.suggestion.curselection():
            if " " in self.suggestion.get(i):
                self.name_entry.insert(0, self.suggestion.get(i))
            else:
                self.name_entry.insert(0, self.suggestion.get(i))

    def stampa_trasmissione_tree(self, raff_rows):
        """Add Tree in tab"""

        tree_style = ttk.Style()
        tree_style.configure("mystyle.Treeview", highlightthickness=0, bd=0,
                             font=('Lucinda Console', 16))  # Modify the font of the body
        tree_style.configure("mystyle.Treeview.Heading",
                             font=('Spectral', 18, 'italic'), width=1, pady=20)  # Modify the font of the headings
        tree_style.layout("mystyle.Treeview",
                          [('mystyle.Treeview.treearea', {'sticky': 'nswe'})])  # Remove the borders
        tree_style.configure('Treeview', rowheight=40)

        self.tree_trasmission = ttk.Treeview(self.trasmissione_container, style="mystyle.Treeview")

        self.tree_trasmission["columns"] = ("Name", "Description", "GIVE", "HAVE", "Date")
        self.tree_trasmission.column("#0", width=0, stretch=NO)
        self.tree_trasmission.column("Name", anchor=W, width=200, stretch=NO)
        self.tree_trasmission.column("Description", anchor=CENTER, width=210, stretch=NO)
        self.tree_trasmission.column("GIVE", anchor=CENTER, width=160, stretch=NO)
        self.tree_trasmission.column("HAVE", anchor=CENTER, width=160, stretch=NO)
        self.tree_trasmission.column("Date", anchor=CENTER, width=125, stretch=NO)

        self.tree_trasmission.heading("#0", text="", anchor=CENTER)
        self.tree_trasmission.heading("Name", text="Name", anchor=W)
        self.tree_trasmission.heading("Description", text="Description", anchor=CENTER)
        self.tree_trasmission.heading("GIVE", text="GIVE", anchor=CENTER)
        self.tree_trasmission.heading("HAVE", text="HAVE", anchor=CENTER)
        self.tree_trasmission.heading("Date", text="Date", anchor=CENTER)

        self.tree_trasmission.tag_configure("oddrow", background="white")
        self.tree_trasmission.tag_configure("evenrow", background="SlateGray2")

        try:
            for index, row in raff_rows.iterrows():
                if f"{row[1]}" == "nan":
                    row1 = ""
                else:
                    row1 = row[1]
                if row[2] < 0:
                    self.tree_trasmission.insert(parent="", index="end", iid=index, text="",
                                                 values=[row[0], row1, row[2] * -1, 0, row[3]],
                                                 tags=("evenrow",))
                else:
                    self.tree_trasmission.insert(parent="", index="end", iid=index, text="",
                                                 values=[row[0], row1, 0, row[2], row[3]],
                                                 tags=("oddrow",))
        except AttributeError:
            for row in raff_rows:
                self.tree_trasmission.insert(parent="", index="end", iid=index, text="",
                                             values=[],
                                             tags=("oddrow",))

        tree_trasmissione_scrolly = Scrollbar(self.trasmissione_container, orient="vertical",
                                              command=self.tree_trasmission.yview)
        self.tree_trasmission.configure(yscrollcommand=tree_trasmissione_scrolly.set)
        tree_trasmissione_scrolly.place(relx=0.988, rely=0, anchor="n", height=500)
        self.tree_trasmission.place(width=875, height=500, relx=0, rely=0)

    def update_trasmissione_tree(self):
        """Update tree for new added values"""

        global global_data_in_out
        global dataset_file
        global_data_in_out = read_csv(f"{dataset_file}", engine="c")

        self.stampa_trasmissione_tree(global_data_in_out[global_data_in_out["Name"] == self.name_entry.get()])

    def enter_search(self):

        for row in self.tree_trasmission.get_children():
            self.tree_trasmission.delete(row)

        Thread(target=self.update_trasmissione_tree()).start()

    def find_name(self, b):
        """Auto complete name typed"""

        self.suggestion_key(b)

        global global_data_in_out
        global dataset_file
        global_data_in_out = read_csv(f"{dataset_file}", engine="c")

        name_in_the_box = str(self.name_entry.get()).upper()

        if global_data_in_out[global_data_in_out["Name"] == str('SIG ' + name_in_the_box)].size > 0:
            self.name_entry.delete(0, END)
            self.name_entry.insert(0, str('SIG ' + name_in_the_box))
        else:
            self.name_entry.delete(0, END)
            self.name_entry.insert(0, name_in_the_box)

    def make_upper_trasmission_entry(self, letter):
        """Make Capital Letters Typed In Trasmission Entry"""

        self.suggestion_key(letter)

        global global_data_in_out
        name_in_the_box = str(self.transfer_entry.get()).upper()

        if global_data_in_out[global_data_in_out["Name"] == str('SIG ' + name_in_the_box)].size > 0:
            self.transfer_entry.delete(0, END)
            self.transfer_entry.insert(0, str('SIG ' + name_in_the_box))
        else:
            self.transfer_entry.delete(0, END)
            self.transfer_entry.insert(0, name_in_the_box)

    def do_transfer(self):
        """Rename the displayed bill"""

        global global_data_in_out
        global dataset_file
        for this_index in global_data_in_out[global_data_in_out["Name"] == self.name_entry.get()].index:
            global_data_in_out.loc[this_index, "Name"] = self.transfer_entry.get().upper()
        self.stampa_trasmissione_tree(
            global_data_in_out[global_data_in_out["Name"] == self.transfer_entry.get().upper()])
        global_data_in_out.to_csv(f"{dataset_file}", index=False)


if __name__ == "__main__":
    Cashier().mainloop()
