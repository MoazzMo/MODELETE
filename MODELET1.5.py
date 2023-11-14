import os
import datetime
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from tkcalendar import DateEntry
from dateutil.relativedelta import relativedelta
from functools import partial
from PIL import Image, ImageTk  # Importing the PIL library
import base64
import io
import ctypes

def format_size(size_in_bytes):
    # Convert bytes to the appropriate unit (KB, MB, GB) with two decimal places
    for unit in ['Bytes', 'KB', 'MB', 'GB']:
        if size_in_bytes < 1024.0:
            return f"{size_in_bytes:.2f} {unit}"
        size_in_bytes /= 1024.0

class FileBrowserApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("MO DEL APP")
        self.geometry("600x500")

        self.selected_files = []
        self.filtered_files = []  # Initialize filtered_files attribute
        self.sort_column = None
        self.sort_descending = False

        # Change the icon of the app
        icon_path = "M:\MO.LOGO.ico"  # Replace "icon.ico" with the actual path to your icon file
        if os.path.exists(icon_path):
            self.iconbitmap(icon_path)
        # Change the taskbar icon (Windows and macOS)
        # icon_image = tk.PhotoImage(file=icon_path)
        # self.iconphoto(True, icon_image)
        # Change the taskbar icon (Windows)
        if os.name == "nt":  # Check if the operating system is Windows
            self.set_taskbar_icon(icon_path)

    def set_taskbar_icon(self, icon_path):
        try:
            icon_path = os.path.abspath(icon_path)
            icon_flags = 0x00000002  # LOAD_LIBRARY_AS_DATAFILE
            ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("MODELETE")  # Change "myappid" to your desired AppID
            ctypes.windll.user32.LoadImageW(None, icon_path, 1, 0, 0, icon_flags)
        except Exception as e:
            print("Error setting taskbar icon:", e)


        # Add the logo image to the app
        self.add_logo()

    def add_logo(self):
        logo_path = "M:\MO.LOGO.png"  # Replace "logo.png" with the actual path to your logo image

        if os.path.exists(logo_path):
            # Load the image using PIL and create a PhotoImage object
            logo_image = Image.open(logo_path)
            logo_image = logo_image.resize((150, 150))  # Resize the logo if necessary
            logo_photo = ImageTk.PhotoImage(logo_image)

            # Create a label to display the logo
            self.logo_label = tk.Label(self, image=logo_photo)
            self.logo_label.image = logo_photo  # Save a reference to the PhotoImage to avoid garbage collection

            # Place the logo label in the UI layout
            self.logo_label.grid(row=0, column=4, padx=5, pady=5, rowspan=5)

        self.create_ui()

    def create_ui(self):
        # File location selection
        location_label = tk.Label(self, text="Select a location:")
        self.location_textctrl = tk.Entry(self)
        browse_button = tk.Button(self, text="Browse", command=self.on_browse)

        # Search button
        search_button = tk.Button(self, text="Search", command=self.on_search)
        search_button.grid(row=1, column=0, padx=5, pady=5)

        # Date range selection
        date_range_frame = tk.LabelFrame(self, text="Date Range")
        start_date_label = tk.Label(date_range_frame, text="Start Date:")
        self.start_date_picker = DateEntry(date_range_frame, width=12, background='darkblue', foreground='white', borderwidth=2, date_pattern='yyyy-mm-dd')

        end_date_label = tk.Label(date_range_frame, text="End Date:")
        self.end_date_picker = DateEntry(date_range_frame, width=12, background='darkblue', foreground='white', borderwidth=2, date_pattern='yyyy-mm-dd')

        # Extension selection
        # Checkbutton and Entry for last number of months

        self.extensions = ['.dcm', '.dcn', '.jpg']
        extensions_label = tk.Label(self, text="Select Extensions:")
        self.extension_vars = [tk.BooleanVar(value=True) for _ in self.extensions]
        extension_checkbuttons = [ttk.Checkbutton(self, text=ext, variable=self.extension_vars[i]) for i, ext in enumerate(self.extensions)]

        # File list control
        self.file_list_ctrl = ttk.Treeview(self, columns=("Name", "Date Modified", "Size"), show="headings")
        self.file_list_ctrl.heading("Name", text="Name")
        self.file_list_ctrl.heading("Date Modified", text="Date Modified")
        self.file_list_ctrl.heading("Size", text="Size")
        self.file_list_ctrl.column("Name", width=300, anchor="center")
        self.file_list_ctrl.column("Date Modified", width=200, anchor="center")
        self.file_list_ctrl.column("Size", width=100, anchor="center")

        self.file_list_ctrl.heading("#1", command=partial(self.sort_column_callback, "Name"))
        self.file_list_ctrl.heading("#2", command=partial(self.sort_column_callback, "Date Modified"))
        self.file_list_ctrl.heading("#3", command=partial(self.sort_column_callback, "Size"))


        # Buttons
        delete_selected_button = tk.Button(self, text="Delete Selected", command=self.on_delete_selected)
        delete_all_button = tk.Button(self, text="Delete All", command=self.on_delete_all)
        select_all_button = tk.Button(self, text="Select All", command=self.on_select_all)

        # Clear button
        clear_button = tk.Button(self, text="Clear", command=self.on_clear)
        clear_button.grid(row=1, column=1, padx=5, pady=5)

        # Grid layout
        self.columnconfigure(0, weight=1)  # Allow column 0 to resize with the window
        self.columnconfigure(1, weight=1)  # Allow column 1 to resize with the window
        self.columnconfigure(2, weight=1)  # Allow column 2 to resize with the window
        self.columnconfigure(3, weight=1)  # Allow column 3 to resize with the window
        self.rowconfigure(5, weight=1)     # Allow row 5 to resize with the window

        location_label.grid(row=0, column=0, padx=5, pady=5)
        self.location_textctrl.grid(row=0, column=1, columnspan=2, padx=5, pady=5, sticky="ew")
        browse_button.grid(row=0, column=3, padx=5, pady=5)

        clear_button.grid(row=1, column=1, padx=5, pady=5)

        date_range_frame.grid(row=2, column=0, columnspan=4, padx=5, pady=5, sticky="ew")
        start_date_label.grid(row=0, column=0, padx=5, pady=5)
        self.start_date_picker.grid(row=0, column=1, padx=5, pady=5)
        end_date_label.grid(row=0, column=2, padx=5, pady=5)
        self.end_date_picker.grid(row=0, column=3, padx=5, pady=5)

        extensions_label.grid(row=3, column=0, columnspan=4, padx=5, pady=5)
        for i, checkbutton in enumerate(extension_checkbuttons):
            checkbutton.grid(row=4, column=i, padx=5, pady=5)

        self.file_list_ctrl.grid(row=5, column=0, columnspan=4, padx=5, pady=5, sticky="nsew")

        delete_selected_button.grid(row=6, column=0, padx=5, pady=5)
        delete_all_button.grid(row=6, column=1, padx=5, pady=5)
        select_all_button.grid(row=6, column=2, padx=5, pady=5)

        # Search button
        search_button = tk.Button(self, text="Search", command=self.on_search)
        search_button.grid(row=1, column=0, padx=5, pady=5)

        # Last Number of Months Modified
        self.last_months_var = tk.BooleanVar()
        self.last_months_checkbutton = ttk.Checkbutton(self, text="Last Num Monthes DESC", variable=self.last_months_var)
        self.last_months_checkbutton.grid(row=1, column=2, padx=5, pady=5)

        self.months_entry = tk.Entry(self, width=5)
        self.months_entry.grid(row=1, column=3, padx=5, pady=5)


        location_label.grid(row=0, column=0, padx=5, pady=5)
        self.location_textctrl.grid(row=0, column=1, columnspan=2, padx=5, pady=5, sticky="ew")
        browse_button.grid(row=0, column=3, padx=5, pady=5)

        date_range_frame.grid(row=2, column=0, columnspan=4, padx=5, pady=5, sticky="ew")
        start_date_label.grid(row=0, column=0, padx=5, pady=5)
        self.start_date_picker.grid(row=0, column=1, padx=5, pady=5)
        end_date_label.grid(row=0, column=2, padx=5, pady=5)
        self.end_date_picker.grid(row=0, column=3, padx=5, pady=5)

        extensions_label.grid(row=3, column=0, columnspan=4, padx=5, pady=5)
        for i, checkbutton in enumerate(extension_checkbuttons):
            checkbutton.grid(row=4, column=i, padx=5, pady=5)

        self.file_list_ctrl.grid(row=5, column=0, columnspan=4, padx=5, pady=5, sticky="nsew")

        delete_selected_button.grid(row=6, column=0, padx=5, pady=5)
        delete_all_button.grid(row=6, column=1, padx=5, pady=5)
        select_all_button.grid(row=6, column=2, padx=5, pady=5)

    def get_selected_extensions(self):
        return [self.extensions[i] for i, var in enumerate(self.extension_vars) if var.get()]

    def get_selected_dates(self):
        start_date = self.start_date_picker.get_date()
        end_date = self.end_date_picker.get_date()
        return start_date, end_date

    def on_search(self):
        location = self.location_textctrl.get()
        start_date = self.start_date_picker.get_date().strftime("%Y-%m-%d")
        end_date = self.end_date_picker.get_date().strftime("%Y-%m-%d")
        extensions = [self.extensions[i] for i in range(len(self.extensions)) if self.extension_vars[i].get()]

        if not os.path.isdir(location):
            messagebox.showerror("Invalid Directory", "Invalid directory path.")
            return

        
        filtered_files = self.get_filtered_files(location, start_date, end_date, extensions)
        self.filtered_files = filtered_files  # Store the filtered files
        self.sort_files()  # Sort the files based on the current sort settings
        self.display_files(filtered_files)  # Display the sorted files



    def on_browse(self):
        directory = filedialog.askdirectory()
        if directory:
            self.location_textctrl.delete(0, tk.END)
            self.location_textctrl.insert(0, directory)

    from dateutil.relativedelta import relativedelta

    def on_search(self):
        location = self.location_textctrl.get()
        start_date_str = self.start_date_picker.get()
        end_date_str = self.end_date_picker.get()
        extensions = [self.extensions[i] for i in range(len(self.extensions)) if self.extension_vars[i].get()]

        try:
            start_date = datetime.datetime.strptime(start_date_str, "%Y-%m-%d").date()
            end_date = datetime.datetime.strptime(end_date_str, "%Y-%m-%d").date()
        except ValueError:
            messagebox.showerror("Invalid Date", "Invalid date format. Please use the format YYYY-MM-DD.")
            return

        if not os.path.isdir(location):
            messagebox.showerror("Invalid Directory", "Invalid directory path.")
            return

        # Get the last N months value
        if self.last_months_var.get():
            try:
                num_months = int(self.months_entry.get())
            except ValueError:
                messagebox.showerror("Invalid Input", "Please enter a valid number of months.")
                return

            # Calculate the start date based on the number of months
            end_date = end_date.replace(day=1)
            start_date = (end_date - relativedelta(months=num_months - 1)).replace(day=1)

        filtered_files = self.get_filtered_files(location, start_date, end_date, extensions)
        self.display_files(filtered_files)


    def get_filtered_files(self, directory, start_date, end_date, extensions):
        filtered_files = []
        for root, _, files in os.walk(directory):
            for file in files:
                ext = os.path.splitext(file)[1]
                if ext in extensions:
                    full_path = os.path.join(root, file)  # Store the full path along with the filename
                    if self.is_within_date_range(full_path, start_date, end_date):
                        modified_time = datetime.datetime.fromtimestamp(os.path.getmtime(full_path))
                        size = os.path.getsize(full_path)
                        filtered_files.append((full_path, file, modified_time, size))  # Store the full path, the filename, modified time, and size
        return filtered_files


    def is_within_date_range(self, file_path, start_date, end_date):
        file_date = datetime.datetime.fromtimestamp(os.path.getmtime(file_path)).date()
        return start_date <= file_date <= end_date

    def display_files(self, files):
        self.file_list_ctrl.delete(*self.file_list_ctrl.get_children())
        for index, (full_path, filename, modified_time, size) in enumerate(files):
            size_str = format_size(size)
            self.file_list_ctrl.insert("", index, values=(filename, str(modified_time), size_str))

    def on_select_all(self):
        self.file_list_ctrl.selection_set(*self.file_list_ctrl.get_children())

    def on_delete_selected(self):
        selected_items = self.file_list_ctrl.selection()
        if selected_items:
            self.selected_files = [os.path.join(self.location_textctrl.get(), self.file_list_ctrl.item(item, "values")[0]) for item in selected_items]
            self.confirm_and_delete_files()

    def on_delete_all(self):
        all_items = self.file_list_ctrl.get_children()
        if all_items:
            self.selected_files = [os.path.join(self.location_textctrl.get(), self.file_list_ctrl.item(item, "values")[0]) for item in all_items]
            self.confirm_and_delete_files()

    def confirm_and_delete_files(self):
        total_size, num_files = self.calculate_selected_files_size()
        confirmation = messagebox.askyesno("Delete Confirmation", f"Are you sure you want to delete {num_files} files?\nTotal Size: {total_size} bytes")
        if confirmation:
            self.delete_selected_files()

    def calculate_selected_files_size(self):
        total_size = 0
        num_files = len(self.selected_files)
        for file_path in self.selected_files:
            if os.path.exists(file_path):
                total_size += os.path.getsize(file_path)
        return total_size, num_files

    def delete_selected_files(self):
        num_deleted_files = 0
        for file_path in self.selected_files:
            if os.path.exists(file_path):
                os.remove(file_path)
                num_deleted_files += 1
        messagebox.showinfo("Deletion Complete", f"{num_deleted_files} files have been deleted.")
        self.selected_files = []
        self.on_search()

    def on_clear(self):
        # Clear all entry and date picker widgets
        self.location_textctrl.delete(0, tk.END)
        self.start_date_picker.set_date(None)
        self.end_date_picker.set_date(None)

        # Clear the extension selection
        for var in self.extension_vars:
            var.set(True)

        # Clear the "Last Number of Months Modified" checkbox and entry
        self.last_months_var.set(False)
        self.months_entry.delete(0, tk.END)

        # Clear the file list in the Treeview
        self.file_list_ctrl.delete(*self.file_list_ctrl.get_children())

    def sort_column_callback(self, column_name):
        # Sorting callback function when a column header is clicked
        if self.sort_column == column_name:
            self.sort_descending = not self.sort_descending
        else:
            self.sort_column = column_name
            self.sort_descending = False

        self.sort_files()
        self.display_files(self.filtered_files)

    def sort_files(self):
        # Sorting the files based on the selected column
        if self.sort_column == "Name":
            self.filtered_files.sort(key=lambda item: item[1], reverse=self.sort_descending)
        elif self.sort_column == "Date Modified":
            self.filtered_files.sort(key=lambda item: item[2], reverse=self.sort_descending)
        elif self.sort_column == "Size":
            self.filtered_files.sort(key=lambda item: item[3], reverse=self.sort_descending)

if __name__ == "__main__":
    app = FileBrowserApp()
    app.mainloop()
