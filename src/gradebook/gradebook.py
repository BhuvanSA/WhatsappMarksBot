import threading
from ttkbootstrap.validation import add_regex_validation, add_range_validation
from ttkbootstrap.tableview import Tableview
# from ttkbootstrap.toast import ToastNotification
from ttkbootstrap.constants import BOTH, YES, X, LEFT, RIGHT, DANGER, SUCCESS, PRIMARY
from tkinter import filedialog
import ttkbootstrap as ttk
from PIL import Image
from excelManager import ExcelManager
from seleniumManager import SeleniumManager
from messageGenerator import messege_generator
import openpyxl
Image.CUBIC = Image.BICUBIC


class Gradebook(ttk.Frame):
    def __init__(self, master_window):
        """
        Initializes the Gradebook class.

        Args:
            master_window: The master window object.

        Returns:
            None
        """
        super().__init__(master_window, padding=(20, 10))
        self.pack(fill=BOTH, expand=YES)
        self.XLFILEPATH = ttk.StringVar(value="Select File")
        self.sheet_input_var = ttk.StringVar(value="Select Sheet")
        self.sheet_input_var.trace_add("write", self.load_data)
        self.internals_input_var = ttk.StringVar(value="Select Internals")
        self.usn_start_var = ttk.IntVar(value=1)
        self.usn_start_var.trace_add("write", self.check_data)
        self.usn_end_var = ttk.IntVar(value=10)
        self.usn_end_var.trace_add("write", self.check_data)
        self.usn_end = ttk.IntVar(value=10)

        self.mentor_name = ttk.StringVar(value="")
        self.student_id = ttk.StringVar(value="")
        self.course_name = ttk.StringVar(value="")
        self.final_score = ttk.DoubleVar(value=0)
        self.data = []
        self.colors = master_window.style.colors

        instruction_text = "Please select the excel file and sheet to load the data from."
        instruction = ttk.Label(self, text=instruction_text, width=50)
        instruction.pack(fill=X, pady=10)

        self.create_select_file("Select File: ", self.XLFILEPATH)
        self.sheet_input = self.create_option_selector(
            "Select Sheet: ", self.sheet_input_var)
        self.internals = self.create_option_selector(
            "Select Internals: ", self.internals_input_var)
        self.create_form_entry("Mentor Name: ", self.mentor_name)

        self.range = self.create_range_spinboxes(
            "SLNo Range: ", self.usn_start_var, self.usn_end_var)
        self.meter_var = self.create_meter(
            current=0, total=int(self.usn_end_var.get()) - int(self.usn_start_var.get()) + 1)
        self.create_buttonbox()
        self.table = self.create_table()

    def select_file(self):
        """
        Opens a file dialog to select a file and sets the selected file path to self.XLFILEPATH.
        Then, it calls the load_sheet method to load the selected sheet.
        """
        file_path = filedialog.askopenfilename()
        if file_path:
            self.XLFILEPATH.set(file_path)
            self.load_sheet()

    def load_sheet(self):
        """
        Loads the Excel sheet and populates the input dropdown with sheet names.

        Raises:
            Exception: If the file is not found.

        """
        try:
            workbook = openpyxl.load_workbook(
                self.XLFILEPATH.get(), read_only=True)
            self.sheet_input.configure(values=workbook.sheetnames)
        except Exception as e:
            print(e)
            print("File not found")

    def load_data(self, *args):
        """
        Loads data from an Excel file and updates the UI elements accordingly.

        Args:
            *args: Variable number of arguments.

        Returns:
            None
        """
        try:
            self.excel_manager = ExcelManager(
                self.XLFILEPATH.get(), self.sheet_input_var.get())
            self.internals.configure(
                values=[str(x) for x in range(1, self.excel_manager.max_internal+1)])
        except Exception as e:
            print(e, "\nSheet not found")

        self.range[0].configure(to=int(self.excel_manager.slno_upper_limit))
        self.range[1].configure(to=int(self.excel_manager.slno_upper_limit))
        self.usn_end.set(int(self.excel_manager.slno_upper_limit))
        self.usn_end_var.set(int(self.excel_manager.slno_upper_limit))

    def check_data(self):
        """
        Check the validity of the data entered in the gradebook.

        Args:
            *args: Variable number of arguments.

        Raises:
            Exception: If the data entered is invalid.

        Returns:
            None
        """
        try:
            add_range_validation(self.range[0], 1, self.usn_end_var.get())
            add_range_validation(
                self.range[1], self.usn_start_var.get() + 1, self.usn_end.get())
            self.meter_var.configure(
                amounttotal=(int(self.usn_end_var.get())-int(self.usn_start_var.get()) + 1))
        except:
            ...

    def create_form_entry(self, label, variable):
        """
        Creates a form entry widget with a label and an input field.

        Parameters:
        label (str): The label text for the form entry.
        variable (tkinter.StringVar): The variable to store the input value.

        Returns:
        ttk.Entry: The form input field widget.
        """
        form_field_container = ttk.Frame(self)
        form_field_container.pack(fill=X, expand=YES, pady=5)

        form_field_label = ttk.Label(
            master=form_field_container, text=label, width=15)
        form_field_label.pack(side=LEFT, padx=12)

        form_input = ttk.Entry(
            master=form_field_container, textvariable=variable)
        form_input.pack(side=LEFT, padx=5, fill=X, expand=YES)

        add_regex_validation(form_input, r'^[a-zA-Z0-9_ ]*$')

        return form_input

    def create_option_selector(self, label, variable):
        """
        Creates an option selector widget with a label and a combobox.

        Args:
            label (str): The label text for the option selector.
            variable (tkinter.StringVar): The variable to store the selected option.

        Returns:
            ttk.Combobox: The created combobox widget.
        """
        sheet_selector_container = ttk.Frame(self)
        sheet_selector_container.pack(fill=X, expand=YES, pady=5)

        sheet_selector_label = ttk.Label(
            master=sheet_selector_container, text=label, width=15)
        sheet_selector_label.pack(side=LEFT, padx=12)

        self.sheet_selector = ttk.Combobox(
            master=sheet_selector_container, textvariable=variable)
        self.sheet_selector.pack(side=LEFT, padx=5, fill=X, expand=YES)

        return self.sheet_selector

    def create_select_file(self, label, variable):
        """
        Creates a select file button with a label.

        Args:
            label (str): The label text for the select file button.
            variable: The variable associated with the select file button.

        Returns:
            None
        """
        form_field_container = ttk.Frame(self)
        form_field_container.pack(fill=X, expand=YES, pady=5)

        form_field_label = ttk.Label(
            master=form_field_container, text=label, width=15)
        form_field_label.pack(side=LEFT, padx=12)

        select_file_button = ttk.Button(
            form_field_container, textvariable=self.XLFILEPATH,
            command=self.select_file)

        select_file_button.pack(side=LEFT, padx=5, fill=X, expand=YES)

    def create_range_spinboxes(self, label, variable1, variable2):
        """
        Creates a pair of spinboxes for selecting a range of values.

        Args:
            label (str): The label text for the spinboxes.
            variable1 (tkinter.StringVar): The variable to store the value of the first spinbox.
            variable2 (tkinter.StringVar): The variable to store the value of the second spinbox.

        Returns:
            list: A list containing the two spinbox objects.
        """
        spinbox_container = ttk.Frame(self)
        spinbox_container.pack(fill=X, expand=YES, pady=5)

        spinbox_label = ttk.Label(spinbox_container, text=label, width=15)
        spinbox_label.pack(side=LEFT, padx=12)

        spinbox1 = ttk.Spinbox(spinbox_container, from_=1,
                               to=self.usn_end.get(), textvariable=variable1)
        spinbox1.pack(side=LEFT, padx=5, fill=X, expand=YES)
        spinbox2 = ttk.Spinbox(spinbox_container, from_=1,
                               to=self.usn_end.get(),  textvariable=variable2)
        spinbox2.pack(side=LEFT, padx=5, fill=X, expand=YES)
        return [spinbox1, spinbox2]

    def create_buttonbox(self):
        """
        Create a button box containing Cancel and Start buttons.

        Returns:
            None
        """
        button_container = ttk.Frame(self)
        button_container.pack(fill=X, expand=YES, pady=(15, 10))

        cancel_btn = ttk.Button(
            master=button_container,
            text="Cancel",
            command=self.on_cancel,
            bootstyle=DANGER,
            width=6,
        )

        cancel_btn.pack(side=RIGHT, padx=5)

        submit_btn = ttk.Button(
            master=button_container,
            text="Start",
            command=self.on_submit,
            bootstyle=SUCCESS,
            width=6,
        )

        submit_btn.pack(side=RIGHT, padx=5)

    def create_meter(self, current, total):
        """
        Creates and returns a ttk.Meter widget with the specified current and total values.

        Parameters:
        current (int): The current value for the meter.
        total (int): The total value for the meter.

        Returns:
        ttk.Meter: The created ttk.Meter widget.
        """
        meter = ttk.Meter(
            master=self,
            metersize=150,
            padding=5,
            amounttotal=total,
            amountused=current,
            metertype="full",
            subtext="Complete",
            interactive=False,
        )

        meter.pack(pady=(30, 0), padx=10)

        return meter

    def create_table(self):
        """
        Creates a table view with specified column data and row data.

        Returns:
            table (Tableview): The created table view.
        """
        coldata = [
            {"text": "USN"},
            {"text": "Name", "stretch": False},
            {"text": "Phone Number"},
            {"text": "Status", "stretch": False}
        ]

        table = Tableview(
            master=self,
            coldata=coldata,
            rowdata=self.data,
            paginated=True,
            searchable=True,
            bootstyle=PRIMARY,
            stripecolor=(self.colors.light, None),
        )

        table.pack(fill=BOTH, expand=YES, padx=10, pady=10)
        return table

    def on_submit(self):
        """
        Executes the submission process in a separate thread.
        Opens WhatsApp, sends messages to students, and updates the table accordingly.
        """
        def run_in_thread():
            selenium_manager = SeleniumManager()
            selenium_manager.open_whatsapp()
            for i in range(max(1, int(self.usn_start_var.get())), min(self.excel_manager.max_rows, self.usn_end_var.get())+1):
                if data := self.excel_manager.get_student_data(int(self.internals_input_var.get()), i):
                    message = messege_generator(data, self.mentor_name.get())
                    try:
                        selenium_manager.send_message(
                            message, data['phone_number'])
                        self.meter_var.configure(
                            amountused=i - self.usn_start_var.get() + 1)
                        self.table.insert_row(
                            'end', [data['usn'], data['name'], data['phone_number'], "Sent"])
                        self.table.load_table_data()  # clear_filters=True
                    except Exception as error:
                        self.table.insert_row(
                            'end', [data['usn'], data['name'], data['phone_number'], "Invalid Phone Number"])
                        self.table.load_table_data()  # clear_filters=True

            selenium_manager.logout()
        threading.Thread(target=run_in_thread).start()

    def on_cancel(self):
        """Cancel and close the application."""
        self.quit()


def main():
    """
    Entry point of the program.
    Creates an instance of the Gradebook class and starts the application's main loop.
    """
    app = ttk.Window("Marks Sender", "superhero", resizable=(True, True))
    Gradebook(app)
    app.mainloop()


if __name__ == "__main__":
    main()
