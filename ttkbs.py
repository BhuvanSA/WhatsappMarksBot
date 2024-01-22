from unittest import skip
import pandas as pd
import select
from numpy import var
import selenium
from ttkbootstrap.validation import add_regex_validation, add_range_validation, add_numeric_validation
from ttkbootstrap.tableview import Tableview
from ttkbootstrap.toast import ToastNotification
from ttkbootstrap.constants import BOTH, YES, X, LEFT, RIGHT, DANGER, SUCCESS, PRIMARY
from tkinter import W, filedialog
import ttkbootstrap as ttk
from PIL import Image
from excelManager import ExcelManager
from seleniumManager import SeleniumManager
from messageGenerator import messege_generator
import openpyxl

Image.CUBIC = Image.BICUBIC


class Gradebook(ttk.Frame):
    def __init__(self, master_window):
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
        # self.excel_manager =

        instruction_text = "Please enter your contact information: "
        instruction = ttk.Label(self, text=instruction_text, width=50)
        instruction.pack(fill=X, pady=10)

        self.create_select_file("Select File: ", self.XLFILEPATH)
        self.sheet_input = self.create_option_selector(
            "Select Sheet: ", self.sheet_input_var)
        self.internals = self.create_option_selector(
            "Select Internals: ", self.internals_input_var)
        self.create_form_entry("Mentor Name: ", self.mentor_name)
        # self.create_form_entry("Usn Range: ", self.course_name)
        # self.create_range_entry("Usn Range: ", )
        self.range = self.create_range_spinboxes(
            "SLNo Range: ", self.usn_start_var, self.usn_end_var)
        self.meter_var = self.create_meter(
            current=5, total=self.usn_end_var.get())
        self.create_buttonbox()
        self.table = self.create_table()

    def select_file(self):
        file_path = filedialog.askopenfilename()
        if file_path:
            self.XLFILEPATH.set(file_path)
            self.load_sheet()

    def load_sheet(self):
        try:
            workbook = openpyxl.load_workbook(self.XLFILEPATH.get())
            self.sheet_input.configure(values=workbook.sheetnames)
        except:
            print("File not found")

    def load_data(self, *args):
        # print("I am no outsider I am jaime lannister", self.sheet_input_var.get())

        try:
            self.excel_manager = ExcelManager(
                self.XLFILEPATH.get(), self.sheet_input_var.get())
            self.internals.configure(
                values=[str(x) for x in range(1, self.excel_manager.max_internal+1)])
        except:
            print("Sheet not found")

        self.range[0].configure(to=int(self.excel_manager.slno_upper_limit))
        self.range[1].configure(to=int(self.excel_manager.slno_upper_limit))
        self.usn_end.set(int(self.excel_manager.slno_upper_limit))
        self.usn_end_var.set(int(self.excel_manager.slno_upper_limit))

    def check_data(self, *args):
        add_range_validation(self.range[0], 1, self.usn_end_var.get())
        add_range_validation(
            self.range[1], self.usn_start_var.get() + 1, self.usn_end.get())
        self.meter_var.configure(
            amounttotal=self.usn_end_var.get()-self.usn_start_var.get())

    def create_form_entry(self, label, variable):
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
        spinbox_container = ttk.Frame(self)
        spinbox_container.pack(fill=X, expand=YES, pady=5)

        spinbox_label = ttk.Label(spinbox_container, text=label, width=15)
        spinbox_label.pack(side=LEFT, padx=12)

        spinbox1 = ttk.Spinbox(spinbox_container, from_=1,
                               to=self.usn_end.get(), textvariable=variable1)
        spinbox1.pack(side=LEFT, padx=5, fill=X, expand=YES)
        spinbox2 = ttk.Spinbox(spinbox_container, from_=1,
                               to=self.usn_end.get(), textvariable=variable2)
        spinbox2.pack(side=LEFT, padx=5, fill=X, expand=YES)
        return [spinbox1, spinbox2]

    def create_buttonbox(self):
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
        coldata = [
            {"text": "USN"},
            {"text": "Name", "stretch": False},
            {"text": "Phone Number"},
            {"text": "Reason", "stretch": False}
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

    # def on_submit(self):
    #     """Print the contents to console and return the values."""
    #     name = self.name.get()
    #     student_id = self.student_id.get()
    #     course_name = self.course_name.get()
    #     final_score = self.final_score_input.get()

    #     print("Name:", name)
    #     print("Student ID: ", student_id)
    #     print("Course Name:", course_name)
    #     print("Final Score:", final_score)

    #     toast = ToastNotification(
    #         title="Submission successful!",
    #         message="Your data has been successfully submitted.",
    #         duration=3000,
    #     )

    #     toast.show_toast()

    #     # Refresh table
    #     self.data.append((name, student_id, course_name, final_score))
    #     self.table.destroy()
    #     self.table = self.create_table()

    def on_submit(self):
        selenium_manager = SeleniumManager()
        selenium_manager.open_whatsapp()
        for i in range(1, min(self.excel_manager.max_rows, self.usn_end_var.get())+1):
            if data := self.excel_manager.get_student_data(int(self.internals_input_var.get()), i):
                print(data)
                message = messege_generator(data, self.mentor_name.get())
                print(message)
                selenium_manager.send_message(
                    message, data['phone_number'])
        selenium_manager.logout()

    def on_cancel(self):
        """Cancel and close the application."""
        self.quit()


if __name__ == "__main__":
    app = ttk.Window("Marks Sender", "superhero", resizable=(True, True))
    Gradebook(app)
    app.mainloop()
