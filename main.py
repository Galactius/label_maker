# python-docx, package installed with venv, enables R/W access for Word Docs.
from docx import Document
from docx.shared import Pt
from docx.enum.style import WD_STYLE_TYPE
from docx2pdf import convert
from datetime import date
import os

# tkinter imports
import tkinter as tk
from tkinter import simpledialog
from tkinter import messagebox


class MainApplication(tk.Frame):
    def validate_name_entry(input):
        return isinstance(input, str)
    def validate_room_entry(input):
        return isinstance(input, int)
    def __init__(self, parent, *args, **kwargs):
        # Setup inheritance for tkinter and remove initialized window
        tk.Frame.__init__(self, parent, *args, **kwargs)
        self.parent = parent
        self.parent.title('My Window')
        root.withdraw()

        # TODO: Add ability to add multiple labels to the same page
        # TODO: ADD ability to enter date/time for reservation

        # Get current date/time
        date_str = date.today()

        # Create Dialog Prompts
        # TODO: Implement input validation
        dlg_name = simpledialog.askstring('name', 'Enter Guest Name: ', parent=parent)
        dlg_room_no = simpledialog.askinteger('room', 'Enter Guest Room Number: ', parent=parent)
        string = dlg_name + "  " + str(dlg_room_no)

        # Create confirmation message and end tkinter root
        confirm = messagebox.showinfo('Confirmation', 'Success! Generating label: ' + string)
        root.destroy()

        # Setup Document and Paragraph for labels
        document = Document()
        label = document.add_paragraph()

        # Setup styling for thicker label
        # TODO add styling for smaller date on label
        font_styles = document.styles
        font_charstyle = font_styles.add_style('thick_label', WD_STYLE_TYPE.CHARACTER)
        font_object = font_charstyle.font
        font_object.size = Pt(20)
        font_object.bold = True
        font_object.name = 'Times New Roman'

        # Styling for thinner label
        font_styles2 = document.styles
        font_charstyle2 = font_styles2.add_style('thin_label', WD_STYLE_TYPE.CHARACTER)
        font_object2 = font_charstyle2.font
        font_object2.size = Pt(20)
        font_object2.bold = True
        font_object2.name = 'Times New Roman'

        # Add text to docx and save
        label.add_run(string, style='thick_label')
        label.add_run("\n")
        label.add_run("\n")
        label.add_run(string, style='thin_label')
        cur_document = "labels - " + str(date_str) + ".docx"
        document.save(cur_document)

        # test message for test branch

        # Covert document to pdf for printing (to physical printer) then delete printed pdf.
        # needed to resolve XPS error when printing .docx files
        convert(cur_document, "labels.pdf")
        os.system("lp labels.pdf")
        os.remove("labels.pdf")


if __name__ == '__main__':
    root = tk.Tk()
    MainApplication(root)
    root.mainloop()
