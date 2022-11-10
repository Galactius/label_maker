# python-docx, package installed with venv, enables R/W access for Word Docs.
from docx import Document
from docx.shared import Pt
from docx.enum.style import WD_STYLE_TYPE
from docx2pdf import convert
import os

# tkinter imports
import tkinter as tk
from tkinter import simpledialog
from tkinter import messagebox


class MainApplication(tk.Frame):
    def __init__(self, parent, *args, **kwargs):
        # Setup inheritance for tkinter and remove initialized window
        tk.Frame.__init__(self, parent, *args, **kwargs)
        self.parent = parent
        self.parent.title('My Window')
        root.withdraw()

        # TODO: Implement General Error Checking
        # TODO: General Improvements to

        # Create Dialog Prompts
        # TODO: Implement input validation
        dlg_name = simpledialog.askstring('name', 'Enter Guest Name: ', parent=parent)
        dlg_room_no = simpledialog.askinteger('room', 'Enter Guest Room Number: ', parent=parent)
        print(dlg_name + "  " + str(dlg_room_no))
        string = dlg_name + "  " + str(dlg_room_no)

        # Create confirmation message and end tkinter root
        # TODO: Implement error message box/window.
        confirm = messagebox.showinfo('Confirmation', 'Generating label: ' + string)
        root.destroy()

        # Setup Document and Paragraph for labels
        document = Document()
        label = document.add_paragraph()

        # Setup styling for text (used in paragraph)
        # TODO: Create separate styling for second label type.
        font_styles = document.styles
        font_charstyle = font_styles.add_style('thick_label', WD_STYLE_TYPE.CHARACTER)
        font_object = font_charstyle.font
        font_object.size = Pt(20)
        font_object.bold = True
        font_object.name = 'Times New Roman'

        # Add text to doc and save
        label.add_run(string, style='thick_label')
        label.add_run("\n")
        label.add_run("\n")
        label.add_run(string, style='test')
        document.save('test.docx')

        # Covert document to pdf for printing (to physical printer) then delete printed pdf.
        # Reason I converted to pdf is because I was receiving an XPS format error
        # while printing out the generated word docs.
        convert("test.docx", "test.pdf")
        os.system("lp test.pdf")
        os.remove("test.pdf")


if __name__ == '__main__':
    root = tk.Tk()
    MainApplication(root)
    root.mainloop()
