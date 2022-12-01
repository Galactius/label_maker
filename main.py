# python-docx, package installed with venv, enables R/W access for Word Docs.
from docx import Document
from docx.shared import Pt
from docx.enum.style import WD_STYLE_TYPE
from docx2pdf import convert
from datetime import date
import os
# from waiting import wait

# tkinter imports
import tkinter as tk
from tkinter import simpledialog
from tkinter import messagebox
from tkinter import Label
from tkinter import Entry
from tkinter import Button


class MainApplication(tk.Frame):
    def validate_name_entry(input):
        return isinstance(input, str)
    def validate_room_entry(input):
        return isinstance(input, int)
    def __init__(self, parent, *args, **kwargs):
        def handle_submit(self):
            print("Form Submitted...")
            guest_names.append(ent_name.get())
            guest_rooms.append(ent_room_number.get())
            guest_arrival.append(ent_arrival_date.get())
            guest_depart.append(ent_depart_date.get())
            print("Data Stored...")
            # print("guest_names: ",*guest_names)
            # print("guest_rooms: ", *guest_rooms)
            # print("guest_arrival: ", *guest_arrival)
            # print("guest_depart: ", *guest_depart)

            # msg_box = messagebox.askquestion('New Label', 'Would you like to enter another label for printing? (Up to 4)',
            #                                  icon='warning')
            # if msg_box == 'yes':
            #
            #     root.destroy()
            # else:
            #     confirm = messagebox.showinfo('Confirmation', 'Success! Generating label(s)')
            #
            #     exit()

        # Setup inheritance for tkinter
        tk.Frame.__init__(self, parent, *args, **kwargs)
        self.parent = parent

        # Create lists to store data from form
        guest_names = []
        guest_rooms = []
        guest_arrival = []
        guest_depart = []

        # Generate form
        self.parent.title("Daniel's Label Maker v0.4b")
        root.geometry('500x500')

        lbl_form_title = Label(root,text = "Populate Label Information Below:", width = 30, font = ('bold', 25))
        lbl_form_title.pack()

        lbl_name = Label(root, text="Guest Name: ", width = 15, font = 8)
        lbl_name.pack()

        ent_name = Entry(root)
        ent_name.pack()

        lbl_room_number = Label(root, text="Room Number: ", width = 15, font = 8)
        lbl_room_number.pack()

        ent_room_number = Entry(root)
        ent_room_number.pack()

        lbl_arrival_date = Label(root, text="Arrival Date (MM/DD/YYYY)", width = 25, font = 8)
        lbl_arrival_date.pack()

        ent_arrival_date = Entry(root)
        ent_arrival_date.pack()

        lbl_depart_date = Label(root, text = "Date of Departure (MM/DD/YYYY)", width = 25, font = 8)
        lbl_depart_date.pack()

        ent_depart_date = Entry(root)
        ent_depart_date.pack()

        btn_submit = Button(root, text="Submit", width = 20, fg='black')
        btn_submit.pack()
        btn_submit.bind("<Button-1>", handle_submit)

        # TODO fix issue with handle_submit running before submit button being pressed.

        # string = guest_names[0] + guest_rooms[0]

        # Create confirmation message and end tkinter root
        #confirm = messagebox.showinfo('Confirmation', 'Success! Generating label(s)')
        #root.destroy()

        # I could do a simple non dialog form that instead of having 1 simple dialog at the end for confirmation, has that confirmation
        # but before that, a messagebox/simple Y/N button to ask if the user wants to add another label to the current page. Then when they hit no
        # it gives the confirmation and continues execution as normal. If they click yes (add another label) then simply add the current form to a list
        # that holds the names and another for the room number/dates, then empty the form and prompt again.

        # while True:
        #   spawn simple form (w/ submit button)
        #   store entered data into separate lists once they hit submit
        #   clear form
        #   spawn simple y/n button dialog asking if they want to add a new label to the page
        #   if (yes button pressed)
        #       do continue while loop
        #   if (no button pressed)
        #       do break while loop
        # spawn confirmation messagebox

        # TODO: Add ability to add multiple labels to the same page
        # TODO: ADD ability to enter date/time for reservation

        # create simple form:
        print("guest_names: ", *guest_names)
        print("guest_rooms: ", *guest_rooms)
        print("guest_arrival: ", *guest_arrival)
        print("guest_depart: ", *guest_depart)

        # Get current date/time
        date_str = date.today()



        '''
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

        # Covert document to pdf for printing (to physical printer) then delete printed pdf.
        # needed to resolve XPS error when printing .docx files
        convert(cur_document, "labels.pdf")
        # os.system("lp labels.pdf")
        os.remove("labels.pdf")
    '''


if __name__ == '__main__':
    root = tk.Tk()
    MainApplication(root)
    root.mainloop()
