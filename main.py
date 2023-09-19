from pdf2docx import Converter
from tkinter import *
from PIL import Image, ImageTk
from tkinter import filedialog
import os

selected_input_path = None

def get_location_PDF():
    global selected_input_path
    pdf_path = filedialog.askopenfilename(filetypes=[('PDF files', '*.pdf')])
    if pdf_path:
        my_PDF_entry.delete(0, END)
        my_PDF_entry.insert(0, pdf_path)
        selected_input_path = pdf_path

def put_location_docx():
    output_place = filedialog.askdirectory()
    if output_place:
        my_docx_entry.delete(0, END)
        my_docx_entry.insert(0, output_place)

def convert_pdf_to_docx():
    if len(my_PDF_entry.get()) == 0 or len(my_docx_entry.get()) == 0:
        status_label.config(text="Please enter all info!")
    else:
        global selected_input_path
        if selected_input_path:
            extension = os.path.splitext(selected_input_path)[-1]
            docx_path = selected_input_path.replace(extension, ".docx")
            try:
                cv = Converter(pdf_file=selected_input_path)
                cv.convert(docx_filename=docx_path)
                cv.close()

                output_filename = os.path.basename(docx_path)
                output_location = os.path.join(my_docx_entry.get(), output_filename)
                os.rename(docx_path, output_location)
                status_label.config(text="Conversion from pdf to text is successful.")
            except Exception as e:
                status_label.config(text=f"Error: {str(e)}")
        else:
            status_label.config(text="Please select an input pdf first.")

my_window = Tk()
my_window.title("PDF Converter")
my_window.minsize(width=300, height=200)
my_window.config(background="#08e8de")

my_icon = PhotoImage(file='pdficon.png')
my_window.iconphoto(False, my_icon)

my_img = Image.open("pdfimage.png")
new_width = 150
new_height = 100
my_img = my_img.resize((new_width, new_height))
tk_img = ImageTk.PhotoImage(my_img)
label = Label(my_window, image=tk_img)
label.grid(row=0, column=0, columnspan=2)
label.config(border=0)

my_PDF_label = Label(text="Select the PDF File:")
my_PDF_label.grid(row=1, column=0, sticky="w")
my_PDF_label.config(background="#08e8de")

my_PDF_entry = Entry(width=30)
my_PDF_entry.grid(row=1, column=1)
my_PDF_entry.config(background="#066eff")

my_button_icon = Image.open("buttonicon.jpg")
second_width = 20
second_height = 20
my_button_icon = my_button_icon.resize((second_width,second_height))
tk_icon = ImageTk.PhotoImage(my_button_icon)

my_location_button = Button(image=tk_icon, command=get_location_PDF)
my_location_button.grid(row=1, column=2)
my_location_button.config(border=0)

my_docx_label = Label(text="Select the Location:")
my_docx_label.grid(row=2, column=0, sticky="w")
my_docx_label.config(background="#08e8de")

my_docx_entry = Entry(width=30)
my_docx_entry.grid(row=2, column=1)
my_docx_entry.config(background="#066eff")

my_button_icon_2 = Image.open("buttonicon.jpg")
second_width_2 = 20
second_height_2 = 20
my_button_icon_2 = my_button_icon_2.resize((second_width_2,second_height_2))
tk_icon_2 = ImageTk.PhotoImage(my_button_icon_2)

my_location_button_2 = Button(image=tk_icon_2, command=put_location_docx)
my_location_button_2.grid(row=2, column=2)
my_location_button_2.config(border=0)

apply_button = Button(text="Apply", command=convert_pdf_to_docx)
apply_button.grid(row=3, column=0, columnspan=2)
apply_button.config(background="#066eff")

status_label = Label(text="")
status_label.grid(row=4, column=0, columnspan=2)
status_label.config(background="#08e8de")

my_window.mainloop()













