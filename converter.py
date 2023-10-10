import os
import sys
import comtypes.client
import tkinter as tk
from tkinter import filedialog
from tkinter import ttk
from tkinter import messagebox
from ttkthemes import ThemedStyle

def convert_word_to_pdf(input_docx, output_pdf):
    try:
        # Ensure that input_docx is a full path to the file
        input_docx = os.path.abspath(input_docx)

        # Get the current working directory
        current_directory = os.getcwd()

        # Set the output_pdf to be in the current directory
        output_pdf = os.path.join(current_directory, output_pdf)

        # Initialize COM objects
        word = comtypes.client.CreateObject("Word.Application")
        doc = word.Documents.Open(input_docx)

        # Save as PDF
        doc.SaveAs(output_pdf, FileFormat=17)  # 17 represents PDF format

        # Close the Word document
        doc.Close()
        word.Quit()

        message = f"Conversion completed: {output_pdf}"
        messagebox.showinfo("Success", message)
        return True  # Indicates successful conversion
    except Exception as e:
        error_message = f"Error: {str(e)}"
        messagebox.showerror("Error", error_message)
        return False  # Indicates conversion failure

def select_input_file():
    input_file = filedialog.askopenfilename(filetypes=[("Word Files", "*.docx")])
    input_entry.delete(0, tk.END)
    input_entry.insert(0, input_file)

def select_output_file():
    output_file = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF Files", "*.pdf")])
    output_entry.delete(0, tk.END)
    output_entry.insert(0, output_file)

def convert_button_click():
    input_docx = input_entry.get()
    output_pdf = output_entry.get()

    if not input_docx or not output_pdf:
        error_label.config(text="Please select input and output files.")
        return

    # Disable the Convert button during conversion
    convert_button.config(state=tk.DISABLED)

    try:
        # Perform the conversion
        success = convert_word_to_pdf(input_docx, output_pdf)

        if success:
            error_label.config(text=f"Conversion completed: {output_pdf}")
        else:
            error_label.config(text="Conversion failed.")
    finally:
        # Re-enable the Convert button after conversion
        convert_button.config(state=tk.NORMAL)

# Create the main window
root = tk.Tk()
root.title("Word to PDF Converter")

# Use ThemedStyle for a modern look
style = ThemedStyle(root)
style.set_theme("clam")

# Create a menu bar
menu_bar = tk.Menu(root)
root.config(menu=menu_bar)

# File menu
file_menu = tk.Menu(menu_bar, tearoff=0)
menu_bar.add_cascade(label="File", menu=file_menu)
file_menu.add_command(label="Open Word Document", command=select_input_file)
file_menu.add_separator()
file_menu.add_command(label="Exit", command=root.quit)

# Status bar
status_var = tk.StringVar()
status_bar = ttk.Label(root, textvariable=status_var, anchor=tk.W)
status_bar.pack(side=tk.BOTTOM, fill=tk.X)

# Input file selection
input_frame = ttk.Frame(root)
input_frame.pack(pady=10)
input_label = ttk.Label(input_frame, text="Select Word Document:")
input_label.pack(side=tk.LEFT, padx=5)
input_entry = ttk.Entry(input_frame, width=40)
input_entry.pack(side=tk.LEFT)
input_button = ttk.Button(input_frame, text="Browse", command=select_input_file)
input_button.pack(side=tk.LEFT, padx=5)

# Output file selection
output_frame = ttk.Frame(root)
output_frame.pack(pady=10)
output_label = ttk.Label(output_frame, text="Save as PDF:")
output_label.pack(side=tk.LEFT, padx=5)
output_entry = ttk.Entry(output_frame, width=40)
output_entry.pack(side=tk.LEFT)
output_button = ttk.Button(output_frame, text="Browse", command=select_output_file)
output_button.pack(side=tk.LEFT, padx=5)

# Convert button
convert_button = ttk.Button(root, text="Convert", command=convert_button_click)
convert_button.pack(pady=10)

# Error label with initial "Welcome!" message
error_label = ttk.Label(root, text="Welcome!", foreground="green")
error_label.pack(pady=5)

root.mainloop()
