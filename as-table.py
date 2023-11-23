import fitz  # PyMuPDF
import pyperclip
import tkinter as tk
from tkinter import filedialog
import subprocess

def extract_text_from_mawb_and_booking(pdf_path):
    pdf_document = fitz.open(pdf_path)
    
    try:
        # Assuming there is only one page in the PDF
        page = pdf_document[0]
        page_text = page.get_text("text")

        # Extract the text between "1. MAWB: " and "2. ROUTING"
        start_mawb_idx = page_text.find("1. MAWB: ") + len("1. MAWB: ")
        end_routing_idx = page_text.find("2. ROUTING")
        extracted_mawb_routing = page_text[start_mawb_idx:end_routing_idx].strip()

        # Extract the text between "BOOKING CONFIRMATION" and "GSA for UNITED"
        start_booking_idx = page_text.find("BOOKING CONFIRMATION") + len("BOOKING CONFIRMATION")
        end_gsa_idx = page_text.find("GSA for UNITED")
        extracted_booking_gsa = page_text[start_booking_idx:end_gsa_idx].strip()

        # Replace line breaks with a space
        extracted_mawb_routing = extracted_mawb_routing.replace('\n', ' ')
        extracted_booking_gsa = extracted_booking_gsa.replace('\n', ' ')

        return extracted_mawb_routing, extracted_booking_gsa
    finally:
        pdf_document.close()

def on_button_click():
    file_path = filedialog.askopenfilename(filetypes=[("PDF Files", "*.pdf")])
    if file_path:
        result_mawb_routing, result_booking_gsa = extract_text_from_mawb_and_booking(file_path)
        
        # Add a space before the last 4 characters in the MAWB and ROUTING result
        modified_result_mawb_routing = result_mawb_routing[:-4] + ' ' + result_mawb_routing[-4:]

        print("Result between MAWB and ROUTING:")
        print(modified_result_mawb_routing)

        print("\nResult between BOOKING CONFIRMATION and GSA for UNITED:")
        print(result_booking_gsa)

        # Format the data with ten cells between MAWB and Booking to GSA
        clipboard_text = f"{modified_result_mawb_routing}\t" + "\t".join([""] * 9) + f"\t{result_booking_gsa}"
        pyperclip.copy(clipboard_text)

        # Export both results to a new Notepad file
        with open("results.txt", "w") as file:
            file.write("Result between MAWB and ROUTING:\n")
            file.write(modified_result_mawb_routing + "\n\n")
            file.write("Result between BOOKING CONFIRMATION and GSA for UNITED:\n")
            file.write(result_booking_gsa)

        # Open the Notepad file
        subprocess.Popen(["notepad.exe", "results.txt"])

# Create the main window
root = tk.Tk()
root.title("PDF Text Extractor")

# Create a button to open the file dialog
button = tk.Button(root, text="Select PDF File", command=on_button_click)
button.pack(pady=10)

# Start the main loop
root.mainloop()
