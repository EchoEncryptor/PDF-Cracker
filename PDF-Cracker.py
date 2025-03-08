import itertools
import string
import zipfile
import msoffcrypto
import io
import PyPDF2
import tkinter as tk
from tkinter import filedialog, messagebox

def unlock_zip(zip_path):
    chars = string.ascii_letters + string.digits + string.punctuation  # All possible characters
    
    with zipfile.ZipFile(zip_path) as zip_file:
        for length in range(1, 10):  # Adjust max length as needed
            for password in itertools.product(chars, repeat=length):
                password = "".join(password).encode('utf-8')
                print(f"Trying: {password.decode()}")
                try:
                    with zip_file.open(zip_file.namelist()[0], 'r', pwd=password):
                        if zip_file.testzip() is None:
                            print(f"Password found: {password.decode()}")
                            messagebox.showinfo("Success", f"Password found: {password.decode()}")
                            return password.decode()
                except:
                    continue
    
    print("Password not found within given length!")
    messagebox.showerror("Failed", "Password not found within given length!")
    return None

def unlock_word(doc_path):
    chars = string.ascii_letters + string.digits + string.punctuation
    
    with open(doc_path, "rb") as file:
        office_file = msoffcrypto.OfficeFile(file)
        
        for length in range(1, 10):  # Adjust max length as needed
            for password in itertools.product(chars, repeat=length):
                password = "".join(password)
                print(f"Trying: {password}")
                try:
                    office_file.load_key(password=password)
                    decrypted = io.BytesIO()
                    office_file.decrypt(decrypted)
                    if decrypted.getvalue():
                        print(f"Password found: {password}")
                        messagebox.showinfo("Success", f"Password found: {password}")
                        return password
                except:
                    continue
    
    print("Password not found within given length!")
    messagebox.showerror("Failed", "Password not found within given length!")
    return None

def unlock_pdf(pdf_path):
    chars = string.ascii_letters + string.digits + string.punctuation
    
    with open(pdf_path, "rb") as pdf_file:
        pdf_reader = PyPDF2.PdfReader(pdf_file)
        
        for length in range(1, 10):  # Adjust max length as needed
            for password in itertools.product(chars, repeat=length):
                password = "".join(password)
                print(f"Trying: {password}")
                if pdf_reader.decrypt(password):
                    try:
                        text = pdf_reader.pages[0].extract_text()
                        if text:
                            print(f"Password found: {password}")
                            messagebox.showinfo("Success", f"Password found: {password}")
                            return password
                    except:
                        pass
    
    print("Password not found within given length!")
    messagebox.showerror("Failed", "Password not found within given length!")
    return None

def select_file():
    root = tk.Tk()
    root.withdraw()
    file_path = filedialog.askopenfilename(filetypes=[("All Files", "*.*"), ("ZIP Files", "*.zip"), ("Word Files", "*.docx"), ("PDF Files", "*.pdf")])
    if not file_path:
        messagebox.showerror("Error", "No file selected!")
        return
    
    if file_path.endswith(".zip"):
        unlock_zip(file_path)
    elif file_path.endswith(".docx"):
        unlock_word(file_path)
    elif file_path.endswith(".pdf"):
        unlock_pdf(file_path)
    else:
        messagebox.showerror("Error", "Unsupported file type!")

if __name__ == "__main__":
    select_file()