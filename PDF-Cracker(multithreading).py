import itertools
import string
import zipfile
import msoffcrypto
import io
import PyPDF2
import tkinter as tk
from tkinter import filedialog, messagebox
from concurrent.futures import ThreadPoolExecutor, as_completed

def try_password(password, file_type, file_path):
    password = "".join(password)
    print(f"Trying: {password}")
    
    if file_type == "zip":
        try:
            with zipfile.ZipFile(file_path) as zip_file:
                with zip_file.open(zip_file.namelist()[0], 'r', pwd=password.encode('utf-8')):
                    if zip_file.testzip() is None:
                        print(f"Password found: {password}")
                        messagebox.showinfo("Success", f"Password found: {password}")
                        return password
        except:
            return None
    
    elif file_type == "word":
        try:
            with open(file_path, "rb") as file:
                office_file = msoffcrypto.OfficeFile(file)
                office_file.load_key(password=password)
                decrypted = io.BytesIO()
                office_file.decrypt(decrypted)
                if decrypted.getvalue():
                    print(f"Password found: {password}")
                    messagebox.showinfo("Success", f"Password found: {password}")
                    return password
        except:
            return None
    
    elif file_type == "pdf":
        try:
            with open(file_path, "rb") as pdf_file:
                pdf_reader = PyPDF2.PdfReader(pdf_file)
                if pdf_reader.decrypt(password):
                    text = pdf_reader.pages[0].extract_text()
                    if text:
                        print(f"Password found: {password}")
                        messagebox.showinfo("Success", f"Password found: {password}")
                        return password
        except:
            return None
    
    return None

def unlock_file(file_type, file_path):
    chars = string.ascii_letters + string.digits + string.punctuation  # All possible characters
    max_workers = 8  # Adjust based on CPU cores
    
    with ThreadPoolExecutor(max_workers=max_workers) as executor:
        for length in range(1, 10):  # Adjust max length as needed
            passwords = itertools.product(chars, repeat=length)
            future_to_password = {executor.submit(try_password, pw, file_type, file_path): pw for pw in passwords}
            
            for future in as_completed(future_to_password):
                result = future.result()
                if result:
                    return result
    
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
        unlock_file("zip", file_path)
    elif file_path.endswith(".docx"):
        unlock_file("word", file_path)
    elif file_path.endswith(".pdf"):
        unlock_file("pdf", file_path)
    else:
        messagebox.showerror("Error", "Unsupported file type!")

if __name__ == "__main__":
    select_file()
