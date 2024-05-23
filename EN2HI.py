import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
import pyodbc

class ExcelHindiNameUpdater:
    def __init__(self, root):
        self.root = root
        self.root.title("Ahmad - Excel Hindi Name Updater (Simple)")
        self.root.geometry("500x400")
        self.center_window(750, 550)
        self.root.resizable(False, False)  # Prevent window from being resized
        
        self.style = ttk.Style()
        self.style.theme_use('clam')
        
        self.frame = ttk.Frame(root, padding="20")
        self.frame.pack(fill=tk.BOTH, expand=True)

        self.title_label = ttk.Label(self.frame, text="Ahmad - Excel Hindi Name Updater", font=("Helvetica", 16))
        self.title_label.pack(pady=10)

        self.upload_button = ttk.Button(self.frame, text="Upload Excel", command=self.upload_excel)
        self.upload_button.pack(pady=10)

        self.log_label = ttk.Label(self.frame, text="Log:", font=("Helvetica", 10))
        self.log_label.pack(pady=10)

        self.log_text = tk.Text(self.frame, height=10, wrap="word",font=("Lucida Console", 10))
        self.log_text.pack(fill=tk.BOTH, expand=True)

        self.status_label = ttk.Label(self.frame, text="", font=("Helvetica", 10))
        self.status_label.pack(pady=10)

    def center_window(self, width, height):
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        x = (screen_width - width) // 2
        y = (screen_height - height) // 2
        self.root.geometry(f'{width}x{height}+{x}+{y}')

    def log(self, message):
        self.log_text.insert(tk.END, message + "\n")
        self.log_text.see(tk.END)
        self.root.update()

    def upload_excel(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xls;*.xlsx")])
        if file_path:
            confirm = messagebox.askyesno("Confirm", "Do you want to process this file?")
            if confirm:
                self.update_excel_with_hindi_names(file_path)
                
    def update_excel_with_hindi_names(self, file_path):
        try:
            self.status_label.config(text="Processing...")
            self.log("Reading Excel file...")
            df = pd.read_excel(file_path)
            if 'EnglishName' not in df.columns:
                messagebox.showerror("Error", "No column named 'EnglishName' found in the Excel file.")
                self.status_label.config(text="")
                return
            self.log("Excel file read successfully.")
            df = df.where(pd.notnull(df), None)
            self.log("Updating Hindi names...")
            df['HindiName'] = df['EnglishName'].apply(self.get_hindi_name_from_database)
            self.log("Saving updated Excel file...")
            df.to_excel(file_path, index=False)
            messagebox.showinfo("Success", "Excel file updated successfully!")
            self.status_label.config(text="File updated successfully!")
            self.log("Excel file updated and saved successfully.")
        except Exception as e:
            messagebox.showerror("Error", str(e))
            self.status_label.config(text="Error occurred!")
            self.log(f"Error: {str(e)}")

    def get_hindi_name_from_database(self, english_name):
        connection_string = (
            "DRIVER={ODBC Driver 13 for SQL Server};"
            "SERVER='Your Server Name'; "
            "DATABASE= 'Your Database Name'; "
            "UID='Your User ID';"
	    "PWD='Your Password';"
        )

        try:
            connection = pyodbc.connect(connection_string)
            cursor = connection.cursor()
            cursor.execute("SELECT name_hin FROM krutidev WHERE name = ?", (english_name,))
            result = cursor.fetchone()
            connection.close()
            if result:
                self.log(f"Found Hindi name for '{english_name}'")
                return result[0]
            else:
                self.log(f"No Hindi name found for '{english_name}'")
                return "Sorry No Data"  # If no Hindi name is found, return this message
        except pyodbc.Error as e:
            self.log(f"Database error: {e}")
            return f"Database error: {e}"  # If an error occurs, return this message

if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelHindiNameUpdater(root)
    root.mainloop()
