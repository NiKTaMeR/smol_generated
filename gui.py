import tkinter as tk
from tkinter import filedialog, messagebox
import data_ingestion
import data_parsing
import mrr_per_client
import mrr_status

class SaaSMetricsCalculatorGUI(tk.Tk):
    def __init__(self):
        super().__init__()

        self.title("SaaS Metrics Calculator")
        self.geometry("400x300")

        self.create_widgets()

    def create_widgets(self):
        self.ingest_data_button = tk.Button(self, text="Ingest Data", command=self.ingest_data)
        self.ingest_data_button.pack(pady=10)

        self.parse_information_button = tk.Button(self, text="Parse Information", command=self.parse_information)
        self.parse_information_button.pack(pady=10)

        self.create_mrr_per_client_button = tk.Button(self, text="Create MRR Per Client Table", command=self.create_mrr_per_client_table)
        self.create_mrr_per_client_button.pack(pady=10)

        self.create_mrr_status_button = tk.Button(self, text="Create MRR Status Table", command=self.create_mrr_status_table)
        self.create_mrr_status_button.pack(pady=10)

    def ingest_data(self):
        file_types = [("Excel files", "*.xlsx"), ("Access Database files", "*.accdb")]
        file_path = filedialog.askopenfilename(filetypes=file_types)

        if file_path:
            data_ingestion.ingest_data(file_path)
            messagebox.showinfo("Success", "Data ingested successfully.")
        else:
            messagebox.showerror("Error", "No file selected.")

    def parse_information(self):
        success, message = data_parsing.parse_information()

        if success:
            messagebox.showinfo("Success", message)
            self.after(1000, self.wait_for_user_input)
        else:
            messagebox.showerror("Error", message)

    def wait_for_user_input(self):
        user_input = messagebox.askokcancel("User Input", "I'm done with the results, go ahead.")
        if user_input:
            data_parsing.close_excel_file()
        else:
            self.after(1000, self.wait_for_user_input)

    def create_mrr_per_client_table(self):
        success, message = mrr_per_client.create_mrr_per_client_table()

        if success:
            messagebox.showinfo("Success", "MRR Per Client table created successfully.")
        else:
            messagebox.showerror("Error", message)

    def create_mrr_status_table(self):
        success, message = mrr_status.create_mrr_status_table()

        if success:
            messagebox.showinfo("Success", "MRR Status table created successfully.")
        else:
            messagebox.showerror("Error", message)

if __name__ == "__main__":
    app = SaaSMetricsCalculatorGUI()
    app.mainloop()