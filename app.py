import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk
import os
import threading
import core

class Application(tk.Tk):
    def __init__(self):
        super().__init__()

        self.title("4G KPI Report Generator")
        self.geometry("600x400")
        self.configure(padx=20, pady=20)

        # UI Variables
        self.raw_file_path = tk.StringVar()
        self.template_file_path = tk.StringVar()
        self.event_name = tk.StringVar()
        self.status_var = tk.StringVar(value="Ready")

        self.create_widgets()

    def create_widgets(self):
        style = ttk.Style()
        style.configure('TLabel', font=('Arial', 11))
        style.configure('TButton', font=('Arial', 10))

        # Title
        title_label = ttk.Label(self, text="4G KPI Report Generator", font=('Arial', 16, 'bold'))
        title_label.pack(pady=(0, 20))

        # File Selection Frame
        file_frame = ttk.Frame(self)
        file_frame.pack(fill='x', pady=10)

        ttk.Label(file_frame, text="1. Select Raw Data File (Excel):").pack(anchor='w')

        file_input_frame = ttk.Frame(file_frame)
        file_input_frame.pack(fill='x', pady=5)

        file_entry = ttk.Entry(file_input_frame, textvariable=self.raw_file_path, state='readonly', width=50)
        file_entry.pack(side='left', expand=True, fill='x', padx=(0, 10))

        browse_btn = ttk.Button(file_input_frame, text="Browse", command=self.browse_file)
        browse_btn.pack(side='right')


        # Template Selection Frame
        template_frame = ttk.Frame(self)
        template_frame.pack(fill='x', pady=10)

        ttk.Label(template_frame, text="2. Select Template File (Optional):").pack(anchor='w')

        template_input_frame = ttk.Frame(template_frame)
        template_input_frame.pack(fill='x', pady=5)

        template_entry = ttk.Entry(template_input_frame, textvariable=self.template_file_path, state='readonly', width=50)
        template_entry.pack(side='left', expand=True, fill='x', padx=(0, 10))

        template_browse_btn = ttk.Button(template_input_frame, text="Browse", command=self.browse_template_file)
        template_browse_btn.pack(side='right')

        # Event Name Frame
        event_frame = ttk.Frame(self)
        event_frame.pack(fill='x', pady=10)

        ttk.Label(event_frame, text="3. Enter Event Name:").pack(anchor='w')
        event_entry = ttk.Entry(event_frame, textvariable=self.event_name, width=50)
        event_entry.pack(anchor='w', pady=5)

        # Generate Button
        self.generate_btn = ttk.Button(self, text="Generate Report", command=self.start_generation, style='Accent.TButton')
        self.generate_btn.pack(pady=30)

        # Status Label
        status_label = ttk.Label(self, textvariable=self.status_var, foreground="gray")
        status_label.pack(side='bottom', pady=10)

    def browse_file(self):
        filename = filedialog.askopenfilename(
            title="Select Raw Data File",
            filetypes=(("Excel files", "*.xlsx *.xls"), ("All files", "*.*"))
        )
        if filename:
            self.raw_file_path.set(filename)


    def browse_template_file(self):
        filename = filedialog.askopenfilename(
            title="Select Template File",
            filetypes=(("Excel files", "*.xlsx *.xls"), ("All files", "*.*"))
        )
        if filename:
            self.template_file_path.set(filename)

    def start_generation(self):
        raw_file = self.raw_file_path.get()
        event_name = self.event_name.get().strip()
        template_file = self.template_file_path.get()

        if not raw_file:
            messagebox.showerror("Error", "Please select a raw data file.")
            return

        if not event_name:
            messagebox.showerror("Error", "Please enter an event name.")
            return

        # Disable button and update status
        self.generate_btn.config(state='disabled')
        self.status_var.set("Processing... Please wait.")
        self.update()

        # Run logic in a separate thread so UI doesn't freeze
        threading.Thread(target=self.generate_report_thread, args=(raw_file, template_file, event_name), daemon=True).start()

    def generate_report_thread(self, raw_file, template_file, event_name):
        try:
            output_path = core.generate_report(raw_file, template_file, event_name)
            self.after(0, self.generation_success, output_path)
        except Exception as e:
            self.after(0, self.generation_error, str(e))

    def generation_success(self, output_path):
        self.generate_btn.config(state='normal')
        self.status_var.set("Done!")
        msg = f"Report successfully generated!\n\nSaved to your Downloads folder at:\n{output_path}"
        messagebox.showinfo("Success", msg)

    def generation_error(self, error_msg):
        self.generate_btn.config(state='normal')
        self.status_var.set("Error occurred.")
        messagebox.showerror("Error", f"Failed to generate report:\n{error_msg}")

if __name__ == "__main__":
    app = Application()
    app.mainloop()
