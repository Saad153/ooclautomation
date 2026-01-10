

import customtkinter
import threading
from tkinter import filedialog, messagebox
from automation import start_automation_process, read_excel_all_records
import time

customtkinter.set_appearance_mode("light")
customtkinter.set_default_color_theme('blue')

def run_automation(shipment_path, dmf_path, status_label):
    try:
        status_label.configure(text="Reading files...")
        shipment_records = read_excel_all_records(shipment_path, sheet_name="Sheet1")
        dmf_records = read_excel_all_records(dmf_path, sheet_name="Master")
        status_label.configure(text="Starting automation...")
        start_automation_process(shipment_records, dmf_records)
        status_label.configure(text="Automation completed.")
    except Exception as e:
        status_label.configure(text=f"Error: {e}")
        messagebox.showerror("Error", str(e))

def start_gui():
    root = customtkinter.CTk()
    root.title('OOCL Automation')
    root.geometry("500x400")

    # Title label
    label = customtkinter.CTkLabel(master=root,
                                text="OOCL Automation",
                                font=('Helvetica', 20))
    label.place(relx=0.5, rely=0.08, anchor='n')

    # Shipment file selector
    shipment_entry = customtkinter.CTkEntry(master=root, width=300, height=25, placeholder_text="Shipment Plan File")
    shipment_entry.place(relx=0.38, rely=0.18, anchor='n')

    def browse_shipment():
        path = filedialog.askopenfilename(title="Select Shipment Plan File", filetypes=[("Excel Files", "*.xlsx;*.xls")])
        if path:
            shipment_entry.delete(0, 'end')
            shipment_entry.insert(0, path)

    shipment_browse = customtkinter.CTkButton(master=root, text="Browse", command=browse_shipment, width=120, height=25, border_width=0, corner_radius=8)
    shipment_browse.place(relx=0.82, rely=0.18, anchor='n')

    # DMF file selector
    dmf_entry = customtkinter.CTkEntry(master=root, width=300, height=25, placeholder_text="DMF File")
    dmf_entry.place(relx=0.38, rely=0.28, anchor='n')

    def browse_dmf():
        path = filedialog.askopenfilename(title="Select DMF File", filetypes=[("Excel Files", "*.xlsx;*.xls")])
        if path:
            dmf_entry.delete(0, 'end')
            dmf_entry.insert(0, path)

    dmf_browse = customtkinter.CTkButton(master=root, text="Browse", command=browse_dmf, width=120, height=25, border_width=0, corner_radius=8)
    dmf_browse.place(relx=0.82, rely=0.28, anchor='n')

    # Status label
    status_label = customtkinter.CTkLabel(master=root, text="Idle", text_color="blue")
    status_label.place(relx=0.5, rely=0.7, anchor='n')

    def on_start():
        shipment_path = shipment_entry.get()
        dmf_path = dmf_entry.get()
        if not shipment_path or not dmf_path:
            messagebox.showwarning("Missing File", "Please select both files.")
            return
        status_label.configure(text="Running automation...")
        threading.Thread(target=run_automation, args=(shipment_path, dmf_path, status_label), daemon=True).start()

    execute_button = customtkinter.CTkButton(master=root,
                                    text="Start Automation",
                                    command=on_start,
                                    width=430,
                                    height=35,
                                    border_width=0,
                                    corner_radius=8,
                                    fg_color="#4CAF50",
                                    text_color="white")
    execute_button.place(relx=0.5, rely=0.45, anchor='n')

    root.mainloop()

try:
    start_gui()
except Exception as e:
    print(e)
    while True:
        pass