import os
import threading
import customtkinter as ctk
from tkinter import filedialog, messagebox
import comtypes.client
import comtypes

# 1. Tech Stack & Libraries Setup
# Set the appearance mode to "System" and color theme to "blue"
ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")

class App(ctk.CTk):
    def __init__(self):
        super().__init__()

        # 2. UI Layout & Components
        self.title("Bulk PowerPoint to PDF Converter")
        self.geometry("600x400")
        self.resizable(False, False)

        # State variables
        self.selected_files = []
        self.output_folder = ""

        # Main frame to center elements vertically
        self.main_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.main_frame.pack(expand=True, fill="both", padx=40, pady=20)

        # Title Label
        self.title_label = ctk.CTkLabel(
            self.main_frame, 
            text="Bulk PowerPoint to PDF Converter", 
            font=ctk.CTkFont(size=22, weight="bold")
        )
        self.title_label.pack(pady=(0, 20))

        # Select Files Section
        self.select_files_btn = ctk.CTkButton(
            self.main_frame, 
            text="1. Select Source Folder", 
            command=self.select_source_folder
        )
        self.select_files_btn.pack(pady=(10, 5))
        
        self.files_label = ctk.CTkLabel(self.main_frame, text="0 files selected", text_color="gray")
        self.files_label.pack(pady=(0, 10))

        # Select Output Folder Section
        self.select_folder_btn = ctk.CTkButton(
            self.main_frame, 
            text="2. Select Output Folder", 
            command=self.select_output_folder
        )
        self.select_folder_btn.pack(pady=(10, 5))
        
        self.folder_label = ctk.CTkLabel(self.main_frame, text="No destination selected", text_color="gray")
        self.folder_label.pack(pady=(0, 10))

        # Start Conversion Button
        self.start_btn = ctk.CTkButton(
            self.main_frame, 
            text="3. START CONVERSION", 
            fg_color="#28a745", # Green distinct color
            hover_color="#218838",
            command=self.start_conversion,
            height=40,
            font=ctk.CTkFont(size=14, weight="bold")
        )
        self.start_btn.pack(pady=(15, 10), fill="x", padx=50)

        # Progress Bar
        self.progress_bar = ctk.CTkProgressBar(self.main_frame)
        self.progress_bar.set(0)
        self.progress_bar.pack(pady=(10, 5), fill="x", padx=50)

        # Status Label
        self.status_label = ctk.CTkLabel(self.main_frame, text="Ready", text_color="gray")
        self.status_label.pack(pady=(0, 10))

    def select_source_folder(self):
        folder = filedialog.askdirectory(title="Select Source Folder")
        if folder:
            self.selected_files = []
            for root, _, files in os.walk(folder):
                for file in files:
                    # Look for all PowerPoint formats and ignore temporary ~$ files created by PowerPoint
                    ppt_extensions = ('.ppt', '.pptx', '.pptm', '.pps', '.ppsx', '.ppsm', '.pot', '.potx', '.potm')
                    if file.lower().endswith(ppt_extensions) and not file.startswith('~$'):
                        self.selected_files.append(os.path.join(root, file))
            
            self.files_label.configure(text=f"{len(self.selected_files)} PowerPoint files found")

    def select_output_folder(self):
        folder = filedialog.askdirectory(title="Select Output Folder")
        if folder:
            self.output_folder = folder
            # Truncate long paths for the label
            display_path = folder if len(folder) < 50 else f"...{folder[-47:]}"
            self.folder_label.configure(text=display_path)

    def set_gui_state(self, state):
        """Helper to toggle button states during conversion"""
        self.select_files_btn.configure(state=state)
        self.select_folder_btn.configure(state=state)
        self.start_btn.configure(state=state)

    def start_conversion(self):
        if not self.selected_files:
            messagebox.showwarning("Missing Information", "No PowerPoint files found in the selected source folder.")
            return
        if not self.output_folder:
            messagebox.showwarning("Missing Information", "Please select a destination folder.")
            return
        
        # 3. Application Logic - UI State
        self.set_gui_state("disabled")
        self.progress_bar.set(0)
        self.status_label.configure(text="Initializing PowerPoint COM...")

        # 3. Application Logic - Threading (Daemon thread)
        thread = threading.Thread(target=self.conversion_worker, daemon=True)
        thread.start()

    def conversion_worker(self):
        """The background thread function doing the heavy lifting."""
        # Initialize COM in this thread
        comtypes.CoInitialize()
        powerpoint = None
        
        try:
            # Create PowerPoint Application object
            powerpoint = comtypes.client.CreateObject("Powerpoint.Application")
            
            total_files = len(self.selected_files)
            
            for index, file_path in enumerate(self.selected_files, start=1):
                filename = os.path.basename(file_path)
                
                # Update GUI safely from background thread
                self.after(0, lambda f=filename, i=index, t=total_files: 
                           self.status_label.configure(text=f"Converting {i} of {t}: {f}..."))
                self.after(0, lambda v=(index - 1) / total_files: self.progress_bar.set(v))

                # 3. Path Resolution (Must be absolute for COM)
                abs_input_path = os.path.abspath(file_path)
                
                base_name = os.path.splitext(filename)[0]
                output_filename = f"{base_name}.pdf"
                abs_output_path = os.path.abspath(os.path.join(self.output_folder, output_filename))

                # Open the presentation silently
                presentation = powerpoint.Presentations.Open(abs_input_path, WithWindow=False)
                
                # 3. Conversion Code (Save as PDF: format type 32)
                presentation.SaveAs(abs_output_path, 32)
                
                # Close the presentation
                presentation.Close()

            # Set progress to 100% at the end
            self.after(0, lambda: self.progress_bar.set(1.0))
            self.after(0, lambda: self.status_label.configure(text="All conversions completed successfully!"))
            
            # Show success message (scheduled on main thread)
            self.after(0, lambda: messagebox.showinfo("Success", f"Successfully converted {total_files} files!"))

        except Exception as e:
            # Handle any COM or file access errors
            error_msg = str(e)
            self.after(0, lambda: self.status_label.configure(text="Error occurred during conversion."))
            self.after(0, lambda err=error_msg: messagebox.showerror("Conversion Error", f"An error occurred:\n\n{err}"))

        finally:
            # 3. Memory Management - Crucial to clean up ghost PowerPoint processes
            if powerpoint:
                powerpoint.Quit()
            comtypes.CoUninitialize()
            
            # Re-enable the GUI buttons
            self.after(0, lambda: self.set_gui_state("normal"))

if __name__ == "__main__":
    app = App()
    app.mainloop()
