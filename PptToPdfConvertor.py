import os
import comtypes.client
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from threading import Thread

class PowerPointToPDFConverter:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("PowerPoint to PDF Converter")

        # Set window size
        self.root.geometry("450x450")  # Adjust the size as needed

        # Set window background color
        self.root.configure(bg='#69A5BF')  # Use your preferred color code
                
        # Label to display selected files
        self.selected_files_label = ttk.Label(self.root, text="Selected Files:", style='Custom.TLabel')
        self.selected_files_label.pack(side=tk.TOP, pady=(10, 0))

        # Listbox to display selected files
        self.file_listbox = tk.Listbox(self.root, selectmode=tk.MULTIPLE)
        self.file_listbox.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        # Add Files button
        self.add_file_button = ttk.Button(self.root, text="Add Files", command=self.add_files, style='Custom.TButton')
        self.add_file_button.pack(pady=5)

        # Remove Selected button
        self.remove_file_button = ttk.Button(self.root, text="Remove Selected", command=self.remove_selected_files, style='Custom.TButton')
        self.remove_file_button.pack(pady=5)

        # Convert button
        self.convert_button = ttk.Button(self.root, text="Convert", command=self.start_conversion_thread, style='Custom.TButton')
        self.convert_button.pack(pady=5)

        # Configure styles
        self.configure_styles()

        # Results variables
        self.successful_conversions = []
        self.failed_conversions = []

        # Developer label
        self.developer_label = ttk.Label(self.root, text="Developed by Muhammad Jehanzaib", style='Custom.TLabel')
        self.developer_label.pack(side=tk.BOTTOM, pady=5)

        # Progress bar
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(self.root, variable=self.progress_var, mode='determinate', style='Custom.Horizontal.TProgressbar')
        self.progress_bar.pack(fill=tk.X, pady=2)
        
        # Status label
        self.status_label = ttk.Label(self.root, text="", style='Custom.TLabel')
        self.status_label.pack()

    def configure_styles(self):
        # Configure styles for consistent appearance
        self.style = ttk.Style()

        self.style.configure('Custom.TLabel', font=('Arial', 12), background='#69A5BF', foreground='black')
        self.style.configure('Custom.TButton', font=('Arial', 10), background='#2F65C6', foreground='black')
        
        # Configure the TProgressbar style
        self.style.configure('TProgressbar',thickness=20,
                            troughcolor='#69A5BF',
                            background='#2F65C6',
                            troughrelief='flat')

        # Configure the Custom.TProgressbar style
        self.style.configure('Custom.Horizontal.TProgressbar', troughcolor='#69A5BF', background='#2F65C6')

    def add_files(self):
        files = filedialog.askopenfilenames(filetypes=[("PowerPoint files", "*.ppt;*.pptx")])
        for file_path in files:
            self.file_listbox.insert(tk.END, file_path)

    def remove_selected_files(self):
        selected_indices = self.file_listbox.curselection()
        for i in reversed(selected_indices):
            self.file_listbox.delete(i)

    def start_conversion_thread(self):
        # Confirm with the user before starting the conversion
        confirmed = messagebox.askokcancel("Confirmation", "Are you sure you want to convert the selected files?")
        if confirmed:
            # Start a new thread for conversion
            thread = Thread(target=self.convert_files_threaded)
            thread.start()

    def convert_pptx_to_pdf(self, input_pptx_path, output_pdf_path, custom_output_name=None):
        try:
            # Convert folder paths to Windows format
            input_folder_path = os.path.abspath(input_pptx_path)
            output_folder_path = os.path.abspath(output_pdf_path)

            # Check if the input file exists
            if not os.path.exists(input_pptx_path):
                raise FileNotFoundError(f"File not found: {input_pptx_path}")

            # Create PowerPoint application object
            powerpoint = comtypes.client.CreateObject("Powerpoint.Application")

            # Set visibility to minimize
            powerpoint.Visible = 1

            # Open the PowerPoint slides
            slides = powerpoint.Presentations.Open(input_pptx_path)

            # Get base file name
            file_name = os.path.splitext(os.path.basename(input_pptx_path))[0]

            # Create output file path with the inputted file name
            output_file_path = os.path.join(output_folder_path, file_name + ".pdf")

            # Save as PDF (formatType = 32)
            slides.SaveAs(output_file_path, 32)

            # Close the slide deck
            slides.Close()

            return output_file_path  # Return the output file path on success

        except Exception as e:
            # Return None on failure
            return None

        finally:
            # Quit PowerPoint application
            powerpoint.Quit()

    def convert_files_threaded(self):
        try:
            self.progress_var.set(0)  # Reset progress bar
            self.progress_bar["maximum"] = 100
            self.progress_bar.start()  # Start determinate progress bar
            self.set_status("Converting files. Please wait...")

            output_folder = filedialog.askdirectory()
            total_files = self.file_listbox.size()

            # Adjusted to handle the case where no files are selected
            if total_files == 0:
                self.show_warning("No Files", "Please add PowerPoint files to convert.")
                return

            failed_files = []

            for index, input_file in enumerate(self.file_listbox.get(0, tk.END), 1):
                try:
                    if input_file and os.path.exists(input_file):
                        input_file = os.path.normpath(input_file)

                        # Start individual progress bar for each file
                        self.progress_bar["maximum"] = 100
                        self.progress_var.set(0)
                        self.progress_bar.start()
                        self.set_status(f"Converting file {index} of {total_files}")

                        output_file_path = self.convert_pptx_to_pdf(input_file, output_folder)

                        if output_file_path:
                            # Conversion successful
                            self.progress_var.set(100)  # Set progress to 100% for individual file
                            self.root.update_idletasks()
                            self.successful_conversions.append(f"File {index}/{total_files}: {input_file} -> {output_file_path}")
                        else:
                            # Conversion failed
                            failed_files.append(f"File {index}/{total_files}: {input_file}")
                            # print(f"File {index}/{total_files}: {input_file} - Conversion failed.")

                except Exception as e:
                    failed_files.append(f"File {index}/{total_files}: {input_file} - Error: {e}")

            # Clear the selected files list
            self.file_listbox.delete(0, tk.END)

            # Display detailed results
            self.display_results(successful_conversions=self.successful_conversions, failed_files=failed_files)

        except Exception as e:
            self.show_error("Error", f"Error converting files: {e}")

        finally:
            self.progress_bar.stop()
            self.progress_var.set(100)
            self.set_status("Conversion completed.")

    def display_results(self, successful_conversions, failed_files):
        result_message = "Conversion Results:\n\n"

        if successful_conversions:
            result_message += "Successful Conversions:\n"
            result_message += "\n".join(successful_conversions) + "\n\n"

        if failed_files:
            result_message += "Failed Conversions:\n"
            result_message += "\n".join(failed_files)

        # Show results in messagebox
        self.show_info("Conversion Results", result_message)

    def show_info(self, title, message):
        messagebox.showinfo(title, message)

    def show_warning(self, title, message):
        messagebox.showwarning(title, message)

    def show_error(self, title, message):
        messagebox.showerror(title, message)

    def set_status(self, message):
        self.status_label.config(text=message)

    def run(self):
        self.root.mainloop()

if __name__ == "__main__":
    converter = PowerPointToPDFConverter()
    converter.run()
