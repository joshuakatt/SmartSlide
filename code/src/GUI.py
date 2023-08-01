import tkinter as tk
from tkinter import filedialog, messagebox

class Application(tk.Frame):
    def __init__(self, master=None):
        super().__init__(master)
        self.master = master
        self.pack()
        self.create_widgets()

    def update_log(self, text):
        # Insert the text at the end and then see it
        self.log_text.insert(tk.END, text + '\n')
        self.log_text.see(tk.END)

    def create_widgets(self):
        self.browse_ppt = tk.Button(self)
        self.browse_ppt["text"] = "Browse PPT"
        self.browse_ppt.pack(side="top")

        self.end_presentation = tk.Button(self)
        self.end_presentation["text"] = "End Presentation"
        self.end_presentation["command"] = self.end_ppt
        self.end_presentation.pack(side="top")

        self.slide_label = tk.Label(self, text="Current Slide: 0")
        self.slide_label.pack(side="top")

        self.status_label = tk.Label(self, text="Status: Not Listening")
        self.status_label.pack(side="top")

        self.clear_button = tk.Button(self)
        self.clear_button["text"] = "Clear"
        self.clear_button["command"] = self.clear_log
        self.clear_button.pack(side="top")

        self.log_text = tk.Text(self)
        self.log_text.pack(side="top")

    def end_ppt(self):
        self.status_label["text"] = "Status: Presentation Ended"

    def clear_log(self):
        self.log_text.delete(1.0, tk.END)

    def update_log(self, text):
        self.log_text.insert(tk.END, text + '\n')

    def update_slide(self, slide_number):
        self.slide_label["text"] = f"Current Slide: {slide_number}"
