import os
import pandas as pd
import tkinter
from tkinter import filedialog
import customtkinter
import spacy
from pdfminer.high_level import extract_text
import re
from docx2python import docx2python
from typing import TypedDict, List


class ResultDict(TypedDict):
    names: List
    phones: List
    emails: List
    skills: List


# Main component
class App(customtkinter.CTk):
    def __init__(self) -> None:
        super().__init__()
        
        # language model
        self.nlp = spacy.load("en_core_web_sm")

        # dictionary with results
        self.result: ResultDict = {
            "names": [],
            "phones": [],
            "emails": [],
            "skills": []
        }
        
        # folders
        self.input_folder_path = tkinter.StringVar(self, value="Choose the folder with CVs:")
        self.output_folder_path = tkinter.StringVar(self, value="Choose the output folder:")
        
        # GUI setup
        self.title("CV-Parser")
        self.geometry("680x420")
        self.minsize(720, 600)
        self.maxsize(1000, 620)
                
        self.grid_rowconfigure((0, 1, 2, 3, 4, 5), weight=0)
        self.grid_columnconfigure((0, 1), weight=1)
        
        self.heading = customtkinter.CTkLabel(self, text="Choose input and output folders\n then enter a comma separated list of skills and click 'Start'", font=("sans-serif", 22))
        self.heading.grid(row=0, column=0, columnspan=2, padx=20, pady=(50, 0))
                
        self.input_choice_frame = ChoiceFrame(self, choice_func=self.handle_input, header_text=self.input_folder_path)
        self.input_choice_frame.grid(row=1, column=0, padx=10, pady=(50, 20))
        
        self.output_choice_frame = ChoiceFrame(self, choice_func=self.handle_output, header_text=self.output_folder_path)
        self.output_choice_frame.grid(row=1, column=1, padx=10, pady=(50, 20))
        
        self.skills_list = tkinter.StringVar(self)
        self.skills_entry = customtkinter.CTkEntry(self, font=("sans-serif", 17), textvariable=self.skills_list, width=400)
        self.skills_entry.grid(row=2, column=0, columnspan=2, pady=(50, 0))
        
        self.start_button = customtkinter.CTkButton(self, height=40, text="Start", font=("sans-serif", 17, "bold"), fg_color="#146C27", hover_color="#18832F", command=self.start_parse)
        self.start_button.grid(row=3, column=0, columnspan=2, pady=(30, 0))
        
        self.info_label = customtkinter.CTkLabel(self, text="", font=("sans-serif", 18), text_color="#18832F")
        self.info_label.grid(row=4, column=0, columnspan=2, pady=(30, 0))
        
        self.reset_btn = customtkinter.CTkButton(self, text="", font=("sans-serif", 17, "bold"), fg_color="transparent", border_width=0, text_color="white", height=0, width=0, hover_color="#111111", command=self.start_over)
        self.reset_btn.grid(row=5, column=0, columnspan=2, pady=(40, 0))

    def handle_input(self):
        folder = filedialog.askdirectory(initialdir="/", title="Select Folder")
        self.input_folder_path.set(folder or "Choose the folder with CVs:")

    def handle_output(self):
        folder = filedialog.askdirectory(initialdir="/", title="Select Folder")
        self.output_folder_path.set(folder or "Choose the output folder:")
        
    def convert(self, file):
        if file.endswith(".docx"):
            text_context = docx2python(file)
            text = text_context.text
            text_context.close()
            return text
        elif file.endswith(".pdf"):
            text = extract_text(file)
            return text
        else:
            return ""
    
    def parse_text(self, text, skills):
        skills_list = skills.split(",")
        skills_list = re.compile("|".join(skills_list))
        phone_num = re.compile("(\d{3}[-\.\s]??\d{3}[-\.\s]??\d{4}|\(\d{3}\)\s*\d{3}[-\.\s]??\d{4}|\d{3}[-\.\s]??\d{4})")
        found_skills = set(re.findall(skills_list, text.lower()))
        if len(found_skills) > 0:
            doc = self.nlp(text)
            name = [node.text for node in doc.ents if node.label_ == "PERSON"]
            email = [node for node in doc if node.like_email is True]
            phone = re.findall(phone_num, text.lower())
            self.result["names"].append(name[0] if len(name) > 0 else None)
            self.result["emails"].append(email[0] if len(email) > 0 else None)
            self.result["phones"].append(phone)
            self.result["skills"].append(found_skills)
        else:
            return
        
    def start_parse(self):
        skills = self.skills_list.get()
        folder = self.input_folder_path.get()
        self.start_button.configure(text="Parsing...")
        if len(skills) > 0 and self.output_folder_path.get() != "Choose the output folder:":
            for file in os.listdir(folder):
                file_path = os.path.join(folder, file)
                text = self.convert(file_path)
                self.parse_text(text, skills.lower())
            self.start_button.configure(text="Done!")        
            self.info_label.configure(text="Check the output folder")
            self.reset_btn.configure(text="Start Over", width=150, height=40)
            result_df = pd.DataFrame(self.result)
            result_df.to_csv(os.path.join(self.output_folder_path.get(), "parsed.csv"))
        else:
            self.start_button.configure(text="Start")
            return
        self.skills_list.set("")
        
    def start_over(self):
        self.input_folder_path.set("Choose the folder with CVs:")
        self.output_folder_path.set("Choose the output folder:")
        self.start_button.configure(text="Start")
        self.info_label.configure(text="")
        self.reset_btn.configure(text="", width=0, height=0)
        return
        
    
# Frame component
class ChoiceFrame(customtkinter.CTkFrame):
    def __init__(self, *args, choice_func=None, header_text="Choice", **kwargs):
        super().__init__(*args, **kwargs)
        self.header_text = header_text
        self.header = customtkinter.CTkLabel(self, textvariable=self.header_text, font=("sans-serif", 18))
        self.header.grid(row=0, column=0, padx=10, pady=(10, 0))
        self.button = customtkinter.CTkButton(self, height=35, text="Open", font=("sans-serif", 17), command=choice_func)
        self.button.grid(row=1, column=0, pady=20, padx=20)
        

customtkinter.set_appearance_mode("System")
customtkinter.set_default_color_theme("dark-blue")
app = App()
app.mainloop()


