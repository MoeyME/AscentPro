import tkinter as tk
from tkinter import messagebox, filedialog, simpledialog, scrolledtext, ttk
import json
import os
import sys
from datetime import datetime
import traceback
import logging
import csv

# Set up logging
logging.basicConfig(filename='app.log', level=logging.DEBUG, 
                    format='%(asctime)s - %(levelname)s - %(message)s')

# Log system information
logging.info(f"Python version: {sys.version}")
logging.info(f"Tkinter version: {tk.TkVersion}")
logging.info(f"Operating System: {os.name}")

# Global exception handler
def global_exception_handler(exc_type, exc_value, exc_traceback):
    logging.error("Uncaught exception", exc_info=(exc_type, exc_value, exc_traceback))
    messagebox.showerror("Critical Error", f"An unexpected error occurred: {exc_value}\n\nPlease check the app.log file for more details.")

sys.excepthook = global_exception_handler

import tkinter.ttk as ttk
from tkinter.constants import *
USE_TTKBOOTSTRAP = False
logging.info("Using standard ttk")

logging.info("Finished importing modules")

class AscentPro:
    def __init__(self, master):
        logging.info("Initializing AscentPro")
        try:
            self.master = master
            self.master.title("AscentPro - Career Progression Tracker")
            self.master.geometry("1400x900")
            logging.info("Master window initialized")

            if USE_TTKBOOTSTRAP:
                logging.info("Attempting to set ttkbootstrap style")
                self.style = ttk.Style("darkly")
                logging.info("ttkbootstrap style set successfully")
            else:
                logging.info("Setting standard ttk style")
                self.style = ttk.Style()
                self.style.theme_use('clam')
                logging.info("Standard ttk style set successfully")

            # Data structures
            logging.info("Initializing data structures")
            self.initialize_data_structures()
            logging.info("Data structures initialized successfully")

            # Initialize StringVar instances for each skill type's category and subskill
            self.category_vars = {
                "Technical Skills": tk.StringVar(),
                "Soft Skills": tk.StringVar(),
                "Software Skills": tk.StringVar(),
                "Certifications": tk.StringVar()  # Add this line
            }
            self.subskill_vars = {
                "Technical Skills": tk.StringVar(),
                "Soft Skills": tk.StringVar(),
                "Software Skills": tk.StringVar(),
                "Certifications": tk.StringVar()  # Add this line
            }

            # Get the directory of the script
            script_dir = os.path.dirname(os.path.abspath(__file__))
            logging.info(f"Script directory: {script_dir}")
            
            # Set full paths for data and config files
            self.data_file = os.path.join(script_dir, "career_progression_data.json")
            self.config_file = os.path.join(script_dir, "config.json")
            
            logging.info(f"Data file will be saved at: {self.data_file}")
            logging.info(f"Config file will be saved at: {self.config_file}")

            # UI components
            logging.info("Initializing UI components")
            self.initialize_ui_components()
            logging.info("UI components initialized successfully")

            # Load configuration and data
            logging.info("Loading configuration")
            self.load_config()
            logging.info("Loading data")
            self.load_data()
            logging.info("Configuration and data loaded successfully")

            # Create GUI components
            logging.info("Creating widgets")
            self.create_widgets()
            logging.info("Widgets created successfully")
            
            logging.info("Refreshing member dropdown")
            self.refresh_member_dropdown()  # Populate the member dropdown
            logging.info("Member dropdown refreshed successfully")
            
        except Exception as e:
            logging.error(f"Error in AscentPro initialization: {str(e)}")
            logging.exception("Detailed traceback:")
            raise

    def initialize_data_structures(self):
        self.team_members = {}
        self.skills_data = {
            "Technical Skills": [],
            "Soft Skills": [],
            "Software Skills": [],
            "Certifications": []
        }
        self.categories = {
            "Technical Skills": {},
            "Soft Skills": {},
            "Software Skills": {},
            "Certifications": {}
        }
        self.category_vars = {}
        self.subskill_vars = {}

    def initialize_ui_components(self):
        self.category_vars = {}
        self.category_dropdowns = {}
        self.subskill_vars = {}
        self.subskill_dropdowns = {}
        self.skills_listboxes = {}
        self.member_var = tk.StringVar()
        self.member_dropdown = None
        self.software_skills_listbox = None
        self.technical_skills_category_dropdown = None
        self.soft_skills_category_dropdown = None
        logging.info("UI components initialized")

    def load_config(self):
        logging.info("Loading configuration")
        try:
            with open(self.config_file, 'r') as config_file:
                config = json.load(config_file)
                if 'window_size' in config:
                    self.master.geometry(config['window_size'])
        except FileNotFoundError:
            logging.warning("Config file not found. Using default settings.")
        except json.JSONDecodeError:
            logging.error("Error reading config file. Using default settings.")

    def save_config(self):
        logging.info("Saving configuration")
        config = {
            'window_size': self.master.geometry()
        }
        try:
            with open(self.config_file, 'w') as config_file:
                json.dump(config, config_file, indent=2)
            logging.info("Configuration saved successfully.")
        except Exception as e:
            logging.error(f"Error saving config: {str(e)}")
            messagebox.showerror("Save Error", f"Unable to save configuration: {str(e)}")

    def load_data(self):
        logging.info("Loading data")
        try:
            with open(self.data_file, 'r') as file:
                data = json.load(file)
                self.team_members = data.get('team_members', {})
                self.skills_data = data.get('skills_data', {
                    "Technical Skills": {},
                    "Soft Skills": {},
                    "Software Skills": [],
                    "Certifications": []
                })
                self.subskills_data = data.get('subskills_data', {
                    "Technical Skills": {},
                    "Soft Skills": {}
                })
                self.meetings = data.get('meetings', [])
                # Convert old format of skills if necessary
                for member in self.team_members.values():
                    for skill_type in ["technical_skills", "soft_skills"]:
                        if isinstance(member.get(skill_type), dict):
                            new_skills = []
                            for category, skills in member[skill_type].items():
                                for skill in skills:
                                    new_skills.append(f"{category}: {skill}")
                            member[skill_type] = new_skills
            logging.info(f"Successfully loaded data from {self.data_file}")
        except FileNotFoundError:
            logging.warning(f"File {self.data_file} not found. Starting with empty data.")
            self.initialize_empty_data()
        except json.JSONDecodeError:
            logging.error(f"Error decoding JSON from {self.data_file}. Starting with empty data.")
            self.initialize_empty_data()
        except Exception as e:
            logging.error(f"Error loading data: {str(e)}")
            self.initialize_empty_data()

    def initialize_empty_data(self):
        logging.info("Initializing empty data structures")
        self.team_members = {}
        self.skills_data = {
            "Technical Skills": {},
            "Soft Skills": {},
            "Software Skills": [],
            "Certifications": []
        }
        self.subskills_data = {
            "Technical Skills": {},
            "Soft Skills": {}
        }
        self.meetings = []

    def save_data(self):
        logging.info("Saving data")
        try:
            with open(self.data_file, 'w') as file:
                json.dump({
                    'team_members': self.team_members,
                    'skills_data': self.skills_data,
                    'subskills_data': self.subskills_data,
                    'meetings': self.meetings
                }, file, indent=2)
            logging.info(f"Successfully saved data to {self.data_file}")
        except Exception as e:
            logging.error(f"Error saving data: {str(e)}")
            self.handle_save_error()

    def handle_save_error(self):
        logging.info("Handling save error")
        message = f"Unable to save to {self.data_file}. Would you like to choose a different save location?"
        if messagebox.askyesno("Save Error", message):
            new_file = filedialog.asksaveasfilename(
                defaultextension=".json",
                filetypes=[("JSON files", "*.json"), ("All files", "*.*")]
            )
            if new_file:
                self.data_file = new_file
                self.save_data()
            else:
                logging.info("Save cancelled by user")
                messagebox.showinfo("Save Cancelled", "Data was not saved.")
        else:
            logging.info("Save cancelled by user")
            messagebox.showinfo("Save Cancelled", "Data was not saved.")

    def create_widgets(self):
        logging.info("Creating widgets")
        # Create menu
        menubar = tk.Menu(self.master)
        self.master.config(menu=menubar)

        file_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="File", menu=file_menu)
        file_menu.add_command(label="Load Config File", command=self.load_config_file)
        file_menu.add_command(label="Save Config File", command=self.save_config)
        file_menu.add_separator()
        file_menu.add_command(label="Exit", command=self.close_app)

        # Create Notebook
        self.notebook = ttk.Notebook(self.master)
        self.notebook.pack(fill=tk.BOTH, expand=tk.YES, padx=10, pady=10)

        # Create frames for each tab
        self.team_members_frame = ttk.Frame(self.notebook, padding="10")
        self.member_details_frame = ttk.Frame(self.notebook, padding="10")
        self.skills_management_frame = ttk.Frame(self.notebook, padding="10")
        self.meetings_frame = ttk.Frame(self.notebook, padding="10")

        # Add tabs to notebook
        self.notebook.add(self.team_members_frame, text="Team Members")
        self.notebook.add(self.member_details_frame, text="Member Details")
        self.notebook.add(self.skills_management_frame, text="Skills Management")
        self.notebook.add(self.meetings_frame, text="Meetings")

        # Initialize tabs
        self.create_team_members_tab()
        self.create_member_details_tab()
        self.create_skills_management_tab()
        self.create_meetings_tab()

        # Initialize meetings data
        self.meetings = []

    def create_team_members_tab(self):
        logging.info("Creating team members tab")
        # Treeview for displaying team members
        self.team_tree = ttk.Treeview(self.team_members_frame, columns=("Name", "Job Title", "Join Date", "Birthday"), show="headings")
        self.team_tree.heading("Name", text="Name")
        self.team_tree.heading("Job Title", text="Job Title")
        self.team_tree.heading("Join Date", text="Join Date")
        self.team_tree.heading("Birthday", text="Birthday")
        self.team_tree.pack(fill=tk.BOTH, expand=tk.YES, padx=5, pady=5)

        # Buttons for adding, modifying, deleting members
        button_frame = ttk.Frame(self.team_members_frame)
        button_frame.pack(fill=tk.X, padx=5, pady=5)

        add_button = ttk.Button(button_frame, text="Add Member", command=self.add_team_member, style="success.TButton" if USE_TTKBOOTSTRAP else "TButton")
        add_button.pack(side=tk.LEFT, padx=5)

        modify_button = ttk.Button(button_frame, text="Modify Member", command=self.modify_team_member, style="primary.TButton" if USE_TTKBOOTSTRAP else "TButton")
        modify_button.pack(side=tk.LEFT, padx=5)

        delete_button = ttk.Button(button_frame, text="Delete Member", command=self.delete_team_member, style="danger.TButton" if USE_TTKBOOTSTRAP else "TButton")
        delete_button.pack(side=tk.LEFT, padx=5)

        refresh_button = ttk.Button(button_frame, text="Refresh List", command=self.refresh_team_list, style="info.TButton" if USE_TTKBOOTSTRAP else "TButton")
        refresh_button.pack(side=tk.LEFT, padx=5)

        self.refresh_team_list()

    def create_member_details_tab(self):
        logging.info("Creating member details tab")
        # Create a canvas with scrollbar for member details
        canvas = tk.Canvas(self.member_details_frame)
        scrollbar = ttk.Scrollbar(self.member_details_frame, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)

        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(
                scrollregion=canvas.bbox("all")
            )
        )

        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        # Bind mousewheel to scroll
        canvas.bind_all("<MouseWheel>", lambda event: canvas.yview_scroll(int(-1*(event.delta/120)), "units"))

        # Add Member Selection Dropdown
        ttk.Label(scrollable_frame, text="Select Team Member:").grid(row=0, column=0, padx=10, pady=5, sticky="w")
        self.member_var = tk.StringVar()
        self.member_dropdown = ttk.Combobox(scrollable_frame, textvariable=self.member_var, state="readonly")
        self.member_dropdown.grid(row=1, column=0, padx=10, pady=5, sticky="ew")
        self.member_dropdown.bind("<<ComboboxSelected>>", self.load_member_data)

        # Initialize skill dropdowns
        for skill_type in ["Technical Skills", "Soft Skills", "Certifications"]:
            dropdown_name = f"{skill_type.lower().replace(' ', '_')}_dropdown"
            setattr(self, dropdown_name, ttk.Combobox(scrollable_frame, state="readonly"))

        # Personal Section
        personal_frame = ttk.LabelFrame(scrollable_frame, text="Personal Information", padding="10")
        personal_frame.grid(row=2, column=0, padx=10, pady=10, sticky="nsew")

        # Professional Section
        professional_frame = ttk.LabelFrame(scrollable_frame, text="Professional Information", padding="10")
        professional_frame.grid(row=3, column=0, padx=10, pady=10, sticky="nsew")

        # Skills Section
        skills_frame = ttk.LabelFrame(scrollable_frame, text="Skills and Certifications", padding="10")
        skills_frame.grid(row=4, column=0, padx=10, pady=10, sticky="nsew")

        # Progression Section
        progression_frame = ttk.LabelFrame(scrollable_frame, text="Progression", padding="10")
        progression_frame.grid(row=5, column=0, padx=10, pady=10, sticky="nsew")

        # Configure grid weights
        scrollable_frame.grid_rowconfigure(2, weight=1)
        scrollable_frame.grid_rowconfigure(3, weight=1)
        scrollable_frame.grid_rowconfigure(4, weight=1)
        scrollable_frame.grid_rowconfigure(5, weight=1)
        scrollable_frame.grid_columnconfigure(0, weight=1)

        # Personal Information Fields
        ttk.Label(personal_frame, text="Hobbies:").grid(row=0, column=0, padx=5, pady=5, sticky="e")
        self.hobbies_entry = ttk.Entry(personal_frame, width=50)
        self.hobbies_entry.grid(row=0, column=1, padx=5, pady=5, sticky="ew")

        ttk.Label(personal_frame, text="Interests:").grid(row=1, column=0, padx=5, pady=5, sticky="e")
        self.interests_entry = ttk.Entry(personal_frame, width=50)
        self.interests_entry.grid(row=1, column=1, padx=5, pady=5, sticky="ew")

        ttk.Label(personal_frame, text="Family Information:").grid(row=2, column=0, padx=5, pady=5, sticky="e")
        self.family_entry = ttk.Entry(personal_frame, width=50)
        self.family_entry.grid(row=2, column=1, padx=5, pady=5, sticky="ew")

        ttk.Label(personal_frame, text="Other Personal Details:").grid(row=3, column=0, padx=5, pady=5, sticky="ne")
        self.other_personal_text = scrolledtext.ScrolledText(personal_frame, width=50, height=5)
        self.other_personal_text.grid(row=3, column=1, padx=5, pady=5, sticky="ew")

        # Professional Information Fields
        ttk.Label(professional_frame, text="Job Title:").grid(row=0, column=0, padx=5, pady=5, sticky="e")
        self.job_title_entry = ttk.Entry(professional_frame, width=50)
        self.job_title_entry.grid(row=0, column=1, padx=5, pady=5, sticky="ew")

        ttk.Label(professional_frame, text="Join Date (DD-MM-YYYY):").grid(row=1, column=0, padx=5, pady=5, sticky="e")
        self.join_date_entry = ttk.Entry(professional_frame, width=50)
        self.join_date_entry.grid(row=1, column=1, padx=5, pady=5, sticky="ew")

        ttk.Label(professional_frame, text="Birthday (DD-MM-YYYY):").grid(row=2, column=0, padx=5, pady=5, sticky="e")
        self.birthday_entry = ttk.Entry(professional_frame, width=50)
        self.birthday_entry.grid(row=2, column=1, padx=5, pady=5, sticky="ew")

        # Skills and Certifications
        self.create_skill_section(skills_frame, "Technical Skills", 0)
        self.create_skill_section(skills_frame, "Soft Skills", 1)
        self.create_skill_section(skills_frame, "Certifications", 2)

        # Progression Fields
        ttk.Label(progression_frame, text="Goals:").grid(row=0, column=0, padx=5, pady=5, sticky="ne")
        self.goals_text = scrolledtext.ScrolledText(progression_frame, width=60, height=5)
        self.goals_text.grid(row=0, column=1, padx=5, pady=5, sticky="ew")
        self.goals_text.bind("<Double-Button-1>", self.move_goal_to_achievement)

        ttk.Label(progression_frame, text="Development Plan:").grid(row=1, column=0, padx=5, pady=5, sticky="ne")
        self.dev_plan_text = scrolledtext.ScrolledText(progression_frame, width=60, height=5)
        self.dev_plan_text.grid(row=1, column=1, padx=5, pady=5, sticky="ew")

        ttk.Label(progression_frame, text="Achieved Goals:").grid(row=2, column=0, padx=5, pady=5, sticky="ne")
        self.achieved_goals_text = scrolledtext.ScrolledText(progression_frame, width=60, height=5)
        self.achieved_goals_text.grid(row=2, column=1, padx=5, pady=5, sticky="ew")

        # Update Button
        update_button = ttk.Button(
            scrollable_frame, 
            text="Update Member", 
            command=self.update_team_member, 
            style="warning.TButton" if USE_TTKBOOTSTRAP else "TButton"
        )
        update_button.grid(row=6, column=0, pady=20)

        # Configure weights for frames
        personal_frame.grid_columnconfigure(1, weight=1)
        professional_frame.grid_columnconfigure(1, weight=1)
        skills_frame.grid_columnconfigure(1, weight=1)
        progression_frame.grid_columnconfigure(1, weight=1)

    def create_skill_section(self, parent, skill_type, row):
        logging.info(f"Creating {skill_type} section")
        ttk.Label(parent, text=f"{skill_type}:").grid(row=row, column=0, padx=5, pady=5, sticky="ne")
        
        frame = ttk.Frame(parent)
        frame.grid(row=row, column=1, padx=5, pady=5, sticky="ew")
        
        category_dropdown = None
        subskill_dropdown = None
        if skill_type in ["Technical Skills", "Soft Skills", "Software Skills"]:
            # Category dropdown
            ttk.Label(frame, text="Category:").grid(row=0, column=0, padx=5, pady=5, sticky="e")
            category_dropdown = ttk.Combobox(frame, textvariable=self.category_vars[skill_type], state="readonly", width=20)
            category_dropdown.grid(row=0, column=1, padx=5, pady=5)
        
            # Subskill dropdown
            ttk.Label(frame, text="Subskill:").grid(row=0, column=2, padx=5, pady=5, sticky="e")
            subskill_dropdown = ttk.Combobox(frame, textvariable=self.subskill_vars[skill_type], state="readonly", width=20)
            subskill_dropdown.grid(row=0, column=3, padx=5, pady=5)
        
            category_dropdown.bind("<<ComboboxSelected>>", lambda e, st=skill_type: self.on_category_select(st))
            
            # Add button
            add_button = ttk.Button(frame, text="Add", command=lambda: self.add_skill_to_member(skill_type, self.category_vars[skill_type], self.subskill_vars[skill_type]))
            add_button.grid(row=0, column=4, padx=5, pady=5)
        else:
            # Single dropdown for Certifications
            ttk.Label(frame, text="Certification:").grid(row=0, column=0, padx=5, pady=5, sticky="e")
            skill_dropdown = ttk.Combobox(frame, textvariable=self.category_vars[skill_type], state="readonly", width=40)
            skill_dropdown.grid(row=0, column=1, columnspan=3, padx=5, pady=5)
            
            # Add button
            add_button = ttk.Button(frame, text="Add", command=lambda: self.add_skill_to_member(skill_type, self.category_vars[skill_type]))
            add_button.grid(row=0, column=4, padx=5, pady=5)
        
        # Listbox for skills
        listbox = tk.Listbox(frame, width=50, height=5)
        listbox.grid(row=1, column=0, columnspan=5, padx=5, pady=5, sticky="ew")
        
        # Remove button
        remove_button = ttk.Button(frame, text="Remove", command=lambda: self.remove_skill_from_member(skill_type, listbox))
        remove_button.grid(row=2, column=0, columnspan=5, padx=5, pady=5)
        
        # Store references
        if skill_type in ["Technical Skills", "Soft Skills", "Software Skills"]:
            self.category_dropdowns[skill_type] = category_dropdown
            self.subskill_dropdowns[skill_type] = subskill_dropdown
        else:
            self.category_dropdowns[skill_type] = skill_dropdown
        setattr(self, f"{skill_type.lower().replace(' ', '_')}_listbox", listbox)

        # Populate dropdowns
        self.populate_skill_dropdowns(skill_type)
        logging.info(f"{skill_type} section created successfully")

    def on_category_select(self, skill_type):
        print(f"on_category_select called with skill_type: {skill_type}")  # Debug print
        logging.info(f"{skill_type} category selected")
        try:
            if skill_type not in self.category_vars or skill_type not in self.subskill_dropdowns:
                logging.warning(f"Category or subskill dropdown not found for {skill_type}")
                return

            category = self.category_vars[skill_type].get()
            if not category:
                logging.warning(f"No category selected for {skill_type}")
                return

            self.refresh_subskill_dropdown(skill_type)
        except Exception as e:
            logging.error(f"Error in on_category_select: {str(e)}")
            messagebox.showerror("Error", f"An error occurred while updating subskills: {str(e)}")

    def refresh_subskill_dropdown(self, skill_type):
        logging.info(f"Refreshing subskill dropdown for {skill_type}")
        try:
            category = self.category_vars[skill_type].get()
            subskills = self.subskills_data.get(skill_type, {}).get(category, [])
            self.subskill_dropdowns[skill_type]['values'] = subskills
            if subskills:
                self.subskill_dropdowns[skill_type].set(subskills[0])
                self.subskill_vars[skill_type].set(subskills[0])
            else:
                self.subskill_dropdowns[skill_type].set('')
                self.subskill_vars[skill_type].set('')
        except Exception as e:
            logging.error(f"Error in refresh_subskill_dropdown: {str(e)}")
            messagebox.showerror("Error", f"An error occurred while refreshing subskills: {str(e)}")

    def add_skill_to_member(self, skill_type, *args):
        logging.info(f"Adding {skill_type} to member")
        member_name = self.member_var.get()
        if not member_name:
            messagebox.showerror("Error", "Please select a team member.")
            return
        
        member = self.team_members[member_name]
        skill_list = member[skill_type.lower().replace(' ', '_')]
        
        if skill_type in ["Technical Skills", "Soft Skills"]:
            category_var, subskill_var = args
            category = category_var.get()
            subskill = subskill_var.get()
            if not category or not subskill:
                messagebox.showerror("Error", f"Please select a category and subskill for {skill_type}.")
                return
            skill = f"{category}: {subskill}"
        else:
            skill_var, = args
            skill = skill_var.get()
            if not skill:
                messagebox.showerror("Error", f"Please select a {skill_type}.")
                return
        
        if skill in skill_list:
            messagebox.showinfo("Info", f"{skill} is already in the member's {skill_type} list.")
            return
        
        skill_list.append(skill)
        self.refresh_member_skills(skill_type)
        self.save_data()
        messagebox.showinfo("Success", f"{skill} has been added to {member_name}'s {skill_type}.")

    def remove_skill_from_member(self, skill_type, listbox):
        logging.info(f"Removing {skill_type} from member")
        selection = listbox.curselection()
        if not selection:
            messagebox.showerror("Error", f"Please select a {skill_type} to remove.")
            return
        
        skill = listbox.get(selection[0])
        member_name = self.member_var.get()
        
        if messagebox.askyesno("Confirm Removal", f"Are you sure you want to remove {skill} from {member_name}'s {skill_type}?"):
            member = self.team_members[member_name]
            skill_list = member[skill_type.lower().replace(' ', '_')]
            skill_list.remove(skill)
            self.refresh_member_skills(skill_type)
            self.save_data()
            messagebox.showinfo("Success", f"{skill} has been removed from {member_name}'s {skill_type}.")

    def refresh_member_skills(self, skill_type):
        logging.info(f"Refreshing member's {skill_type}")
        member_name = self.member_var.get()
        if not member_name:
            return
        
        member = self.team_members[member_name]
        skill_key = skill_type.lower().replace(' ', '_')
        skill_list = member.get(skill_key, [])
        
        listbox_name = f"{skill_key}_listbox"
        listbox = getattr(self, listbox_name, None)
        if listbox is not None:
            listbox.delete(0, tk.END)
            for skill in skill_list:
                listbox.insert(tk.END, skill)
        else:
            logging.warning(f"Listbox for {skill_type} not found")

    def load_member_data(self, event):
        logging.info("Loading member data")
        name = self.member_var.get()
        if name in self.team_members:
            member = self.team_members[name]
            # ... (existing code for loading personal and professional information)

            # Populate category dropdowns
            self.populate_category_dropdowns()

            # Populate Certifications dropdown
            if hasattr(self, 'certifications_dropdown'):
                self.certifications_dropdown['values'] = self.skills_data["Certifications"]
                self.certifications_dropdown['state'] = 'readonly'

            # Skills and Certifications
            for skill_type in ["Technical Skills", "Soft Skills", "Certifications"]:
                if hasattr(self, f"{skill_type.lower().replace(' ', '_')}_listbox"):
                    self.refresh_member_skills(skill_type)

            # ... (existing code for loading progression fields)
        else:
            self.clear_member_details()

    def populate_category_dropdowns(self):
        logging.info("Populating category dropdowns")
        for skill_type in ["Technical Skills", "Soft Skills"]:
            try:
                category_dropdown = self.category_dropdowns.get(skill_type)
                subskill_dropdown = self.subskill_dropdowns.get(skill_type)
                
                if not category_dropdown or not subskill_dropdown:
                    logging.warning(f"{skill_type} dropdowns not found. Skipping.")
                    continue
                
                categories = list(self.subskills_data.get(skill_type, {}).keys())
                category_dropdown['values'] = categories
                if categories:
                    category_dropdown.current(0)
                    selected_category = category_dropdown.get()
                    subskills = self.subskills_data.get(skill_type, {}).get(selected_category, [])
                    subskill_dropdown['values'] = subskills
                    if subskills:
                        subskill_dropdown.current(0)
                else:
                    category_dropdown.set('')
                    subskill_dropdown['values'] = []
                    subskill_dropdown.set('')
                logging.info(f"Successfully populated {skill_type} dropdowns")
            except Exception as e:
                logging.error(f"Error populating {skill_type} dropdowns: {str(e)}")
                messagebox.showwarning("Warning", f"Unable to populate {skill_type} dropdowns. Some features may not work correctly.")

    def populate_category_dropdown(self, skill_type, category_dropdown):
        logging.info(f"Populating {skill_type} category dropdown")
        categories = list(self.subskills_data[skill_type].keys())
        category_dropdown['values'] = categories
        if categories:
            category_dropdown.current(0)

    def create_skills_management_tab(self):
        logging.info("Creating skills management tab")
        # Create a notebook within the skills management tab
        notebook = ttk.Notebook(self.skills_management_frame)
        notebook.pack(fill=tk.BOTH, expand=tk.YES, padx=10, pady=10)

        # Create frames for each skill category
        technical_skills_frame = ttk.Frame(notebook)
        soft_skills_frame = ttk.Frame(notebook)
        certifications_frame = ttk.Frame(notebook)

        # Add frames to notebook
        notebook.add(technical_skills_frame, text="Technical Skills")
        notebook.add(soft_skills_frame, text="Soft Skills")
        notebook.add(certifications_frame, text="Certifications")

        # Initialize skill management sections
        self.create_hierarchical_skills_section(technical_skills_frame, "Technical Skills")
        self.create_hierarchical_skills_section(soft_skills_frame, "Soft Skills")
        self.create_certifications_section(certifications_frame)

        # Add CSV upload button
        upload_button = ttk.Button(self.skills_management_frame, text="Upload Skills CSV", command=self.upload_skills_csv)
        upload_button.pack(pady=10)

    def create_hierarchical_skills_section(self, frame, skill_type):
        logging.info(f"Creating {skill_type} section")
        
        # Category selection
        category_frame = ttk.Frame(frame)
        category_frame.pack(fill=tk.X, pady=5)
        ttk.Label(category_frame, text=f"{skill_type} Categories:").pack(side=tk.LEFT, padx=5)
        self.category_vars[skill_type] = tk.StringVar()
        self.category_dropdowns[skill_type] = ttk.Combobox(category_frame, textvariable=self.category_vars[skill_type], state="readonly", width=40)
        self.category_dropdowns[skill_type].pack(side=tk.LEFT, padx=5)
        self.category_dropdowns[skill_type].bind("<<ComboboxSelected>>", lambda e: self.on_category_select(skill_type))
        
        # Add category button
        add_category_button = ttk.Button(category_frame, text="Add Category", command=lambda: self.add_category(skill_type), style="info.TButton" if USE_TTKBOOTSTRAP else "TButton")
        add_category_button.pack(side=tk.LEFT, padx=5)

        # Subskill entry and addition
        subskill_frame = ttk.Frame(frame)
        subskill_frame.pack(fill=tk.X, pady=5)
        ttk.Label(subskill_frame, text="Add Subskills:").pack(side=tk.LEFT, padx=5)
        self.subskill_entry = ttk.Entry(subskill_frame, width=40)
        self.subskill_entry.pack(side=tk.LEFT, padx=5)
        add_subskills_button = ttk.Button(subskill_frame, text="Add Subskills", command=lambda: self.add_multiple_subskills(skill_type), style="success.TButton" if USE_TTKBOOTSTRAP else "TButton")
        add_subskills_button.pack(side=tk.LEFT, padx=5)

        # Subskill selection
        subskill_select_frame = ttk.Frame(frame)
        subskill_select_frame.pack(fill=tk.X, pady=5)
        ttk.Label(subskill_select_frame, text="Subskills:").pack(side=tk.LEFT, padx=5)
        self.subskill_vars[skill_type] = tk.StringVar()
        self.subskill_dropdowns[skill_type] = ttk.Combobox(subskill_select_frame, textvariable=self.subskill_vars[skill_type], state="readonly", width=40)
        self.subskill_dropdowns[skill_type].pack(side=tk.LEFT, padx=5)

        # Delete subskill button
        delete_subskill_button = ttk.Button(subskill_select_frame, text="Delete Subskill", command=lambda: self.delete_subskill(skill_type), style="danger.TButton" if USE_TTKBOOTSTRAP else "TButton")
        delete_subskill_button.pack(side=tk.LEFT, padx=5)

        # Delete category button
        delete_category_button = ttk.Button(subskill_select_frame, text="Delete Category", command=lambda: self.delete_category(skill_type), style="danger.TButton" if USE_TTKBOOTSTRAP else "TButton")
        delete_category_button.pack(side=tk.LEFT, padx=5)

        # Listbox for added skills
        ttk.Label(frame, text=f"Added {skill_type}:").pack(pady=5)
        self.skills_listboxes[skill_type] = tk.Listbox(frame, selectmode=tk.SINGLE, width=50, height=10)
        self.skills_listboxes[skill_type].pack(pady=5)

        # Remove skill button
        remove_skill_button = ttk.Button(frame, text="Remove Skill", command=lambda: self.remove_hierarchical_skill(skill_type), style="danger.TButton" if USE_TTKBOOTSTRAP else "TButton")
        remove_skill_button.pack(pady=5)

        # Refresh categories
        self.refresh_categories(skill_type)

    def create_soft_skills_section(self, frame):
        logging.info("Creating soft skills section")
        # Listbox for soft skill categories
        ttk.Label(frame, text="Soft Skill Categories:").pack(pady=5)
        self.soft_categories_listbox = tk.Listbox(frame, selectmode=tk.SINGLE, width=50, height=10)
        self.soft_categories_listbox.pack(pady=5)
        self.soft_categories_listbox.bind('<<ListboxSelect>>', self.on_soft_category_select)

        # Buttons for category management
        category_buttons_frame = ttk.Frame(frame)
        category_buttons_frame.pack(pady=5)

        add_category_button = ttk.Button(category_buttons_frame, text="Add Category", command=self.add_soft_category, style="success.TButton" if USE_TTKBOOTSTRAP else "TButton")
        add_category_button.pack(side=tk.LEFT, padx=5)

        remove_category_button = ttk.Button(category_buttons_frame, text="Remove Category", command=self.remove_soft_category, style="danger.TButton" if USE_TTKBOOTSTRAP else "TButton")
        remove_category_button.pack(side=tk.LEFT, padx=5)

        modify_category_button = ttk.Button(category_buttons_frame, text="Modify Category", command=self.modify_soft_category, style="primary.TButton" if USE_TTKBOOTSTRAP else "TButton")
        modify_category_button.pack(side=tk.LEFT, padx=5)

        delete_category_button = ttk.Button(category_buttons_frame, text="Delete Category", command=self.delete_soft_skill_category, style="danger.TButton" if USE_TTKBOOTSTRAP else "TButton")
        delete_category_button.pack(side=tk.LEFT, padx=5)

        # Listbox for skills in selected category
        ttk.Label(frame, text="Skills in Selected Category:").pack(pady=5)
        self.soft_skills_listbox = tk.Listbox(frame, selectmode=tk.MULTIPLE, width=50, height=10)
        self.soft_skills_listbox.pack(pady=5)

        # Buttons for skill management
        skill_buttons_frame = ttk.Frame(frame)
        skill_buttons_frame.pack(pady=5)

        add_skill_button = ttk.Button(skill_buttons_frame, text="Add Skill", command=self.add_soft_skill, style="success.TButton" if USE_TTKBOOTSTRAP else "TButton")
        add_skill_button.pack(side=tk.LEFT, padx=5)

        remove_skill_button = ttk.Button(skill_buttons_frame, text="Remove Skill", command=self.remove_soft_skill, style="danger.TButton" if USE_TTKBOOTSTRAP else "TButton")
        remove_skill_button.pack(side=tk.LEFT, padx=5)

        modify_skill_button = ttk.Button(skill_buttons_frame, text="Modify Skill", command=self.modify_soft_skill, style="primary.TButton" if USE_TTKBOOTSTRAP else "TButton")
        modify_skill_button.pack(side=tk.LEFT, padx=5)

        self.refresh_soft_categories()

    def create_certifications_section(self, frame):
        logging.info("Creating certifications section")
        # Listbox for certifications
        ttk.Label(frame, text="Certifications:").pack(pady=5)
        self.certifications_listbox = tk.Listbox(frame, selectmode=tk.MULTIPLE, width=50, height=15)
        self.certifications_listbox.pack(pady=5)

        # Buttons for certifications management
        cert_buttons_frame = ttk.Frame(frame)
        cert_buttons_frame.pack(pady=5)

        add_certification_button = ttk.Button(cert_buttons_frame, text="Add Certification", command=self.add_certification, style="success.TButton" if USE_TTKBOOTSTRAP else "TButton")
        add_certification_button.pack(side=tk.LEFT, padx=5)

        remove_certification_button = ttk.Button(cert_buttons_frame, text="Remove Certification", command=self.remove_certification, style="danger.TButton" if USE_TTKBOOTSTRAP else "TButton")
        remove_certification_button.pack(side=tk.LEFT, padx=5)

        modify_certification_button = ttk.Button(cert_buttons_frame, text="Modify Certification", command=self.modify_certification, style="primary.TButton" if USE_TTKBOOTSTRAP else "TButton")
        modify_certification_button.pack(side=tk.LEFT, padx=5)

        self.refresh_certifications()

    def refresh_team_list(self):
        logging.info("Refreshing team list")
        # Clear existing entries
        for item in self.team_tree.get_children():
            self.team_tree.delete(item)
        # Insert all team members
        for name, data in self.team_members.items():
            self.team_tree.insert("", "end", values=(name, data.get('job_title', ''), data.get('join_date', ''), data.get('birthday', '')))

    def refresh_member_dropdown(self):
        logging.info("Refreshing member dropdown")
        members = list(self.team_members.keys())
        self.member_dropdown['values'] = members
        if members:
            self.member_dropdown.current(0)
            self.load_member_data(None)
        else:
            self.clear_member_details()
            self.member_dropdown.set('')  # Clear the dropdown text when there are no members

    def refresh_tech_categories(self):
        logging.info("Refreshing technical skill categories")
        self.tech_categories_listbox.delete(0, tk.END)
        for category in self.skills_data["Technical Skills"].keys():
            self.tech_categories_listbox.insert(tk.END, category)
        self.refresh_technical_skills_listbox()

    def refresh_technical_skills_listbox(self, selected_category=None):
        logging.info("Refreshing technical skills listbox")
        self.tech_skills_listbox.delete(0, tk.END)
        if not selected_category:
            selection = self.tech_categories_listbox.curselection()
            if selection:
                selected_category = self.tech_categories_listbox.get(selection[0])
            else:
                return
        for skill in self.skills_data["Technical Skills"].get(selected_category, []):
            self.tech_skills_listbox.insert(tk.END, skill)

    def refresh_soft_skills(self):
        logging.info("Refreshing soft skills")
        self.soft_skills_listbox.delete(0, tk.END)
        for skill in self.skills_data["Soft Skills"]:
            self.soft_skills_listbox.insert(tk.END, skill)

    def refresh_certifications(self):
        logging.info("Refreshing certifications")
        self.certifications_listbox.delete(0, tk.END)
        for cert in self.skills_data["Certifications"]:
            self.certifications_listbox.insert(tk.END, cert)

    def populate_technical_skills(self):
        logging.info("Populating technical skills listbox")
        self.technical_skills_listbox.delete(0, tk.END)
        for category, skills in self.skills_data["Technical Skills"].items():
            for skill in skills:
                self.technical_skills_listbox.insert(tk.END, skill)

    def populate_soft_skills(self):
        logging.info("Populating soft skills listbox")
        self.soft_skills_listbox.delete(0, tk.END)
        for skill in self.skills_data["Soft Skills"]:
            self.soft_skills_listbox.insert(tk.END, skill)

    def populate_certifications(self):
        logging.info("Populating certifications listbox")
        self.certifications_listbox.delete(0, tk.END)
        for cert in self.skills_data["Certifications"]:
            self.certifications_listbox.insert(tk.END, cert)

    def add_team_member(self):
        logging.info("Adding team member")
        dialog = TeamMemberDialog(self.master, "Add Team Member")
        self.master.wait_window(dialog.top)
        if dialog.result:
            name, job_title, join_date, birthday = dialog.result
            if name in self.team_members:
                messagebox.showerror("Error", "A team member with this name already exists.")
                return
            # Add team member
            self.team_members[name] = {
                "job_title": job_title,
                "join_date": join_date,
                "birthday": birthday,
                "technical_skills": [],
                "soft_skills": [],
                "software_skills": [],
                "certifications": [],
                "goals": [],
                "development_plan": "",
                "hobbies": [],
                "interests": [],
                "family": "",
                "achievements": [],
                "issues_concerns": "",
                "other_personal_details": ""
            }
            self.save_data()
            self.refresh_team_list()
            self.refresh_member_dropdown()
            messagebox.showinfo("Success", f"{name} has been added to the team.")

    def modify_team_member(self):
        logging.info("Modifying team member")
        selected = self.team_tree.selection()
        if not selected:
            messagebox.showerror("Error", "Please select a team member to modify.")
            return
        item = self.team_tree.item(selected[0])
        name = item['values'][0]
        data = self.team_members[name]
        dialog = TeamMemberDialog(self.master, "Modify Team Member", name, data['job_title'], data['join_date'], data['birthday'])
        self.master.wait_window(dialog.top)
        if dialog.result:
            new_name, job_title, join_date, birthday = dialog.result
            if new_name != name and new_name in self.team_members:
                messagebox.showerror("Error", "A team member with the new name already exists.")
                return
            # Update member details
            if new_name != name:
                self.team_members[new_name] = self.team_members.pop(name)
                name = new_name
            self.team_members[name]['job_title'] = job_title
            self.team_members[name]['join_date'] = join_date
            self.team_members[name]['birthday'] = birthday
            self.save_data()
            self.refresh_team_list()
            self.refresh_member_dropdown()
            messagebox.showinfo("Success", f"{name}'s information has been updated.")

    def delete_team_member(self):
        logging.info("Deleting team member")
        selected = self.team_tree.selection()
        if not selected:
            messagebox.showerror("Error", "Please select a team member to delete.")
            return
        item = self.team_tree.item(selected[0])
        name = item['values'][0]
        if messagebox.askyesno("Confirm Deletion", f"Are you sure you want to delete '{name}' from the team?"):
            del self.team_members[name]
            self.save_data()
            self.refresh_team_list()
            self.refresh_member_dropdown()
            messagebox.showinfo("Deleted", f"{name} has been deleted from the team.")

    def load_member_data(self, event):
        logging.info("Loading member data")
        name = self.member_var.get()
        if name in self.team_members:
            member = self.team_members[name]
            # Personal Information
            self.hobbies_entry.delete(0, tk.END)
            self.hobbies_entry.insert(0, ", ".join(member.get('hobbies', [])))
            self.interests_entry.delete(0, tk.END)
            self.interests_entry.insert(0, ", ".join(member.get('interests', [])))
            self.family_entry.delete(0, tk.END)
            self.family_entry.insert(0, member.get('family', ''))
            self.other_personal_text.delete("1.0", tk.END)
            self.other_personal_text.insert(tk.END, member.get('other_personal_details', ''))

            # Professional Information
            self.job_title_entry.delete(0, tk.END)
            self.job_title_entry.insert(0, member.get('job_title', ''))
            self.join_date_entry.delete(0, tk.END)
            self.join_date_entry.insert(0, member.get('join_date', ''))
            self.birthday_entry.delete(0, tk.END)
            self.birthday_entry.insert(0, member.get('birthday', ''))

            # Skills and Certifications
            self.refresh_member_skills("Technical Skills")
            self.refresh_member_skills("Soft Skills")
            self.refresh_member_skills("Software Skills")
            self.refresh_member_skills("Certifications")

            # Progression Fields
            self.goals_text.delete("1.0", tk.END)
            self.goals_text.insert(tk.END, "\n".join(member.get('goals', [])))
            self.dev_plan_text.delete("1.0", tk.END)
            self.dev_plan_text.insert(tk.END, member.get('development_plan', ''))
            self.achieved_goals_text.delete("1.0", tk.END)
            self.achieved_goals_text.insert(tk.END, "\n".join(member.get('achieved_goals', [])))
        else:
            self.clear_member_details()

        # Populate category dropdowns
        self.populate_category_dropdowns()

    def populate_skill_dropdowns(self, skill_type):
        logging.info(f"Populating {skill_type} dropdowns")
        try:
            if skill_type in ["Technical Skills", "Soft Skills", "Software Skills"]:
                category_dropdown = self.category_dropdowns.get(skill_type)
                subskill_dropdown = self.subskill_dropdowns.get(skill_type)
                
                if category_dropdown and subskill_dropdown:
                    if skill_type not in self.subskills_data:
                        self.subskills_data[skill_type] = {}
                    categories = list(self.subskills_data[skill_type].keys())
                    category_dropdown['values'] = categories
                    if categories:
                        category_dropdown.current(0)
                        self.category_vars[skill_type].set(categories[0])
                        self.on_category_select(skill_type)
                    else:
                        self.category_vars[skill_type].set('')
                        self.subskill_vars[skill_type].set('')
                        subskill_dropdown['values'] = []
            elif skill_type == "Certifications":
                dropdown = getattr(self, f"{skill_type.lower()}_dropdown", None)
                if dropdown:
                    dropdown['values'] = self.skills_data.get("Certifications", [])
                    if self.skills_data.get("Certifications"):
                        dropdown.current(0)
            logging.info(f"Successfully populated {skill_type} dropdowns")
        except Exception as e:
            logging.error(f"Error in populate_skill_dropdowns: {str(e)}")
            messagebox.showerror("Error", f"An error occurred while populating {skill_type} dropdowns: {str(e)}")

    def update_team_member(self):
        logging.info("Updating team member")
        name = self.member_var.get()
        if name not in self.team_members:
            messagebox.showerror("Error", "Please select a valid team member.")
            return
        member = self.team_members[name]

        # Personal Information
        hobbies = [hobby.strip() for hobby in self.hobbies_entry.get().split(',') if hobby.strip()]
        interests = [interest.strip() for interest in self.interests_entry.get().split(',') if interest.strip()]
        family = self.family_entry.get().strip()
        other_personal = self.other_personal_text.get("1.0", tk.END).strip()

        # Professional Information
        job_title = self.job_title_entry.get().strip()
        join_date = self.join_date_entry.get().strip()
        birthday = self.birthday_entry.get().strip()

        # Validate dates
        if not self.validate_date(join_date):
            messagebox.showerror("Invalid Date", "Join Date is not in the correct format (DD-MM-YYYY).")
            return
        if not self.validate_date(birthday):
            messagebox.showerror("Invalid Date", "Birthday is not in the correct format (DD-MM-YYYY).")
            return

        # Skills and Certifications
        for skill_type in ["Technical Skills", "Soft Skills", "Software Skills", "Certifications"]:
            listbox = getattr(self, f"{skill_type.lower().replace(' ', '_')}_listbox")
            member[skill_type.lower().replace(' ', '_')] = list(listbox.get(0, tk.END))

        # Progression Fields
        goals = [goal.strip() for goal in self.goals_text.get("1.0", tk.END).strip().split('\n') if goal.strip()]
        development_plan = self.dev_plan_text.get("1.0", tk.END).strip()
        achieved_goals = [goal.strip() for goal in self.achieved_goals_text.get("1.0", tk.END).strip().split('\n') if goal.strip()]

        # Update member details
        member['hobbies'] = hobbies
        member['interests'] = interests
        member['family'] = family
        member['other_personal_details'] = other_personal
        member['job_title'] = job_title
        member['join_date'] = join_date
        member['birthday'] = birthday
        member['goals'] = goals
        member['development_plan'] = development_plan
        member['achieved_goals'] = achieved_goals

        self.save_data()
        self.refresh_team_list()
        messagebox.showinfo("Success", f"{name}'s information has been updated.")

    def move_goal_to_achievement(self, event):
        logging.info("Moving goal to achievement")
        widget = event.widget
        try:
            index = widget.index(f"@{event.x},{event.y}")
            line_num = int(index.split('.')[0])
            goals = widget.get("1.0", tk.END).strip().split('\n')
            if 1 <= line_num <= len(goals):
                achieved_goal = goals.pop(line_num - 1)
                widget.delete("1.0", tk.END)
                widget.insert(tk.END, "\n".join(goals))
                member = self.team_members.get(self.member_var.get(), {})
                member['achievements'].append(achieved_goal)
                self.save_data()
                messagebox.showinfo("Goal Achieved", f"The goal '{achieved_goal}' has been moved to achievements!")
        except Exception as e:
            logging.error(f"Error moving goal to achievement: {str(e)}")
            messagebox.showerror("Error", "Failed to move goal to achievement. Please try again.")

    def on_category_select(self, skill_type):
        logging.info(f"{skill_type} category selected")
        if skill_type not in self.category_vars or skill_type not in self.subskill_dropdowns:
            logging.warning(f"Category or subskill dropdown not found for {skill_type}")
            return

        category = self.category_vars[skill_type].get()
        if not category:
            logging.warning(f"No category selected for {skill_type}")
            return

        subskills = self.categories.get(skill_type, {}).get(category, [])
        self.subskill_dropdowns[skill_type]['values'] = subskills
        if subskills:
            self.subskill_dropdowns[skill_type].set(subskills[0])
        else:
            self.subskill_dropdowns[skill_type].set('')

        # Update the subskill_vars
        if skill_type in self.subskill_vars:
            self.subskill_vars[skill_type].set(self.subskill_dropdowns[skill_type].get())
        else:
            logging.warning(f"Subskill variable not found for {skill_type}")

    def add_hierarchical_skill(self, skill_type):
        logging.info(f"Adding {skill_type}")
        category = self.category_vars[skill_type].get()
        subskill = self.subskill_vars[skill_type].get()
        if not category:
            messagebox.showerror("Error", f"Please select a category for {skill_type}.")
            return
        if not subskill:
            # If no subskill is selected, prompt to add a new one
            new_subskill = simpledialog.askstring("Add Subskill", f"Enter a new subskill for '{category}':")
            if new_subskill:
                subskill = new_subskill.strip()
                if subskill not in self.subskills_data[skill_type][category]:
                    self.subskills_data[skill_type][category].append(subskill)
                    self.save_data()
                    self.refresh_categories(skill_type)
                    logging.info(f"Added new subskill '{subskill}' to category '{category}' in {skill_type}")
                else:
                    messagebox.showwarning("Duplicate", f"Subskill '{subskill}' already exists in this category.")
            else:
                return
        
        skill = f"{category}: {subskill}"
        member_name = self.member_var.get()
        if not member_name:
            messagebox.showerror("Error", "Please select a team member first.")
            return
        if skill in self.team_members[member_name][skill_type.lower().replace(' ', '_')]:
            messagebox.showerror("Error", f"This {skill_type} already exists for the selected member.")
            return
        self.team_members[member_name][skill_type.lower().replace(' ', '_')].append(skill)
        self.save_data()
        self.refresh_skills_listbox(skill_type)
        messagebox.showinfo("Success", f"{skill} has been added to {member_name}'s {skill_type}.")

    def remove_hierarchical_skill(self, skill_type):
        logging.info(f"Removing {skill_type}")
        selection = self.skills_listboxes[skill_type].curselection()
        if not selection:
            messagebox.showerror("Error", f"Please select a {skill_type} to remove.")
            return
        skill = self.skills_listboxes[skill_type].get(selection[0])
        member_name = self.member_var.get()
        if not member_name:
            messagebox.showerror("Error", "Please select a team member first.")
            return
        if messagebox.askyesno("Confirm Deletion", f"Are you sure you want to delete '{skill}' from {member_name}'s {skill_type}?"):
            self.team_members[member_name][skill_type.lower().replace(' ', '_')].remove(skill)
            self.save_data()
            self.refresh_skills_listbox(skill_type)
            messagebox.showinfo("Deleted", f"{skill} has been deleted from {member_name}'s {skill_type}.")

    def add_category(self, skill_type):
        logging.info(f"Adding category to {skill_type}")
        new_category = simpledialog.askstring("Add Category", f"Enter new {skill_type} category:")
        if new_category:
            new_category = new_category.strip()
            if new_category in self.categories[skill_type]:
                messagebox.showerror("Error", f"Category '{new_category}' already exists in {skill_type}.")
            else:
                self.categories[skill_type][new_category] = []
                self.save_data()
                self.refresh_categories(skill_type)
                messagebox.showinfo("Success", f"Category '{new_category}' added to {skill_type}.")
                
                # Prompt to add subskills
                self.add_subskills_to_category(skill_type, new_category)

    def add_subskills_to_category(self, skill_type, category):
        while True:
            new_subskill = simpledialog.askstring("Add Subskill", f"Enter a subskill for '{category}' (or cancel to finish):")
            if new_subskill:
                new_subskill = new_subskill.strip()
                if new_subskill not in self.categories[skill_type][category]:
                    self.categories[skill_type][category].append(new_subskill)
                    logging.info(f"Added subskill '{new_subskill}' to category '{category}' in {skill_type}")
                else:
                    messagebox.showwarning("Duplicate", f"Subskill '{new_subskill}' already exists in this category.")
            else:
                break
        
        self.save_data()
        self.refresh_categories(skill_type)

    def add_multiple_subskills(self, skill_type):
        category = self.category_vars[skill_type].get()
        if not category:
            messagebox.showerror("Error", "Please select a category first.")
            return
        
        subskills_text = self.subskill_entry.get().strip()
        if not subskills_text:
            messagebox.showerror("Error", "Please enter subskills separated by commas.")
            return
        
        subskills = [s.strip() for s in subskills_text.split(',') if s.strip()]
        added_subskills = []
        
        for subskill in subskills:
            if subskill not in self.subskills_data[skill_type][category]:
                self.subskills_data[skill_type][category].append(subskill)
                added_subskills.append(subskill)
                logging.info(f"Added subskill '{subskill}' to category '{category}' in {skill_type}")
            else:
                messagebox.showwarning("Duplicate", f"Subskill '{subskill}' already exists in this category.")
        
        if added_subskills:
            self.save_data()
            self.refresh_categories(skill_type)
            self.subskill_entry.delete(0, tk.END)
            messagebox.showinfo("Success", f"Added {len(added_subskills)} new subskills to '{category}'.")
        else:
            messagebox.showinfo("No Changes", "No new subskills were added.")

    def delete_subskill(self, skill_type):
        logging.info(f"Deleting subskill from {skill_type}")
        category = self.category_vars[skill_type].get()
        subskill = self.subskill_vars[skill_type].get()
        if not category or not subskill:
            messagebox.showerror("Error", f"Please select a category and subskill to delete from {skill_type}.")
            return
        if messagebox.askyesno("Confirm Deletion", f"Are you sure you want to delete '{subskill}' from '{category}' in {skill_type}?"):
            self.subskills_data[skill_type][category].remove(subskill)
            self.save_data()
            self.refresh_categories(skill_type)
            messagebox.showinfo("Deleted", f"Subskill '{subskill}' has been deleted from '{category}' in {skill_type}.")

    def delete_category(self, skill_type):
        logging.info(f"Deleting category from {skill_type}")
        category = self.category_vars[skill_type].get()
        if not category:
            messagebox.showerror("Error", f"Please select a category to delete from {skill_type}.")
            return
        if messagebox.askyesno("Confirm Deletion", f"Are you sure you want to delete the category '{category}' and all its subskills from {skill_type}?"):
            del self.categories[skill_type][category]
            self.save_data()
            self.refresh_categories(skill_type)
            messagebox.showinfo("Deleted", f"Category '{category}' and all its subskills have been deleted from {skill_type}.")

    def refresh_categories(self, skill_type):
        logging.info(f"Refreshing {skill_type} categories")
        self.category_dropdowns[skill_type]['values'] = list(self.categories[skill_type].keys())
        if self.category_dropdowns[skill_type]['values']:
            self.category_dropdowns[skill_type].current(0)
            self.on_category_select(skill_type)

    def refresh_skills_listbox(self, skill_type):
        logging.info(f"Refreshing {skill_type} listbox")
        member_name = self.member_var.get()
        self.skills_listboxes[skill_type].delete(0, tk.END)
        if member_name:
            for skill in self.team_members[member_name][skill_type.lower().replace(' ', '_')]:
                self.skills_listboxes[skill_type].insert(tk.END, skill)

    def refresh_soft_categories(self):
        logging.info("Refreshing soft skill categories")
        self.soft_categories_listbox.delete(0, tk.END)
        for category in self.skills_data["Soft Skills"].keys():
            self.soft_categories_listbox.insert(tk.END, category)
        self.refresh_soft_skills_listbox()

    def refresh_soft_skills_listbox(self, selected_category=None):
        logging.info("Refreshing soft skills listbox")
        self.soft_skills_listbox.delete(0, tk.END)
        if not selected_category:
            selection = self.soft_categories_listbox.curselection()
            if selection:
                selected_category = self.soft_categories_listbox.get(selection[0])
            else:
                return
        for skill in self.skills_data["Soft Skills"].get(selected_category, []):
            self.soft_skills_listbox.insert(tk.END, skill)

    def on_soft_category_select(self, event):
        logging.info("Soft skill category selected")
        selection = event.widget.curselection()
        if selection:
            category = event.widget.get(selection[0])
            self.refresh_soft_skills_listbox(category)

    def add_soft_category(self):
        logging.info("Adding soft skill category")
        category = simpledialog.askstring("Add Category", "Enter the new soft skill category:")
        if category:
            category = category.strip()
            if category in self.skills_data["Soft Skills"]:
                messagebox.showerror("Error", "This category already exists.")
                return
            self.skills_data["Soft Skills"][category] = []
            self.save_data()
            self.refresh_soft_categories()
            messagebox.showinfo("Success", f"Category '{category}' has been added.")

    def remove_soft_category(self):
        logging.info("Removing soft skill category")
        selection = self.soft_categories_listbox.curselection()
        if not selection:
            messagebox.showerror("Error", "Please select a category to remove.")
            return
        category = self.soft_categories_listbox.get(selection[0])
        if messagebox.askyesno("Confirm Deletion", f"Are you sure you want to delete the category '{category}' and all its skills?"):
            del self.skills_data["Soft Skills"][category]
            self.save_data()
            self.refresh_soft_categories()
            messagebox.showinfo("Deleted", f"Category '{category}' and its skills have been deleted.")

    def modify_soft_category(self):
        logging.info("Modifying soft skill category")
        selection = self.soft_categories_listbox.curselection()
        if not selection:
            messagebox.showerror("Error", "Please select a category to modify.")
            return
        old_category = self.soft_categories_listbox.get(selection[0])
        new_category = simpledialog.askstring("Modify Category", f"Enter the new name for category '{old_category}':")
        if new_category:
            new_category = new_category.strip()
            if new_category in self.skills_data["Soft Skills"]:
                messagebox.showerror("Error", "This category name already exists.")
                return
            self.skills_data["Soft Skills"][new_category] = self.skills_data["Soft Skills"].pop(old_category)
            self.save_data()
            self.refresh_soft_categories()
            messagebox.showinfo("Success", f"Category '{old_category}' has been renamed to '{new_category}'.")

    def add_soft_skill(self):
        logging.info("Adding soft skill")
        selection = self.soft_categories_listbox.curselection()
        if not selection:
            messagebox.showerror("Error", "Please select a category first.")
            return
        category = self.soft_categories_listbox.get(selection[0])
        skill = simpledialog.askstring("Add Skill", f"Enter the new soft skill for category '{category}':")
        if skill:
            skill = skill.strip()
            if skill in self.skills_data["Soft Skills"][category]:
                messagebox.showerror("Error", "This skill already exists in the selected category.")
                return
            self.skills_data["Soft Skills"][category].append(skill)
            self.save_data()
            self.refresh_soft_skills_listbox(category)
            messagebox.showinfo("Success", f"Skill '{skill}' has been added to category '{category}'.")

    def remove_soft_skill(self):
        logging.info("Removing soft skill")
        category_selection = self.soft_categories_listbox.curselection()
        skill_selection = self.soft_skills_listbox.curselection()
        if not category_selection or not skill_selection:
            messagebox.showerror("Error", "Please select a category and a skill to remove.")
            return
        category = self.soft_categories_listbox.get(category_selection[0])
        skills = self.soft_skills_listbox.curselection()
        skills_to_remove = [self.soft_skills_listbox.get(i) for i in skills]
        if messagebox.askyesno("Confirm Deletion", f"Are you sure you want to delete the selected skill(s) from category '{category}'?"):
            for skill in skills_to_remove:
                try:
                    self.skills_data["Soft Skills"][category].remove(skill)
                    logging.info(f"Removed skill '{skill}' from category '{category}'")
                except ValueError:
                    logging.error(f"Skill '{skill}' not found in category '{category}'")
            self.save_data()
            self.refresh_soft_skills_listbox(category)
            messagebox.showinfo("Deleted", f"Selected skill(s) have been deleted from category '{category}'.")

    def modify_soft_skill(self):
        logging.info("Modifying soft skill")
        category_selection = self.soft_categories_listbox.curselection()
        skill_selection = self.soft_skills_listbox.curselection()
        if not category_selection or not skill_selection:
            messagebox.showerror("Error", "Please select a category and a skill to modify.")
            return
        category = self.soft_categories_listbox.get(category_selection[0])
        old_skill = self.soft_skills_listbox.get(skill_selection[0])
        new_skill = simpledialog.askstring("Modify Skill", f"Enter the new name for skill '{old_skill}':")
        if new_skill:
            new_skill = new_skill.strip()
            if new_skill in self.skills_data["Soft Skills"][category]:
                messagebox.showerror("Error", "This skill already exists in the selected category.")
                return
            index = self.skills_data["Soft Skills"][category].index(old_skill)
            self.skills_data["Soft Skills"][category][index] = new_skill
            self.save_data()
            self.refresh_soft_skills_listbox(category)
            messagebox.showinfo("Success", f"Skill '{old_skill}' has been renamed to '{new_skill}'.")

    def delete_soft_skill_category(self):
        logging.info("Deleting soft skill category")
        selection = self.soft_categories_listbox.curselection()
        if not selection:
            messagebox.showerror("Error", "Please select a category to delete.")
            return
        category = self.soft_categories_listbox.get(selection[0])
        if messagebox.askyesno("Confirm Deletion", f"Are you sure you want to delete the category '{category}' and all its skills?"):
            del self.skills_data["Soft Skills"][category]
            self.save_data()
            self.refresh_soft_categories()
            messagebox.showinfo("Deleted", f"Category '{category}' and all its skills have been deleted.")

    def add_certification(self):
        logging.info("Adding certification")
        cert = simpledialog.askstring("Add Certification", "Enter the new certification:")
        if cert:
            cert = cert.strip()
            if cert in self.skills_data["Certifications"]:
                messagebox.showerror("Error", "This certification already exists.")
                return
            self.skills_data["Certifications"].append(cert)
            self.save_data()
            self.refresh_certifications()
            messagebox.showinfo("Success", f"Certification '{cert}' has been added.")

    def remove_certification(self):
        logging.info("Removing certification")
        selection = self.certifications_listbox.curselection()
        if not selection:
            messagebox.showerror("Error", "Please select a certification to remove.")
            return
        certs = self.certifications_listbox.curselection()
        certs_to_remove = [self.certifications_listbox.get(i) for i in certs]
        if messagebox.askyesno("Confirm Deletion", f"Are you sure you want to delete the selected certification(s)?"):
            for cert in certs_to_remove:
                try:
                    self.skills_data["Certifications"].remove(cert)
                    logging.info(f"Removed certification '{cert}'")
                except ValueError:
                    logging.error(f"Certification '{cert}' not found.")
            self.save_data()
            self.refresh_certifications()
            messagebox.showinfo("Deleted", f"Selected certification(s) have been deleted.")

    def modify_certification(self):
        logging.info("Modifying certification")
        selection = self.certifications_listbox.curselection()
        if not selection:
            messagebox.showerror("Error", "Please select a certification to modify.")
            return
        old_cert = self.certifications_listbox.get(selection[0])
        new_cert = simpledialog.askstring("Modify Certification", f"Enter the new name for certification '{old_cert}':")
        if new_cert:
            new_cert = new_cert.strip()
            if new_cert in self.skills_data["Certifications"]:
                messagebox.showerror("Error", "This certification already exists.")
                return
            index = self.skills_data["Certifications"].index(old_cert)
            self.skills_data["Certifications"][index] = new_cert
            self.save_data()
            self.refresh_certifications()
            messagebox.showinfo("Success", f"Certification '{old_cert}' has been renamed to '{new_cert}'.")

    def validate_date(self, date_text):
        try:
            datetime.strptime(date_text, '%d/%m/%Y')
            return True
        except ValueError:
            return False

    def clear_member_details(self):
        logging.info("Clearing member details")
        # Personal Information
        self.hobbies_entry.delete(0, tk.END)
        self.interests_entry.delete(0, tk.END)
        self.family_entry.delete(0, tk.END)
        self.other_personal_text.delete("1.0", tk.END)
        # Professional Information
        self.job_title_entry.delete(0, tk.END)
        self.join_date_entry.delete(0, tk.END)
        self.birthday_entry.delete(0, tk.END)
        # Skills and Certifications
        for skill_type in ["Technical Skills", "Soft Skills", "Certifications"]:
            listbox = getattr(self, f"{skill_type.lower().replace(' ', '_')}_listbox")
            listbox.delete(0, tk.END)
        # Progression Fields
        self.goals_text.delete("1.0", tk.END)
        self.dev_plan_text.delete("1.0", tk.END)
        self.achieved_goals_text.delete("1.0", tk.END)

    def load_config_file(self):
        logging.info("Loading config file via dialog")
        selected_file = filedialog.askopenfilename(
            title="Select Configuration File",
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")]
        )
        if selected_file:
            self.config_file = selected_file
            self.load_config()
            messagebox.showinfo("Config Loaded", f"Configuration loaded from {selected_file}")

    def close_app(self):
        logging.info("Closing application")
        self.save_config()
        logging.info("Application closed by user")
        self.master.destroy()

    def create_meetings_tab(self):
        logging.info("Creating meetings tab")
        # Create a frame for the meetings list
        meetings_list_frame = ttk.Frame(self.meetings_frame)
        meetings_list_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=10, pady=10)

        # Create a frame for the meeting details
        meeting_details_frame = ttk.Frame(self.meetings_frame)
        meeting_details_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=10, pady=10)

        # Meetings list
        self.meetings_listbox = tk.Listbox(meetings_list_frame, width=50)
        self.meetings_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.meetings_listbox.bind('<<ListboxSelect>>', self.on_meeting_select)

        # Scrollbar for meetings list
        meetings_scrollbar = ttk.Scrollbar(meetings_list_frame, orient=tk.VERTICAL, command=self.meetings_listbox.yview)
        meetings_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.meetings_listbox.config(yscrollcommand=meetings_scrollbar.set)

        # Meeting details
        ttk.Label(meeting_details_frame, text="Date (DD/MM/YYYY):").grid(row=0, column=0, sticky="e", padx=5, pady=5)
        self.meeting_date_entry = ttk.Entry(meeting_details_frame, width=12)
        self.meeting_date_entry.grid(row=0, column=1, sticky="w", padx=5, pady=5)

        ttk.Label(meeting_details_frame, text="Title:").grid(row=1, column=0, sticky="e", padx=5, pady=5)
        self.meeting_title_entry = ttk.Entry(meeting_details_frame, width=40)
        self.meeting_title_entry.grid(row=1, column=1, sticky="w", padx=5, pady=5)

        ttk.Label(meeting_details_frame, text="Highlights:").grid(row=2, column=0, sticky="ne", padx=5, pady=5)
        self.meeting_highlights_text = scrolledtext.ScrolledText(meeting_details_frame, width=40, height=5)
        self.meeting_highlights_text.grid(row=2, column=1, sticky="w", padx=5, pady=5)

        ttk.Label(meeting_details_frame, text="Detailed Notes:").grid(row=3, column=0, sticky="ne", padx=5, pady=5)
        self.meeting_notes_text = scrolledtext.ScrolledText(meeting_details_frame, width=40, height=10)
        self.meeting_notes_text.grid(row=3, column=1, sticky="w", padx=5, pady=5)

        ttk.Label(meeting_details_frame, text="Action Items:").grid(row=4, column=0, sticky="ne", padx=5, pady=5)
        self.meeting_actions_text = scrolledtext.ScrolledText(meeting_details_frame, width=40, height=5)
        self.meeting_actions_text.grid(row=4, column=1, sticky="w", padx=5, pady=5)

        # Buttons
        buttons_frame = ttk.Frame(meeting_details_frame)
        buttons_frame.grid(row=5, column=0, columnspan=2, pady=10)

        ttk.Button(buttons_frame, text="Add Meeting", command=self.add_meeting).pack(side=tk.LEFT, padx=5)
        ttk.Button(buttons_frame, text="Update Meeting", command=self.update_meeting).pack(side=tk.LEFT, padx=5)
        ttk.Button(buttons_frame, text="Delete Meeting", command=self.delete_meeting).pack(side=tk.LEFT, padx=5)

        # Populate meetings list
        self.refresh_meetings_list()

    def refresh_meetings_list(self):
        logging.info("Refreshing meetings list")
        self.meetings_listbox.delete(0, tk.END)
        for meeting in sorted(self.meetings, key=lambda x: x['date'], reverse=True):
            self.meetings_listbox.insert(tk.END, f"{meeting['date']} - {meeting['title']}")

    def on_meeting_select(self, event):
        logging.info("Meeting selected")
        selection = self.meetings_listbox.curselection()
        if selection:
            index = selection[0]
            if 0 <= index < len(self.meetings):
                meeting = self.meetings[index]
                self.meeting_date_entry.delete(0, tk.END)
                self.meeting_date_entry.insert(0, meeting['date'])
                self.meeting_title_entry.delete(0, tk.END)
                self.meeting_title_entry.insert(0, meeting['title'])
                self.meeting_highlights_text.delete("1.0", tk.END)
                self.meeting_highlights_text.insert(tk.END, meeting['highlights'])
                self.meeting_notes_text.delete("1.0", tk.END)
                self.meeting_notes_text.insert(tk.END, meeting['notes'])
                self.meeting_actions_text.delete("1.0", tk.END)
                self.meeting_actions_text.insert(tk.END, meeting['action_items'])
            else:
                logging.error(f"Meeting index {index} out of range. Total meetings: {len(self.meetings)}")
                messagebox.showerror("Error", "Selected meeting not found. Please refresh the meetings list.")
                self.refresh_meetings_list()

    def add_meeting(self):
        logging.info("Adding new meeting")
        date_str = self.meeting_date_entry.get()
        try:
            # Validate the date format
            datetime.strptime(date_str, "%d/%m/%Y")
        except ValueError:
            messagebox.showerror("Invalid Date", "Please enter the date in DD/MM/YYYY format.")
            return
        
        meeting = {
            'date': date_str,
            'title': self.meeting_title_entry.get(),
            'highlights': self.meeting_highlights_text.get("1.0", tk.END).strip(),
            'notes': self.meeting_notes_text.get("1.0", tk.END).strip(),
            'action_items': self.meeting_actions_text.get("1.0", tk.END).strip()
        }
        self.meetings.append(meeting)
        self.save_data()
        self.refresh_meetings_list()
        messagebox.showinfo("Success", "Meeting added successfully.")

    def update_meeting(self):
        logging.info("Updating meeting")
        selection = self.meetings_listbox.curselection()
        if selection:
            index = selection[0]
            date_str = self.meeting_date_entry.get()
            try:
                # Validate the date format
                datetime.strptime(date_str, "%d/%m/%Y")
            except ValueError:
                messagebox.showerror("Invalid Date", "Please enter the date in DD/MM/YYYY format.")
                return
            
            self.meetings[index] = {
                'date': date_str,
                'title': self.meeting_title_entry.get(),
                'highlights': self.meeting_highlights_text.get("1.0", tk.END).strip(),
                'notes': self.meeting_notes_text.get("1.0", tk.END).strip(),
                'action_items': self.meeting_actions_text.get("1.0", tk.END).strip()
            }
            self.save_data()
            self.refresh_meetings_list()
            messagebox.showinfo("Success", "Meeting updated successfully.")
        else:
            messagebox.showerror("Error", "Please select a meeting to update.")

    def delete_meeting(self):
        logging.info("Deleting meeting")
        selection = self.meetings_listbox.curselection()
        if selection:
            index = selection[0]
            if messagebox.askyesno("Confirm Deletion", "Are you sure you want to delete this meeting?"):
                del self.meetings[index]
                self.save_data()
                self.refresh_meetings_list()
                messagebox.showinfo("Success", "Meeting deleted successfully.")
        else:
            messagebox.showerror("Error", "Please select a meeting to delete.")

    def upload_skills_csv(self):
        logging.info("Uploading skills CSV")
        file_path = filedialog.askopenfilename(filetypes=[("CSV Files", "*.csv")])
        if not file_path:
            return

        try:
            with open(file_path, 'r') as csv_file:
                csv_reader = csv.reader(csv_file, delimiter='\t')
                next(csv_reader)  # Skip the header row
                for row in csv_reader:
                    if len(row) == 1:  # Handle single-column rows
                        row = row[0].split(',')
                    if len(row) != 3:
                        raise ValueError(f"Invalid row format: {row}")
                    category, subcategory, skill = [item.strip() for item in row]

                    if category not in self.categories:
                        self.categories[category] = {}
                    if subcategory not in self.categories[category]:
                        self.categories[category][subcategory] = []
                    if skill not in self.categories[category][subcategory]:
                        self.categories[category][subcategory].append(skill)

            self.save_data()
            self.refresh_categories("Technical Skills")
            messagebox.showinfo("Success", "Skills CSV uploaded and processed successfully.")
        except Exception as e:
            logging.error(f"Error processing CSV: {str(e)}")
            messagebox.showerror("Error", f"An error occurred while processing the CSV file: {str(e)}")

class TeamMemberDialog:
    def __init__(self, parent, title, name="", job_title="", join_date="", birthday=""):
        top = self.top = tk.Toplevel(parent)
        top.title(title)
        top.grab_set()
        self.result = None

        ttk.Label(top, text="Full Name:").grid(row=0, column=0, padx=10, pady=5, sticky="e")
        self.name_entry = ttk.Entry(top, width=40)
        self.name_entry.grid(row=0, column=1, padx=10, pady=5)
        self.name_entry.insert(0, name)

        ttk.Label(top, text="Job Title:").grid(row=1, column=0, padx=10, pady=5, sticky="e")
        self.job_title_entry = ttk.Entry(top, width=40)
        self.job_title_entry.grid(row=1, column=1, padx=10, pady=5)
        self.job_title_entry.insert(0, job_title)

        ttk.Label(top, text="Join Date (DD/MM/YYYY):").grid(row=2, column=0, padx=10, pady=5, sticky="e")
        self.join_date_entry = ttk.Entry(top, width=40)
        self.join_date_entry.grid(row=2, column=1, padx=10, pady=5)
        self.join_date_entry.insert(0, join_date)

        ttk.Label(top, text="Birthday (DD/MM/YYYY):").grid(row=3, column=0, padx=10, pady=5, sticky="e")
        self.birthday_entry = ttk.Entry(top, width=40)
        self.birthday_entry.grid(row=3, column=1, padx=10, pady=5)
        self.birthday_entry.insert(0, birthday)

        # Correctly indented button frame
        button_frame = ttk.Frame(top)
        button_frame.grid(row=4, column=0, columnspan=2, pady=10)

        ok_button = ttk.Button(button_frame, text="OK", command=self.on_ok, style="success.TButton" if USE_TTKBOOTSTRAP else "TButton")
        ok_button.pack(side=tk.LEFT, padx=5)

        cancel_button = ttk.Button(button_frame, text="Cancel", command=self.on_cancel, style="danger.TButton" if USE_TTKBOOTSTRAP else "TButton")
        cancel_button.pack(side=tk.LEFT, padx=5)

    def on_ok(self):
        name = self.name_entry.get().strip()
        job_title = self.job_title_entry.get().strip()
        join_date = self.join_date_entry.get().strip()
        birthday = self.birthday_entry.get().strip()

        if not name:
            messagebox.showerror("Input Error", "Full Name is required.")
            return
        if not self.validate_date(join_date):
            messagebox.showerror("Input Error", "Join Date is not in the correct format (DD/MM/YYYY).")
            return
        if not self.validate_date(birthday):
            messagebox.showerror("Input Error", "Birthday is not in the correct format (DD/MM/YYYY).")
            return

        self.result = (name, job_title, join_date, birthday)
        self.top.destroy()

    def on_cancel(self):
        self.top.destroy()

    def validate_date(self, date_text):
        try:
            datetime.strptime(date_text, '%d-%m-%Y')
            return True
        except ValueError:
            return False

def main():
    logging.info("Starting AscentPro application")
    try:
        logging.info("Creating Tk root window")
        root = tk.Tk()
        logging.info("Tk root window created successfully")
        
        logging.info("Initializing AscentPro")
        app = AscentPro(root)
        logging.info("AscentPro initialized successfully")
        
        logging.info("Entering mainloop")
        root.mainloop()
    except KeyboardInterrupt:
        logging.info("Application closed by user")
    except Exception as e:
        logging.exception("An unexpected error occurred:")
        error_message = f"An unexpected error occurred: {str(e)}\n\nPlease check the app.log file for more details."
        logging.error(error_message)
        try:
            messagebox.showerror("Critical Error", error_message)
        except:
            print(error_message)  # Fallback if messagebox fails
    finally:
        logging.info("AscentPro application finished executing")

if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        logging.exception("An unhandled exception occurred:")
        messagebox.showerror("Critical Error", f"An unhandled exception occurred: {str(e)}\n\nPlease check the app.log file for more details.")
