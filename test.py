import os
import sys
import customtkinter as ctk
from PIL import Image, ImageTk

ctk.set_appearance_mode("Dark")  # Set initial theme
ctk.set_default_color_theme("blue")

class ExcelCreationToolGUI(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Nexus Config Tool")
        self.geometry("400x800")
        self.base_dir = sys._MEIPASS if getattr(sys, 'frozen', False) and hasattr(sys, '_MEIPASS') else os.path.dirname(__file__)
        
        # Set the window icon
        icon_path = os.path.join(self.base_dir, "Data", "GUI", "FM_icon.ico")
        self.iconbitmap(icon_path)  # Use the iconbitmap method with the path to your .ico file

        self.setup_ui()

    def setup_ui(self):
        image_path = os.path.join(self.base_dir, "Data", "GUI")
        
        self.main_frame = ctk.CTkFrame(self, corner_radius=10)
        self.main_frame.pack(pady=20, padx=20, fill="both", expand=True)

        # Load images for dark and light mode
        self.dark_mode_image = self.create_resized_photoimage(os.path.join(image_path, "dark.png"), (100, 50))
        self.light_mode_image = self.create_resized_photoimage(os.path.join(image_path, "light.png"), (100, 50))
        self.logo_image = self.create_resized_photoimage(os.path.join(image_path, "FMg.png"), (300, 100))

        # Set the theme button image based on the current theme mode
        current_button_image = self.light_mode_image if ctk.get_appearance_mode() == "Light" else self.dark_mode_image

        # Initialize the theme toggle label
        self.theme_button = ctk.CTkLabel(self.main_frame, text="", image=current_button_image)
        self.theme_button.place(x=10, y=10, anchor="nw")
        
        self.logo = ctk.CTkLabel(self.main_frame, text="", image=self.logo_image)
        self.logo.place(x=50, y=70)

        # Bind a mouse click event to the label
        self.theme_button.bind("<Button-1>", self.toggle_theme)

        # Keep a reference to the image to avoid garbage collection
        self.theme_button.image = current_button_image
        
        # Additional UI elements...
                # Barge Number with placeholder
        self.barge_number_entry = ctk.CTkEntry(self.main_frame, placeholder_text="FHxxx")
        ctk.CTkLabel(self.main_frame, text="Barge Number:").place(x=40, y=200, anchor="nw")
        self.barge_number_entry.place(x=160, y=200, anchor="nw")

        # Fjord Control Password with placeholder
        self.fjord_control_password_entry = ctk.CTkEntry(self.main_frame, placeholder_text="Password", show="*")
        ctk.CTkLabel(self.main_frame, text="FC Password:").place(x=40, y=250, anchor="nw")
        self.fjord_control_password_entry.place(x=160, y=250, anchor="nw")

        # Send Interval with placeholder
        self.send_interval_entry = ctk.CTkEntry(self.main_frame, placeholder_text="2000ms")
        ctk.CTkLabel(self.main_frame, text="Send Interval:").place(x=40, y=300, anchor="nw")
        self.send_interval_entry.place(x=160, y=300, anchor="nw")
        self.send_interval_entry.insert(0, "2000")  # Set default value

        # EMS PLC Type OptionMenu
        self.ems_plc_type_var = ctk.StringVar(value="Beckhoff")  # Default value
        ctk.CTkLabel(self.main_frame, text="EMS PLC:").place(x=140, y=350, anchor="ne")
        self.ems_plc_type_optionmenu = ctk.CTkSegmentedButton(self.main_frame, variable=self.ems_plc_type_var, values=["Beckhoff", "Wago"])
        self.ems_plc_type_optionmenu.place(x=150, y=350, anchor="nw")

        # ESS PLC Type OptionMenu
        self.ess_plc_type_var = ctk.StringVar(value="Beckhoff")  # Default value
        ctk.CTkLabel(self.main_frame, text="ESS PLC Type:").place(x=40, y=400, anchor="nw")
        self.ess_plc_type_optionmenu = ctk.CTkSegmentedButton(self.main_frame, variable=self.ess_plc_type_var, values=["Beckhoff", "Wago"])
        self.ess_plc_type_optionmenu.place(x=160, y=400, anchor="nw")

        # Number of Generators Combobox
        ctk.CTkLabel(self.main_frame, text="Number of Generators:").place(x=40, y=450, anchor="nw")
        self.num_generators_combobox = ctk.CTkComboBox(self.main_frame, values=[str(i) for i in range(1, 4)])
        self.num_generators_combobox.place(x=160, y=450, anchor="nw")

        # Create File Button
        self.create_file_button = ctk.CTkButton(self.main_frame, text="Create File", command=self.create_file)
        self.create_file_button.place(x=100, y=500, anchor="nw")

    def create_resized_photoimage(self, image_path, size):
        """Load image, resize it, and convert to PhotoImage."""
        original_image = Image.open(image_path)
        resized_image = original_image.resize(size, Image.Resampling.LANCZOS)
        return ImageTk.PhotoImage(resized_image)


    def show_gen_settings(event=None):
        # Placeholder function to demonstrate dynamic UI changes based on user input
        print("Generator settings updated.")

    def toggle_theme(self, event):
        """Toggle between dark and light mode and update label image."""
        new_mode = "Light" if ctk.get_appearance_mode() == "Dark" else "Dark"
        ctk.set_appearance_mode(new_mode)
        
        # Update the label image based on the new mode
        current_button_image = self.light_mode_image if new_mode == "Light" else self.dark_mode_image
        self.theme_button.configure(image=current_button_image)
        self.theme_button.image = current_button_image  # Keep a reference
        
    def create_file(self):
        # Retrieve values from CTkEntry widgets
        barge_number = self.barge_number_entry.get()
        fjord_control_password = self.fjord_control_password_entry.get()
        send_interval = self.send_interval_entry.get()

        # Example: Print the values
        print(f"Barge Number: {barge_number}")
        print(f"FC Password: {fjord_control_password}")
        print(f"Send Interval: {send_interval}")

if __name__ == "__main__":
    app = ExcelCreationToolGUI()
    app.mainloop()