
import customtkinter
from aux_functions import choose_store
from process_functions import download

def setup_gui(self):
    self.title("Comparador de Planilhas")
    self.iconbitmap('images/icon.ico')
    self.geometry(f"{550}x{290}")
    self.grid_columnconfigure(0, weight=1)
    self.grid_rowconfigure(0, weight=0)
    self.grid_rowconfigure((1, 2, 3), weight=1)

    create_menubar(self)
    create_buttons(self)
    create_dark_mode_switch(self)
    customtkinter.set_appearance_mode("Light")  # Modes: "System" (standard), "Dark", "Light"
    customtkinter.set_default_color_theme("dark-blue")  # Themes: "blue" (standard), "green", "dark-blue"

def create_menubar(self):
    self.menubar_frame = customtkinter.CTkFrame(self, height=30, corner_radius=0)
    self.menubar_frame.grid(row=0, column=0, columnspan=3, sticky="nsew")
    self.menubar_frame.grid_rowconfigure(0, weight=1)

    self.chooser_label = customtkinter.CTkLabel(self.menubar_frame, text="Escolha uma loja:", anchor="w")
    self.chooser_label.grid(row=0, column=1, columnspan=2, padx=5, pady=(10, 10))
    self.chooser_optionemenu = customtkinter.CTkOptionMenu(self.menubar_frame, values=["Loja Castelo", "Loja Cidade Nova", "Loja Planalto", "Loja Contagem", "Loja Nova Lima", "Loja E-commerce"],
                                                            command=lambda selected_store: choose_store(self, selected_store))
    self.chooser_optionemenu.grid(row=0, column=3, columnspan=4, padx=5, pady=(10, 10))

def create_buttons(self):
    self.button_1 = customtkinter.CTkButton(self, text="Adicionar REDE", command=lambda: download(self, 'xlsx'))
    self.button_1.grid(row=1, column=0, padx=20, pady=10)
    self.button_2 = customtkinter.CTkButton(self, text="Adicionar w3erp", command=lambda: download(self, 'csv'))
    self.button_2.grid(row=2, column=0, padx=20, pady=10)

def create_dark_mode_switch(self):
    self.switch = customtkinter.CTkSwitch(master=self.menubar_frame, text="DarkMode", command=change_mode)
    self.switch.grid(row=0, column=10, padx=30, pady=(5, 5))

def change_mode():
    mode = (customtkinter.get_appearance_mode())
    if mode == "Light":
        customtkinter.set_appearance_mode("Dark")
    elif mode == "Dark":
        customtkinter.set_appearance_mode("Light")