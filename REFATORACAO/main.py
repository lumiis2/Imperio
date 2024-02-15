import customtkinter
import os
from gui_functions import setup_gui
from aux_functions import set_var

class App(customtkinter.CTk):
    def __init__(self):
        super().__init__()
        self.setup_ui()

    def setup_ui(self):
        setup_gui(self)
        set_var(self)

if __name__ == "__main__":
    script_directory = os.path.dirname(os.path.abspath(__file__))
    os.chdir(script_directory)

    app = App()
    app.mainloop()