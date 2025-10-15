import sys
import os

# Add the project root to the Python path to allow for absolute imports
# This makes the app runnable from anywhere
project_root = os.path.abspath(os.path.dirname(__file__))
sys.path.insert(0, project_root)

from UI.login_ui import App

if __name__ == "__main__":
    app = App()
    app.mainloop()