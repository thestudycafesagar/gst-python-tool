import sys
import os

# Add the project directory to sys.path
sys.path.append(r"c:\Users\HP\Desktop\Rohit Python Tools\rohit combo\rohit combo")
sys.path.append(r"c:\Users\HP\Desktop\Rohit Python Tools\rohit combo\rohit combo\GST\GST 2B Downloader")

import traceback
try:
    from main import App
    app = App()
    print("App instantiated successfully.")
except Exception as e:
    print("ERROR:")
    traceback.print_exc()
